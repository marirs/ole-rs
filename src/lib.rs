mod constants;
mod directory;
pub mod error;
mod header;

use crate::{
    constants::Readable,
    directory::{DirectoryEntry, DirectoryEntryRaw, ObjectType},
    header::{parse_raw_header, OleHeader},
};
use derivative::Derivative;
use error::{HeaderErrorType, OleError};
use tokio::io::AsyncReadExt;

#[derive(Clone, Derivative)]
#[derivative(Debug)]
pub struct OleFile {
    header: OleHeader,
    #[derivative(Debug = "ignore")]
    sectors: Vec<Vec<u8>>,
    #[derivative(Debug = "ignore")]
    sector_allocation_table: Vec<u32>,
    #[derivative(Debug = "ignore")]
    short_sector_allocation_table: Vec<u32>,
    #[derivative(Debug = "ignore")]
    directory_stream_data: Vec<u8>,
    directory_entries: Vec<DirectoryEntry>,
    #[derivative(Debug = "ignore")]
    mini_stream: Vec<[u8; 64]>,
}

impl OleFile {
    pub fn list_streams(&self) -> Vec<String> {
        self.directory_entries
            .iter()
            .filter_map(|entry| {
                if entry.object_type == ObjectType::Stream {
                    Some(entry.name.clone())
                } else {
                    None
                }
            })
            .collect()
    }

    pub fn open_stream(&mut self, stream_name: &str) -> Result<Vec<u8>, OleError> {
        let potential_entries = self
            .directory_entries
            .iter()
            .filter(|entry| entry.name == stream_name);

        for directory_entry in potential_entries {
            if directory_entry.object_type == ObjectType::Stream {
                let mut data = vec![];
                let mut collected_bytes = 0;
                // the unwrap is safe because the location is guaranteed to exist for this object type
                let mut next_sector = directory_entry.starting_sector_location.unwrap();

                if directory_entry.stream_size < self.header.standard_stream_min_size as u64 {
                    // it's in the mini-FAT
                    loop {
                        if next_sector == constants::CHAIN_END {
                            break;
                        } else {
                            let mut sector_data: Vec<u8> = vec![];
                            for byte in self.mini_stream[next_sector as usize] {
                                sector_data.push(byte);
                                collected_bytes += 1;
                                if collected_bytes == directory_entry.stream_size {
                                    break;
                                }
                            }
                            data.extend(sector_data)
                        }
                        next_sector = self.short_sector_allocation_table[next_sector as usize];
                    }
                } else {
                    // it's in the FAT
                    loop {
                        if next_sector == constants::CHAIN_END {
                            break;
                        } else {
                            let mut sector_data: Vec<u8> = vec![];
                            for byte in &self.sectors[next_sector as usize] {
                                sector_data.push(*byte);
                                collected_bytes += 1;
                                if collected_bytes == directory_entry.stream_size {
                                    break;
                                }
                            }
                            data.extend(sector_data)
                        }
                        next_sector = self.sector_allocation_table[next_sector as usize];
                    }
                }
                // println!("data.len(): {}", data.len());
                return Ok(data);
            }
        }
        Err(OleError::DirectoryEntryNotFound)
    }


    pub async fn from_file<P: AsRef<std::path::Path>>(file: P) -> Result<Self, OleFile> {
        let f = tokio::fs::File::open(file)?;
        Self::parse(f)?
    }

    async fn parse<R>(mut read: R) -> Result<Self, OleError>
    where
        R: Readable,
    {
        // read the header
        let raw_file_header = parse_raw_header(&mut read).await?;
        // println!("raw_file_header: {:#?}", raw_file_header);
        let file_header = OleHeader::from_raw(raw_file_header);
        // println!("file_header: {:#?}", file_header);
        let sector_size = file_header.sector_size as usize;

        //we have to read the remainder of the header if the sector size isn't what we tried to read
        if sector_size > constants::HEADER_LENGTH {
            let should_read_size = sector_size - constants::HEADER_LENGTH;
            let mut should_read = vec![0u8; should_read_size];
            let did_read_size = read.read(&mut should_read).await?;
            if did_read_size != should_read_size {
                return Err(OleError::InvalidHeader(HeaderErrorType::NotEnoughBytes(
                    should_read_size,
                    did_read_size,
                )));
            } else if should_read != vec![0u8; should_read_size] {
                return Err(OleError::InvalidHeader(HeaderErrorType::Parsing(
                    "all bytes must be zero for larger header sizes",
                    "n/a".to_string(),
                )));
            }
        }

        let mut sectors = vec![];
        loop {
            let mut buf = vec![0u8; sector_size];
            match read.read(&mut buf).await {
                Ok(actually_read_size) if actually_read_size == sector_size => {
                    sectors.push((&buf[0..actually_read_size]).to_vec());
                }
                Ok(wrong_size) if wrong_size != 0 => {
                    // TODO: we might have to handle the case where the
                    //       last sector isn't actually complete. Not sure yet.
                    //       the spec says the entire file has to be present here,
                    //       with equal sectors, so I'm doing it this way.
                    return Err(OleError::UnexpectedEof(format!(
                        "short read when parsing sector number: {}",
                        sectors.len()
                    )));
                }
                Ok(_empty) => {
                    break;
                }
                Err(error) => {
                    return Err(OleError::StdIo(error));
                }
            }
        }

        // println!("read_sectors: {}", sectors.len());
        let mut self_to_init = OleFile {
            header: file_header,
            sectors,
            sector_allocation_table: vec![],
            short_sector_allocation_table: vec![],
            directory_stream_data: vec![],
            directory_entries: vec![],
            mini_stream: vec![],
        };

        self_to_init.initialize_sector_allocation_table()?;
        self_to_init.initialize_short_sector_allocation_table()?;
        self_to_init.initialize_directory_stream()?;
        self_to_init.initialize_mini_stream()?;

        Ok(self_to_init)
    }

    fn initialize_sector_allocation_table(&mut self) -> Result<(), OleError> {
        for sector_index in self.header.sector_allocation_table_head.iter() {
            // println!("sector_index: {:#x?}", *sector_index);
            if *sector_index == constants::UNALLOCATED_SECTOR
                || *sector_index == constants::CHAIN_END
            {
                break;
            }
            let sector = self.sectors[*sector_index as usize]
                .chunks_exact(4)
                .map(|quad| u32::from_le_bytes([quad[0], quad[1], quad[2], quad[3]]));
            self.sector_allocation_table.extend(sector);
        }

        if self.header.master_sector_allocation_table_len > 0 {
            return Err(OleError::CurrentlyUnimplemented(
                "MSAT/DI-FAT unsupported todo impl me".to_string(),
            ));
        }

        Ok(())
    }
    fn initialize_short_sector_allocation_table(&mut self) -> Result<(), OleError> {
        if self.header.short_sector_allocation_table_len == 0
            || self.header.short_sector_allocation_table_first_sector == constants::CHAIN_END
        {
            return Ok(()); //no mini stream here
        }

        let mut next_index = self.header.short_sector_allocation_table_first_sector;
        let mut short_sector_allocation_table_raw_data: Vec<u8> = vec![];
        loop {
            // println!("next_index: {:#x?}", next_index);
            if next_index == constants::CHAIN_END {
                break;
            } else {
                short_sector_allocation_table_raw_data
                    .extend(self.sectors[next_index as usize].iter());
            }
            next_index = self.sector_allocation_table[next_index as usize];
        }

        // println!("short_sector_allocation_table_raw_data: {}", short_sector_allocation_table_raw_data.len());

        self.short_sector_allocation_table.extend(
            short_sector_allocation_table_raw_data
                .chunks_exact(4)
                .map(|quad| u32::from_le_bytes([quad[0], quad[1], quad[2], quad[3]])),
        );

        // println!("short_sector_allocation_table: {:#x?}", self.short_sector_allocation_table);

        Ok(())
    }

    fn initialize_directory_stream(&mut self) -> Result<(), OleError> {
        let mut next_directory_index = self.header.sector_allocation_table_first_sector;
        self.directory_stream_data
            .extend(self.sectors[next_directory_index as usize].iter());

        loop {
            next_directory_index =
                self.sector_allocation_table[next_directory_index as usize].clone();
            // println!("next: {:x?}", next);
            if next_directory_index == constants::CHAIN_END {
                break;
            } else {
                self.directory_stream_data
                    .extend(self.sectors[next_directory_index as usize].iter());
            }
        }

        self.initialize_directory_entries()?;

        Ok(())
    }

    fn initialize_directory_entries(&mut self) -> Result<(), OleError> {
        if self.directory_stream_data.len() % constants::SIZE_OF_DIRECTORY_ENTRY != 0 {
            return Err(OleError::InvalidDirectoryEntry(
                "directory_stream_size",
                format!(
                    "size of directory stream data is not correct? {}",
                    self.directory_stream_data.len()
                ),
            ));
        }

        self.directory_entries = Vec::with_capacity(
            self.directory_stream_data.len() / constants::SIZE_OF_DIRECTORY_ENTRY,
        );
        for (index, unparsed_entry) in self
            .directory_stream_data
            .chunks(constants::SIZE_OF_DIRECTORY_ENTRY)
            .enumerate()
        {
            // println!("unparsed_entry: {}", unparsed_entry.len());
            let raw_directory_entry = DirectoryEntryRaw::parse(unparsed_entry)?;
            match DirectoryEntry::from_raw(&self.header, raw_directory_entry, index) {
                Ok(directory_entry) => self.directory_entries.push(directory_entry),
                Err(OleError::UnknownOrUnallocatedDirectoryEntry) => continue,
                Err(anything_else) => return Err(anything_else),
            }
        }

        Ok(())
    }
    fn initialize_mini_stream(&mut self) -> Result<(), OleError> {
        let (mut next_sector, mini_stream_size) = {
            let root_entry = &self.directory_entries[0];
            match root_entry.starting_sector_location {
                None => return Ok(()), //no mini-stream here
                Some(starting_sector_location) => {
                    (starting_sector_location, root_entry.stream_size)
                }
            }
        };

        let mut raw_mini_stream_data: Vec<u8> = vec![];
        loop {
            // println!("next_sector: {:x?}", next_sector);
            if next_sector == constants::CHAIN_END {
                break;
            } else {
                raw_mini_stream_data.extend(self.sectors[next_sector as usize].iter());
            }
            next_sector = self.sector_allocation_table[next_sector as usize].clone();
        }
        raw_mini_stream_data.truncate(mini_stream_size as usize);
        // println!("raw_mini_stream_data {:#?}", raw_mini_stream_data.len());

        //mini streams are sectors of 64 bytes, and the size is guaranteed to be an exact multiple.
        raw_mini_stream_data.chunks_exact(64).for_each(|chunk| {
            //the unwrap is safe because the chunk is guaranteed to be 64 bytes.
            self.mini_stream.push(<[u8; 64]>::try_from(chunk).unwrap());
        });

        Ok(())
    }
}
