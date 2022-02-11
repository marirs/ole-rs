use crate::{
    constants::{self, Readable},
    error::{HeaderErrorType, Error},
    Result,
};
use derivative::Derivative;
use std::array::TryFromSliceError;
use tokio::io::AsyncReadExt;

#[derive(Clone, Derivative)]
#[derivative(Debug)]
pub struct OleHeader {
    pub major_version: u16,
    minor_version: u16,
    pub sector_size: u16,
    mini_sector_size: u16,
    directory_sectors_len: u32,
    pub standard_stream_min_size: u32,
    // sector allocation table AKA "FAT"
    pub sector_allocation_table_first_sector: u32,
    pub sector_allocation_table_len: u32,
    // short sector allocation table AKA "mini-FAT"
    pub short_sector_allocation_table_first_sector: u32,
    pub short_sector_allocation_table_len: u32,
    // master sector allocation table AKA "DI-FAT"
    master_sector_allocation_table_first_sector: u32,
    pub master_sector_allocation_table_len: u32,
    // the first 109 FAT sector locations
    #[derivative(Debug = "ignore")]
    pub sector_allocation_table_head: Vec<u32>,
}

impl OleHeader {
    pub fn from_raw(raw_file_header: RawFileHeader) -> Self {
        let major_version = u16::from_le_bytes(raw_file_header.major_version);
        let minor_version = u16::from_le_bytes(raw_file_header.minor_version);
        let sector_size = 2u16.pow(u16::from_le_bytes(raw_file_header.sector_size) as u32);
        let mini_sector_size =
            2u16.pow(u16::from_le_bytes(raw_file_header.mini_sector_size) as u32);
        let directory_sectors_len = u32::from_le_bytes(raw_file_header.directory_sectors_len);
        let standard_stream_min_size = u32::from_le_bytes(raw_file_header.standard_stream_min_size);
        let sector_allocation_table_first_sector =
            u32::from_le_bytes(raw_file_header.sector_allocation_table_first_sector);
        let sector_allocation_table_len =
            u32::from_le_bytes(raw_file_header.sector_allocation_table_len);
        let short_sector_allocation_table_first_sector =
            u32::from_le_bytes(raw_file_header.short_sector_allocation_table_first_sector);
        let short_sector_allocation_table_len =
            u32::from_le_bytes(raw_file_header.short_sector_allocation_table_len);
        let master_sector_allocation_table_first_sector =
            u32::from_le_bytes(raw_file_header.master_sector_allocation_table_first_sector);
        let master_sector_allocation_table_len =
            u32::from_le_bytes(raw_file_header.master_sector_allocation_table_len);
        let sector_allocation_table_head = raw_file_header.sector_allocation_table_head;

        OleHeader {
            major_version,
            minor_version,
            sector_size,
            mini_sector_size,
            directory_sectors_len,
            standard_stream_min_size,
            sector_allocation_table_first_sector,
            sector_allocation_table_len,
            short_sector_allocation_table_first_sector,
            short_sector_allocation_table_len,
            master_sector_allocation_table_first_sector,
            master_sector_allocation_table_len,
            sector_allocation_table_head,
        }
    }
}

/**
 * https://github.com/libyal/libolecf/blob/main/documentation/OLE%20Compound%20File%20format.asciidoc
 * https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-CFB/%5bMS-CFB%5d.pdf
 */
#[derive(Clone, Derivative)]
#[derivative(Debug)]
pub struct RawFileHeader {
    /**
    Revision number of the file format
    (minor version)
     */
    minor_version: [u8; 2],
    /**
    Version number of the file format
    (major version)
     */
    major_version: [u8; 2],
    /**
    Size of a sector in the compound document file in power-of-two
     */
    sector_size: [u8; 2],
    /**
    Size of a short-sector (mini-sector) in the short-stream container stream in power-of-two
     */
    mini_sector_size: [u8; 2],
    /**
    This integer field contains the count of the number of
    directory sectors in the compound file.
     */
    directory_sectors_len: [u8; 4],
    /**
    Total number of sectors used for the sector allocation table (SAT).
    The SAT is also referred to as the FAT (chain).
     */
    sector_allocation_table_len: [u8; 4],
    /**
    Sector identifier (SID) of first sector of the directory stream (chain).
     */
    sector_allocation_table_first_sector: [u8; 4],
    /**
    Minimum size of a standard stream (in bytes, most used size is 4096 bytes),
    streams smaller than this value are stored as short-streams
     */
    standard_stream_min_size: [u8; 4],
    /**
    Sector identifier (SID) of first sector of the short-sector allocation table (SSAT).
    The SSAT is also referred to as Mini-FAT.
     */
    short_sector_allocation_table_first_sector: [u8; 4],
    /**
    Total number of sectors used for the short-sector allocation table (SSAT).
     */
    short_sector_allocation_table_len: [u8; 4],
    /**
    Sector identifier (SID) of first sector of the master sector allocation table (MSAT).
    The MSAT is also referred to as Double Indirect FAT (DIF).
     */
    master_sector_allocation_table_first_sector: [u8; 4],
    /**
    Total number of sectors used for the master sector allocation table (MSAT).
     */
    master_sector_allocation_table_len: [u8; 4],
    /**
    This array of 32-bit integer fields contains the first 109 FAT sector locations of
    the compound file.
     */
    #[derivative(Debug = "ignore")]
    sector_allocation_table_head: Vec<u32>,
}
pub async fn parse_raw_header<R>(read: &mut R) -> Result<RawFileHeader>
where
    R: Readable,
{
    let mut header = [0u8; constants::HEADER_LENGTH];
    let bytes_read = read.read(&mut header).await?;
    if bytes_read != constants::HEADER_LENGTH {
        return Err(Error::OleInvalidHeader(HeaderErrorType::NotEnoughBytes(
            constants::HEADER_LENGTH,
            bytes_read,
        )));
    }

    //https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-CFB/%5bMS-CFB%5d.pdf
    //Identification signature for the compound file structure, and MUST be
    // set to the value 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1.
    let _: [u8; 8] = (&header[0..8])
        .try_into()
        .map_err(|err: TryFromSliceError| {
            Error::OleInvalidHeader(HeaderErrorType::Parsing("signature", err.to_string()))
        })
        .and_then(|signature: [u8; 8]| {
            if signature != constants::MAGIC_BYTES {
                Err(Error::OleInvalidHeader(HeaderErrorType::WrongMagicBytes(
                    signature.into(),
                )))
            } else {
                Ok(signature)
            }
        })?;

    //https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-CFB/%5bMS-CFB%5d.pdf
    //Reserved and unused class ID that MUST be set to all zeroes
    let _: [u8; 16] = (&header[8..24])
        .try_into()
        .map_err(|err: TryFromSliceError| {
            Error::OleInvalidHeader(HeaderErrorType::Parsing(
                "class_identifier",
                err.to_string(),
            ))
        })
        .and_then(|class_identifier| {
            if class_identifier != [0u8; 16] {
                Err(Error::OleInvalidHeader(HeaderErrorType::Parsing(
                    "class_identifier",
                    "non-zero entries in class_identifier field".to_string(),
                )))
            } else {
                Ok(class_identifier)
            }
        })?;
    //https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-CFB/%5bMS-CFB%5d.pdf
    //says this SHOULD be set to 0x003E.
    let minor_version: [u8; 2] = (&header[24..26])
        .try_into()
        .map_err(|err: TryFromSliceError| {
            Error::OleInvalidHeader(HeaderErrorType::Parsing("minor_version", err.to_string()))
        })
        .and_then(|minor_version| {
            if minor_version != constants::CORRECT_MINOR_VERSION {
                Err(Error::OleInvalidHeader(HeaderErrorType::Parsing(
                    "minor_version",
                    format!("incorrect minor version {:x?}", minor_version),
                )))
            } else {
                Ok(minor_version)
            }
        })?;
    //https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-CFB/%5bMS-CFB%5d.pdf
    //This field MUST be set to either
    // 0x0003 (version 3) or 0x0004 (version 4).
    let major_version: [u8; 2] = (&header[26..28])
        .try_into()
        .map_err(|err: TryFromSliceError| {
            Error::OleInvalidHeader(HeaderErrorType::Parsing("major_version", err.to_string()))
        })
        .and_then(|major_version: [u8; 2]| match major_version {
            constants::MAJOR_VERSION_3 | constants::MAJOR_VERSION_4 => Ok(major_version),
            _ => Err(Error::OleInvalidHeader(HeaderErrorType::Parsing(
                "major_version",
                format!("incorrect major version {:x?}", major_version),
            ))),
        })?;
    //https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-CFB/%5bMS-CFB%5d.pdf
    //This field MUST be set to 0xFFFE. This field is a byte order mark for all integer
    // fields, specifying little-endian byte order.
    let _: [u8; 2] = (&header[28..30])
        .try_into()
        .map_err(|err: TryFromSliceError| {
            Error::OleInvalidHeader(HeaderErrorType::Parsing(
                "byte_order_identifier",
                err.to_string(),
            ))
        })
        .and_then(
            |byte_order_identifier: [u8; 2]| match byte_order_identifier {
                [0xFE, 0xFF] => Ok(byte_order_identifier),
                _ => Err(Error::OleInvalidHeader(HeaderErrorType::Parsing(
                    "byte_order_identifier",
                    format!(
                        "incorrect byte order identifier {:x?}",
                        byte_order_identifier
                    ),
                ))),
            },
        )?;
    //https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-CFB/%5bMS-CFB%5d.pdf
    //This field MUST be set to 0x0009, or 0x000c, depending on the Major
    // Version field. This field specifies the sector size of the compound file as a power of 2.
    //  If Major Version is 3, the Sector Shift MUST be 0x0009, specifying a sector size of 512 bytes.
    //  If Major Version is 4, the Sector Shift MUST be 0x000C, specifying a sector size of 4096 bytes.
    let sector_size: [u8; 2] = (&header[30..32])
        .try_into()
        .map_err(|err: TryFromSliceError| {
            Error::OleInvalidHeader(HeaderErrorType::Parsing("sector_size", err.to_string()))
        })
        .and_then(|sector_size: [u8; 2]| match major_version {
            constants::MAJOR_VERSION_3 if sector_size == constants::SECTOR_SIZE_VERSION_3 => {
                Ok(sector_size)
            }
            constants::MAJOR_VERSION_4 if sector_size == constants::SECTOR_SIZE_VERSION_4 => {
                Ok(sector_size)
            }
            _ => Err(Error::OleInvalidHeader(HeaderErrorType::Parsing(
                "sector_size",
                format!(
                    "incorrect sector size {:x?} for major version {:x?}",
                    sector_size, major_version
                ),
            ))),
        })?;
    //https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-CFB/%5bMS-CFB%5d.pdf
    //This field MUST be set to 0x0006. This field specifies the sector size of
    // the Mini Stream as a power of 2. The sector size of the Mini Stream MUST be 64 bytes.
    let mini_sector_size: [u8; 2] = (&header[32..34])
        .try_into()
        .map_err(|err: TryFromSliceError| {
            Error::OleInvalidHeader(HeaderErrorType::Parsing(
                "mini_sector_size",
                err.to_string(),
            ))
        })
        .and_then(|mini_sector_size: [u8; 2]| match mini_sector_size {
            [0x06, 0x00] => Ok(mini_sector_size),
            _ => Err(Error::OleInvalidHeader(HeaderErrorType::Parsing(
                "mini_sector_size",
                format!("incorrect mini sector size {:x?}", mini_sector_size),
            ))),
        })?;
    let _: [u8; 6] = (&header[34..40])
        .try_into()
        .map_err(|err: TryFromSliceError| {
            Error::OleInvalidHeader(HeaderErrorType::Parsing("first_reserved", err.to_string()))
        })
        .and_then(|reserved| {
            if reserved != [0u8; 6] {
                Err(Error::OleInvalidHeader(HeaderErrorType::Parsing(
                    "first_reserved",
                    "non-zero entries in reserved field".to_string(),
                )))
            } else {
                Ok(reserved)
            }
        })?;
    //https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-CFB/%5bMS-CFB%5d.pdf
    //If Major Version is 3, the Number of Directory Sectors MUST be zero. This field is not
    // supported for version 3 compound files.
    let directory_sectors_len: [u8; 4] = (&header[40..44])
        .try_into()
        .map_err(|err: TryFromSliceError| {
            Error::OleInvalidHeader(HeaderErrorType::Parsing(
                "directory_sectors_len",
                err.to_string(),
            ))
        })
        .and_then(|directory_sectors_len| {
            if directory_sectors_len != [0u8; 4] && major_version == constants::MAJOR_VERSION_3 {
                Err(Error::OleInvalidHeader(HeaderErrorType::Parsing(
                    "directory_sectors_len",
                    "non-zero number of directory sectors with major version 3".to_string(),
                )))
            } else {
                Ok(directory_sectors_len)
            }
        })?;
    let sector_allocation_table_len: [u8; 4] =
        (&header[44..48])
            .try_into()
            .map_err(|err: TryFromSliceError| {
                Error::OleInvalidHeader(HeaderErrorType::Parsing(
                    "sector_allocation_table_len",
                    err.to_string(),
                ))
            })?;
    let sector_allocation_table_first_sector: [u8; 4] =
        (&header[48..52])
            .try_into()
            .map_err(|err: TryFromSliceError| {
                Error::OleInvalidHeader(HeaderErrorType::Parsing(
                    "sector_allocation_table_first_sector",
                    err.to_string(),
                ))
            })?;
    let _: [u8; 4] = (&header[52..56])
        .try_into()
        .map_err(|err: TryFromSliceError| {
            Error::OleInvalidHeader(HeaderErrorType::Parsing(
                "transaction_signature_number",
                err.to_string(),
            ))
        })?;
    //This integer field MUST be set to 0x00001000. This field
    // specifies the maximum size of a user-defined data stream that is allocated from the mini FAT
    // and mini stream, and that cutoff is 4,096 bytes. Any user-defined data stream that is greater than
    // or equal to this cutoff size must be allocated as normal sectors from the FAT.
    let standard_stream_min_size: [u8; 4] = (&header[56..60])
        .try_into()
        .map_err(|err: TryFromSliceError| {
            Error::OleInvalidHeader(HeaderErrorType::Parsing(
                "standard_stream_min_size",
                err.to_string(),
            ))
        })
        .and_then(|standard_stream_min_size| {
            if standard_stream_min_size != constants::CORRECT_STANDARD_STREAM_MIN_SIZE {
                Err(Error::OleInvalidHeader(HeaderErrorType::Parsing(
                    "standard_stream_min_size",
                    format!(
                        "incorrect standard_stream_min_size {:x?}",
                        standard_stream_min_size
                    ),
                )))
            } else {
                Ok(standard_stream_min_size)
            }
        })?;
    let short_sector_allocation_table_first_sector: [u8; 4] = (&header[60..64])
        .try_into()
        .map_err(|err: TryFromSliceError| {
            Error::OleInvalidHeader(HeaderErrorType::Parsing(
                "short_sector_allocation_table_first_sector",
                err.to_string(),
            ))
        })?;
    let short_sector_allocation_table_len: [u8; 4] =
        (&header[64..68])
            .try_into()
            .map_err(|err: TryFromSliceError| {
                Error::OleInvalidHeader(HeaderErrorType::Parsing(
                    "short_sector_allocation_table_len",
                    err.to_string(),
                ))
            })?;
    let master_sector_allocation_table_first_sector: [u8; 4] = (&header[68..72])
        .try_into()
        .map_err(|err: TryFromSliceError| {
            Error::OleInvalidHeader(HeaderErrorType::Parsing(
                "master_sector_allocation_table_first_sector",
                err.to_string(),
            ))
        })?;
    let master_sector_allocation_table_len: [u8; 4] =
        (&header[72..76])
            .try_into()
            .map_err(|err: TryFromSliceError| {
                Error::OleInvalidHeader(HeaderErrorType::Parsing(
                    "master_sector_allocation_table_len",
                    err.to_string(),
                ))
            })?;

    let sector_allocation_table_head = (&header[76..512])
        .chunks_exact(4)
        .map(|quad| u32::from_le_bytes([quad[0], quad[1], quad[2], quad[3]]))
        .collect::<Vec<_>>();

    Ok(RawFileHeader {
        minor_version,
        major_version,
        sector_size,
        mini_sector_size,
        directory_sectors_len,
        sector_allocation_table_len,
        sector_allocation_table_first_sector,
        standard_stream_min_size,
        short_sector_allocation_table_first_sector,
        short_sector_allocation_table_len,
        master_sector_allocation_table_first_sector,
        master_sector_allocation_table_len,
        sector_allocation_table_head,
    })
}
