use crate::{constants, error::Error, header::OleHeader, Result};
use chrono::NaiveDateTime;
use derivative::Derivative;
use std::array::TryFromSliceError;

#[derive(Clone, Derivative, Copy, PartialEq)]
#[derivative(Debug)]
pub enum ObjectType {
    Storage,
    Stream,
    RootStorage,
}

#[derive(Clone, Derivative, Copy)]
#[derivative(Debug)]
pub enum NodeColor {
    Red,
    Black,
}

/**
https://winprotocoldoc.blob.core.windows.net/productionwindowsarchives/MS-CFB/%5bMS-CFB%5d.pdf
The directory entry array is an array of directory entries that are grouped into a directory sector.
Each storage object or stream object within a compound file is represented by a single directory
entry. The space for the directory sectors that are holding the array is allocated from the FAT.
The valid values for a stream ID, which are used in the Child ID, Right Sibling ID, and Left Sibling
ID fields, are 0 through MAXREGSID (0xFFFFFFFA). The special value NOSTREAM (0xFFFFFFFF) is
used as a terminator.
 */
#[derive(Clone, Derivative)]
#[derivative(Debug)]
pub struct DirectoryEntryRaw {
    /**
    Directory Entry Name (64 bytes): This field MUST contain a Unicode string for the storage or
    stream name encoded in UTF-16. The name MUST be terminated with a UTF-16 terminating null
    character. Thus, storage and stream names are limited to 32 UTF-16 code points, including the
    terminating null character. When locating an object in the compound file except for the root
    storage, the directory entry name is compared by using a special case-insensitive uppercase
    mapping, described in Red-Black Tree. The following characters are illegal and MUST NOT be part
    of the name: '/', '\', ':', '!'.
     */
    name: [u8; 64],
    /**
    Directory Entry Name Length (2 bytes): This field MUST match the length of the Directory Entry
    Name Unicode string in bytes. The length MUST be a multiple of 2 and include the terminating null
    character in the count. This length MUST NOT exceed 64, the maximum size of the Directory Entry
    Name field.
     */
    name_len: [u8; 2],
    /**
    Object Type (1 byte): This field MUST be 0x00, 0x01, 0x02, or 0x05, depending on the actual type
    of object. All other values are not valid.
    Name Value
    Unknown or unallocated 0x00
    Storage Object 0x01
    Stream Object 0x02
    Root Storage Object 0x05
     */
    object_type: [u8; 1],
    /**
    This field MUST be 0x00 (red) or 0x01 (black). All other values are not valid.
     */
    color_flag: [u8; 1],
    /**
    This field contains the stream ID of the left sibling. If there is no left
    sibling, the field MUST be set to NOSTREAM (0xFFFFFFFF).
    Value Meaning
    REGSID: 0x00000000 — 0xFFFFFFF9
    Regular stream ID to identify the directory entry.
    MAXREGSID: 0xFFFFFFFA
    Maximum regular stream ID.
    NOSTREAM: 0xFFFFFFFF
    If there is no left sibling.
     */
    left_sibling_id: [u8; 4],
    /**
    This field contains the stream ID of the right sibling. If there is no right
    sibling, the field MUST be set to NOSTREAM (0xFFFFFFFF).
    Value Meaning
    REGSID: 0x00000000 — 0xFFFFFFF9
    Regular stream ID to identify the directory entry.
    MAXREGSID: 0xFFFFFFFA
    Maximum regular stream ID.
    NOSTREAM: 0xFFFFFFFF
    If there is no right sibling.
     */
    right_sibling_id: [u8; 4],
    /**
    This field contains the stream ID of a child object. If there is no child object,
    including all entries for stream objects, the field MUST be set to NOSTREAM (0xFFFFFFFF).
    Value Meaning
    REGSID: 0x00000000 — 0xFFFFFFF9
    Regular stream ID to identify the directory entry.
    MAXREGSID: 0xFFFFFFFA
    Maximum regular stream ID.
    NOSTREAM: 0xFFFFFFFF
    If there is no child object
     */
    child_id: [u8; 4],
    /**
    This field contains an object class GUID, if this entry is for a storage object or
    root storage object. For a stream object, this field MUST be set to all zeroes. A value containing all
    zeroes in a storage or root storage directory entry is valid, and indicates that no object class is
    associated with the storage. If an implementation of the file format enables applications to create
    storage objects without explicitly setting an object class GUID, it MUST write all zeroes by default.
    If this value is not all zeroes, the object class GUID can be used as a parameter to start
    applications.
     */
    class_id: [u8; 16],
    /**
    This field contains the user-defined flags if this entry is for a storage object or
    root storage object. For a stream object, this field SHOULD be set to all zeroes because many
    implementations provide no way for applications to retrieve state bits from a stream object. If an
    implementation of the file format enables applications to create storage objects without explicitly
    setting state bits, it MUST write all zeroes by default.
     */
    state_bits: [u8; 4],
    /**
    This field contains the creation time for a storage object, or all zeroes to
    indicate that the creation time of the storage object was not recorded. The Windows FILETIME
    structure is used to represent this field in UTC. For a stream object, this field MUST be all zeroes.
    For a root storage object, this field MUST be all zeroes, and the creation time is retrieved or set on
    the compound file itself
     */
    creation_time: [u8; 8],
    /**
    This field contains the modification time for a storage object, or all
    zeroes to indicate that the modified time of the storage object was not recorded. The Windows
    FILETIME structure is used to represent this field in UTC. For a stream object, this field MUST be
    all zeroes. For a root storage object, this field MAY<2> be set to all zeroes, and the modified time
    is retrieved or set on the compound file itself
     */
    modification_time: [u8; 8],
    /**
    This field contains the first sector location if this is a stream
    object. For a root storage object, this field MUST contain the first sector of the mini stream, if the
    mini stream exists. For a storage object, this field MUST be set to all zeroes
     */
    starting_sector_location: [u8; 4],
    /**
    This 64-bit integer field contains the size of the user-defined data if this is
    a stream object. For a root storage object, this field contains the size of the mini stream. For a
    storage object, this field MUST be set to all zeroes.
     */
    stream_size: [u8; 8],
}

impl DirectoryEntryRaw {
    pub fn parse(unparsed_entry: &[u8]) -> Result<Self> {
        let name: [u8; 64] =
            unparsed_entry[0..64]
                .try_into()
                .map_err(|err: TryFromSliceError| {
                    Error::OleInvalidDirectoryEntry("name", err.to_string())
                })?;
        let name_len: [u8; 2] =
            unparsed_entry[64..66]
                .try_into()
                .map_err(|err: TryFromSliceError| {
                    Error::OleInvalidDirectoryEntry("name_len", err.to_string())
                })?;
        let object_type: [u8; 1] =
            unparsed_entry[66..67]
                .try_into()
                .map_err(|err: TryFromSliceError| {
                    Error::OleInvalidDirectoryEntry("object_type", err.to_string())
                })?;
        let color_flag: [u8; 1] =
            unparsed_entry[67..68]
                .try_into()
                .map_err(|err: TryFromSliceError| {
                    Error::OleInvalidDirectoryEntry("color_flag", err.to_string())
                })?;
        let left_sibling_id: [u8; 4] =
            unparsed_entry[68..72]
                .try_into()
                .map_err(|err: TryFromSliceError| {
                    Error::OleInvalidDirectoryEntry("left_sibling_id", err.to_string())
                })?;
        let right_sibling_id: [u8; 4] =
            unparsed_entry[72..76]
                .try_into()
                .map_err(|err: TryFromSliceError| {
                    Error::OleInvalidDirectoryEntry("right_sibling_id", err.to_string())
                })?;
        let child_id: [u8; 4] =
            unparsed_entry[76..80]
                .try_into()
                .map_err(|err: TryFromSliceError| {
                    Error::OleInvalidDirectoryEntry("child_id", err.to_string())
                })?;
        let class_id: [u8; 16] =
            unparsed_entry[80..96]
                .try_into()
                .map_err(|err: TryFromSliceError| {
                    Error::OleInvalidDirectoryEntry("class_id", err.to_string())
                })?;
        let state_bits: [u8; 4] =
            unparsed_entry[96..100]
                .try_into()
                .map_err(|err: TryFromSliceError| {
                    Error::OleInvalidDirectoryEntry("state_bits", err.to_string())
                })?;
        let creation_time: [u8; 8] =
            unparsed_entry[100..108]
                .try_into()
                .map_err(|err: TryFromSliceError| {
                    Error::OleInvalidDirectoryEntry("creation_time", err.to_string())
                })?;
        let modification_time: [u8; 8] =
            unparsed_entry[108..116]
                .try_into()
                .map_err(|err: TryFromSliceError| {
                    Error::OleInvalidDirectoryEntry("modification_time", err.to_string())
                })?;
        let starting_sector_location: [u8; 4] =
            unparsed_entry[116..120]
                .try_into()
                .map_err(|err: TryFromSliceError| {
                    Error::OleInvalidDirectoryEntry("starting_sector_location", err.to_string())
                })?;
        let stream_size: [u8; 8] =
            unparsed_entry[120..128]
                .try_into()
                .map_err(|err: TryFromSliceError| {
                    Error::OleInvalidDirectoryEntry("stream_size", err.to_string())
                })?;

        Ok(DirectoryEntryRaw {
            name,
            name_len,
            object_type,
            color_flag,
            left_sibling_id,
            right_sibling_id,
            child_id,
            class_id,
            state_bits,
            creation_time,
            modification_time,
            starting_sector_location,
            stream_size,
        })
    }
}

#[derive(Clone, Derivative)]
#[derivative(Debug)]
pub struct DirectoryEntry {
    index: usize,
    //the index in the directory array
    pub(crate) object_type: ObjectType,
    pub(crate) name: String,
    color: NodeColor,
    pub(crate) left_sibling_id: Option<u32>,
    pub(crate) right_sibling_id: Option<u32>,
    pub(crate) child_id: Option<u32>,

    pub(crate) class_id: Option<String>,

    //TODO: do we need this?
    #[derivative(Debug = "ignore")]
    _state_bits: [u8; 4],

    creation_time: Option<NaiveDateTime>,
    modification_time: Option<NaiveDateTime>,
    pub(crate) starting_sector_location: Option<u32>,
    pub(crate) stream_size: u64,
}

impl DirectoryEntry {
    pub(crate) fn from_raw(
        ole_file_header: &OleHeader,
        raw_directory_entry: DirectoryEntryRaw,
        index: usize,
    ) -> Result<Self> {
        // first, check to see if the directory entry is even allocated...
        let object_type = match raw_directory_entry.object_type {
            constants::OBJECT_TYPE_UNKNOWN_OR_UNALLOCATED => {
                Err(Error::OleUnknownOrUnallocatedDirectoryEntry)
            }
            constants::OBJECT_TYPE_ROOT_STORAGE => Ok(ObjectType::RootStorage),
            constants::OBJECT_TYPE_STORAGE => Ok(ObjectType::Storage),
            constants::OBJECT_TYPE_STREAM => Ok(ObjectType::Stream),
            anything_else => Err(Error::OleInvalidDirectoryEntry(
                "object_type",
                format!("invalid value: {:x?}", anything_else),
            )),
        }?;

        let name_len = u16::from_le_bytes(raw_directory_entry.name_len);
        let name_raw = &raw_directory_entry.name[0..(name_len as usize)]
            .chunks(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect::<Vec<_>>();
        let mut name = String::from_utf16(&name_raw[..])?;
        //drop the null terminator
        let _ = name.pop();
        let color = match raw_directory_entry.color_flag {
            constants::NODE_COLOR_RED => Ok(NodeColor::Red),
            constants::NODE_COLOR_BLACK => Ok(NodeColor::Black),
            anything_else => Err(Error::OleInvalidDirectoryEntry(
                "node_color",
                format!("invalid value: {:x?}", anything_else),
            )),
        }?;

        let left_sibling_id = match raw_directory_entry.left_sibling_id {
            constants::NO_STREAM => Ok(None),
            potential_value => {
                let potential_value = u32::from_le_bytes(potential_value);
                if potential_value > constants::MAX_REG_STREAM_ID_VALUE {
                    Err(Error::OleInvalidDirectoryEntry(
                        "left_sibling_id",
                        format!("invalid value: {:x?}", potential_value),
                    ))
                } else {
                    Ok(Some(potential_value))
                }
            }
        }?;
        let right_sibling_id = match raw_directory_entry.right_sibling_id {
            constants::NO_STREAM => Ok(None),
            potential_value => {
                let potential_value = u32::from_le_bytes(potential_value);
                if potential_value > constants::MAX_REG_STREAM_ID_VALUE {
                    Err(Error::OleInvalidDirectoryEntry(
                        "right_sibling_id",
                        format!("invalid value: {:x?}", potential_value),
                    ))
                } else {
                    Ok(Some(potential_value))
                }
            }
        }?;
        let child_id = match raw_directory_entry.child_id {
            constants::NO_STREAM => Ok(None),
            potential_value => {
                let potential_value = u32::from_le_bytes(potential_value);
                if potential_value > constants::MAX_REG_STREAM_ID_VALUE {
                    Err(Error::OleInvalidDirectoryEntry(
                        "child_id",
                        format!("invalid value: {:x?}", potential_value),
                    ))
                } else {
                    Ok(Some(potential_value))
                }
            }
        }?;
        //TODO: the spec says there are some validations we should carry out on these times, but I'm passing them on unmodified.
        let creation_time = match i64::from_le_bytes(raw_directory_entry.creation_time) {
            0 => None,
            time => epochs::windows_file(time),
        };
        let modification_time = match i64::from_le_bytes(raw_directory_entry.modification_time) {
            0 => None,
            time => epochs::windows_file(time),
        };

        // This field contains the first sector location if this is a stream
        // object. For a root storage object, this field MUST contain the first sector of the mini stream, if the
        // mini stream exists. For a storage object, this field MUST be set to all zeroes.
        let starting_sector_location =
            // this code previously checked that storage entries have a zero starting sector location
            // but there are known cases where this trips in real files, so removed this check.
            match (object_type, raw_directory_entry.starting_sector_location) {
                (ObjectType::Storage, _assumed_zero) => None,
                (_, location) => Some(u32::from_le_bytes(location)),
            };

        let stream_size = if ole_file_header.major_version == constants::MAJOR_VERSION_3_VALUE {
            /*
            For a version 3 compound file 512-byte sector size, the value of this field MUST be less than
            or equal to 0x80000000. (Equivalently, this requirement can be stated: the size of a stream or
            of the mini stream in a version 3 compound file MUST be less than or equal to 2 gigabytes
            (GB).) Note that as a consequence of this requirement, the most significant 32 bits of this field
            MUST be zero in a version 3 compound file. However, implementers should be aware that
            some older implementations did not initialize the most significant 32 bits of this field, and
            these bits might therefore be nonzero in files that are otherwise valid version 3 compound
            files. Although this document does not normatively specify parser behavior, it is recommended
            that parsers ignore the most significant 32 bits of this field in version 3 compound files,
            treating it as if its value were zero, unless there is a specific reason to do otherwise (for
            example, a parser whose purpose is to verify the correctness of a compound file).
             */
            let mut stream_size_modified = raw_directory_entry.stream_size;
            stream_size_modified[4] = 0x00;
            stream_size_modified[5] = 0x00;
            stream_size_modified[6] = 0x00;
            stream_size_modified[7] = 0x00;

            stream_size_modified
        } else {
            raw_directory_entry.stream_size
        };
        let stream_size = u64::from_le_bytes(stream_size);
        if stream_size != 0 && object_type == ObjectType::Storage {
            return Err(Error::OleInvalidDirectoryEntry(
                "stream_size",
                "storage object type has non-zero stream size".to_string(),
            ));
        } else if object_type == ObjectType::RootStorage && stream_size % 64 != 0 {
            return Err(Error::OleInvalidDirectoryEntry(
                "stream_size",
                "root storage object type must have stream size % 64 === 0".to_string(),
            ));
        }

        let class_id = match raw_directory_entry.class_id {
            empty if empty == [0x00; 16] => None,
            bytes => {
                let a = i32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]);
                let b = i16::from_le_bytes([bytes[4], bytes[5]]);
                let c = i16::from_le_bytes([bytes[6], bytes[7]]);

                Some(
                    format!(
                        "{:08x}-{:04x}-{:04x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}",
                        a,
                        b,
                        c,
                        bytes[8],
                        bytes[9],
                        bytes[10],
                        bytes[11],
                        bytes[12],
                        bytes[13],
                        bytes[14],
                        bytes[15]
                    )
                    .to_uppercase(),
                )
            }
        };

        Ok(Self {
            index,
            object_type,
            name,
            color,
            left_sibling_id,
            right_sibling_id,
            child_id,
            class_id,
            _state_bits: raw_directory_entry.state_bits,
            creation_time,
            modification_time,
            starting_sector_location,
            stream_size,
        })
    }
}
