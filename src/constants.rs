use std::marker::Unpin;
use tokio::io::AsyncRead;

pub trait Readable: Unpin + AsyncRead {}
impl Readable for tokio::fs::File {}

pub const HEADER_LENGTH: usize = 512;
pub const MAGIC_BYTES: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
pub const CORRECT_MINOR_VERSION: [u8; 2] = [0x3E, 0x00];
pub const MAJOR_VERSION_3_VALUE: u16 = 3;
pub const MAJOR_VERSION_3: [u8; 2] = [0x03, 0x00];
pub const MAJOR_VERSION_4: [u8; 2] = [0x04, 0x00];
pub const SECTOR_SIZE_VERSION_3: [u8; 2] = [0x09, 0x00];
pub const SECTOR_SIZE_VERSION_4: [u8; 2] = [0x0C, 0x00];
pub const CORRECT_STANDARD_STREAM_MIN_SIZE: [u8; 4] = [0x00, 0x10, 0x00, 0x00];

pub const CHAIN_END: u32 = 0xFFFFFFFE;
pub const UNALLOCATED_SECTOR: u32 = 0xFFFFFFFF;

pub const SIZE_OF_DIRECTORY_ENTRY: usize = 128;

pub const NODE_COLOR_RED: [u8; 1] = [0x00];
pub const NODE_COLOR_BLACK: [u8; 1] = [0x01];

pub const OBJECT_TYPE_UNKNOWN_OR_UNALLOCATED: [u8; 1] = [0x00];
pub const OBJECT_TYPE_STORAGE: [u8; 1] = [0x01];
pub const OBJECT_TYPE_STREAM: [u8; 1] = [0x02];
pub const OBJECT_TYPE_ROOT_STORAGE: [u8; 1] = [0x05];

pub const MAX_REG_STREAM_ID_VALUE: u32 = u32::from_le_bytes(MAX_REG_STREAM_ID_BYTES);
pub const MAX_REG_STREAM_ID_BYTES: [u8; 4] = [0xFA, 0xFF, 0xFF, 0xFF];
pub const NO_STREAM: [u8; 4] = [0xFF, 0xFF, 0xFF, 0xFF];
