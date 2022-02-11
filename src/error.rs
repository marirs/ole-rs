use thiserror::Error;

#[derive(Error, Debug)]
pub enum HeaderErrorType {
    #[error("the magic number was expected but not found, found {0:?} instead")]
    WrongMagicBytes(Vec<u8>),
    #[error("tried to read {0} bytes, found {1} bytes")]
    NotEnoughBytes(usize, usize),
    #[error("ParsingLocation => {0} UnderlyingError => {1}")]
    Parsing(&'static str, String),
}

#[derive(Error, Debug)]
pub enum OleError {
    #[error("InvalidHeader => {0}")]
    InvalidHeader(HeaderErrorType),
    #[error("StdIo => {0}")]
    StdIo(#[from] std::io::Error),
    #[error("UnexpectedEof => {0}")]
    UnexpectedEof(String),
    #[error("CurrentlyUnimplemented => {0}")]
    CurrentlyUnimplemented(String),
    #[error("InvalidDirectoryEntry => {0}")]
    InvalidDirectoryEntry(&'static str, String),
    #[error("FromUtf16 => {0}")]
    FromUtf16(#[from] std::string::FromUtf16Error),
    #[error("UnknownOrUnallocatedDirectoryEntry")]
    UnknownOrUnallocatedDirectoryEntry,
    #[error("DirectoryEntryNotFound")]
    DirectoryEntryNotFound,
}
