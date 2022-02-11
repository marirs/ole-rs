#[derive(thiserror::Error, Debug)]
pub enum Error {
    // OLE Errors
    #[error("InvalidHeader => {0}")]
    OleInvalidHeader(HeaderErrorType),
    #[error("CurrentlyUnimplemented => {0}")]
    CurrentlyUnimplemented(String),
    #[error("InvalidDirectoryEntry => {0}")]
    OleInvalidDirectoryEntry(&'static str, String),
    #[error("UnknownOrUnallocatedDirectoryEntry")]
    OleUnknownOrUnallocatedDirectoryEntry,
    #[error("DirectoryEntryNotFound")]
    OleDirectoryEntryNotFound,
    #[error("UnexpectedEof => {0}")]
    OleUnexpectedEof(String),

    // Std Errors
    #[error("StdIo => {0}")]
    StdIo(#[from] std::io::Error),
    #[error("FromUtf16 => {0}")]
    FromUtf16(#[from] std::string::FromUtf16Error),

    // Generic Error
    #[error("{0}")]
    GenericError(&'static str),
}

#[derive(thiserror::Error, Debug)]
pub enum HeaderErrorType {
    #[error("the magic number was expected but not found, found {0:?} instead")]
    WrongMagicBytes(Vec<u8>),
    #[error("tried to read {0} bytes, found {1} bytes")]
    NotEnoughBytes(usize, usize),
    #[error("ParsingLocation => {0} UnderlyingError => {1}")]
    Parsing(&'static str, String),
}
