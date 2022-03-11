use crate::{
    encryption::{DocumentType, EncryptionHandler},
    OleFile,
};

pub(crate) struct ExcelEncryptionHandler<'a> {
    _ole_file: &'a OleFile,
    _stream_name: String,
}

impl<'a> EncryptionHandler<'a> for ExcelEncryptionHandler<'a> {
    fn doc_type(&self) -> DocumentType {
        DocumentType::Excel
    }

    fn is_encrypted(&self) -> bool {
        false
    }

    fn new(ole_file: &'a OleFile, stream_name: String) -> Self {
        Self {
            _ole_file: ole_file,
            _stream_name: stream_name,
        }
    }
}
