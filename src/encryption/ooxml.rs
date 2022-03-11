use crate::{
    encryption::{DocumentType, EncryptionHandler},
    OleFile,
};

pub(crate) struct OpenOfficeXmlEncryptionHandler<'a> {
    _ole_file: &'a OleFile,
    _stream_name: String,
}

impl<'a> EncryptionHandler<'a> for OpenOfficeXmlEncryptionHandler<'a> {
    fn doc_type(&self) -> DocumentType {
        DocumentType::Ooxml
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
