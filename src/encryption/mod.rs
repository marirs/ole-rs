use crate::{
    encryption::{
        excel::ExcelEncryptionHandler, ooxml::OpenOfficeXmlEncryptionHandler,
        powerpoint::PowerPointEncryptionHandler, word::WordEncryptionHandler,
    },
    OleFile,
};

mod excel;
mod ooxml;
mod powerpoint;
mod word;

lazy_static! {
    pub static ref WORD_DOC_STR: String = "WordDocument".to_lowercase();
    pub static ref POWER_POINT_STR: String = "PowerPoint Document".to_lowercase();
    pub static ref EXCEL_STR: String = "Workbook".to_lowercase();
    pub static ref OOXML_DOC_STR: String = "EncryptionInfo".to_lowercase();
}

#[derive(Debug, Clone, Copy)]
pub enum DocumentType {
    Word,
    Excel,
    PowerPoint,
    Ooxml,
}

pub trait EncryptionHandler<'a> {
    fn doc_type(&self) -> DocumentType;
    fn is_encrypted(&self) -> bool;
    fn new(ole_file: &'a OleFile, stream_name: String) -> Self
    where
        Self: Sized;
}

pub fn is_encrypted(ole_file: &OleFile) -> bool {
    let streams = ole_file.list_streams();
    let mut document_type = None;
    for stream in streams.into_iter() {
        match stream.to_lowercase() {
            word_doc if word_doc == *WORD_DOC_STR => {
                let handler: Box<dyn EncryptionHandler> =
                    Box::new(WordEncryptionHandler::new(ole_file, stream));
                document_type = Some(handler);
                break;
            }
            power_point if power_point == *POWER_POINT_STR => {
                let handler: Box<dyn EncryptionHandler> =
                    Box::new(PowerPointEncryptionHandler::new(ole_file, stream));
                document_type = Some(handler);
                break;
            }
            excel if excel == *EXCEL_STR => {
                let handler: Box<dyn EncryptionHandler> =
                    Box::new(ExcelEncryptionHandler::new(ole_file, stream));
                document_type = Some(handler);
                break;
            }
            ooxml if ooxml == *OOXML_DOC_STR => {
                let handler: Box<dyn EncryptionHandler> =
                    Box::new(OpenOfficeXmlEncryptionHandler::new(ole_file, stream));
                document_type = Some(handler);
                break;
            }
            _ => {}
        }
    }

    if document_type.is_none() {
        false
    } else {
        document_type.as_ref().unwrap().is_encrypted()
    }
}
