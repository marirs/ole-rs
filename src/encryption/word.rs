use crate::{
    encryption::{DocumentType, EncryptionHandler},
    OleFile,
};
use packed_struct::prelude::*;

#[derive(PackedStruct, Debug, PartialEq)]
#[packed_struct(bit_numbering = "lsb0", endian = "lsb", size_bytes = "2")]
pub struct FirstFlags {
    #[packed_field(bits = "8")]
    f_dot: bool,
    #[packed_field(bits = "9")]
    f_glsy: bool,
    #[packed_field(bits = "10")]
    f_complex: bool,
    #[packed_field(bits = "11")]
    f_has_pic: bool,
    #[packed_field(bits = "12..=15")]
    c_quick_saves: Integer<u8, packed_bits::Bits<4>>,
    #[packed_field(bits = "0")]
    f_encrypted: bool,
    #[packed_field(bits = "1")]
    f_which_table_stream: bool,
    #[packed_field(bits = "2")]
    f_read_only_recommended: bool,
    #[packed_field(bits = "3")]
    f_write_reservation: bool,
    #[packed_field(bits = "4")]
    f_ext_char: bool,
    #[packed_field(bits = "5")]
    f_load_override: bool,
    #[packed_field(bits = "6")]
    f_far_east: bool,
    #[packed_field(bits = "7")]
    f_obfuscation: bool,
}

#[derive(PackedStruct, Debug, PartialEq)]
#[packed_struct(bit_numbering = "lsb0", endian = "lsb", size_bytes = "1")]
pub struct SecondFlags {
    #[packed_field(bits = "0")]
    f_mac: bool,
    #[packed_field(bits = "1")]
    f_empty_special: bool,
    #[packed_field(bits = "2")]
    f_load_override_page: bool,
    #[packed_field(bits = "3")]
    reserved_1: bool,
    #[packed_field(bits = "4")]
    reserved_2: bool,
    #[packed_field(bits = "5..=7")]
    f_spare_0: Integer<u8, packed_bits::Bits<3>>,
}

#[derive(PackedStruct, Debug)]
#[packed_struct(endian = "lsb", size_bytes = "32")]
pub struct PackedWordHeader {
    w_ident: u16,
    n_fib: u16,
    _unused: u16,
    lid: u16,
    pn_next: u16,
    #[packed_field(element_size_bytes = "2")]
    first_flags: FirstFlags,
    n_fib_back: u16,
    i_key: u32,
    environment: u8,
    #[packed_field(element_size_bytes = "1")]
    second_flags: SecondFlags,
    reserved_3: u16,
    reserved_4: u16,
    reserved_5: u32,
    reserved_6: u32,
}

pub(crate) struct WordEncryptionHandler<'a> {
    ole_file: &'a OleFile,
    stream_name: String,
}

impl<'a> EncryptionHandler<'a> for WordEncryptionHandler<'a> {
    fn doc_type(&self) -> DocumentType {
        DocumentType::Word
    }

    fn is_encrypted(&self) -> bool {
        let stream = self
            .ole_file
            .open_stream(&[self.stream_name.as_str()])
            .expect("stream has to exist");

        let bytes: Vec<u8> = stream.iter().take(32).copied().collect();
        let word_header = PackedWordHeader::unpack_from_slice(&bytes).expect("unable to unpack?");

        // println!("{word_header:#?}");
        word_header.first_flags.f_encrypted
    }

    fn new(ole_file: &'a OleFile, stream_name: String) -> Self {
        Self {
            ole_file,
            stream_name,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[tokio::test]
    pub async fn there_and_back_again() {
        //b'\xec\xa5\xc1\x00G\x00\t\x04\x00\x00\x00\x13\xbf\x004\x00\
        //         ... \x00\x00\x00\x10\x00\x00\x00\x00\x00\x04\x00\x00\x16\x04\x00\x00'
        let bytes = [
            0xec, 0xa5, 0xc1, 0x00, 0x47, 0x00, 0x09, 0x04, 0x00, 0x00, 0x00, 0x13, 0xbf, 0x00,
            0x34, 0x00, 0x00, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00, 0x04, 0x00, 0x00,
            0x16, 0x04, 0x00, 0x00,
        ];

        let unpacked = PackedWordHeader::unpack(&bytes).unwrap();
        let packed = unpacked.pack().unwrap();

        assert_eq!(packed, bytes);
    }

    #[tokio::test]
    pub async fn test_parsing_functionality() {
        //b'\xec\xa5\xc1\x00G\x00\t\x04\x00\x00\x00\x13\xbf\x004\x00\
        //         ... \x00\x00\x00\x10\x00\x00\x00\x00\x00\x04\x00\x00\x16\x04\x00\x00'
        let bytes = [
            0xec, 0xa5, 0xc1, 0x00, 0x47, 0x00, 0x09, 0x04, 0x00, 0x00, 0x00, 0x13, 0xbf, 0x00,
            0x34, 0x00, 0x00, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00, 0x04, 0x00, 0x00,
            0x16, 0x04, 0x00, 0x00,
        ];

        let unpacked = PackedWordHeader::unpack(&bytes).unwrap();
        println!("{unpacked:#0x?}");
    }
}
