use log::{debug, error, info};
use ole::ftype::OleFileType;
use ole::util::StringUtils;
use ole::OleFile;
use std::cmp::max;
use std::fs;
use std::io::{BufRead, Cursor, Read};
use std::path::Path;

/// OLE object contained into an OLENativeStream structure.
/// (see MS-OLEDS 2.3.6 OLENativeStream)  Filename and paths are
/// decoded to unicode.
pub struct OleNativeStream {
    filename: Option<String>,
    src_path: Option<String>,
    unknown_short: Option<u16>,
    unknown_long_1: Option<u32>,
    unknown_long_2: Option<u32>,
    temp_path: Option<String>,
    native_data_size: u32,
    actual_size: Option<u32>,
    data: Vec<u8>,
    package: bool,
}

impl OleNativeStream {
    /// Constructor for OleNativeStream.
    /// If bindata is provided, it will be parsed using the parse() method.
    /// :param bindata: forwarded to parse, see docu there
    /// :param package: bool, set to True when extracting from an OLE Package
    /// object
    pub fn new(bin_data: Option<Vec<u8>>, package: bool) -> Self {
        let mut instance = OleNativeStream {
            filename: None,
            src_path: None,
            unknown_short: None,
            unknown_long_1: None,
            unknown_long_2: None,
            temp_path: None,
            native_data_size: 0,
            actual_size: None,
            data: Vec::new(),
            package,
        };
        if let Some(data) = bin_data {
            instance.parse(data);
        }
        instance
    }

    /// Parse binary data containing an OLENativeStream structure,
    /// to extract the OLE object it contains.
    /// (see MS-OLEDS 2.3.6 OLENativeStream)  
    /// **Params**
    /// - data: bytes array or stream, containing OLENativeStream
    /// structure containing an OLE object  
    ///
    /// **returns** None
    pub fn parse(&mut self, data: Vec<u8>) {
        let mut cursor = Cursor::new(data);
        // An Ole package does not have the native data size field.
        if !self.package {
            self.native_data_size = read_u32(&mut cursor);
            debug!(
                "OLE native data size = {0:08X} ({} bytes)",
                self.native_data_size
            );
        }
        // Probably an ole type specifier.
        self.unknown_short = Some(read_u16(&mut cursor));
        let mut filename_buf: Vec<u8> = Vec::new();
        cursor.read_until(0x00, &mut filename_buf).unwrap();
        // The filename.
        self.filename = Some(String::from_utf8(filename_buf).unwrap());
        // The source path
        let mut source_path_buf: Vec<u8> = Vec::new();
        cursor.read_until(0x00, &mut source_path_buf).unwrap();
        self.src_path = Some(String::from_utf8(source_path_buf).unwrap());
        // Most probably time stamps.
        self.unknown_long_1 = Some(read_u32(&mut cursor));
        self.unknown_long_2 = Some(read_u32(&mut cursor));
        // The temp path
        let mut temp_path_buf: Vec<u8> = Vec::new();
        cursor.read_until(0x00, &mut temp_path_buf).unwrap();
        self.temp_path = Some(String::from_utf8(temp_path_buf).unwrap());
        // Size the rest of the data.
        self.actual_size = Some(read_u32(&mut cursor));
        cursor.read(&mut self.data).unwrap();
    }
}

/// find embedded objects in given file
pub fn process_file(filepath: &str) {
    let sane_filename = sanitize_filepath(filepath);
    // let base_dir = Path::new(filepath).parent().unwrap();
    // let filename_prefix = base_dir.join(sane_filename.clone());

    println!("{}", vec!["-"; 79].join(""));
    println!("File: {}", filepath);
    // let index = 1;

    // Look for ole files inside file.
    for ole in find_ole(filepath) {
        for parts_path in ole.list_streams() {
            let stream_path = Path::new("/").join(parts_path.clone());
            debug!("Checking stream {}", stream_path.display());
            if parts_path.to_lowercase() == "\x01ole10native".to_string() {
                println!(
                    "Extract file embedded in OLE object from stream {}",
                    stream_path.display()
                );
                println!("Parsing OLE Package");
                let stream = ole.open_stream(&vec![parts_path.as_str()]).unwrap();
                let opkg = OleNativeStream::new(Some(stream.clone()), false);

                println!("Filename = {}", opkg.filename.as_ref().cloned().unwrap());
                println!("Source path = {}", opkg.src_path.as_ref().cloned().unwrap());
                println!("Temp path = {}", opkg.temp_path.as_ref().cloned().unwrap());
                let mut fname = String::new();
                for embedded_fname in get_sane_embedded_filenames(
                    opkg.filename.as_ref().cloned().unwrap(),
                    opkg.src_path.as_ref().cloned().unwrap(),
                    opkg.temp_path.as_ref().cloned().unwrap(),
                ) {
                    fname = format!("{}_{}", sane_filename, embedded_fname);
                    println!("{}", fname);
                    if !Path::new(fname.as_str()).is_file() {
                        break;
                    }
                }
                // Dump
                println!("Saving to file {}", fname.clone());
                fs::write(fname, stream).unwrap();
            }
        }
    }
}

/// Get some sane filenames out of path information, preserving file suffix.
/// Returns several canddiates, first with suffix, then without, then random
/// with suffix and finally one last attempt ignoring max_len using arg
/// `noname_index`.
/// In some malware examples, filename (on which we relied sofar exclusively
/// for this) is empty or " ", but src_path and tmp_path contain paths with
/// proper file names. Try to extract filename from any of those.
/// Preservation of suffix is especially important since that controls how
/// windows treats the file.
pub fn get_sane_embedded_filenames(
    filename: String,
    source_path: String,
    temp_path: String,
) -> Vec<String> {
    let mut filenames = Vec::new();
    let mut suffixes = Vec::new();
    let mut candidates_without_suffixes = Vec::new();
    for mut candidate in [filename, source_path, temp_path] {
        let mut index = max(
            match candidate.rfind("/") {
                Some(t) => t,
                _ => 0,
            },
            match candidate.rfind("\\") {
                Some(t) => t,
                _ => 0,
            },
        );

        candidate = candidate.substring(index + 1..).trim().to_string();

        candidate = sanitize_filepath(candidate.as_str());
        // Skip whitespace only.
        if candidate.len() == 0 {
            continue;
        }
        if candidate.rfind(".").is_none() {
            candidates_without_suffixes.push(candidate);
            continue;
        }
        index = candidate.rfind(".").unwrap();
        if index < candidate.len() - 5 {
            candidates_without_suffixes.push(candidate);
            continue;
        }
        suffixes.push(candidate.substring(index..));
        filenames.push(candidate);
    }
    filenames
}

/// try to open somehow as zip/ole/rtf/... ; yield None if fail
/// yields embedded ole streams in form of OleFileIO.
fn find_ole(filename: &str) -> Vec<OleFile> {
    return match OleFile::from_file_blocking(filename) {
        Ok(t) => {
            return match t.file_type {
                OleFileType::Powerpoint97 => {
                    info!("Is a powerpoint file {}", filename);
                    find_ole_in_ppt(t)
                }
                _ => {
                    // An OLE file of another format.
                    info!("Is an OLE file {}", filename);
                    vec![t.clone()]
                }
            };
        }
        _ => {
            // TODO: Try loading the file as a zip file
            error!("Open failed: {} (or its data) is not an OLE.", filename);
            vec![]
        }
    };
}

/// find ole streams in ppt
/// This may be a bit confusing: we get an ole file (or its name) as input and
/// as output we produce possibly several ole files. This is because the
/// data structure can be pretty nested:
/// A ppt file has many streams that consist of records. Some of these records
/// can contain data which contains data for another complete ole file (which
/// we yield). This embedded ole file can have several streams, one of which
/// can contain the actual embedded file we are looking for (caller will check
/// for these).
fn find_ole_in_ppt(olefile: OleFile) -> Vec<OleFile> {
    // We could just return the default file.
    // let ppt_file: Option<i32> = None;
    // for stream in olefile.directory_entries {
    //     stream
    // }
    vec![olefile]
}

/// Return filename that is save to work with.
/// Removes path components, replaces all non-whitelisted characters (so output
/// is always a pure-ascii string), replaces '..' and '  ' and shortens to
/// given max length, trying to preserve suffix.
/// Might return empty string
fn sanitize_filepath(filepath: &str) -> String {
    let sane_filepath = filepath.replace("..", ".");
    sane_filepath.clone()
}

fn read_u32(cursor: &mut Cursor<Vec<u8>>) -> u32 {
    let mut data_size_buf = vec![0; 4];
    cursor.read(&mut data_size_buf).unwrap();
    let response = u32::from_ne_bytes(if cfg!(target_endian = "big") {
        [
            data_size_buf[0],
            data_size_buf[1],
            data_size_buf[2],
            data_size_buf[3],
        ]
    } else {
        [
            data_size_buf[3],
            data_size_buf[2],
            data_size_buf[1],
            data_size_buf[0],
        ]
    });
    response
}

fn read_u16(cursor: &mut Cursor<Vec<u8>>) -> u16 {
    let mut unknown_short_buf = vec![0; 2];
    cursor.read(&mut unknown_short_buf).unwrap();
    let short = u16::from_ne_bytes(if cfg!(target_endian = "big") {
        [unknown_short_buf[0], unknown_short_buf[1]]
    } else {
        [unknown_short_buf[1], unknown_short_buf[0]]
    });
    short
}
