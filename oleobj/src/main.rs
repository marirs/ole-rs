pub mod ole_object;

use std::process::exit;
use clap::{Arg, Command};
use log::{error, Level};
use simple_logger::init_with_level;
use crate::ole_object::process_file;


pub fn main(){
    // Set up logger
    init_with_level(Level::Debug).unwrap();
    
    // Get arguments.
    let args_matches = Command::new("OleObj")
        .about("A tool to parse OLE objects and files stored into various MS Office file formats (doc, xls, ppt, docx, xlsx, pptx, etc).")
        .version(env!("CARGO_PKG_VERSION"))
        .arg(
            Arg::new("file")
                .long("file")
                .short('f')
                .help("The path to the file to be processed.")
                .takes_value(true)
        ).get_matches();

    let file_path = match args_matches.value_of("file") {
        Some(t) => t,
        _=> {
            error!("File path is required.");
            exit(1);
        }
    };
    
    process_file(file_path);
}