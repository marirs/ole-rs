pub mod oleid;

use std::process::exit;
use clap::{Arg, Command};
use log::{error, Level};
use simple_logger::init_with_level;
use crate::oleid::OleId;

pub fn main() {
    // Set up logging
    init_with_level(Level::Debug).unwrap();
    
    // Get arguments.
    let args_matches = Command::new("oleid")
        .about("A tool to analyze OLE files such as MS Office documents (e.g. Word,
Excel), to detect specific characteristics that could potentially indicate that
the file is suspicious or malicious, in terms of security (e.g. malware).")
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
    
    let mut oleid = OleId::new(file_path);
    let indicators = oleid.check();
    println!("{:#?}", indicators);
}