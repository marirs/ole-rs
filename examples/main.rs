use ole::OleFile;
use std::{env::args, process::exit};

fn main() {
    let args: Vec<String> = args().collect();
    if args.len() == 1 {
        println!("please give filename!");
        exit(1);
    }
    let file = args[1].to_owned();

    let res = OleFile::from_file_blocking(file).expect("file not found");
    println!("{:#?}", &res);
    println!("entries: {:#?}", res.list_streams());
}
