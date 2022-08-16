use ole::OleFile;

#[tokio::main]
async fn main() {
    let file = "data/oledoc1.doc_";
    let res = OleFile::from_file_blocking(file).expect("file not found");
    println!("{:#?}", &res);
    println!("entries: {:#?}", res.list_streams());

    let res = OleFile::from_file(file).await.expect("file not found");
    println!("{:#?}", &res);
    println!("entries: {:#?}", res.list_streams());
}
