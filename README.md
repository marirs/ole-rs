# OLE
[![Linux Arm7](https://github.com/marirs/ole-rs/actions/workflows/linux_arm.yml/badge.svg)](https://github.com/marirs/ole-rs/actions/workflows/linux_arm.yml)
[![Linux x86_64](https://github.com/marirs/ole-rs/actions/workflows/linux_x86_64.yml/badge.svg)](https://github.com/marirs/ole-rs/actions/workflows/linux_x86_64.yml)
[![macOS](https://github.com/marirs/ole-rs/actions/workflows/macos.yml/badge.svg)](https://github.com/marirs/ole-rs/actions/workflows/macos.yml)
[![Windows](https://github.com/marirs/ole-rs/actions/workflows/windows.yml/badge.svg)](https://github.com/marirs/ole-rs/actions/workflows/windows.yml)

A set of OLE parsers and tools to deal with OLE files.

### Requirements

- Rust 1.56+ (edition: 2021)

### Tools
- **OleId** : A tool to analyze OLE files such as MS Office documents (e.g. Word,
  Excel), to detect specific characteristics that could potentially indicate that
  the file is suspicious or malicious, in terms of security (e.g. malware).
- **OleObj** : A tool to parse OLE objects and files stored into various MS Office file formats (doc, xls, ppt, docx, xlsx, pptx, etc).
- **Ole-Common** : A crate that reads and parses OLE files.
## 1. OleId
This is a tool to analyze MS Office documents(eg. Word, Excel) to detect specific characteristics common in malicious files.
### Usage
```
oleid [options] <filename> 

Options

--file: The filepath to the file to process.
```

## 2.OleObj
This is a tool to parse OLE objects and files stored into various MS Office file formats (doc, xls, ppt, docx, xlsx, pptx, etc).
### Usage
```
oleobj [options] <filename> 

Options

--file: The filepath to the file to process.
```

## 3. Ole-Common
### Example Usage

- add dependency (default feature is to use async)
```toml
[dependencies]
ole-common = { git = "https://github.com/marirs/ole-rs.git", branch = "master" }
```
- example code
```rust
use olecommon::OleFile;

fn main() {
    let file = "data/oledoc1.doc_";
    let res = OleFile::from_file(file).await.expect("file not found");
    println!("{:#?}", &res);
    println!("entries: {:#?}", res.list_streams());
}
```

- dependency with blocking
```toml
[dependencies]
ole-common = { git = "https://github.com/marirs/ole-rs.git", branch = "master", default-features = false, features = ["blocking"] }
```

- example code
```rust
use olecommon::OleFile;

fn main() {
    let file = "data/oledoc1.doc_";
    let res = OleFile::from_file_blocking(file).expect("file not found");
    println!("{:#?}", &res);
    println!("entries: {:#?}", res.list_streams());
}
```

- Running the Example Code
```bash
cargo r --example ole_cli --features="blocking" data/oledoc1.doc_
```

---
License: MIT or Apache