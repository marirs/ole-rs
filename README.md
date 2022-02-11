# OLE

A set of OLE parsers and tools to deal with OLE files.

### Requirements

- Rust 1.56+ (edition: 2021)

### Example Usage

- add dependency (default feature is to use async)
```toml
[dependencies]
ole = { git = "https://github.com/marirs/ole-rs.git", branch = "master" }
```

- dependency with blocking
```toml
[dependencies]
ole = { git = "https://github.com/marirs/ole-rs.git", branch = "master", default-features = false, features = ["blocking"] }
```

- example code
```rust
use ole::OleFile;

fn main() {
    let file = "data/oledoc1.doc_";
    let res = OleFile::from_file_blocking(file).expect("file not found");
    println!("{:#?}", &res);
    println!("entries: {:#?}", res.list_streams());
}
```

---
License: MIT or Apache