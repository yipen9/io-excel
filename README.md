# Cargo.toml 引入

```toml
io-excel = "0.1.0"
serde = { version = "1.0", features = ["derive"] }
calamine = "0.26.1"
rust_xlsxwriter = "0.79.0"
```

# code demo

Record::excel("中文名称.xlsx", "Sheet1")，分别为 file_path 和 sheet name

```rust
use io_excel::IOExcel;
#[derive(IOExcel, Debug)]
pub struct Record {
    #[column(name = "省份")]
    pub province: String,

    #[column(name = "城市")]
    pub city: String,
}

fn main() {
    let record_list = Record::excel("中文名称.xlsx", "Sheet1").unwrap();
    for record in record_list {
        eprintln!("{:#?}", record);
    }
}

```

# 输出结果

```rust
Record {
    province: "湖北",
    city: "武汉",
}
Record {
    province: "湖北",
    city: "孝感",
}
```
