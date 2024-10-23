# Cargo.toml 引入

```toml
io-excel = "0.1.3"
serde = { version = "1.0", features = ["derive"] }
calamine = "0.26.1"
rust_xlsxwriter = "0.79.0"
```

# Code **demo**

## Read Excel For Vec

Record::read_excel("中文名称.xlsx", "Sheet1")，分别为 file_path 和 sheet name

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
    let record_list = Record::read_excel("中文名称.xlsx", "Sheet1").unwrap();
    for record in &record_list {
        eprintln!("{:#?}", record);
    }
}

```

### 输出结果

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

## Write Excel With Vec

Record::write_excel("中文名称.xlsx", "Sheet1", record_list).unwrap();分别为 file_path 和 sheet name and record_list

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
    let record_list = Record::read_excel("中文名称.xlsx", "Sheet1").unwrap();
    Record::write_excel("中文名称2.xlsx", "第二个中文名称", &record_list).unwrap();
}
```

## Write Excel Which Some Column Is None

```rust
use io_excel::IOExcel;
#[derive(IOExcel, Debug)]
pub struct Record {
    #[column(name = "省份")]
    pub province: String,

    #[column(name = "城市")]
    pub city: Option<String>,

    #[column(name = "版本号")]
    pub name: u32,
}

fn main() {
    let record_list = Record::read_excel("中文名称.xlsx", "Sheet1").unwrap();
    for record in &record_list {
        eprintln!("{:#?}", record);
    }
    Record::write_excel("中文名称2.xlsx", "第二个中文名称", &record_list).unwrap();
}
```
