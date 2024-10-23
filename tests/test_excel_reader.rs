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

    Record::write_excel("中文名称2.xlsx", "第二个中文名称", &record_list).unwrap();
}
