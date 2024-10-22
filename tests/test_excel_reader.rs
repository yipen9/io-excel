use io_excel::IOExcel;

#[derive(IOExcel)]
pub struct StoreRecord {
    #[column(name = "B商家ID")]
    pub b_vender_id: String,
    #[column(name = "B商家名称")]
    pub b_vender_name: String,
    #[column(name = "B门店ID")]
    pub b_store_id: String,
    #[column(name = "B门店编码")]
    pub b_store_code: String,
    #[column(name = "B门店名称")]
    pub b_store_name: String,
    #[column(name = "C商家ID")]
    pub c_vender_id: String,
    #[column(name = "C商家名称")]
    pub c_vender_name: String,
    #[column(name = "C门店ID")]
    pub c_store_id: String,
    #[column(name = "C门店编码")]
    pub c_store_code: String,
    #[column(name = "C门店名称")]
    pub c_store_name: String,
    #[column(name = "C店新名称")]
    pub c_new_store_name: String,
    #[column(name = "919改C店名称")]
    pub c_store_name_919: String,
}

fn main() {
    eprintln!("1234");
    let record_list = StoreRecord::excel("BC门店映射.xlsx", "BC店映射表").unwrap();
    for record in record_list {
        eprintln!("{:#?}", record.b_store_code);
    }
}
