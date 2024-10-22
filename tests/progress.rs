#[test]
fn tests() {
    let t = trybuild::TestCases::new();
    t.pass("tests/test_excel_reader.rs");
}
