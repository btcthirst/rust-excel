mod excel;

fn main()  {
    let _store: Vec<excel::Record>;
    let path = format!("../../test_data/test.xlsx");
    println!("path: {}",&path);
    
    _store = excel::from_excel_to_struct(path).expect("Vec rec");    

    //println!("{:#?}, lenth {:?}",store, store.len())
    let path = format!("../../test_data/test.xlsx");
    excel::from_excel_simple(path)
}