use calamine::{deserialize_as_f64_or_none, deserialize_as_i64_or_none, open_workbook, RangeDeserializerBuilder, Reader, Xlsx};
use serde::Deserialize;

#[derive(Deserialize,Debug,Clone)]
struct Record {
    #[serde(default)]
    name: String,
    #[serde(default)]
    unit_of_meas: String,    
    #[serde(default,deserialize_with = "deserialize_as_f64_or_none")]
    price_in: Option<f64>,
    #[serde(default,deserialize_with = "deserialize_as_f64_or_none")]
    sum_in: Option<f64>,
    #[serde(default,deserialize_with = "deserialize_as_f64_or_none")]
    price_out: Option<f64>,
    #[serde(default,deserialize_with = "deserialize_as_i64_or_none")]
    percent: Option<i64>,   
    #[serde(default,deserialize_with = "deserialize_as_f64_or_none")]
    sum_out: Option<f64>,
    #[serde(default,deserialize_with = "deserialize_as_f64_or_none")]
    value: Option<f64>,
}

fn main()  {
    let store: Vec<Record>;
    let path = format!("../../test_data/test.xlsx");
    println!("path: {}",&path);
    
    let mut res = read_excel(path).expect("Vec rec");
    res.retain(|r| r.name != ""); // deleted record with name equal to empty
    res.retain(|r| !(r.price_in == None && r.price_out == None));
    store = res;
    

   println!("{:#?}, lenth {:?}",store, store.len())
}

fn read_excel(path: String) -> Result<Vec<Record>, Box<dyn std::error::Error>>{
    let mut res: Vec<Record> = Vec::new();
    let mut excel: Xlsx<_> = open_workbook(path)?;

    let sheet_names = excel.sheet_names();
    println!("{:?}",sheet_names);
    let range = excel
        .worksheet_range(&sheet_names[0])
        .map_err(|_| calamine::Error::Msg("Cannot find sheet_names[0]"))?;

    let iter_records =
        RangeDeserializerBuilder::with_headers(&["name", "unit_of_meas", "price_in", "sum_in","price_out", "percent", "sum_out", "value"]).from_range(&range)?;

    for result in iter_records {
        //let rec: Record = result?;
        
        res.push(result?);
        
    }

    Ok(res)
}
