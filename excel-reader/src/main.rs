use calamine::{deserialize_as_f64_or_none, deserialize_as_i64_or_none, open_workbook, RangeDeserializerBuilder, Reader, Xlsx};
use serde::Deserialize;

#[derive(Deserialize,Debug)]
struct Record {
    name: String,
    #[serde(deserialize_with = "deserialize_as_i64_or_none")]
    value: Option<i64>,
    #[serde(deserialize_with = "deserialize_as_f64_or_none")]
    price: Option<f64>,
    #[serde(deserialize_with = "deserialize_as_f64_or_none")]
    sum: Option<f64>,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = format!("{}/testdata/test.xlsx", env!("CARGO_MANIFEST_DIR"));
    print!("path: {}",&path);
    let mut excel: Xlsx<_> = open_workbook(path)?;

    let range = excel
        .worksheet_range("Хортиця")
        .map_err(|_| calamine::Error::Msg("Cannot find Хортиця"))?;

    let iter_records =
        RangeDeserializerBuilder::with_headers(&["name", "value", "price", "sum"]).from_range(&range)?;

    for result in iter_records {
        let record: Record = result?;
        println!("\n row={:#?}", record);
    }

    Ok(())
}

