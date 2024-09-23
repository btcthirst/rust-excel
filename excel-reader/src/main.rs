use calamine::{deserialize_as_f64_or_none, open_workbook, RangeDeserializerBuilder, Reader, Xlsx};
use serde::Deserialize;

#[derive(Deserialize)]
struct Record {
    metric: String,
    #[serde(deserialize_with = "deserialize_as_f64_or_none")]
    value: Option<f64>,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = format!("{}/testdata/test.xlsx", env!("CARGO_MANIFEST_DIR"));
    print!("path: {}",&path);
    let mut excel: Xlsx<_> = open_workbook(path)?;

    let range = excel
        .worksheet_range("Sheet1")
        .map_err(|_| calamine::Error::Msg("Cannot find Sheet1"))?;

    let iter_records =
        RangeDeserializerBuilder::with_headers(&["metric", "value"]).from_range(&range)?;

    for result in iter_records {
        let record: Record = result?;
        println!("\n metric={:?}, value={:?}", record.metric, record.value);
    }

    Ok(())
}

