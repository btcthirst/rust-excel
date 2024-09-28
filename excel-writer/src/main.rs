use rust_xlsxwriter::*;

#[derive(Debug)]
struct Somename {
    name: String,
    value: i64,
    price_in: f64,
    price_out: f64
}
fn main() -> Result<(), XlsxError> {

    let mut data:Vec<Somename> = Vec::new();
    data.push(Somename{
            name: String::from("пиво ЛЬВІВСЬКЕ світле 0.5 л ск/б"),
            value: 15,
            price_in: 30.1,
            price_out: 35.0
        });
        
    data.push(Somename{
            name: String::from("пиво ЛЬВІВСЬКЕ 1715 1 л "),
            value: 10,
            price_in: 50.15,
            price_out: 65.0
        });
    data.push(Somename{
            name: String::from("пиво ЛЬВІВСЬКЕ оксамит 0.5 л з/б"),
            value: 9,
            price_in: 30.35,
            price_out: 36.0
        });
   
        

    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // Create some formats to use in the worksheet.
    let bold_format = Format::new().set_bold();
    let decimal_format = Format::new().set_num_format("0.00");
    let date_format = Format::new().set_num_format("yyyy-mm-dd");
    let merge_format = Format::new()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center);

    // Add a worksheet to the workbook.
    let worksheet = workbook.add_worksheet();
    // add worksheet name
    let _ = worksheet.set_name("Пивний двір");

    // Set the column width for clarity.
    worksheet.set_column_width(0, 35)?;
    let mut row = 0;
    worksheet.write_with_format(row, 0, "назва",&bold_format)?;
    worksheet.write_with_format(row, 1, "кількість",&bold_format)?;
    worksheet.write_with_format(row, 2, "ціна приходу",&bold_format)?;
    worksheet.write_with_format(row, 3, "сума",&bold_format)?;
    worksheet.write_with_format(row, 4, "ціна продажу",&bold_format)?;
    row +=1;
    
    for el in data {
            worksheet.write(row, 0, el.name)?;
            worksheet.write_with_format(row, 1, el.value,&decimal_format)?;
            worksheet.write_with_format(row, 2, el.price_in,&decimal_format)?;
            let formula = format!("= B{}*C{}",row+1,row+1);
            worksheet.write_formula_with_format(row, 3, Formula::new(&formula),&decimal_format)?;
            worksheet.write_with_format(row, 4, el.price_out,&decimal_format)?;
            row +=1;
    }
            
       
    /*
    // Write a string without formatting.
    worksheet.write(0, 0, "Hello")?;

    // Write a string with the bold format defined above.
    worksheet.write_with_format(1, 0, "World", &bold_format)?;

    // Write some numbers.
    worksheet.write(2, 0, 1)?;
    worksheet.write(3, 0, 2.34)?;

    // Write a number with formatting.
    worksheet.write_with_format(4, 0, 3.00, &decimal_format)?;

    // Write a formula.
    worksheet.write(5, 0, Formula::new("=SIN(PI()/4)"))?;

    // Write a date.
    let date = ExcelDateTime::from_ymd(2023, 1, 25)?;
    worksheet.write_with_format(6, 0, &date, &date_format)?;

    // Write some links.
    worksheet.write(7, 0, Url::new("https://www.rust-lang.org"))?;
    worksheet.write(8, 0, Url::new("https://www.rust-lang.org").set_text("Rust"))?;

    // Write some merged cells.
    worksheet.merge_range(9, 0, 9, 1, "Merged cells", &merge_format)?;

    // Insert an image.
    //let image = Image::new("examples/rust_logo.png")?;
    //worksheet.insert_image(1, 2, &image)?;
    */
    // Save the file to disk.
    workbook.save("../excel-reader/testdata/test1.xlsx")?;

    Ok(())
}

