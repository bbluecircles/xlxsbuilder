use std::{collections::{HashMap, HashSet}, error::Error, fs::File,  io::{BufReader, Read}, path::Path, ptr::null};

use rust_xlsxwriter::*;

use serde::{Deserialize};
use serde_json::{Value, json, from_str};
mod modules;


fn createWorksheetFromWorkbook(parsed_json: Value, mut workbook: Workbook) -> Result<(), XlsxError>
{
    let worksheet = workbook.add_worksheet();

    if let Value::Array(array) = parsed_json {
        let mut keys: HashSet<String> = HashSet::new();

        for item in &array {
            if let Value::Object(map) = item {
                for key in map.keys() {
                    keys.insert(key.clone());
                }
            }
        }
        
        let mut headers: Vec<&String> = keys.iter().collect();
        headers.sort();

        let heading_format = Format::new()
            .set_border_bottom(FormatBorder::Medium)
            .set_background_color(Color::Theme(8, 4))
            .set_font_color(Color::Theme(0, 1))
            .set_bold()
            .set_font_size(18);

        for (col_num, header) in headers.iter().enumerate() {
            println!("Header: {}", header);
            worksheet.write_string_with_format(0, col_num as u16, header.to_owned(), &heading_format)?;
        }


        for (row_num, item) in array.iter().enumerate() {
            if let Value::Object(map) = item {
                for (col_num, header) in headers.iter().enumerate() {
                    if let Some(value) = map.get(*header) {
                        match value {
                            Value::String(s) => worksheet.write_string((row_num + 1) as u32, col_num as u16, s)?,
                            Value::Number(n) => worksheet.write_number((row_num + 1) as u32, col_num as u16, n.as_f64().unwrap())?,
                            Value::Bool(b) => worksheet.write_boolean((row_num + 1) as u32, col_num as u16, *b)?,
                            Value::Null => worksheet.write_string((row_num + 1) as u32, col_num as u16, "")?,
                            Value::Object(o) => worksheet.write_string((row_num + 1) as u32, col_num as u16, "")?,
                            Value::Array(a) => worksheet.write_string((row_num + 1) as u32, col_num as u16, "")?
                        };
                    }
                }
            }
        }

        // Auto fit the sheet 
        worksheet.autofit();
    }

    workbook.save("output.xlsx")?;

    Ok(())
}


fn read_json_from_file(filename: &str) -> Result<Value, Box<dyn Error>> {
    println!("Filename: {}", filename);
    let json_path = Path::new(filename);

    let file = File::open(&json_path)?;
    let reader = BufReader::new(file);

    let u = serde_json::from_reader(reader)?;

    Ok(u)
}

fn main() {
    let filename = "test.json";
    let u = read_json_from_file(filename).unwrap();
    
    println!("{}", u);

    let mut workbook = Workbook::new();
    createWorksheetFromWorkbook(u, workbook);

}
