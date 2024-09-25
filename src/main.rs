use std::{collections::{HashMap, HashSet}, error::Error, ffi::{CStr, CString}, fs::{self, File}, io::{BufReader, Read}, os::raw::{c_char, c_void}, path::Path, ptr::null};

use rust_xlsxwriter::*;
use chrono::NaiveDate;

use serde::{Deserialize, Deserializer};
use serde_json::{Value, json, from_str};
mod modules;

#[derive(Deserialize, Debug)]
struct Variables {
    format: String,
    value: String
}

#[derive(Deserialize, Debug)]
struct JsonTable {
    table_name: String,
    variables: Vec<Variables>,
    headers: Vec<String>,
    data: Vec<Vec<f64>>,
    columns: u16,
    rows: u32
}

#[derive(Deserialize, Debug)]
struct Tables {
    tables: Vec<JsonTable>
}

enum CellDataType {
    Number(f64),
    Date(NaiveDate),
    String(String)
}

// fn determine_data_type(s: &str) -> Result<(CellDataType, Option<Format>), XlsxError> {
//     let s = s.trim();

//     if s.contains('$') {
//         // Remove '$' and commas, then try to parse as a number.
//         let s_clean = s.replace('$', "").replace(",", "").trim();
//         if let Ok(value) = s_clean.parse::<f64>() {
//             let mut currency_format = Format::new();
//             currency_format.set_num_format("$#,##0.00");
//             return Ok((CellDataType::Number(value), Some(currency_format)));
//         }
//     }

//     if s.ends_with('%') {
//         let s_clean = s.trim_end_matches('%').replace(",", "").trim();
//         if let Ok(value) = s_clean.parse::<f64>() {
//             let mut percent_format = Format::new();
//             percent_format.set_num_format("0.00%");
//             return Ok((CellDataType::Number(value / 100.00), Some(percent_format)));
//         }
//     }

//     if let Ok(date) = NaiveDate::parse_from_str(s, "%Y-%m-%d") {
//         let mut date_format = Format::new();
//         date_format.set_num_format("yyyy-mm-dd");
//         return Ok((CellDataType::Date(date), Some(date_format)));
//     }

//     if let Ok(value) = s.replace(",", "").parse::<f64>() {
//         return Ok((CellDataType::Number(value), None));
//     }

//     Ok((CellDataType::String(s.to_string()), None))
// }

fn create_table(table_info: &JsonTable, worksheet: & mut Worksheet, col_start: u16, row_start: u32) {
    let columns: Vec<String> = table_info.headers.clone();

    let _ = worksheet.write_column(row_start + 1, col_start, columns);
    let _ = worksheet.write_row_matrix(row_start + 1, col_start + 1, table_info.data.clone());

    // Create column headers
    let heading_columns_to_table_column_headers: Vec<TableColumn> = table_info.variables.iter().map(|ti| TableColumn::new().set_header(ti.value.clone()).set_format(Format::new().set_num_format(ti.format.clone())).set_total_function(TableFunction::Sum)).collect();

    // Create a new table and set heading columns
    let table = Table::new()
        .set_style(TableStyle::Medium27)
        .set_columns(&heading_columns_to_table_column_headers)
        .set_total_row(true);

    let table_cell_height = (row_start + table_info.rows) - 1;
    let table_cell_width = (col_start + table_info.columns) - 1;


    // Add the table to the worksheet
    let _ = worksheet.add_table(row_start, col_start, table_cell_height, table_cell_width, &table);

    // Resize columns based on content
    for (i, header) in table_info.variables.iter().enumerate() {
        let col = col_start + i as u16;
        let mut max_length = header.value.len();

        // Check the length of each data item in the column
        for row in &table_info.data {
            for  cell in row {
                let cell_length = format!("{:?}", cell).len(); 
                if cell_length > max_length {
                    max_length = cell_length;
                }
            }
        }

        // Set the column width based on the max_length (with adjusted scaling factor)
        let width = (max_length as f64 + 2.8) * 1.2;
        let _ = worksheet.set_column_width(col, width);
    }
}

fn create_new_workbook(config: Tables) /*-> Result<Vec<u8>, Box<dyn Error>>*/ {
    let mut workbook = Workbook::new();
    let mut worksheet = workbook.add_worksheet();

    const MAX_COLUMNS: u16 = 10;
    const TABLE_GAP: u16 = 3;

    let mut col_start: u16 = 0;
    let mut row_start: u32 = 0;
    let mut max_table_height: u32 = 0;


    for (index, table_config) in config.tables.into_iter().enumerate() {
        let table_width = table_config.columns as u16;
        let table_height = table_config.rows as u32 + 2;

        if col_start + table_width > MAX_COLUMNS {
            col_start = 0;
            row_start += max_table_height + 2;
            max_table_height = table_height;
        } else {
            if table_height > max_table_height {
                max_table_height = table_height;
            }
        }

        create_table(&table_config, &mut worksheet, col_start, row_start);

        // Move col_start for the next able
        col_start += table_width + TABLE_GAP;
    }
    // Save workbook.
   //let buffer = workbook.save_to_buffer()?;
   //Ok(buffer)
   let _ = workbook.save("tables.xlsx");
}


fn read_json_from_file(filename: &str) -> Result<String, Box<dyn Error>> {
    println!("Filename: {}", filename);
    let json_path = Path::new(filename);

    let json_string = fs::read_to_string(json_path)?;

    Ok(json_string)
}

// #[no_mangle]
// pub extern "C" fn create_workbook_from_json(
//     json_string: *const c_char,
//     out_buffer: *mut *mut u8,
//     out_size: *mut usize
// )  -> *mut c_char {
//     if json_string.is_null() {
//         let error_message = CString::new("Null pointer received").unwrap();

//         return error_message.into_raw();
//     }

//     let c_str = unsafe { CStr::from_ptr(json_string) };
//     let json_str = match c_str.to_str() {
//         Ok(s) => s,
//         Err(_) => {
//             let error_message = CString::new("Invalid UTF8 sequence").unwrap();
//             return error_message.into_raw();
//         }
//     };

//     // Deserialize JSON
//     let to_deserialized: Tables = match from_str(json_str) {
//         Ok(s) => s,
//         Err(_) => {
//             let error_message = CString::new("Could not deserialize JSON").unwrap();
//             return error_message.into_raw();
//         }
//     };

//     // Process the workbook creation
//     match create_new_workbook(to_deserialized) {
//         Ok(buffer) => {
//             unsafe {
//                 *out_size = buffer.len();
//                 let buf = libc::malloc(buffer.len()) as *mut u8;
//                 if buf.is_null() {
//                     let error_message = CString::new("Failed to allocate memory").unwrap();
//                     return error_message.into_raw();
//                 }
//                 std::ptr::copy_nonoverlapping(buffer.as_ptr(), buf, buffer.len());
//                 *out_buffer = buf;
//             }
//             std::ptr::null_mut()
//         },
//         Err(e) => {
//             let error_message = CString::new(format!("Error creating workbook: {:?}", e)).unwrap();

//             return error_message.into_raw();
//         }
//     }
// }

#[no_mangle]
pub extern "C" fn free_buffer(buffer: *mut u8, size: usize) {
    if !buffer.is_null() {
        unsafe {
            libc::free(buffer as *mut c_void);
        }
    }
}

#[no_mangle]
pub extern "C" fn free_rust_string(s: *mut c_char) {
    if !s.is_null() {
        unsafe {
            CString::from_raw(s);
        }
    }
}

fn main() {
    let filename = "test2.json";
    let json_string = read_json_from_file(&filename).unwrap();

    //println!("JSON STRING: {:?}", json_string);

    let to_deserialized: Tables = match from_str(&json_string) {
        Ok(s) => s,
        Err(e) => panic!("{}", format!("ERR: {:?}", e))
    };

    create_new_workbook(to_deserialized);
}
