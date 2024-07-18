use std::{collections::{HashMap, HashSet}, error::Error, fs::File,  io::{BufReader, Read}, path::Path, ptr::null};

use rust_xlsxwriter::*;

use serde::Deserialize;
use serde_json::{Value, json, from_str};
mod modules;

#[derive(Deserialize, Debug)]
struct WorkbookProtectionSettings {
    should_protect: bool,
    password: String
}
#[derive(Deserialize, Debug)]
struct WorkbookConfigurationSettings {
    protection: WorkbookProtectionSettings
}

#[derive(Deserialize, Debug)]
enum WorkSheetColumnDataType {
    String,
    Number
}

impl WorkSheetColumnDataType {
    fn new_string() -> Self {
        WorkSheetColumnDataType::String
    }
    fn new_number() -> Self {
        WorkSheetColumnDataType::Number
    }
}

impl Clone for WorkSheetColumnDataType {
    fn clone(&self) -> Self {
        match self {
            WorkSheetColumnDataType::String => {
                WorkSheetColumnDataType::String 
            }
            WorkSheetColumnDataType::Number => {
                WorkSheetColumnDataType::Number
            }
        }
    }
}

#[derive(Deserialize, Debug)]
enum WorkSheetColumnFormats {

}

#[derive(Deserialize, Debug)]
struct WorksheetSettings {

}

#[derive(Deserialize, Debug)]
struct WorksheetColumnProperties {
    col_name: String,
    col_data_type: WorkSheetColumnDataType,
    col_formats: WorkSheetColumnFormats,
    col_formula: String,
    col_pattern: String
}

#[derive(Deserialize, Debug)]
struct WorksheetConfiguration {
    worksheet_name: String,
    worksheet_properties: WorksheetSettings,
    worksheet_column_properties: Vec<WorksheetColumnProperties>,
    worksheet_data: Vec<Value>
}

#[derive(Deserialize, Debug)]
struct WorkbookConfiguration {
    workbook_name: String,
    workbook_settings: WorkbookConfigurationSettings,
    worksheets: Vec<WorksheetConfiguration>
}

struct WorkbookWrapper {
    workbook: Workbook,
    filename: String
}

struct PreparedWorkSheetColumn {
    index: u16,
    format: Format,
    data_type: WorkSheetColumnDataType
}

impl PreparedWorkSheetColumn {
    fn new(index: u16, format: &Format, data_type: WorkSheetColumnDataType) -> Self {
        PreparedWorkSheetColumn { index, format: format.clone(), data_type }
    }
}

impl Clone for PreparedWorkSheetColumn {
    fn clone(&self) -> Self {
        PreparedWorkSheetColumn {
            index: self.index,
            format: self.format.clone(),
            data_type: self.data_type.clone()
        }
    }
}


fn insert_value_into_cell(row_index: u32, col_index: u16, worksheet: &mut Worksheet, value: &Value, format: &Format) {
    match value {
        Value::String(s) => worksheet.write_string_with_format(row_index as u32, col_index as u16, s, format),
        Value::Number(n) => worksheet.write_number_with_format(row_index as u32, col_index as u16, n.as_f64().unwrap(), format),
        Value::Bool(b) => worksheet.write_boolean_with_format(row_index as u32, col_index as u16, *b, format),
        Value::Null => worksheet.write_string_with_format(row_index as u32, col_index as u16, "", format),
        Value::Object(o) => worksheet.write_string_with_format(row_index as u32, col_index as u16, "", format),
        Value::Array(a) => worksheet.write_string_with_format(row_index as u32, col_index as u16, "", format)
    };
}


fn insert_cell_to_worksheet(
    row_index: u32, 
    col_index: u16, 
    worksheet: &mut Worksheet,
    data_type: &WorkSheetColumnDataType,
    value: &Value,
    format: &Format
) {
    match data_type {
        WorkSheetColumnDataType::String => insert_value_into_cell(row_index, col_index, worksheet, value, format),
        WorkSheetColumnDataType::Number => insert_value_into_cell(row_index, col_index, worksheet, value, format)
    };
}

fn create_column_for_worksheet_from_configuration(
    worksheet: &mut Worksheet,
    column_configuration: &WorksheetColumnProperties,
    column_index: u16
) -> PreparedWorkSheetColumn {
    let WorksheetColumnProperties {
        col_name,
        col_data_type,
        col_formats,
        col_formula,
        col_pattern
    } = column_configuration;

    let name_to_value = Value::String(col_name.clone());

    let header_format = Format::new()
        .set_border_bottom(FormatBorder::Medium)
        .set_background_color(Color::Theme(8, 4))
        .set_font_color(Color::Theme(0, 1))
        .set_bold()
        .set_font_size(18);

    insert_cell_to_worksheet(0, column_index, worksheet, col_data_type, &name_to_value, &header_format);

    PreparedWorkSheetColumn::new(column_index, &header_format, col_data_type.clone())
}

fn create_columns_for_worksheet_from_configuration(
    worksheet: &mut Worksheet, 
    column_configurations: &Vec<WorksheetColumnProperties>
) -> Vec<PreparedWorkSheetColumn> {
    column_configurations.into_iter()
        .enumerate()
        .map(|(index, column_properies)| create_column_for_worksheet_from_configuration(
            worksheet,
            column_properies,
            index as u16
        ))
        .collect()
}

fn prepare_worksheet_from_configuration(
    workbook: &mut Workbook, 
    worksheet_configuration: &WorksheetConfiguration, 
    workbook_protection_settings: &WorkbookProtectionSettings
) {
    let WorksheetConfiguration { 
        worksheet_name,
        worksheet_properties,
        worksheet_column_properties, 
        worksheet_data 
    } = worksheet_configuration;

    let WorkbookProtectionSettings { should_protect, password } = workbook_protection_settings;

    let worksheet= workbook.add_worksheet();

    // Set the name
    worksheet.set_name(worksheet_name);
    
    // If a password protection enabled, set it.
    if *should_protect {
        worksheet.protect_with_password(&password);
    }

    let worksheet_columns: Vec<PreparedWorkSheetColumn> = create_columns_for_worksheet_from_configuration(worksheet, worksheet_column_properties);

    // If column properties are found, we'll setup the headers here.
    if worksheet_column_properties.len() > 0 {

    }



}

fn create_workbook_from_json(json_string: &str) -> Result<WorkbookWrapper, Error>
{
    let workbook_configuration: WorkbookConfiguration = serde_json::from_str(json_string).unwrap_or_else(|e| {
        println!("Failed to deserialize JSON: {}", e);
        std::process::exit(1);
    });

    let new_workbook = Workbook::new();

    let WorkbookConfigurationSettings { protection } = workbook_configuration.workbook_settings;

    let worksheet_configurations: Vec<WorksheetConfiguration> = workbook_configuration.worksheets;

    worksheet_configurations.iter().for_each(|worksheet_configuration| prepare_worksheet_from_configuration(&mut new_workbook, worksheet_configuration, protection));
}


fn create_worksheet_from_workbook(parsed_json: Value, mut workbook: Workbook) -> Result<(), XlsxError>
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
    create_worksheet_from_workbook(u, workbook);

}
