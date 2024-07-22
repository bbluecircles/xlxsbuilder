use std::{collections::{HashMap, HashSet}, error::Error, fs::{self, File},  io::{BufReader, Read}, path::Path, ptr::null};

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
#[serde(rename_all = "lowercase")]
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
struct WorksheetColumn {
    #[serde(rename = "col_data_type")]
    column_type: WorkSheetColumnDataType
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
    col_formula: String,
    col_pattern: String
}

#[derive(Deserialize, Debug)]
struct WorksheetConfiguration {
    worksheet_name: String,
    worksheet_properties: WorksheetSettings,
    worksheet_column_properties: Vec<WorksheetColumnProperties>,
    worksheet_data: Value
}

#[derive(Deserialize, Debug)]
struct WorkbookConfiguration {
    workbook_name: String,
    workbook_settings: WorkbookConfigurationSettings,
    worksheets: Vec<WorksheetConfiguration>
}

#[derive(Debug)]
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
        col_formula,
        col_pattern
    } = column_configuration;

    let name_to_value = Value::String(col_name.clone());

    let header_format = Format::new()
        .set_border_bottom(FormatBorder::Medium)
        .set_background_color(Color::Theme(1, 3))
        .set_font_color(Color::Theme(0, 1))
        .set_bold()
        .set_font_size(12);

    insert_cell_to_worksheet(1, column_index, worksheet, col_data_type, &name_to_value, &header_format);

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
) -> Result<(), XlsxError> {
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


    let mut worksheet_columns: Vec<PreparedWorkSheetColumn> = vec![];

    // If column properties are found, we'll setup the headers here.
    if worksheet_column_properties.len() > 0 {
        worksheet_columns = create_columns_for_worksheet_from_configuration(worksheet, worksheet_column_properties);
    }

    // Freeze the top row only.
    worksheet.set_freeze_panes(1, 0)?;

    // @Todo -> Handle case when worksheet_column_properties is empty (create headings from JSON keys)

    // If worksheet_data is Array iterate through each object and insert as row
    if let Value::Array(array) = worksheet_data {
        let mut keys: HashSet<String> = HashSet::new();

        // if keys.len() != worksheet_columns.len() {
        //     std::process::exit(1);
        // }

        for item in array {
            if let Value::Object(map) = item {
                for key in map.keys() {
                    keys.insert(key.clone());
                }
            }
        }

        let keys_to_vec: Vec<&String> = keys.iter().collect();

        let title = "Sample Title";
        let worksheet_watermark = Image::new("logo.png")?;
    
        worksheet.set_row_height_pixels(0, 32)?;

        // Create format for the first row
        let heading_format = Format::new().set_align(FormatAlign::Right);

        // Merge all columns in the first row
        worksheet.merge_range(0, 0, 0, (keys_to_vec.len() - 1 as usize).try_into().unwrap(), "", &heading_format)?;
        
        // Format for sheet title
        let title_format = Format::new()
            .set_align(FormatAlign::VerticalCenter)
            .set_align(FormatAlign::Left);

        worksheet.write_string_with_format(0, 0, title, &title_format)?;

        // Create format for the image
        let image_format = Format::new()
            .set_align(FormatAlign::VerticalCenter)
            .set_align(FormatAlign::Right);

        worksheet.embed_image_with_format(0, 0, &worksheet_watermark, &image_format)?;

        for (row_num, data_item) in array.iter().enumerate() {
            if let Value::Object(map) = data_item {
                for (data_item_col_index, data_item_key) in keys_to_vec.iter().enumerate() {
                    if let Some(value) = map.get(*data_item_key) {
                        match value {
                            Value::String(s) => worksheet.write_string((row_num + 2) as u32, data_item_col_index as u16, s),
                            Value::Number(n) => worksheet.write_number((row_num + 2) as u32, data_item_col_index as u16, n.as_f64().unwrap()),
                            Value::Bool(b) => worksheet.write_boolean((row_num + 2) as u32, data_item_col_index as u16, *b),
                            Value::Null => worksheet.write_string((row_num + 2) as u32, data_item_col_index as u16, ""),
                            Value::Object(o) => worksheet.write_string((row_num + 2) as u32, data_item_col_index as u16, ""),
                            Value::Array(a) => worksheet.write_string((row_num + 2) as u32, data_item_col_index as u16, "")
                        };
                    }
                }
            }
        }

        // Auto fit the worksheet
        worksheet.autofit();


    }
    
    Ok(())
}

fn create_workbook_from_json(json_string: &str) -> Result<(), XlsxError>
{
    let workbook_configuration: WorkbookConfiguration = serde_json::from_str(json_string).unwrap_or_else(|e| {
        println!("Failed to deserialize JSON: {}", e);
        std::process::exit(1);
    });

    let mut new_workbook = Workbook::new();

    let WorkbookConfigurationSettings { protection } = workbook_configuration.workbook_settings;

    let worksheet_configurations: Vec<WorksheetConfiguration> = workbook_configuration.worksheets;

    worksheet_configurations.iter().for_each(|worksheet_configuration| prepare_worksheet_from_configuration(&mut new_workbook, worksheet_configuration, &protection).unwrap());

    // Save new workbook
    new_workbook.save("output.xlsx");

    Ok(())
}


fn read_json_from_file(filename: &str) -> Result<String, Box<dyn Error>> {
    println!("Filename: {}", filename);
    let json_path = Path::new(filename);

    let json_string = fs::read_to_string(json_path)?;

    Ok(json_string)
}

fn main() {
    let filename = "example.json";
    let json_string = read_json_from_file(&filename).unwrap();

    let _ = create_workbook_from_json(&json_string);

}
