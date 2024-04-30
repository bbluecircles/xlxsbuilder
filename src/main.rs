use std::{collections::HashMap, hash::Hash};

use serde::{Deserialize};
use serde_json::{Value, json};
mod modules;

struct WorkbookColumn {
    field: String,
    format: String,
    size: u16
}

impl WorkbookColumn {
    pub fn new(
        field: String,
        format: String,
        size: u16
    ) {
        WorkbookColumn { field, format, size };
    }
}

struct Workbook {
    columns: Vec<WorkbookColumn>,
    rows: HashMap<String, String>,
    name: String
}

impl Workbook {
    pub fn new(
        columns: Vec<WorkbookColumn>,
        rows: HashMap<String, String>,
        name: String
    ) {
        Workbook { columns, rows, name };
    }
}

fn createColsAndRowsFromJson(jsonList: Value) -> Workbook {
    let cols: Vec<String> = Vec::new();
    let workbook_cols: Vec<WorkbookColumn> = Vec::new();
    let mut rows: HashMap<String, String> = HashMap::new();

    if let Some(obj) = jsonList.as_object() {
        let json_to_keys: Vec<&String> = obj.keys().collect();
        cols = json_to_keys.into_iter().cloned().collect();
    }
    if let Some(array) = jsonList.as_array() {
        for element in array {
            let mut map: HashMap<_, _> = HashMap::new();
            for col in cols {
                if let Some(propVal) = element.get(col) {
                    rows.insert(col as String, propVal.to_string());
                }
            }
        }
    }

    // Create workbook cols
    for col in cols {
        let new_col: WorkbookColumn = WorkbookColumn {
            field: col,
            format: "".to_owned(),
            size: 50
        };
        workbook_cols.push(new_col);
    }

    Workbook {
        columns: workbook_cols, 
        rows: rows,
        name: "sheet test".to_owned()
    }
}

fn main() {
    println!("Hello, world!");
}
