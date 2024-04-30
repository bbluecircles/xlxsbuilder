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
    let workbookCols: Vec<WorkbookColumn> = Vec::new();
    let mut rows: HashMap<String, String> = HashMap::new();

    if let Some(obj) = jsonList.as_object() {
        cols = obj.keys().collect();
    }
    if let Some(array) = jsonList.as_array() {
        for element in array {
            let mut map = HashMap::new();
            for col in cols {
                if let Some(propVal) = element.get(col) {
                    rows.insert(col as String, propVal.to_string());
                }
            }
        }
    }

    // Create workbook cols
    for col in cols {
        workbookCols.push(WorkbookColumn::new(
            col,
            "".to_owned(),
            50
        ));
    }

    let wb: Workbook = Workbook::new(workbookCols, rows, "sheet test");
}

fn main() {
    println!("Hello, world!");
}
