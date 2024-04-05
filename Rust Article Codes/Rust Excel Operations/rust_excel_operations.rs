use std::error::Error;

use calamine::{open_workbook, Data, Range, Reader, Xlsx};

fn main() {
    let path = format!("{}/src/data100000.xlsx", env!("CARGO_MANIFEST_DIR"));

    let sheet_name: &str = "Sheet1";

    let excel_file_data: Range<Data> =
        ExcelFileOperations::read_file(&path, Some(sheet_name)).unwrap();

    let _column_names: Vec<String> = excel_file_data.get_column_names().unwrap();

    let search_column_name: &[&str] = &["Hire Date"];

    let _filtered_columns_values: Result<Vec<Vec<Data>>, Box<dyn Error>> =
        excel_file_data.get_column_values_by_names(search_column_name);
}

trait ExcelFileOperationTrait {
    fn read_file(path: &str, sheet_name: Option<&str>) -> Result<Range<Data>, Box<dyn Error>>;
}

pub struct ExcelFileOperations {}

impl ExcelFileOperationTrait for ExcelFileOperations {
    fn read_file(path: &str, sheet_name: Option<&str>) -> Result<Range<Data>, Box<dyn Error>> {
        // Workbook'u aç
        let mut workbook: Xlsx<_> = open_workbook(path)?;

        // Sheet adını belirle
        let sheet_name = match sheet_name {
            Some(name) => name.to_owned(),
            None => {
                // Eğer sheet adı verilmediyse, ilk sheet'i kullan
                workbook
                    .sheet_names()
                    .get(0)
                    .ok_or("No sheets found in workbook")?
                    .clone()
            }
        };

        // Belirtilen sayfadaki veri aralığını al
        let worksheet_range = workbook.worksheet_range(&sheet_name)?;

        Ok(worksheet_range)
    }
}

trait RangeExtensions {
    fn get_column_names(&self) -> Option<Vec<String>>;
    fn get_column_values_by_names(
        &self,
        search_words: &[&str],
    ) -> Result<Vec<Vec<Data>>, Box<dyn Error>>;
}

impl RangeExtensions for Range<Data> {
    fn get_column_names(&self) -> Option<Vec<String>> {
        self.rows().next().map(|row| {
            row.iter()
                .map(|cell| cell.to_string())
                .collect::<Vec<String>>()
        })
    }

    fn get_column_values_by_names(
        &self,
        search_words: &[&str],
    ) -> Result<Vec<Vec<Data>>, Box<dyn Error>> {
        // Get Column Name List
        let column_names = self.get_column_names().ok_or("Column names not found")?;

        // Find The Index Of The Column
        let column_indices: Vec<usize> = column_names
            .iter()
            .enumerate()
            .filter_map(|(i, name)| {
                if search_words.contains(&name.as_str()) {
                    Some(i)
                } else {
                    None
                }
            })
            .collect();

        if column_indices.is_empty() {
            return Err("No matching columns found".into());
        }

        let columns_data: Vec<Vec<Data>> = column_indices
            .iter()
            .map(|&index| {
                self.rows()
                    .skip(1)
                    .filter_map(|row| row.get(index).map(|cell| cell.to_owned()))
                    .collect()
            })
            .collect();

        Ok(columns_data)
    }
}
