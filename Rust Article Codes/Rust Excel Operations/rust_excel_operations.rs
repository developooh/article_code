use std::error::Error;

use calamine::{open_workbook, Data, Range, Reader, Xlsx};

 fn main() {

    /* This line defines the path to the Excel file. It uses the format! 
    macro to construct the path by appending "/src/data100000.xlsx" to the 
    directory of the Cargo manifest file (env!("CARGO_MANIFEST_DIR")).
    Define the path to the Excel file */
    let path = format!("{}/src/data100000.xlsx", env!("CARGO_MANIFEST_DIR"));

    /* Specify the sheet name
       Here, we specify the name of the Excel sheet we want to work with. 
       In this case, it's "Sheet1". */
    let sheet_name = "Sheet1";

    // Read data from the Excel file
    let excel_file_data = ExcelFileOperations::read_file(&path, Some(sheet_name)).unwrap();

    /* Get column names from the Excel file
    This line gets the column names from the Excel file data obtained 
    in the previous step. It calls the get_column_names method of the 
    excel_file_data variable and unwraps the Option to retrieve the column names.*/
    let column_names = excel_file_data.get_column_names().unwrap();

    // Specify the column name to search
    let search_column_name = &["Hire Date"];

    /* Filter columns based on the specified column name
    This line filters columns based on the specified column name 
    ("Hire Date"). It calls the get_column_values_by_names method of the 
    excel_file_data variable to filter the columns and stores the result 
    in the filtered_columns variable.*/
    let filtered_column_values = excel_file_data.get_column_values_by_names(search_column_name);
}

trait ExcelFileOperationTrait {
    // Signature of the function that reads an Excel file
    fn read_file(path: &str, sheet_name: Option<&str>) -> Result<Range<Data>, Box<dyn Error>>;
}

pub struct ExcelFileOperations {}

impl ExcelFileOperationTrait for ExcelFileOperations {

    // Implementation of the function that reads an Excel file
    fn read_file(path: &str, sheet_name: Option<&str>) -> Result<Range<Data>, Box<dyn Error>> {
        // Open the workbook
        let mut workbook: Xlsx<_> = open_workbook(path)?;

        // Determine the sheet name
        let sheet_name = match sheet_name {
            Some(name) => name.to_owned(),
            None => {
                // If no sheet name is provided, use the first sheet
                workbook
                    .sheet_names()
                    .get(0)
                    .ok_or("No sheets found in workbook")?
                    .clone()
            }
        };

        // Get the range of data from the specified sheet
        let worksheet_range = workbook.worksheet_range(&sheet_name)?;

        // Return the data range with Ok
        Ok(worksheet_range)
    }
}

trait RangeExtensions {
    // Returns the column names as an option
    fn get_column_names(&self) -> Option<Vec<String>>;
    
    // Returns column values based on column names
    fn get_column_values_by_names(
        &self,
        search_words: &[&str],
    ) -> Result<Vec<Vec<Data>>, Box<dyn Error>>;
}


impl RangeExtensions for Range<Data> {
    // Implementation of the trait for the Range<Data> type
    fn get_column_names(&self) -> Option<Vec<String>> {
        // Get the first row of the range
        self.rows().next().map(|row| {
            // Convert cells to strings and collect them into a vector
            row.iter()
                .map(|cell| cell.to_string())
                .collect::<Vec<String>>()
        })
    }

// This function is a method that can be called on a Range<Data> instance. 
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

        // If no matching columns are found, return an error
        if column_indices.is_empty() {
            return Err("No matching columns found".into());
        }

        // Extract column data based on column indices
        let columns_data: Vec<Vec<Data>> = column_indices
            .iter()
            .map(|&index| {
                // Extract data from the column skipping the first row (column names)
                self.rows()
                    .skip(1)
                    .filter_map(|row| row.get(index).map(|cell| cell.to_owned()))
                    .collect()
            })
            .collect();

        Ok(columns_data)
    }
}
