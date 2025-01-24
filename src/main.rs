use calamine::{open_workbook, Reader, Xlsx};
use clap::{App, Arg};
use csv::Writer;
use std::collections::HashMap;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    let matches = App::new("XLSX to CSV Converter")
        .arg(
            Arg::with_name("input")
                .short('i')
                .long("input")
                .takes_value(true)
                .required(true)
                .help("Input XLSX file"),
        )
        .arg(
            Arg::with_name("output")
                .short('o')
                .long("output")
                .takes_value(true)
                .help("Output CSV file (optional). if no value , use sheet name as output"),
        )
        .arg(
            Arg::with_name("sheet")
                .short('s')
                .long("sheet")
                .takes_value(true)
                .help("Sheet name (optional)"),
        )
        .arg(
            Arg::with_name("use_sheet_name")
                .short('u')
                .long("use-sheet-name")
                .takes_value(false)
                .help("Use sheet name as output (optional)"),
        )
        .arg(
            Arg::with_name("all_in_one")
                .short('a')
                .long("all-in-one")
                .takes_value(false)
                .help("All sheets will be merged in one csv. All.csv as default outfile name if --input is empty (optional)"),
        )
        .arg(
            Arg::with_name("first_line_only_once")
                .short('f')
                .long("first-line-only-once")
                .takes_value(false)
                .help("When set all in one , you can set first_line_only_once (optional)"),
        )
        .get_matches();

    let input_file = matches.value_of("input").expect("not input file");
    let output_file = matches.value_of("output");
    let sheet_name = matches.value_of("sheet");
    let use_sheet_name = matches.is_present("use_sheet_name");
    let all_in_one = matches.is_present("all_in_one");
    let first_line_only_once = matches.is_present("first_line_only_once");

    let mut workbook: Xlsx<_> = open_workbook(input_file).expect("Cannot open file");

    let sheet_names: Vec<String>;
    if sheet_name == None {
        // Retrieve all sheet names
        sheet_names = workbook.sheet_names().to_vec();
    } else {
        sheet_names = vec![String::from(sheet_name.unwrap())];
    }

    let mut all_data: HashMap<String, Vec<Vec<String>>> = HashMap::new();

    for sheet_name in sheet_names {
        let mut data: Vec<Vec<String>> = Vec::new();
        if let Ok(range) = workbook.worksheet_range(&sheet_name) {
            for row in range.rows() {
                let csv_row: Vec<String> = row.iter().map(|cell| cell.to_string().trim().to_string()).collect();
                data.push(csv_row);
            }
            all_data.insert(sheet_name.clone(), data.clone());
        } else {
            eprintln!("Sheet '{}' not found in the workbook.", sheet_name);
        }
        if all_in_one {
            continue;
        }

        let output_file_name: String;
        if use_sheet_name || output_file == None {
            output_file_name = sheet_name + ".csv";
        } else {
            output_file_name = output_file.unwrap().to_string();
        }
        csv_writer(data, output_file_name.as_str()).expect("Cannot write file correctly");
        println!(
            "Conversion complete! Output written to {}",
            output_file_name
        );
    }

    if all_in_one {
        let output_file_name: String;
        if output_file == None {
            output_file_name = "all.csv".to_string();
        } else {
            output_file_name = output_file.unwrap().to_string();
        }
        csv_write_all(all_data, output_file_name.as_str(), first_line_only_once)
            .expect("Cannot write file correctly");
        println!(
            "Conversion complete! Output written to {}",
            output_file_name
        );
    }

    Ok(())
}

fn csv_writer(data: Vec<Vec<String>>, output_file_name: &str) -> Result<(), Box<dyn Error>> {
    let mut csv_writer = Writer::from_path(output_file_name)?;

    let max_length = data.iter().map(|row| row.len()).max().unwrap_or(0);

    for csv_row in data {
        if csv_row.iter().all(|cell| cell.trim().is_empty()) {
            continue;
        }
        let adjusted_row: Vec<_> = if csv_row.len() < max_length {
            let mut adjusted_row = csv_row.iter().map(|cell| cell.trim().to_string()).collect::<Vec<_>>();
            adjusted_row.resize(max_length, String::new());
            adjusted_row
        } else {
            csv_row.iter().map(|cell| cell.trim().to_string()).collect()
        };

        csv_writer.write_record(&adjusted_row)?;
    }
    csv_writer
        .flush()
        .expect("Csv_writer Cannot flush correctly");
    Ok(())
}

fn csv_write_all(
    all_data: HashMap<String, Vec<Vec<String>>>,
    output_file_name: &str,
    first_line_only_once: bool,
) -> Result<(), Box<dyn Error>> {
    let mut csv_writer = Writer::from_path(output_file_name)?;

    let mut first_line_written = false;

    let mut data_merged: Vec<Vec<String>> = Vec::new();
    for (_, data) in &all_data {
        for (idx, csv_row) in data.iter().enumerate() {
            if first_line_written && idx == 0 {
                continue;
            }

            if csv_row.iter().all(|cell| cell.trim().is_empty()) {
                continue;
            }

            data_merged.push(csv_row.clone());
        }
        if first_line_only_once {
            first_line_written = true;
        };
    }

    let max_length = data_merged.iter().map(|row| row.len()).max().unwrap_or(0);
    for csv_row in data_merged {
        if csv_row.iter().all(|cell| cell.trim().is_empty()) {
            continue;
        }
        let adjusted_row: Vec<_> = if csv_row.len() < max_length {
            let mut adjusted_row = csv_row.clone();
            adjusted_row.resize(max_length, String::new());
            adjusted_row
        } else {
            csv_row
        };

        csv_writer.write_record(&adjusted_row)?;
    }

    csv_writer
        .flush()
        .expect("Csv_writer Cannot flush correctly");
    Ok(())
}
