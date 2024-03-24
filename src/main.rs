use calamine::{open_workbook, Reader, Xlsx};
use clap::{App, Arg};
use csv::Writer;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    let matches = App::new("XLSX to CSV Converter")
        .arg(
            Arg::with_name("input")
                .required(true)
                .help("Input XLSX file"),
        )
        .arg(
            Arg::with_name("output")
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
        .get_matches();

    let input_file = matches.value_of("input").expect("not input file");
    let output_file = matches.value_of("output");
    let sheet_name = matches.value_of("sheet");
    let use_sheet_name = matches.is_present("use_sheet_name");

    let mut workbook: Xlsx<_> = open_workbook(input_file).expect("Cannot open file");

    let sheet_names: Vec<String>;
    if sheet_name == None {
        // Retrieve all sheet names
        sheet_names = workbook.sheet_names().to_vec();
    } else {
        sheet_names = vec![String::from(sheet_name.unwrap())];
    }

    for sheet_name in sheet_names {
        let output_file_name: String;
        let mut data: Vec<Vec<String>> = Vec::new();
        if let Ok(range) = workbook.worksheet_range(&sheet_name) {
            for row in range.rows() {
                let csv_row: Vec<String> = row.iter().map(|cell| cell.to_string()).collect();
                data.push(csv_row);
            }
        } else {
            eprintln!("Sheet '{}' not found in the workbook.", sheet_name);
        }

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

    Ok(())
}

fn csv_writer(data: Vec<Vec<String>>, output_file_name: &str) -> Result<(), Box<dyn Error>> {
    let mut csv_writer = Writer::from_path(output_file_name)?;

    let max_length = data.iter().map(|row| row.len()).max().unwrap_or(0);

    for csv_row in data {
        if csv_row.iter().all(|cell| cell.is_empty()) {
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
