use calamine::Error;
use calamine::{open_workbook, DataType, Range, Reader, Xlsx};
use clap::{App, Arg};
use std::collections::HashMap;
use std::fs;

fn get_keys(start: u32, range: &Range<DataType>) -> Vec<String> {
    let col = range.get_size().1 as u32;
    let mut keys: Vec<String> = vec![];
    for n in 0..col {
        keys.push(range.get_value((start, n)).unwrap().to_string());
    }
    keys
}

fn get_range(path: &str) -> Result<Range<DataType>, Error> {
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let sheet_name = workbook.sheet_names()[0].clone();
    let range = workbook
        .worksheet_range(&sheet_name)
        .ok_or(Error::Msg("Cannot read first sheet."))??;
    Ok(range)
}

fn parse_row(keys: &Vec<String>, range: &Range<DataType>) -> HashMap<String, String> {
    let mut line: HashMap<String, String> = HashMap::new();
    for n in 0..range.get_size().1 {
        line.insert(keys[n].to_string(), range.get((0, n)).unwrap().to_string());
    }
    line
}

fn main() {
    let matches = App::new("xlsx2json")
        .version("0.1")
        .author("Angela-1 <mail>")
        .about("xlsx2json can generate json file from xlsx file with title as keys.")
        .arg(
            Arg::with_name("INPUT")
                .help("Sets the input file to use")
                .required(true)
                .index(1),
        )
        .arg(
            Arg::with_name("OUTPUT")
                .help("Sets the output file to use")
                .index(2)
        )
        .arg(
            Arg::with_name("start")
                .short("s")
                .long("start")
                .value_name("NUMBER")
                .help("Sets title line number"),
        )
        
        .get_matches();

    let file = matches
        .value_of("INPUT")
        .expect("Please provide a xlsx file");

    let title_line = matches.value_of("start").unwrap_or("0").parse().unwrap();

    let default_result_filename = file.to_string() + ".json";
    let dest = matches.value_of("OUTPUT").unwrap_or(&default_result_filename);

    let range = get_range(&file).ok().expect("fail to get range.");
    let keys = get_keys(title_line, &range);

    let mut result: Vec<HashMap<String, String>> = Vec::new();

    let (row, col) = range.get_size();
    let data_line = title_line + 1;
    for n in data_line..row as u32 {
        let r = range.range((n, 0), (n, col as u32 - 1));
        result.push(parse_row(&keys, &r));
    }
    let s = serde_json::ser::to_string(&result).unwrap();
    fs::write(dest, s).expect("Unable to write file");
    println!("Done");
}
