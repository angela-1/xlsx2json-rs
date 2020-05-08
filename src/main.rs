use calamine::Error;
use calamine::{open_workbook, DataType, Range, Reader, Xlsx};
use clap::{App, Arg};
use std::collections::HashMap;
use std::fs;

/**
 * 获取xlsx文件的标题行作为json对象的键值
 */
fn get_keys(start: u32, range: &Range<DataType>) -> Vec<String> {
    let col = range.get_size().1 as u32;
    let mut keys: Vec<String> = vec![];
    for n in 0..col {
        keys.push(range.get_value((start, n)).unwrap().to_string());
    }
    keys
}

/**
 * 获取xlsx文件数据范围
 */
fn get_range(path: &str) -> Result<Range<DataType>, Error> {
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let sheet_name = workbook.sheet_names()[0].clone();
    let range = workbook
        .worksheet_range(&sheet_name)
        .ok_or(Error::Msg("Cannot read first sheet"))??;
    Ok(range)
}

/**
 * 解析一行生成一个hashmap对象
 */
fn parse_row(keys: &Vec<String>, range: &Range<DataType>) -> HashMap<String, String> {
    let mut line: HashMap<String, String> = HashMap::new();
    for n in 0..range.get_size().1 {
        line.insert(keys[n].to_string(), range.get((0, n)).unwrap().to_string());
    }
    line
}

/**
 * 转换函数
 */
fn transform() {
    let arguments = App::new("xlsx2json")
        .version("0.0.2")
        .author("Angela-1 <ruoshui_engr@163.com>")
        .about("A command line tool for parse xlsx file to json file.")
        .arg(
            Arg::with_name("input")
                .value_name("XLSX_FILE")
                .help("Input xlsx file")
                .required(true)
                .takes_value(true),
        )
        .arg(
            Arg::with_name("output")
                .short("o")
                .long("output")
                .value_name("JSON_FILE")
                .help("Output json file"),
        )
        .arg(
            Arg::with_name("row")
                .short("r")
                .long("row")
                .value_name("NUMBER")
                .help("Row number of key, zero index start, default 0"),
        )
        .get_matches();

    let file = arguments
        .value_of("input")
        .expect("Please provide a xlsx file");

    // 获取键值的行号
    let key_row: u32 = arguments.value_of("row").unwrap_or("0").parse().unwrap();

    // 默认输出json文件名
    let default_result_filename = file.to_string() + ".json";
    let dest = arguments
        .value_of("output")
        .unwrap_or(&default_result_filename);

    let range = get_range(&file).ok().expect("fail to get range");
    // 获取json对象键值
    let keys = get_keys(key_row, &range);

    let mut key_value_map: Vec<HashMap<String, String>> = Vec::new();

    let (row, col): (usize, usize) = range.get_size();
    let data_row = key_row + 1;
    for n in data_row..row as u32 {
        let rec = range.range((n, 0), (n, col as u32 - 1));
        key_value_map.push(parse_row(&keys, &rec));
    }

    // 将hashmap转换为带格式的json字符串
    let json_string = serde_json::to_string_pretty(&key_value_map).unwrap();
    fs::write(dest, json_string).expect("Unable to write file");
    println!("Done");
}

fn main() {
    transform();
}
