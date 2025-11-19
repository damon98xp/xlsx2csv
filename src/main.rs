use csv::WriterBuilder;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;
use std::env;
use std::error::Error;
use std::fs::File;
use std::io::{self, BufReader, BufWriter, Read, Seek, Write};
use zip::read::ZipArchive;

type BoxResult<T> = Result<T, Box<dyn Error>>;

#[derive(Clone, Copy)]
enum CellType {
    SharedString,
    InlineStr,
    Bool,
    Number,
    Error,
    PlainStr,
}

fn main() -> BoxResult<()> {
    let (file_path, sheet_name, output_path) = parse_args()?;

    let file = File::open(&file_path)?;
    let mut archive = ZipArchive::new(file)?;

    let rels = load_relationships(&mut archive)?;
    let sheets = load_sheets(&mut archive, &rels)?;
    let shared_strings = load_shared_strings(&mut archive)?;

    let targets: Vec<(String, String)> = match sheet_name {
        Some(name) => sheets
            .into_iter()
            .filter(|(sheet, _)| sheet == &name)
            .collect(),
        None => sheets,
    };

    if targets.is_empty() {
        return Err("Cannot find requested sheet".into());
    }

    let writer: Box<dyn Write> = match output_path {
        Some(path) => Box::new(BufWriter::new(File::create(path)?)),
        None => Box::new(io::stdout()),
    };
    let mut wtr = WriterBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_writer(writer);

    for (sheet, path) in targets {
        if let Err(err) = convert_sheet(&mut archive, &path, &shared_strings, &mut wtr) {
            if is_broken_pipe(&*err) {
                return Ok(());
            }
            return Err(format!("Failed to read sheet '{sheet}': {err}").into());
        }
    }

    if let Err(err) = wtr.flush() {
        if is_broken_pipe(&err) {
            return Ok(());
        }
        return Err(err.into());
    }
    Ok(())
}

fn parse_args() -> BoxResult<(String, Option<String>, Option<String>)> {
    let mut args = env::args().skip(1); // skip binary name

    let file_path = args.next().unwrap_or_else(|| {
        eprintln!("Usage: xlsx2csv <file.xlsx> [sheet-name] [-o output.csv]");
        std::process::exit(2);
    });

    let mut sheet_name: Option<String> = None;
    let mut output_path: Option<String> = None;

    while let Some(arg) = args.next() {
        match arg.as_str() {
            "-o" | "--output" => {
                output_path = Some(args.next().unwrap_or_else(|| {
                    eprintln!("Expected output path after {}", arg);
                    std::process::exit(2);
                }));
            }
            _ if sheet_name.is_none() => {
                sheet_name = Some(arg);
            }
            _ => {
                eprintln!("Unexpected argument: {}", arg);
                eprintln!("Usage: xlsx2csv <file.xlsx> [sheet-name] [-o output.csv]");
                std::process::exit(2);
            }
        }
    }

    Ok((file_path, sheet_name, output_path))
}

fn load_relationships<R: Read + Seek>(archive: &mut ZipArchive<R>) -> BoxResult<HashMap<String, String>> {
    let mut map = HashMap::new();
    let Ok(file) = archive.by_name("xl/_rels/workbook.xml.rels") else {
        return Ok(map);
    };

    let mut reader = Reader::from_reader(BufReader::new(file));
    reader.trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if eq_local(e.name().as_ref(), b"Relationship") => {
                let mut id = None;
                let mut target = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"Id" => id = Some(attr.unescape_value()?.into_owned()),
                        b"Target" => target = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }
                if let (Some(id), Some(target)) = (id, target) {
                    map.insert(id, target);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(map)
}

fn load_sheets<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    rels: &HashMap<String, String>,
) -> BoxResult<Vec<(String, String)>> {
    let file = archive.by_name("xl/workbook.xml")?;
    let mut reader = Reader::from_reader(BufReader::new(file));
    reader.trim_text(true);
    let mut buf = Vec::new();

    let mut sheets = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if eq_local(e.name().as_ref(), b"sheet") => {
                let mut name = None;
                let mut rel_id = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"name" => name = Some(attr.unescape_value()?.into_owned()),
                        b"r:id" => rel_id = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }
                if let (Some(name), Some(rel_id)) = (name, rel_id) {
                    if let Some(target) = rels.get(&rel_id) {
                        sheets.push((name, normalize_sheet_path(target)));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    if sheets.is_empty() {
        return Err("No sheets found in workbook".into());
    }

    Ok(sheets)
}

fn normalize_sheet_path(target: &str) -> String {
    let cleaned = target.trim_start_matches('/');
    if cleaned.starts_with("xl/") {
        cleaned.to_string()
    } else {
        format!("xl/{}", cleaned)
    }
}

fn load_shared_strings<R: Read + Seek>(archive: &mut ZipArchive<R>) -> BoxResult<Vec<String>> {
    let mut strings = Vec::new();
    let Ok(file) = archive.by_name("xl/sharedStrings.xml") else {
        return Ok(strings);
    };

    let mut reader = Reader::from_reader(BufReader::new(file));
    reader.trim_text(false);
    let mut buf = Vec::new();
    let mut current = String::new();
    let mut in_string = false;
    let mut in_phonetic = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if eq_local(e.name().as_ref(), b"si") => {
                current.clear();
                in_string = true;
            }
            Event::End(e) if eq_local(e.name().as_ref(), b"si") => {
                if in_string {
                    strings.push(current.clone());
                }
                in_string = false;
            }
            Event::Start(e) if eq_local(e.name().as_ref(), b"rPh") => {
                in_phonetic = true;
            }
            Event::End(e) if eq_local(e.name().as_ref(), b"rPh") => {
                in_phonetic = false;
            }
            Event::Text(t) => {
                if in_string && !in_phonetic {
                    current.push_str(&t.unescape()?);
                }
            }
            Event::CData(t) => {
                if in_string && !in_phonetic {
                    current.push_str(&String::from_utf8_lossy(t.as_ref()));
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(strings)
}

fn convert_sheet<R: Read + Seek, W: Write>(
    archive: &mut ZipArchive<R>,
    path: &str,
    shared_strings: &[String],
    writer: &mut csv::Writer<W>,
) -> BoxResult<()> {
    let file = archive.by_name(path)?;
    let mut reader = Reader::from_reader(BufReader::new(file));
    reader.trim_text(false);

    let mut buf = Vec::new();
    let mut current_row: Vec<String> = Vec::new();
    let mut current_value = String::new();
    let mut current_col: Option<usize> = None;
    let mut cell_type = CellType::Number;
    let mut in_value_tag = false;
    let mut in_inline = false;
    let mut in_phonetic = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if eq_local(e.name().as_ref(), b"row") => {
                current_row.clear();
            }
            Event::End(e) if eq_local(e.name().as_ref(), b"row") => {
                writer.write_record(&current_row)?;
            }
            Event::Start(e) if eq_local(e.name().as_ref(), b"c") => {
                current_value.clear();
                current_col = None;
                cell_type = CellType::Number;
                in_value_tag = false;
                in_inline = false;
                in_phonetic = false;

                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"t" => {
                            let v = attr.unescape_value()?.into_owned();
                            cell_type = match v.as_str() {
                                "s" => CellType::SharedString,
                                "b" => CellType::Bool,
                                "inlineStr" => CellType::InlineStr,
                                "str" => CellType::PlainStr,
                                "e" => CellType::Error,
                                _ => CellType::Number,
                            };
                        }
                        b"r" => {
                            let reference = attr.unescape_value()?.into_owned();
                            current_col = column_index(&reference);
                        }
                        _ => {}
                    }
                }
            }
            Event::Empty(e) if eq_local(e.name().as_ref(), b"c") => {
                place_cell(&mut current_row, current_col, String::new());
            }
            Event::Start(e) if eq_local(e.name().as_ref(), b"v") => {
                in_value_tag = true;
            }
            Event::End(e) if eq_local(e.name().as_ref(), b"v") => {
                in_value_tag = false;
            }
            Event::Start(e) if eq_local(e.name().as_ref(), b"is") => {
                in_inline = true;
            }
            Event::End(e) if eq_local(e.name().as_ref(), b"is") => {
                in_inline = false;
            }
            Event::Start(e) if eq_local(e.name().as_ref(), b"rPh") => {
                in_phonetic = true;
            }
            Event::End(e) if eq_local(e.name().as_ref(), b"rPh") => {
                in_phonetic = false;
            }
            Event::Text(t) => {
                if (in_value_tag || in_inline) && !in_phonetic {
                    current_value.push_str(&t.unescape()?);
                }
            }
            Event::CData(t) => {
                if (in_value_tag || in_inline) && !in_phonetic {
                    current_value.push_str(&String::from_utf8_lossy(t.as_ref()));
                }
            }
            Event::End(e) if eq_local(e.name().as_ref(), b"c") => {
                let value = match cell_type {
                    CellType::SharedString => match current_value.trim().parse::<usize>() {
                        Ok(idx) => shared_strings.get(idx).cloned().unwrap_or_default(),
                        Err(_) => current_value.clone(),
                    },
                    CellType::Bool => match current_value.trim() {
                        "1" => "true".to_string(),
                        "0" => "false".to_string(),
                        other => other.to_string(),
                    },
                    CellType::InlineStr | CellType::PlainStr | CellType::Error => {
                        current_value.clone()
                    }
                    CellType::Number => current_value.clone(),
                };
                place_cell(&mut current_row, current_col, value);
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(())
}

fn place_cell(row: &mut Vec<String>, col_idx: Option<usize>, value: String) {
    let idx = col_idx.unwrap_or_else(|| row.len());
    if row.len() <= idx {
        row.resize(idx + 1, String::new());
    }
    row[idx] = value;
}

fn column_index(cell_ref: &str) -> Option<usize> {
    let mut col = 0usize;
    let mut has_column = false;
    for c in cell_ref.chars() {
        if c.is_ascii_alphabetic() {
            has_column = true;
            col = col * 26 + (c.to_ascii_uppercase() as usize - b'A' as usize + 1);
        } else {
            break;
        }
    }
    if has_column {
        Some(col.saturating_sub(1))
    } else {
        None
    }
}

fn eq_local(name: &[u8], expected: &[u8]) -> bool {
    let local = name
        .rsplit(|&b| b == b':')
        .next()
        .unwrap_or(name);
    local == expected
}

fn is_broken_pipe(err: &(dyn std::error::Error + 'static)) -> bool {
    if let Some(io_err) = err.downcast_ref::<io::Error>() {
        if io_err.kind() == io::ErrorKind::BrokenPipe {
            return true;
        }
    }
    if let Some(csv_err) = err.downcast_ref::<csv::Error>() {
        if let csv::ErrorKind::Io(io_err) = csv_err.kind() {
            if io_err.kind() == io::ErrorKind::BrokenPipe {
                return true;
            }
        }
    }
    if let Some(source) = err.source() {
        return is_broken_pipe(source);
    }
    false
}
