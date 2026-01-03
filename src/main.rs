use clap::Parser;
use csv::{QuoteStyle, WriterBuilder};
use quick_xml::events::Event;
use quick_xml::Reader;
use regex::Regex;
use std::collections::HashMap;
use std::error::Error;
use std::fs::File;
use std::io::{self, BufReader, BufWriter, Read, Seek, Write};
use zip::read::ZipArchive;

type BoxResult<T> = Result<T, Box<dyn Error>>;

const VERSION: &str = env!("CARGO_PKG_VERSION");

#[derive(Parser)]
#[command(name = "xlsx2csv")]
#[command(about = "xlsx to csv converter", version = VERSION)]
struct Args {
    /// xlsx file path, use '-' to read from STDIN
    xlsxfile: String,

    /// output csv file path
    outfile: Option<String>,

    /// export all sheets
    #[arg(short = 'a', long)]
    all: bool,

    /// encoding of output csv (default: utf-8)
    #[arg(short = 'c', long, default_value = "utf-8")]
    outputencoding: String,

    /// delimiter - columns delimiter in csv, 'tab' or 'x09' for a tab (default: comma ',')
    #[arg(short = 'd', long, default_value = ",")]
    delimiter: String,

    /// include hyperlinks
    #[arg(long)]
    hyperlinks: bool,

    /// Escape \r\n\t characters
    #[arg(short = 'e', long)]
    escape: bool,

    /// Replace \r\n\t with space
    #[arg(long = "no-line-breaks")]
    no_line_breaks: bool,

    /// exclude sheets named matching given pattern, only effects when -a option is enabled
    #[arg(short = 'E', long = "exclude_sheet_pattern")]
    exclude_sheet_pattern: Vec<String>,

    /// override date/time format (ex. %Y/%m/%d)
    #[arg(short = 'f', long)]
    dateformat: Option<String>,

    /// override time format (ex. %H/%M/%S)
    #[arg(short = 't', long)]
    timeformat: Option<String>,

    /// override float format (ex. %.15f)
    #[arg(long)]
    floatformat: Option<String>,

    /// force scientific notation to float
    #[arg(long = "sci-float")]
    sci_float: bool,

    /// only include sheets named matching given pattern, only effects when -a option is enabled
    #[arg(short = 'I', long = "include_sheet_pattern")]
    include_sheet_pattern: Vec<String>,

    /// Exclude hidden sheets from the output, only effects when -a option is enabled
    #[arg(long)]
    exclude_hidden_sheets: bool,

    /// Ignores format for specific data types
    #[arg(long = "ignore-formats")]
    ignore_formats: Vec<String>,

    /// line terminator - lines terminator in csv, '\n' '\r\n' or '\r' (default: \n)
    #[arg(short = 'l', long, default_value = "\n")]
    lineterminator: String,

    /// merge cells
    #[arg(short = 'm', long)]
    merge_cells: bool,

    /// sheet name to convert
    #[arg(short = 'n', long)]
    sheetname: Option<String>,

    /// skip empty lines
    #[arg(short = 'i', long)]
    ignoreempty: bool,

    /// skip trailing empty columns
    #[arg(long)]
    skipemptycolumns: bool,

    /// sheet delimiter used to separate sheets, pass '' if you do not need delimiter, or 'x07' or '\f' for form feed (default: '--------')
    #[arg(short = 'p', long, default_value = "--------")]
    sheetdelimiter: String,

    /// quoting - fields quoting in csv, 'none' 'minimal' 'nonnumeric' or 'all' (default: minimal)
    #[arg(short = 'q', long, default_value = "minimal")]
    quoting: String,

    /// sheet number to convert
    #[arg(short = 's', long)]
    sheet: Option<usize>,

    /// include hidden rows
    #[arg(long)]
    include_hidden_rows: bool,
}

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
    let args = Args::parse();

    // Handle version flag (already handled by clap)

    // Validate encoding
    if args.outputencoding != "utf-8" {
        eprintln!("Warning: Only UTF-8 encoding is supported in this Rust implementation");
    }

    // Parse delimiter
    let delimiter = parse_delimiter(&args.delimiter)?;

    // Parse line terminator
    let line_terminator = parse_escape_sequence(&args.lineterminator)?;

    // Parse sheet delimiter
    let sheet_delimiter = if args.sheetdelimiter.is_empty() {
        None
    } else {
        Some(parse_escape_sequence(&args.sheetdelimiter)?)
    };

    // Parse quoting style
    let quote_style = parse_quote_style(&args.quoting)?;

    // Determine if we're reading from stdin
    if args.xlsxfile == "-" {
        return Err("Reading from STDIN is not yet supported in this implementation".into());
    }

    let file = File::open(&args.xlsxfile)?;
    let mut archive = ZipArchive::new(file)?;

    let rels = load_relationships(&mut archive)?;
    let sheets = load_sheets(&mut archive, &rels)?;
    let shared_strings = load_shared_strings(&mut archive)?;

    // Filter sheets based on arguments
    let targets = filter_sheets(
        sheets,
        &args.sheetname,
        args.sheet,
        args.all,
        &args.include_sheet_pattern,
        &args.exclude_sheet_pattern,
    )?;

    if targets.is_empty() {
        return Err("No sheets found matching criteria".into());
    }

    // Setup output writer
    let writer: Box<dyn Write> = match &args.outfile {
        Some(path) if path != "-" => Box::new(BufWriter::new(File::create(path)?)),
        _ => Box::new(io::stdout()),
    };

    let mut wtr = WriterBuilder::new()
        .has_headers(false)
        .flexible(true)
        .delimiter(delimiter)
        .quote_style(quote_style)
        .terminator(csv::Terminator::Any(line_terminator.as_bytes()[0]))
        .from_writer(writer);

    let mut first_sheet = true;
    for (sheet_name, path) in targets {
        // Write sheet delimiter if not first sheet
        if !first_sheet {
            if let Some(ref delim) = sheet_delimiter {
                if let Err(err) = wtr.write_record(&[delim]) {
                    let boxed: Box<dyn Error> = Box::new(err);
                    if is_broken_pipe(&*boxed) {
                        return Ok(());
                    }
                }
            }
        }
        first_sheet = false;

        if let Err(err) = convert_sheet(
            &mut archive,
            &path,
            &shared_strings,
            &mut wtr,
            &args,
        ) {
            if is_broken_pipe(&*err) {
                return Ok(());
            }
            return Err(format!("Failed to read sheet '{sheet_name}': {err}").into());
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

fn parse_delimiter(s: &str) -> BoxResult<u8> {
    match s {
        "tab" | "\\t" | "x09" => Ok(b'\t'),
        s if s.len() == 1 => Ok(s.as_bytes()[0]),
        _ => Err(format!("Invalid delimiter: {}", s).into()),
    }
}

fn parse_escape_sequence(s: &str) -> BoxResult<String> {
    Ok(s.replace("\\n", "\n")
        .replace("\\r", "\r")
        .replace("\\t", "\t")
        .replace("\\f", "\x0C")
        .replace("x07", "\x07")
        .replace("x09", "\t"))
}

fn parse_quote_style(s: &str) -> BoxResult<QuoteStyle> {
    match s {
        "none" => Ok(QuoteStyle::Never),
        "minimal" => Ok(QuoteStyle::Necessary),
        "nonnumeric" => Ok(QuoteStyle::NonNumeric),
        "all" => Ok(QuoteStyle::Always),
        _ => Err(format!("Invalid quoting style: {}", s).into()),
    }
}

fn filter_sheets(
    sheets: Vec<(String, String)>,
    sheetname: &Option<String>,
    sheet_id: Option<usize>,
    all: bool,
    include_patterns: &[String],
    exclude_patterns: &[String],
) -> BoxResult<Vec<(String, String)>> {
    // If specific sheet name is requested
    if let Some(name) = sheetname {
        let targets: Vec<(String, String)> = sheets
            .into_iter()
            .filter(|(sheet, _)| sheet == name)
            .collect();
        if targets.is_empty() {
            return Err(format!("Cannot find sheet named '{}'", name).into());
        }
        return Ok(targets);
    }

    // If specific sheet ID is requested
    if let Some(id) = sheet_id {
        if id == 0 || id > sheets.len() {
            return Err(format!("Sheet ID {} out of range (1-{})", id, sheets.len()).into());
        }
        return Ok(vec![sheets[id - 1].clone()]);
    }

    // If all sheets or pattern matching
    let mut targets = if all {
        sheets
    } else {
        // Default to first sheet if not using -a
        sheets.into_iter().take(1).collect()
    };

    // Apply include patterns if specified
    if !include_patterns.is_empty() {
        let patterns: Vec<Regex> = include_patterns
            .iter()
            .map(|p| Regex::new(p))
            .collect::<Result<Vec<_>, _>>()?;

        targets = targets
            .into_iter()
            .filter(|(name, _)| patterns.iter().any(|p| p.is_match(name)))
            .collect();
    }

    // Apply exclude patterns if specified
    if !exclude_patterns.is_empty() {
        let patterns: Vec<Regex> = exclude_patterns
            .iter()
            .map(|p| Regex::new(p))
            .collect::<Result<Vec<_>, _>>()?;

        targets = targets
            .into_iter()
            .filter(|(name, _)| !patterns.iter().any(|p| p.is_match(name)))
            .collect();
    }

    if targets.is_empty() {
        return Err("No sheets found matching criteria".into());
    }

    Ok(targets)
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
    args: &Args,
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
                // Skip empty rows if requested
                if args.ignoreempty && current_row.iter().all(|s| s.is_empty()) {
                    continue;
                }

                // Skip trailing empty columns if requested
                let row_to_write = if args.skipemptycolumns {
                    let mut trimmed = current_row.clone();
                    while trimmed.last().map_or(false, |s| s.is_empty()) {
                        trimmed.pop();
                    }
                    trimmed
                } else {
                    current_row.clone()
                };

                writer.write_record(&row_to_write)?;
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
                let mut value = match cell_type {
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

                // Apply line break handling if requested
                if args.no_line_breaks {
                    value = value.replace('\r', " ").replace('\n', " ").replace('\t', " ");
                } else if args.escape {
                    value = value
                        .replace('\r', "\\r")
                        .replace('\n', "\\n")
                        .replace('\t', "\\t");
                }

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
