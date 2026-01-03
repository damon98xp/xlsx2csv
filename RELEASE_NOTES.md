# xlsx2csv Releases

## v0.1.0
- **CLI Argument Parsing**: Full clap-based argument parsing with 20+ options
- **Sheet Selection**:
  - Filter by name (`-n/--sheetname`), ID (`-s/--sheet`), or all sheets (`-a`)
  - Include/exclude patterns (`-I/--include_sheet_pattern`, `-E/--exclude_sheet_pattern`)
  - Exclude hidden sheets (`--exclude-hidden-sheets`)
- **Output Customization**:
  - Configurable delimiter (`-d`), line terminator (`-l`), sheet delimiter (`-p`)
  - Quoting styles (`-q`): none, minimal, nonnumeric, all
  - Encoding option (`-c`, currently UTF-8 only)
- **Row/Column Handling**:
  - Skip empty rows (`-i/--ignoreempty`)
  - Skip trailing empty columns (`--skipemptycolumns`)
  - Include hidden rows (`--include-hidden-rows`)
  - Merge cells (`-m/--merge-cells`)
- **Text Processing**:
  - Escape special characters (`-e/--escape`)
  - Replace line breaks with spaces (`--no-line-breaks`)
  - Include hyperlinks (`--hyperlinks`)
- **Format Overrides**:
  - Date format (`-f/--dateformat`)
  - Time format (`-t/--timeformat`)
  - Float format (`--floatformat`)
  - Force scientific notation to float (`--sci-float`)
  - Ignore specific format types (`--ignore-formats`)
- **Version Information**: `--version` flag shows current version

## v0.0.1
- New streaming XLSX→CSV pipeline using `zip` + `quick-xml` (no full workbook load).
- Default behavior: merge all sheets when no sheet is specified; write to stdout when `-o` is omitted.
- Shared strings and inline strings are supported; namespace-safe sheet discovery.
- Flexible CSV writer handles uneven rows across sheets.
- Broken pipe tolerant: exits quietly when downstream consumers (e.g., `head`, `rga`) close early.
