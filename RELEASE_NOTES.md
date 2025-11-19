# xlsx2csv Releases

## v0.0.1
- New streaming XLSX→CSV pipeline using `zip` + `quick-xml` (no full workbook load).
- Default behavior: merge all sheets when no sheet is specified; write to stdout when `-o` is omitted.
- Shared strings and inline strings are supported; namespace-safe sheet discovery.
- Flexible CSV writer handles uneven rows across sheets.
- Broken pipe tolerant: exits quietly when downstream consumers (e.g., `head`, `rga`) close early.
