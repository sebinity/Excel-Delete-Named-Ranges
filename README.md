# Excel-Delete-Named-Ranges
Delete all Names Ranges within an Excel Workbook

Switches:
- -Path: Path to File
- -NoBackup: Do not create a .bak file (optional)
- -OutputPath: Specify a different output path (optional)

What it does:
- Opens the .xlsx ZIP package
- Reads /xl/workbook.xml
- Removes the <definedNames> element (which contains all named ranges)
- Writes the updated workbook.xml back with correct UTF-8 encoding
- Leaves the rest of the workbook intact
