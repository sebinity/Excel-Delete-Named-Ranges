# Excel-Delete-Named-Ranges
Delete all Names Ranges within an Excel Workbook

Examples:
- Single file: .\Remove-NamedRanges.ps1 -Path "C:\Files\Workbook.xlsx" -Backup
- Folder (non-recursive): .\Remove-NamedRanges.ps1 -Path "C:\Files" -Backup
- Folder (recursive): .\Remove-NamedRanges.ps1 -Path "C:\Files" -Recurse -Backup

What it does:
- Opens the .xlsx ZIP package
- Reads /xl/workbook.xml
- Removes the <definedNames> element (which contains all named ranges)
- Writes the updated workbook.xml back with correct UTF-8 encoding
- Leaves the rest of the workbook intact
