# WPS Print Script

This PowerShell script allows you to print Word and Excel files using WPS Office.

## Prerequisites

- Windows operating system
- WPS Office installed
- PowerShell 5.1 or later

## Usage

```powershell
.\print-wps.ps1 -FilePath "path\to\your\file.xlsx" [options]
```

### Parameters

- `-FilePath` (Required): Path to the Word or Excel file you want to print
- `-PrinterName` (Optional): Name of the printer to use. If not specified, the default printer will be used
- `-PageSize` (Optional): Paper size to use. Values: "A4" (default), "Letter", "Legal", "A3"
- `-Orientation` (Optional): Page orientation. Values: "Portrait" (default), "Landscape"
- `-WpsDir` (Optional): WPS Office installation directory. If not specified, the script will automatically detect it from the Windows Registry

### Supported File Types

- Excel: .xlsx, .xls, .et
- Word: .docx, .doc, .wps

### Examples

Print Excel file using default settings:
```powershell
.\print-wps.ps1 -FilePath "C:\Documents\report.xlsx"
```

Print Word file with specific printer and page settings:
```powershell
.\print-wps.ps1 -FilePath "C:\Documents\document.docx" -PrinterName "HP LaserJet Pro" -PageSize "A4" -Orientation "Landscape"
```

Print using custom WPS Office installation directory:
```powershell
.\print-wps.ps1 -FilePath "C:\Documents\report.xlsx" -WpsDir "C:\Program Files\WPS Office\office6"
```

## Notes

- The script automatically detects the file type and uses the appropriate WPS Office component
- WPS Office installation path is automatically detected from the Windows Registry
- Make sure WPS Office has the necessary permissions to access the printer
- The script uses COM automation to control WPS Office
- If the automatic path detection fails, you can specify the WPS Office directory manually using `-WpsDir`
- powershell exec policy, run the following cmd in powershell as admin
```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
