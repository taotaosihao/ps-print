param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,
    [string]$PrinterName = $null,
    [string]$PageSize = $null,
    [ValidateSet("Portrait", "Landscape")]
    [string]$Orientation = "Portrait",
    [string]$WpsDir,
    [switch]$Force
)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# Function to get WPS Office path from registry
function Get-WpsPath {
    $registryPaths = @(
        "HKLM:\SOFTWARE\WOW6432Node\Kingsoft\Office\6.0\Common",
        "HKLM:\SOFTWARE\Kingsoft\Office\6.0\Common",
        "HKCU:\SOFTWARE\Kingsoft\Office\6.0\Common"
    )

    foreach ($path in $registryPaths) {
        if (Test-Path $path) {
            $installPath = (Get-ItemProperty -Path $path -ErrorAction SilentlyContinue)."InstallRoot"
            if ($installPath -and (Test-Path $installPath)) {
                $office6Path = Join-Path $installPath "office6"
                if (Test-Path $office6Path) {
                    return $office6Path
                }
            }
        }
    }
    return $null
}

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class PrinterInfo
{
    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int DeviceCapabilities(
        string lpDeviceName,
        string lpPort,
        short fwCapability,
        IntPtr lpOutput,
        IntPtr lpDevMode);

    // Constants for DeviceCapabilities
    public const short DC_PAPERS = 2;
    public const short DC_PAPERNAMES = 3;
    public const short DC_PAPERSIZE = 4;
}
"@

# Function to get printer pages by page size
function Get-PrinterPaperBySize  {
    param (
        [Parameter(Mandatory=$true)]
        [string]$PrinterName,
        [Parameter(Mandatory=$true)]
        [string]$PaperSize
    )

    # Get the printer port
    $printer = Get-CimInstance -Class Win32_Printer | Where-Object { $_.Name -eq $PrinterName }
    if (-not $printer) {
        Write-Error "Printer '$PrinterName' not found."
        return
    }
    $port = $printer.PortName

    # Get count of supported paper types
    $count = [PrinterInfo]::DeviceCapabilities($PrinterName, $port, [PrinterInfo]::DC_PAPERS, [IntPtr]::Zero, [IntPtr]::Zero)
    
    if ($count -le 0) {
        Write-Error "Failed to get paper sizes. Error code: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())"
        return
    }

    $paperNames = $printer.PrinterPaperNames
    # Allocate memory for paper IDs
    $paperIdsPtr = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($count * 2)  # short is 2 bytes
    $paperNamesPtr = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($count * 64)  # Each paper name is 64 bytes
    # $paperSizePtr = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($count * 8)  # POINT structure is 8 bytes

    try {
        # Get paper IDs
        [PrinterInfo]::DeviceCapabilities($PrinterName, $port, [PrinterInfo]::DC_PAPERS, $paperIdsPtr, [IntPtr]::Zero)
        
        # Get paper names
        [PrinterInfo]::DeviceCapabilities($PrinterName, $port, [PrinterInfo]::DC_PAPERNAMES, $paperNamesPtr, [IntPtr]::Zero)
        
        # Get paper sizes
        # [PrinterInfo]::DeviceCapabilities($PrinterName, $port, [PrinterInfo]::DC_PAPERSIZE, $paperSizePtr, [IntPtr]::Zero)

        $results = @()
        
        for ($i = 0; $i -lt $count; $i++) {
            $paperId = [System.Runtime.InteropServices.Marshal]::ReadInt16($paperIdsPtr, ($i * 2))
            # $paperName = [System.Runtime.InteropServices.Marshal]::PtrToStringUni(
            #                     [System.IntPtr]::Add($paperNamesPtr, ($i * 64)), 32)
            $paperName = $paperNames[$i]
            if ($paperNames[$i] -eq $PaperSize) {
                return [PSCustomObject]@{
                    PaperId = $paperId
                    PaperName = $paperName
                } 
            }
            # $results += [PSCustomObject]@{
            #     PaperId = $paperId
            #     PaperName = $paperName
            # }
        }

        return $results
    }
    finally {
        # Free allocated memory
        [System.Runtime.InteropServices.Marshal]::FreeHGlobal($paperIdsPtr)
        [System.Runtime.InteropServices.Marshal]::FreeHGlobal($paperNamesPtr)
        # [System.Runtime.InteropServices.Marshal]::FreeHGlobal($paperSizePtr)
    }
}
# Check if the file exists
if (-not (Test-Path $FilePath)) {
    Write-Error "File not found: $FilePath"
    exit 1
}

# Get the absolute path of the file
$FilePath = (Resolve-Path $FilePath).Path
$extension = [System.IO.Path]::GetExtension($FilePath).ToLower()

# Determine file type and set appropriate application path
$appType = switch ($extension) {
    { $_ -in ".xlsx", ".xls", ".et" } { 
        @{
            type = "Excel"
            exeName = "et.exe"
            comObject = "KET.Application"
        }
    }
    { $_ -in ".docx", ".doc", ".wps" } { 
        @{
            type = "Word"
            exeName = "wps.exe"
            comObject = "KWps.Application"
        }
    }
    default {
        Write-Error "Unsupported file type: $extension. Supported types are: .xlsx, .xls, .et, .docx, .doc, .wps"
        exit 1
    }
}

# If WpsDir not provided, try to get it from registry
if (-not $WpsDir) {
    $WpsDir = Get-WpsPath
    if (-not $WpsDir) {
        Write-Error "WPS Office installation directory not found in registry. Please provide it using -WpsDir parameter."
        exit 1
    }
    Write-Host "Found WPS Office directory: $WpsDir"
}

# Check if WPS directory exists
if (-not (Test-Path $WpsDir)) {
    Write-Error "WPS Office directory not found at: $WpsDir. Please verify WPS Office is installed."
    exit 1
}

# Get the full path to the executable
$WpsPath = Join-Path $WpsDir $appType.exeName

if (-not (Test-Path $WpsPath)) {
    Write-Error "WPS $($appType.type) executable ($($appType.exeName)) not found in directory: $WpsDir"
    exit 1
}

try {
    Write-Host "Attempting to create WPS $($appType.type) Application instance..."
    $wps = New-Object -ComObject $appType.comObject
    if ($null -eq $wps) {
        throw "Failed to create WPS $($appType.type) Application instance"
    }
    $wps.Visible = $false

    Write-Host "Opening document: $FilePath"
    $doc = switch ($appType.type) {
        "Excel" { $wps.Workbooks.Open($FilePath) }
        "Word" { $wps.Documents.Open($FilePath) }
    }
    
    if ($null -eq $doc) {
        throw "Failed to open document"
    }

    # Configure printer settings
    if ($PrinterName) {
        if ($appType.type -eq "Excel") {
            $printer = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$PrinterName'"
            if ($null -eq $printer) {
                throw "Printer '$PrinterName' not found."
            }
            Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
            Write-Host "Default printer set to '$PrinterName'"
        }
        else {
            $wps.ActivePrinter = $PrinterName
        }
    }

    $realPaperSize = $null
    if ($PageSize) {
        $realPaperSize = Get-PrinterPaperBySize -PrinterName $PrinterName -PaperSize $PageSize
        Write-Host "Printer '$PrinterName' supports page size '$($realPaperSize.PaperId)'"
    }

    # Set page orientation and size based on document type
    if ($appType.type -eq "Excel") {
        foreach ($sheet in $doc.Worksheets) {
            $sheet.PageSetup.Orientation = if ($Orientation -eq "Landscape") { 2 } else { 1 }
            
            if ($realPaperSize) {
                $sheet.PageSetup.PaperSize = $realPaperSize.PaperId
            }
        }
    } else {
        $doc.PageSetup.Orientation = if ($Orientation -eq "Landscape") { 1 } else { 0 }
        
        if ($realPaperSize) {
            $doc.PageSetup.PaperSize = $realPaperSize.PaperId
        }
    }

    Write-Host "Printing document..."
    $doc.PrintOut()
    
    Write-Host "Print job sent successfully"
}
catch {
    Write-Error "An error occurred: $_"
    exit 1
}
finally {
    try {
        if ($doc) {
            if ($appType.type -eq "Excel") {
                $doc.Saved = $true  # Prevent save prompt
                $doc.Close()
            } else {
                $doc.Saved = $true  # Prevent save prompt
                $doc.Close()
            }
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
            $doc = $null
        }
        if ($wps) {
            $wps.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wps) | Out-Null
            $wps = $null
        }
    }
    catch {
        Write-Warning "Error during cleanup: $_"
    }
    finally {
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
