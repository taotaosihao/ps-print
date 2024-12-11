param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath,
    [string]$PrinterName = $null,
    [ValidateSet("A4", "Letter", "Legal", "A3")]
    [string]$PageSize = $null,
    [ValidateSet("Portrait", "Landscape")]
    [string]$Orientation = "Portrait",
    [string]$WpsDir,
    [switch]$Force
)

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

    # Set page orientation and size based on document type
    if ($appType.type -eq "Excel") {
        foreach ($sheet in $doc.Worksheets) {
            $sheet.PageSetup.Orientation = if ($Orientation -eq "Landscape") { 2 } else { 1 }
            
            if ($PageSize) {
                switch ($PageSize) {
                    "A4" { $sheet.PageSetup.PaperSize = 9 }
                    "Letter" { $sheet.PageSetup.PaperSize = 1 }
                    "Legal" { $sheet.PageSetup.PaperSize = 5 }
                    "A3" { $sheet.PageSetup.PaperSize = 8 }
                }
            }
        }
    } else {
        $doc.PageSetup.Orientation = if ($Orientation -eq "Landscape") { 1 } else { 0 }
        
        if ($PageSize) {
            switch ($PageSize) {
                "A4" { $doc.PageSetup.PaperSize = 9 }
                "Letter" { $doc.PageSetup.PaperSize = 1 }
                "Legal" { $doc.PageSetup.PaperSize = 5 }
                "A3" { $doc.PageSetup.PaperSize = 8 }
            }
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
