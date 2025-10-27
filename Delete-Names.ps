param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,
    [switch]$Recurse,
    [switch]$Backup
)

# Load the required .NET assembly for ZipFile
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Remove-NamedRangesFromWorkbookXml {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XmlContent
    )

    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.LoadXml($XmlContent)

    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('s', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    # Remove the entire <definedNames> element (if present)
    $definedNames = $doc.SelectSingleNode('/s:workbook/s:definedNames', $ns)
    $changed = $false
    if ($definedNames -ne $null) {
        [void]$definedNames.ParentNode.RemoveChild($definedNames)
        $changed = $true
    }

    # Return the XmlDocument so we can save with proper UTF-8 encoding
    [pscustomobject]@{
        Doc     = $doc
        Changed = $changed
    }
}

function Remove-ExcelNamedRanges {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XlsxFile
    )

    if (-not (Test-Path -LiteralPath $XlsxFile)) {
        Write-Error "File not found: $XlsxFile"
        return
    }

    if ([System.IO.Path]::GetExtension($XlsxFile) -ne '.xlsx') {
        Write-Warning "Skipping non-.xlsx file: $XlsxFile"
        return
    }

    if ($Backup) {
        try {
            Copy-Item -LiteralPath $XlsxFile -Destination ($XlsxFile + '.bak') -ErrorAction Stop
            Write-Host "Backup created: $XlsxFile.bak"
        } catch {
            Write-Warning "Could not create backup for $XlsxFile: $_"
        }
    }

    $zip = $null
    try {
        # Open the .xlsx as a Zip archive in Update mode
        $zip = [System.IO.Compression.ZipFile]::Open($XlsxFile, [System.IO.Compression.ZipArchiveMode]::Update)

        $entry = $zip.GetEntry('xl/workbook.xml')
        if (-not $entry) {
            Write-Error "workbook.xml not found in: $XlsxFile"
            return
        }

        # Read workbook.xml
        $sr = New-Object System.IO.StreamReader($entry.Open())
        $xmlContent = $sr.ReadToEnd()
        $sr.Close()

        # Remove defined names
        $result = Remove-NamedRangesFromWorkbookXml -XmlContent $xmlContent

        if ($result.Changed) {
            # Replace workbook.xml in the ZIP with the modified content
            $entry.Delete()

            $newEntry = $zip.CreateEntry('xl/workbook.xml', [System.IO.Compression.CompressionLevel]::Optimal)
            $ws = $newEntry.Open()
            # Save with UTF-8 (no BOM) so the XML declaration matches the actual encoding
            $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
            $sw = New-Object System.IO.StreamWriter($ws, $utf8NoBom)
            $result.Doc.Save($sw)
            $sw.Flush()
            $sw.Close()
            $ws.Close()

            Write-Host "Removed named ranges from: $XlsxFile"
        } else {
            Write-Host "No named ranges found: $XlsxFile"
        }
    } catch {
        Write-Error "Failed to process $XlsxFile: $_"
    } finally {
        if ($zip) { $zip.Dispose() }
    }
}

# Entry point: process file or folder
if (-not (Test-Path -LiteralPath $Path)) {
    Write-Error "Path not found: $Path"
    exit 1
}

$item = Get-Item -LiteralPath $Path
if ($item.PSIsContainer) {
    Get-ChildItem -LiteralPath $Path -Filter '*.xlsx' -Recurse:$Recurse | ForEach-Object {
        Remove-ExcelNamedRanges -XlsxFile $_.FullName
    }
} else {
    Remove-ExcelNamedRanges -XlsxFile $item.FullName
}
