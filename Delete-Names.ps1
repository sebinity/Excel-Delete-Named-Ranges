param(
  [Parameter(Mandatory=$true)][string]$Path,
  [string]$OutputPath,
  [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Load ZIP support (this one line is enough)
Add-Type -AssemblyName System.IO.Compression.FileSystem

function New-TempDirectory {
  $dir = Join-Path ([IO.Path]::GetTempPath()) ("xlsx_edit_" + [Guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Path $dir | Out-Null
  $dir
}

$fullIn = Resolve-Path $Path | % Path
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = $fullIn }

$tempRoot   = New-TempDirectory
$extractDir = Join-Path $tempRoot 'unzipped'
New-Item -ItemType Directory -Path $extractDir | Out-Null

try {
  # Unzip
  [IO.Compression.ZipFile]::ExtractToDirectory($fullIn, $extractDir)

  # Edit xl/workbook.xml
  $wbXmlPath = Join-Path $extractDir 'xl\workbook.xml'
  if (-not (Test-Path $wbXmlPath)) { throw "Could not find xl\workbook.xml" }

  [xml]$doc = Get-Content -LiteralPath $wbXmlPath
  $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
  $ns.AddNamespace('d', $doc.DocumentElement.NamespaceURI)
  $defined = $doc.SelectSingleNode('/d:workbook/d:definedNames', $ns)
  if ($defined) {
    [void]$defined.ParentNode.RemoveChild($defined)
    $doc.Save($wbXmlPath)
  }

  # Repack
  $tempZip = Join-Path $tempRoot 'rebuilt.zip'
  if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
  [IO.Compression.ZipFile]::CreateFromDirectory($extractDir, $tempZip, [IO.Compression.CompressionLevel]::Optimal, $false)

  # Backup (if overwriting)
  if ($OutputPath -ieq $fullIn -and -not $NoBackup) {
    $bak = "$fullIn.bak"
    if (-not (Test-Path $bak)) { Copy-Item $fullIn $bak }
  }

  if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
  Move-Item $tempZip $OutputPath
  Write-Host "✅ Saved: $OutputPath"
}
catch {
  Write-Error $_.Exception.Message
}
finally {
  if (Test-Path $tempRoot) { Remove-Item $tempRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
