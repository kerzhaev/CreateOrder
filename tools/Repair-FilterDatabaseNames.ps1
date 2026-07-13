param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$resolvedPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$backupPath = "$resolvedPath.before-filter-name-repair"
Copy-Item -LiteralPath $resolvedPath -Destination $backupPath -Force

$archive = [System.IO.Compression.ZipFile]::Open($resolvedPath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    $entry = $archive.GetEntry('xl/workbook.xml')
    if ($null -eq $entry) { throw 'xl/workbook.xml was not found.' }

    $reader = New-Object System.IO.StreamReader($entry.Open())
    $xml = $reader.ReadToEnd()
    $reader.Close()

    # Excel resolves both spellings to the same built-in local name.  Leave the
    # original _FilterDatabase records and remove only their duplicate xlnm form.
    $updatedXml = [regex]::Replace($xml, '<definedName name="_xlnm\._FilterDatabase"[^>]*>.*?</definedName>', '')
    if ($updatedXml -eq $xml) { throw 'No duplicate _xlnm._FilterDatabase names were found.' }

    $entry.Delete()
    $newEntry = $archive.CreateEntry('xl/workbook.xml')
    $writer = New-Object System.IO.StreamWriter($newEntry.Open(), (New-Object System.Text.UTF8Encoding($false)))
    $writer.Write($updatedXml)
    $writer.Close()
}
finally {
    $archive.Dispose()
}

Write-Output "Repaired duplicate filter names. Backup: $backupPath"
