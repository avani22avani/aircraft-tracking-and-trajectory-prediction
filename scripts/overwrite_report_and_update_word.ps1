param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,
    [Parameter(Mandatory = $true)]
    [string]$DestinationPath
)

$ErrorActionPreference = "Stop"

$source = (Resolve-Path -LiteralPath $SourcePath).Path
$destinationDirectory = Split-Path -Parent $DestinationPath
if (!(Test-Path -LiteralPath $destinationDirectory)) {
    New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
}

Copy-Item -LiteralPath $source -Destination $DestinationPath -Force

$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($DestinationPath)
    $document.Repaginate()
    foreach ($field in $document.Fields) {
        try { $field.Update() | Out-Null } catch {}
    }
    foreach ($toc in $document.TablesOfContents) {
        try { $toc.Update() | Out-Null } catch {}
    }
    $document.Repaginate()
    $document.Save()
} finally {
    if ($null -ne $document) {
        $document.Close($true)
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
    }
    if ($null -ne $word) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    }
}

Write-Output $DestinationPath
