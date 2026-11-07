$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression

$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$ReportsDir = Join-Path $Root "reports"
$OutPath = Join-Path $ReportsDir "adsb_tracker_work_report.docx"
New-Item -ItemType Directory -Path $ReportsDir -Force | Out-Null

function Escape-Xml([string]$Text) {
    if ($null -eq $Text) { return "" }
    return [System.Security.SecurityElement]::Escape($Text)
}

function Para([string]$Text, [string]$Style = $null) {
    $styleXml = ""
    if ($Style) { $styleXml = "<w:pPr><w:pStyle w:val=`"$Style`"/></w:pPr>" }
    return "<w:p>$styleXml<w:r><w:t xml:space=`"preserve`">$(Escape-Xml $Text)</w:t></w:r></w:p>"
}

function Bullet([string]$Text) {
    return "<w:p><w:pPr><w:pStyle w:val=`"ListBullet`"/></w:pPr><w:r><w:t xml:space=`"preserve`">$(Escape-Xml "- $Text")</w:t></w:r></w:p>"
}

function Table($Rows) {
    $xml = "<w:tbl><w:tblPr><w:tblStyle w:val=`"TableGrid`"/><w:tblW w:w=`"0`" w:type=`"auto`"/><w:tblLook w:val=`"04A0`"/></w:tblPr>"
    foreach ($row in $Rows) {
        $xml += "<w:tr>"
        foreach ($cell in $row) {
            $xml += "<w:tc><w:tcPr><w:tcW w:w=`"2400`" w:type=`"dxa`"/></w:tcPr>$(Para ([string]$cell))</w:tc>"
        }
        $xml += "</w:tr>"
    }
    $xml += "</w:tbl>"
    return $xml
}

function Count-Lines([string]$RelativePath) {
    $path = Join-Path $Root $RelativePath
    if (!(Test-Path $path)) { return 0 }
    return (Get-Content -Path $path | Measure-Object -Line).Lines
}

function Csv-DataRows([string]$RelativePath) {
    $path = Join-Path $Root $RelativePath
    if (!(Test-Path $path)) { return 0 }
    $lines = (Get-Content -Path $path | Measure-Object -Line).Lines
    return [Math]::Max(0, $lines - 1)
}

$generatedOn = (Get-Date).ToString("dd MMM yyyy, HH:mm")
$projectFiles = @(
    @("app.py", "Flask web app, API routes, preprocessing, database ingestion, CSV export triggers", (Count-Lines "app.py")),
    @("data_source.py", "Simulation replay and optional live Pingstation stream source", (Count-Lines "data_source.py")),
    @("preprocessing/pingstation_raw_decoder.py", "Raw Pingstation frame splitting, decoding, state snapshots, capture logging", (Count-Lines "preprocessing\pingstation_raw_decoder.py")),
    @("preprocessing/pingstation_preprocess.py", "Normalization of altitude, heading, speed, timestamp, aircraft identifiers", (Count-Lines "preprocessing\pingstation_preprocess.py")),
    @("storage/flight_data_store.py", "SQLite trajectory store, latest matrix, retention, CSV exports", (Count-Lines "storage\flight_data_store.py")),
    @("templates/index.html", "Leaflet dashboard, live aircraft list, selection panels, altitude chart", (Count-Lines "templates\index.html"))
)

$artifactRows = @(
    @("Artifact", "Purpose", "Current Evidence"),
    @("logs/flight_matrix_latest.csv", "Latest matrix of tracked aircraft", "$(Csv-DataRows "logs\flight_matrix_latest.csv") aircraft rows"),
    @("logs/training_csv/trajectory_20260506.csv", "Append-only trajectory training CSV", "$(Csv-DataRows "logs\training_csv\trajectory_20260506.csv") sample rows"),
    @("logs/exports/trajectory_ALL_72h_20260506T164125Z.csv", "Manual/export API trajectory extract", "$(Csv-DataRows "logs\exports\trajectory_ALL_72h_20260506T164125Z.csv") exported rows"),
    @("logs/flight_training.db", "SQLite database for latest matrix and trajectory samples", "Created and updated during tracker runs")
)

$moduleRows = @(@("File", "Work Completed", "Lines"))
foreach ($row in $projectFiles) { $moduleRows += ,$row }

$docBody = ""
$docBody += Para "ADS-B Live Tracker: Work Completed Report" "Title"
$docBody += Para "Generated on $generatedOn from project folder: $Root" "Subtitle"
$docBody += Para "Executive Summary" "Heading1"
$docBody += Para "The project has progressed into a working ADS-B/Pingstation tracking application. It includes a Flask backend, a browser-based live map dashboard, a Pingstation raw data decoding pipeline, preprocessing for normalized aircraft fields, and persistent storage for machine-learning-ready trajectory data."
$docBody += Para "The current build can replay captured Pingstation logs for simulation, optionally listen to a live Pingstation TCP stream, display aircraft positions on a Leaflet map, and continuously store latest aircraft state plus historical trajectory samples in SQLite and CSV formats."
$docBody += Para "Major Work Completed" "Heading1"
$docBody += Bullet "Built the Flask application entry point with API routes for live aircraft, latest training matrix, trajectory history, CSV information, and trajectory export."
$docBody += Bullet "Implemented Pingstation data sourcing with simulation replay from captured logs and optional live stream receiver mode."
$docBody += Bullet "Created a raw Pingstation decoder that splits length-prefixed frames, decodes compressed VRS-style fields, extracts fallback callsigns/positions, and maintains aircraft state."
$docBody += Bullet "Added preprocessing to normalize altitude, heading, speed, timestamps, callsign, and ICAO fields into consistent units."
$docBody += Bullet "Added a SQLite-backed training data store with latest-aircraft matrix rows and time-series trajectory samples."
$docBody += Bullet "Added CSV output paths for latest matrix snapshots, append-only daily trajectory data, and on-demand trajectory export."
$docBody += Bullet "Built the web dashboard with a dark radar-style Leaflet map, aircraft markers, rotated airplane icons, route lines, search/filtering, selected-aircraft panels, and altitude charting."
$docBody += Para "System Architecture" "Heading1"
$docBody += Para "The application flow is: Pingstation source or replay log -> raw decoder -> preprocessing -> Flask API response -> frontend map rendering. In parallel, processed records are ingested into the training data store, which updates both SQLite tables and CSV outputs."
$docBody += Bullet "Data input: captured Pingstation log replay by default, controlled through SIMULATION_MODE and replay-related environment variables."
$docBody += Bullet "Live input option: TCP receiver on configurable host/port, with optional raw capture logging."
$docBody += Bullet "Processing layer: raw binary decoding, state merging, stale snapshot filtering, and data normalization."
$docBody += Bullet "Storage layer: SQLite WAL database plus matrix and trajectory CSV files."
$docBody += Bullet "User interface: live map, aircraft list, selected flight details, route tab, and altitude trend chart."
$docBody += Para "Implemented Modules" "Heading1"
$docBody += Table $moduleRows
$docBody += Para "Generated Data And Outputs" "Heading1"
$docBody += Table $artifactRows
$docBody += Para "Backend API Endpoints" "Heading1"
$docBody += Table @(
    @("Endpoint", "Purpose"),
    @("/", "Renders the live tracker dashboard."),
    @("/api/aircraft", "Returns processed aircraft rows and ingests them into storage."),
    @("/api/training/matrix", "Returns latest per-flight matrix rows for training/inspection."),
    @("/api/training/matrix/csv-info", "Reports latest matrix CSV configuration."),
    @("/api/training/trajectory", "Returns historical trajectory samples, optionally filtered by flight key and hours."),
    @("/api/training/trajectory/export", "Writes trajectory history to a timestamped CSV export.")
)
$docBody += Para "Current Configuration Highlights" "Heading1"
$docBody += Bullet "Default server: localhost on port 5005."
$docBody += Bullet "Default database path: logs/flight_training.db."
$docBody += Bullet "Default latest matrix CSV path: logs/flight_matrix_latest.csv."
$docBody += Bullet "Default daily trajectory CSV directory: logs/training_csv."
$docBody += Bullet "Default simulation source: pingstation_data_log/pingstation_data_log/final.txt."
$docBody += Para "Observations" "Heading1"
$docBody += Bullet "The codebase is not currently initialized as a Git repository, so this report is based on files and artifacts present in the workspace rather than commit history."
$docBody += Bullet "The virtual environment exists and dependencies are listed in requirements.txt, but launching the venv Python from this path currently fails on this machine because the workspace path contains an apostrophe."
$docBody += Bullet "Generated logs show that the pipeline has produced matrix, trajectory, and export CSV files, which is useful evidence that storage/export work has been exercised."
$docBody += Para "Suggested Next Steps" "Heading1"
$docBody += Bullet "Initialize Git and commit the current stable version so future work can be tracked cleanly."
$docBody += Bullet "Move or recreate the virtual environment under a path without an apostrophe, or rebuild it after renaming the folder path."
$docBody += Bullet "Add automated tests for decoder parsing, preprocessing unit conversion, and storage export behavior."
$docBody += Bullet "Add a small README with setup instructions, environment variables, API endpoint descriptions, and sample run commands."
$docBody += Bullet "Consider adding screenshots of the live dashboard to future versions of this report."

$documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    $docBody
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1008" w:right="1008" w:bottom="1008" w:left="1008" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>
"@

$stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="40"/></w:rPr><w:pPr><w:spacing w:after="220"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:color w:val="666666"/><w:sz w:val="20"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="280" w:after="120"/></w:pPr><w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="28"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr></w:style>
  <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:qFormat/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders><w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>
"@

$contentTypesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"@

$relsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"@

$documentRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"@

$createdIso = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$coreXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>ADS-B Live Tracker Work Completed Report</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">$createdIso</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$createdIso</dcterms:modified>
</cp:coreProperties>
"@

$appXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>
"@

$entries = [ordered]@{
    "[Content_Types].xml" = $contentTypesXml
    "_rels/.rels" = $relsXml
    "docProps/core.xml" = $coreXml
    "docProps/app.xml" = $appXml
    "word/document.xml" = $documentXml
    "word/_rels/document.xml.rels" = $documentRelsXml
    "word/styles.xml" = $stylesXml
}

if (Test-Path $OutPath) {
    Remove-Item -Path $OutPath -Force
}

$fileStream = [System.IO.File]::Open($OutPath, [System.IO.FileMode]::CreateNew)
try {
    $zip = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $entries.Keys) {
            $entry = $zip.CreateEntry($name)
            $stream = $entry.Open()
            try {
                $writer = New-Object System.IO.StreamWriter($stream, [System.Text.UTF8Encoding]::new($false))
                try {
                    $writer.Write($entries[$name])
                } finally {
                    $writer.Dispose()
                }
            } finally {
                $stream.Dispose()
            }
        }
    } finally {
        $zip.Dispose()
    }
} finally {
    $fileStream.Dispose()
}

Write-Output $OutPath
