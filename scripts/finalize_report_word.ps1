param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,
    [Parameter(Mandatory = $true)]
    [string]$DestinationPath
)

$ErrorActionPreference = "Stop"

function Convert-ToRoman([int]$Number) {
    $map = @(
        @(1000, "m"), @(900, "cm"), @(500, "d"), @(400, "cd"),
        @(100, "c"), @(90, "xc"), @(50, "l"), @(40, "xl"),
        @(10, "x"), @(9, "ix"), @(5, "v"), @(4, "iv"), @(1, "i")
    )
    $result = ""
    foreach ($entry in $map) {
        while ($Number -ge $entry[0]) {
            $result += $entry[1]
            $Number -= $entry[0]
        }
    }
    return $result
}

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

    $headingPages = @{}
    foreach ($paragraph in $document.Paragraphs) {
        $text = $paragraph.Range.Text.Trim([char]13, [char]7, " ")
        if ($text -in @(
            "LIST OF FIGURES",
            "LIST OF TABLES",
            "ABBREVIATIONS",
            "CERTIFICATE",
            "ACKNOWLEDGEMENTS",
            "ABSTRACT",
            "CHAPTER 1: INTRODUCTION",
            "CHAPTER 2: LITERATURE REVIEW",
            "CHAPTER 3: WORK CARRIED OUT",
            "CHAPTER 4: EXPERIMENTAL RESULT AND COMPARISION",
            "CHAPTER 5: CONCLUSION, SUMMARY & SCOPE",
            "REFERENCES",
            "BIBLIOGRAPHY",
            "APPENDIX"
        )) {
            $headingPages[$text] = [int]$paragraph.Range.Information(1)
        }
    }

    $tocLabels = @{
        "List of Figures" = "LIST OF FIGURES"
        "List of Tables" = "LIST OF TABLES"
        "Abbreviations" = "ABBREVIATIONS"
        "Certificate" = "CERTIFICATE"
        "Acknowledgements" = "ACKNOWLEDGEMENTS"
        "Abstract" = "ABSTRACT"
        "CHAPTER 1: INTRODUCTION" = "CHAPTER 1: INTRODUCTION"
        "CHAPTER 2: LITERATURE REVIEW" = "CHAPTER 2: LITERATURE REVIEW"
        "CHAPTER 3: WORK CARRIED OUT" = "CHAPTER 3: WORK CARRIED OUT"
        "CHAPTER 4: EXPERIMENTAL RESULT AND COMPARISION" = "CHAPTER 4: EXPERIMENTAL RESULT AND COMPARISION"
        "CHAPTER 5: CONCLUSION, SUMMARY & SCOPE" = "CHAPTER 5: CONCLUSION, SUMMARY & SCOPE"
        "References" = "REFERENCES"
        "Bibliography" = "BIBLIOGRAPHY"
        "Appendix" = "APPENDIX"
    }

    $toc = $document.Tables.Item(1)
    for ($rowIndex = 2; $rowIndex -le $toc.Rows.Count; $rowIndex++) {
        $label = $toc.Cell($rowIndex, 1).Range.Text.Trim([char]13, [char]7, " ")
        if (!$tocLabels.ContainsKey($label)) {
            continue
        }
        $heading = $tocLabels[$label]
        if (!$headingPages.ContainsKey($heading)) {
            continue
        }
        $page = $headingPages[$heading]
        $displayPage = if ($heading -in @(
            "LIST OF FIGURES",
            "LIST OF TABLES",
            "ABBREVIATIONS",
            "CERTIFICATE",
            "ACKNOWLEDGEMENTS",
            "ABSTRACT"
        )) {
            Convert-ToRoman $page
        } else {
            [string]$page
        }

        $cellRange = $toc.Cell($rowIndex, 2).Range
        $cellRange.End = $cellRange.End - 1
        $cellRange.Text = $displayPage
    }

    $document.Repaginate()
    $document.Save()
} finally {
    if ($null -ne $document) {
        $document.Close($true)
    }
    if ($null -ne $word) {
        $word.Quit()
    }
    if ($null -ne $document) {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
    }
    if ($null -ne $word) {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    }
}

Write-Output $DestinationPath
