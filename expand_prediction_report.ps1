$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression.FileSystem

$sourceDir = Join-Path $PSScriptRoot 'report_updated_unpacked'
$workDir = Join-Path $PSScriptRoot 'report_prediction_expanded'
$outputDocx = Join-Path $PSScriptRoot 'Aircraft_Tracker_Project_Report_Updated_With_Algorithm_Comparison.docx'

if (Test-Path -LiteralPath $workDir) {
    Remove-Item -LiteralPath $workDir -Recurse -Force
}
Copy-Item -LiteralPath $sourceDir -Destination $workDir -Recurse

$documentPath = Join-Path $workDir 'word\document.xml'
$xml = [xml](Get-Content -LiteralPath $documentPath -Raw)
$wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$ns = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
$ns.AddNamespace('w', $wNs)

function New-WElement([string]$name) {
    return $xml.CreateElement('w', $name, $wNs)
}

function Set-WAttribute($node, [string]$name, [string]$value) {
    [void]$node.SetAttribute($name, $wNs, $value)
}

function Get-ParagraphText($paragraph) {
    return (($paragraph.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '')
}

function New-Paragraph(
    [string]$text,
    [string]$style = '',
    [bool]$firstLine = $false,
    [bool]$center = $false,
    [bool]$bold = $false,
    [bool]$italic = $false
) {
    $p = New-WElement 'p'
    if ($style -or $firstLine -or $center) {
        $pPr = New-WElement 'pPr'
        if ($style) {
            $pStyle = New-WElement 'pStyle'
            Set-WAttribute $pStyle 'val' $style
            [void]$pPr.AppendChild($pStyle)
        }
        if ($firstLine) {
            $ind = New-WElement 'ind'
            Set-WAttribute $ind 'firstLine' '567'
            [void]$pPr.AppendChild($ind)
        }
        if ($center) {
            $jc = New-WElement 'jc'
            Set-WAttribute $jc 'val' 'center'
            [void]$pPr.AppendChild($jc)
        }
        [void]$p.AppendChild($pPr)
    }

    $r = New-WElement 'r'
    if ($bold -or $italic) {
        $rPr = New-WElement 'rPr'
        if ($bold) { [void]$rPr.AppendChild((New-WElement 'b')) }
        if ($italic) { [void]$rPr.AppendChild((New-WElement 'i')) }
        [void]$r.AppendChild($rPr)
    }
    $t = New-WElement 't'
    [void]$t.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
    $t.InnerText = $text
    [void]$r.AppendChild($t)
    [void]$p.AppendChild($r)
    return $p
}

function New-TableCell([string]$text, [int]$width, [bool]$header = $false) {
    $tc = New-WElement 'tc'
    $tcPr = New-WElement 'tcPr'
    $tcW = New-WElement 'tcW'
    Set-WAttribute $tcW 'w' ([string]$width)
    Set-WAttribute $tcW 'type' 'dxa'
    [void]$tcPr.AppendChild($tcW)
    $vAlign = New-WElement 'vAlign'
    Set-WAttribute $vAlign 'val' 'center'
    [void]$tcPr.AppendChild($vAlign)
    if ($header) {
        $shd = New-WElement 'shd'
        Set-WAttribute $shd 'val' 'clear'
        Set-WAttribute $shd 'color' 'auto'
        Set-WAttribute $shd 'fill' 'D9E2F3'
        [void]$tcPr.AppendChild($shd)
    }
    [void]$tc.AppendChild($tcPr)

    $p = New-WElement 'p'
    $pPr = New-WElement 'pPr'
    $spacing = New-WElement 'spacing'
    Set-WAttribute $spacing 'before' '0'
    Set-WAttribute $spacing 'after' '0'
    Set-WAttribute $spacing 'line' '200'
    Set-WAttribute $spacing 'lineRule' 'auto'
    [void]$pPr.AppendChild($spacing)
    [void]$p.AppendChild($pPr)

    $r = New-WElement 'r'
    $rPr = New-WElement 'rPr'
    if ($header) { [void]$rPr.AppendChild((New-WElement 'b')) }
    $sz = New-WElement 'sz'
    Set-WAttribute $sz 'val' $(if ($header) { '17' } else { '16' })
    [void]$rPr.AppendChild($sz)
    $szCs = New-WElement 'szCs'
    Set-WAttribute $szCs 'val' $(if ($header) { '17' } else { '16' })
    [void]$rPr.AppendChild($szCs)
    [void]$r.AppendChild($rPr)
    $t = New-WElement 't'
    [void]$t.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
    $t.InnerText = $text
    [void]$r.AppendChild($t)
    [void]$p.AppendChild($r)
    [void]$tc.AppendChild($p)
    return $tc
}

function New-ComparisonTable {
    $widths = @(1150, 1450, 1900, 1900, 1906)
    $rows = @(
        @('Algorithm', 'Basic method', 'Main advantages', 'Main limitations', 'Suitability for this system'),
        @('Constant-velocity kinematics', 'Projects the latest speed, heading and vertical rate forward in time.', 'No training data; extremely fast; transparent; works with one recent state.', 'Very sensitive to the final sample and cannot anticipate turns or acceleration.', 'Excellent fallback and baseline; weaker as the primary estimator.'),
        @('Polynomial regression', 'Fits curved polynomial functions of time to recent coordinates.', 'Can represent gradual curvature and acceleration with little infrastructure.', 'Higher-order fits can oscillate and extrapolate unrealistically outside the sample window.', 'Possible extension, but less stable than linear regression for sparse noisy tracks.'),
        @('Sliding-window OLS linear regression', 'Fits independent straight-line trends to recent latitude, longitude and altitude samples.', 'Fast, interpretable, smooths sample noise, uses several observations and needs no offline training.', 'Assumes approximately constant local motion and loses accuracy during sharp manoeuvres.', 'Selected primary method for short reception gaps.'),
        @('Alpha-beta filter', 'Recursively updates position and velocity using fixed correction gains.', 'Very light, smooth and suitable for online tracking.', 'Gain selection is manual and uncertainty is not represented as fully as in a Kalman filter.', 'Strong lightweight alternative for future comparison.'),
        @('Kalman filter', 'Combines a motion model and measurements using estimated process and measurement uncertainty.', 'Noise-aware, recursive, efficient and well established in tracking applications.', 'Requires a state-space model and careful tuning of covariance matrices.', 'Best conventional next-stage candidate when sensor-noise statistics are available.'),
        @('Extended/Unscented Kalman filter', 'Extends Kalman filtering to nonlinear aircraft motion and observation models.', 'Handles nonlinear turns and geodetic motion more naturally.', 'Greater implementation cost and tuning complexity; can still fail with a poor model.', 'Suitable for manoeuvring targets after a validated dynamic model is defined.'),
        @('Random forest / XGBoost', 'Learns nonlinear corrections from engineered trajectory, route and weather features.', 'Models complex relationships and offers useful feature-importance analysis.', 'Requires a representative labelled dataset and does not naturally model sequence order.', 'Useful for offline correction or ETA tasks, not the simplest live gap bridge.'),
        @('LSTM / GRU network', 'Learns temporal dependencies from sequences of historical aircraft states.', 'Can learn turns, climb/descent patterns and route-specific behaviour.', 'Needs large clean datasets, training hardware and careful validation; less interpretable.', 'Promising research model once sufficient per-route training data is collected.'),
        @('Transformer model', 'Uses attention to model long-range relationships across trajectory sequences and context.', 'Powerful for long sequences and multi-source contextual inputs.', 'Highest data and computation demand; excessive for a short local outage in the current system.', 'Future long-horizon research rather than the present real-time baseline.'),
        @('Hybrid Kalman-ML model', 'Uses a physical filter for the base trajectory and a learned model for residual corrections.', 'Combines physical consistency, uncertainty handling and learned manoeuvre patterns.', 'Most complex architecture; needs both model tuning and substantial training data.', 'Potential long-term research direction after strong baselines are established.')
    )

    $tbl = New-WElement 'tbl'
    $tblPr = New-WElement 'tblPr'
    $tblStyle = New-WElement 'tblStyle'
    Set-WAttribute $tblStyle 'val' 'TableGrid'
    [void]$tblPr.AppendChild($tblStyle)
    $tblW = New-WElement 'tblW'
    Set-WAttribute $tblW 'w' '8306'
    Set-WAttribute $tblW 'type' 'dxa'
    [void]$tblPr.AppendChild($tblW)
    $tblLayout = New-WElement 'tblLayout'
    Set-WAttribute $tblLayout 'type' 'fixed'
    [void]$tblPr.AppendChild($tblLayout)
    $margins = New-WElement 'tblCellMar'
    foreach ($side in @('top', 'left', 'bottom', 'right')) {
        $margin = New-WElement $side
        Set-WAttribute $margin 'w' $(if ($side -in @('left', 'right')) { '80' } else { '45' })
        Set-WAttribute $margin 'type' 'dxa'
        [void]$margins.AppendChild($margin)
    }
    [void]$tblPr.AppendChild($margins)
    [void]$tbl.AppendChild($tblPr)

    $grid = New-WElement 'tblGrid'
    foreach ($width in $widths) {
        $gridCol = New-WElement 'gridCol'
        Set-WAttribute $gridCol 'w' ([string]$width)
        [void]$grid.AppendChild($gridCol)
    }
    [void]$tbl.AppendChild($grid)

    for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
        $tr = New-WElement 'tr'
        if ($rowIndex -eq 0) {
            $trPr = New-WElement 'trPr'
            $tblHeader = New-WElement 'tblHeader'
            Set-WAttribute $tblHeader 'val' 'true'
            [void]$trPr.AppendChild($tblHeader)
            [void]$tr.AppendChild($trPr)
        }
        for ($columnIndex = 0; $columnIndex -lt $widths.Count; $columnIndex++) {
            [void]$tr.AppendChild((New-TableCell $rows[$rowIndex][$columnIndex] $widths[$columnIndex] ($rowIndex -eq 0)))
        }
        [void]$tbl.AppendChild($tr)
    }
    return $tbl
}

function Set-CellText($cell, [string]$text) {
    $textNodes = $cell.SelectNodes('.//w:t', $ns)
    if ($textNodes.Count -eq 0) { return }
    $textNodes[0].InnerText = $text
    for ($i = 1; $i -lt $textNodes.Count; $i++) { $textNodes[$i].InnerText = '' }
}

$body = $xml.SelectSingleNode('//w:body', $ns)
$bodyChildren = @($body.ChildNodes)
$sectionStart = $null
$sectionEnd = $null
foreach ($node in $bodyChildren) {
    if ($node.LocalName -ne 'p') { continue }
    $text = Get-ParagraphText $node
    if ($text -eq '3.6  Trajectory-Prediction Algorithm') { $sectionStart = $node }
    if ($text -eq '3.7  Aircraft Classification Logic') { $sectionEnd = $node; break }
}
if (-not $sectionStart -or -not $sectionEnd) {
    throw 'Could not locate Sections 3.6 and 3.7 in the report.'
}

$cursor = $sectionStart.NextSibling
while ($cursor -and $cursor -ne $sectionEnd) {
    $next = $cursor.NextSibling
    [void]$body.RemoveChild($cursor)
    $cursor = $next
}

$content = [System.Collections.Generic.List[System.Xml.XmlNode]]::new()
$content.Add((New-Paragraph '3.6.1  Prediction Objective and Activation' 'Heading3'))
$content.Add((New-Paragraph 'The prediction subsystem is designed to bridge a short interruption in ADS-B reception, not to generate a complete strategic flight plan. An aircraft is eligible for prediction only after it has previously supplied genuine observations during the current session and its ICAO address is no longer present in the active receiver snapshot. The predictor therefore never replaces valid live data. When reception resumes, the newly received state immediately becomes authoritative and the predicted display record is removed. The prediction horizon is limited by FLIGHT_PREDICTION_MAX_AGE_SEC, which is 180 seconds by default, because uncertainty grows rapidly when an aircraft turns, changes speed, climbs or descends.' $false))

$content.Add((New-Paragraph '3.6.2  Sliding-Window Ordinary Least-Squares Model' 'Heading3'))
$content.Add((New-Paragraph 'For each temporarily missing aircraft, FlightTrajectoryStore retrieves up to the twenty most recent genuine trajectory samples. At least four valid timestamped samples are required. The timestamps are converted to elapsed time values, and three independent straight-line models are fitted: latitude as a function of time, longitude as a function of time and altitude as a function of time. Fitting a fresh local model for each aircraft and each outage makes the method an online statistical-learning baseline rather than a globally trained neural model.' $false))
$content.Add((New-Paragraph 'b = SUM[(t_i - mean(t))(y_i - mean(y))] / SUM[(t_i - mean(t))^2]' '' $false $true $false $true))
$content.Add((New-Paragraph 'a = mean(y) - b mean(t),     predicted y(t_future) = a + b t_future' '' $false $true $false $true))
$content.Add((New-Paragraph 'In these equations, t_i is the elapsed time of observation i, y_i is the latitude, longitude or altitude value, b is the least-squares slope and a is the intercept. The future latitude, longitude and altitude are evaluated at the current prediction time. The fitted latitude and longitude slopes are converted to horizontal velocity and bearing, while the altitude slope provides vertical velocity. Geographic bounds and finite-number checks are applied before a predicted state is returned.' $true))

$content.Add((New-Paragraph '3.6.3  Why the Regression Is More Stable Than Two-Point Extrapolation' 'Heading3'))
$content.Add((New-Paragraph 'A two-point method calculates motion from only the final pair of observations. Any coordinate quantisation, timestamp jitter or single erroneous measurement directly changes its estimated speed and direction. Ordinary least squares instead chooses the line that minimises the sum of squared residuals across several recent observations. Positive and negative measurement errors tend to offset one another, and one noisy point has only partial influence on the fitted slope. The sliding window also prevents very old flight phases from dominating the current estimate. Consequently, the displayed continuation is usually smoother and less sensitive to one abnormal ADS-B update, while still adapting as new samples enter the window.' $false))

$content.Add((New-Paragraph '3.6.4  Kinematic Fallback, Output and Recovery' 'Heading3'))
$content.Add((New-Paragraph 'If fewer than four suitable samples are available, the timestamps have no usable spread or the regression produces a non-finite result, the system falls back to constant-velocity kinematics. Missing motion values are estimated from the latest two stored positions where possible. Distance is calculated as speed multiplied by outage duration, latitude and longitude are projected along the latest heading, and altitude is updated from the vertical rate. Every generated record is marked isPredicted and includes predictionAgeSec and the source label PREDICTED. The interface draws its continuation as a red dotted line. Predicted rows are deliberately excluded from trajectory_samples and the Excel training records so that later models are trained only on genuine receiver observations.' $false))

$content.Add((New-Paragraph 'Fig. 3.6: Regression-based short-outage prediction. Recent latitude, longitude and altitude samples define least-squares trends; constant-velocity kinematics provide a fallback when a reliable fit cannot be produced.' 'FigCaption'))

$content.Add((New-Paragraph '3.6.5  Comparison of Suitable Prediction Algorithms' 'Heading3'))
$content.Add((New-Paragraph 'Aircraft-trajectory prediction can be approached using deterministic motion equations, recursive state estimators, conventional machine-learning regressors and deep sequence models. Table 3.3 compares the principal algorithms suitable for this project. Suitability is assessed against the present requirements: online operation, short data outages, limited training data, low latency, modest memory consumption and the need to explain the method clearly during academic evaluation.' $false))
$content.Add((New-Paragraph 'Table 3.3: Comparison of suitable aircraft-trajectory prediction algorithms' 'TabCaption'))
$content.Add((New-ComparisonTable))

$content.Add((New-Paragraph '3.6.6  Justification for Selecting Sliding-Window Linear Regression' 'Heading3'))
$content.Add((New-Paragraph 'Sliding-window OLS regression was selected as the primary method because it provides the best balance for the current implementation rather than because it is universally the most accurate trajectory algorithm. It works immediately with each aircraft without a pre-collected labelled dataset, executes in negligible time for a twenty-sample window, has very small memory requirements, and exposes coefficients whose physical meaning can be explained and audited. Compared with constant-velocity extrapolation, it uses several observations and therefore suppresses isolated noise. Compared with Kalman-family filters, it avoids covariance-model selection and tuning at this stage. Compared with tree ensembles, LSTM/GRU networks and Transformers, it does not require extensive historical trajectories, route labels, weather inputs or offline training hardware.' $false))
$content.Add((New-Paragraph 'The choice is therefore scientifically appropriate for short reception gaps in a single-station educational tracker. It also establishes a reproducible baseline against which more advanced models can be measured. A future research study should compare constant velocity, OLS regression, Kalman filtering and at least one trained sequence model on held-out trajectories using Haversine position error, altitude mean absolute error, heading error, root-mean-square error, inference latency and memory consumption. Results should be reported separately for straight cruise, turns, climbs, descents and different prediction horizons.' $true))

foreach ($node in $content) {
    [void]$body.InsertBefore($node, $sectionEnd)
}

# Insert the new table into the List of Tables and renumber the later Chapter 3 table.
$listOfTables = $xml.SelectNodes('//w:body/w:tbl', $ns)[1]
$rows = @($listOfTables.SelectNodes('./w:tr', $ns))
$existingTable33Row = $null
foreach ($row in $rows) {
    $cells = @($row.SelectNodes('./w:tc', $ns))
    if ($cells.Count -ge 2 -and (Get-ParagraphText $cells[0]) -eq 'Table 3.3') {
        $existingTable33Row = $row
        break
    }
}
if ($existingTable33Row) {
    $newRow = $existingTable33Row.CloneNode($true)
    $newCells = @($newRow.SelectNodes('./w:tc', $ns))
    Set-CellText $newCells[0] 'Table 3.3'
    Set-CellText $newCells[1] 'Comparison of suitable aircraft-trajectory prediction algorithms'
    [void]$listOfTables.InsertBefore($newRow, $existingTable33Row)

    $oldCells = @($existingTable33Row.SelectNodes('./w:tc', $ns))
    Set-CellText $oldCells[0] 'Table 3.4'
}

foreach ($p in $xml.SelectNodes('//w:body/w:p', $ns)) {
    $text = Get-ParagraphText $p
    if ($text -eq 'Table 3.3 summarises the principal hardware and software tools and technologies used in the implementation of the system, together with the role each played in the project.') {
        $textNode = $p.SelectSingleNode('.//w:t', $ns)
        $textNode.InnerText = $text.Replace('Table 3.3', 'Table 3.4')
    }
    elseif ($text -eq 'Table 3.3: Principal tools, technologies and their roles in the project') {
        $textNode = $p.SelectSingleNode('.//w:t', $ns)
        $textNode.InnerText = $text.Replace('Table 3.3', 'Table 3.4')
    }
}

$writerSettings = [System.Xml.XmlWriterSettings]::new()
$writerSettings.Encoding = [System.Text.UTF8Encoding]::new($false)
$writerSettings.Indent = $false
$writer = [System.Xml.XmlWriter]::Create($documentPath, $writerSettings)
$xml.Save($writer)
$writer.Close()

# Ask Word-compatible applications to refresh the TOC and page-number fields on open.
$settingsPath = Join-Path $workDir 'word\settings.xml'
$settingsXml = [xml](Get-Content -LiteralPath $settingsPath -Raw)
$settingsNs = [System.Xml.XmlNamespaceManager]::new($settingsXml.NameTable)
$settingsNs.AddNamespace('w', $wNs)
$settingsRoot = $settingsXml.SelectSingleNode('/w:settings', $settingsNs)
$updateFields = $settingsXml.SelectSingleNode('/w:settings/w:updateFields', $settingsNs)
if (-not $updateFields) {
    $updateFields = $settingsXml.CreateElement('w', 'updateFields', $wNs)
    [void]$settingsRoot.AppendChild($updateFields)
}
$updateFields.SetAttribute('val', $wNs, 'true')
$settingsWriter = [System.Xml.XmlWriter]::Create($settingsPath, $writerSettings)
$settingsXml.Save($settingsWriter)
$settingsWriter.Close()

if (Test-Path -LiteralPath $outputDocx) {
    Remove-Item -LiteralPath $outputDocx -Force
}
[System.IO.Compression.ZipFile]::CreateFromDirectory($workDir, $outputDocx, [System.IO.Compression.CompressionLevel]::Optimal, $false)

Write-Output $outputDocx
