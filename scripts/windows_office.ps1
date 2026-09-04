param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("CheckOffice", "UpdatePresentation", "ReplaceImage", "ReplaceExcelChart", "ExportPdf", "ExtractExcelRange")]
    [string]$Action,
    [string]$PresentationPath,
    [string]$OperationsPath,
    [string]$ImagePath,
    [string]$WorkbookPath,
    [string]$WorksheetName,
    [int]$ChartNumber = 0,
    [int]$SlideNumber = 0,
    [string]$TargetShapeName,
    [string]$PdfPath,
    [int]$StartRow = 0,
    [int]$StartColumn = 0,
    [int]$RowCount = 0,
    [int]$ColumnCount = 0,
    [string]$OutputMode = "display",
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$ReportFont = "Arial"

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Require-Value {
    param([string]$Name, [object]$Value)
    if ($null -eq $Value -or ([string]$Value).Trim().Length -eq 0) {
        throw "Missing required parameter -$Name for action $Action."
    }
}

function Require-ExistingFile {
    param([string]$Name, [string]$PathValue)
    Require-Value $Name $PathValue
    if (-not (Test-Path -LiteralPath $PathValue -PathType Leaf)) {
        throw "File not found for -$Name`: $PathValue"
    }
    return (Resolve-Path -LiteralPath $PathValue).Path
}

function Open-PowerPointPresentation {
    param([object]$PowerPoint, [string]$PathValue)
    return $PowerPoint.Presentations.Open($PathValue, $false, $false, $false)
}

function Set-TextRangeReportFont {
    param([object]$TextRange)
    if ($null -eq $TextRange) { return }
    try { $TextRange.Font.Name = $ReportFont } catch { }
    try { $TextRange.Font.NameFarEast = $ReportFont } catch { }
    try { $TextRange.Font.NameComplexScript = $ReportFont } catch { }
}

function Set-ShapeReportFont {
    param([object]$Shape)
    if ($null -eq $Shape) { return }
    try {
        if ($Shape.Type -eq 6) { # msoGroup
            for ($index = 1; $index -le $Shape.GroupItems.Count; $index++) {
                $child = $null
                try {
                    $child = $Shape.GroupItems.Item($index)
                    Set-ShapeReportFont $child
                }
                finally { Release-ComObject $child }
            }
        }
    } catch { }
    try {
        if ($Shape.HasTable -eq -1) {
            for ($row = 1; $row -le $Shape.Table.Rows.Count; $row++) {
                for ($column = 1; $column -le $Shape.Table.Columns.Count; $column++) {
                    $cell = $null
                    $cellShape = $null
                    $range = $null
                    try {
                        $cell = $Shape.Table.Cell($row, $column)
                        $cellShape = $cell.Shape
                        $range = $cellShape.TextFrame.TextRange
                        Set-TextRangeReportFont $range
                    }
                    finally {
                        Release-ComObject $range
                        Release-ComObject $cellShape
                        Release-ComObject $cell
                    }
                }
            }
        }
    } catch { }
    try {
        if ($Shape.HasTextFrame -eq -1 -and $Shape.TextFrame.HasText -eq -1) {
            $range = $null
            try {
                $range = $Shape.TextFrame.TextRange
                Set-TextRangeReportFont $range
            }
            finally { Release-ComObject $range }
        }
    } catch { }
}

function Set-PresentationReportFont {
    param([object]$Presentation)
    foreach ($slide in @($Presentation.Slides)) {
        try {
            foreach ($shape in @($slide.Shapes)) {
                try { Set-ShapeReportFont $shape }
                finally { Release-ComObject $shape }
            }
        }
        finally { Release-ComObject $slide }
    }
}

function Invoke-CheckOffice {
    $excel = $null
    $powerPoint = $null
    $excelInitialCount = 0
    $powerPointInitialCount = 0
    try {
        $excel = New-Object -ComObject Excel.Application
        $excelInitialCount = $excel.Workbooks.Count
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $powerPointInitialCount = $powerPoint.Presentations.Count
        [pscustomobject]@{
            ExcelVersion = [string]$excel.Version
            PowerPointVersion = [string]$powerPoint.Version
            Status = "OK"
        } | ConvertTo-Json -Compress
    }
    finally {
        if ($null -ne $powerPoint -and $powerPointInitialCount -eq 0) {
            try { $powerPoint.Quit() } catch { }
        }
        if ($null -ne $excel -and $excelInitialCount -eq 0) {
            try { $excel.Quit() } catch { }
        }
        Release-ComObject $powerPoint
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Invoke-UpdatePresentation {
    $resolvedPresentation = Require-ExistingFile "PresentationPath" $PresentationPath
    $resolvedOperations = Require-ExistingFile "OperationsPath" $OperationsPath
    $operationsJson = [System.IO.File]::ReadAllText($resolvedOperations, [System.Text.Encoding]::UTF8)
    $operations = ConvertFrom-Json -InputObject $operationsJson
    $powerPoint = $null
    $presentation = $null
    $initialCount = 0
    try {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $initialCount = $powerPoint.Presentations.Count
        $presentation = Open-PowerPointPresentation $powerPoint $resolvedPresentation
        $operationIndex = 0
        foreach ($operation in $operations) {
            $operationIndex++
            $slide = $null
            $shape = $null
            try {
                $slide = $presentation.Slides.Item([int]$operation.slide)
                $shape = $slide.Shapes.Item([string]$operation.shape)
                switch ([string]$operation.kind) {
                    "table" {
                        $matrix = @($operation.matrix)
                        for ($rowOffset = 0; $rowOffset -lt $matrix.Count; $rowOffset++) {
                            $rowValues = @($matrix[$rowOffset])
                            for ($columnOffset = 0; $columnOffset -lt $rowValues.Count; $columnOffset++) {
                                $targetRow = [int]$operation.start_row + $rowOffset
                                $targetColumn = [int]$operation.start_column + $columnOffset
                                $cell = $null
                                $cellShape = $null
                                $textRange = $null
                                try {
                                    $cell = $shape.Table.Cell($targetRow, $targetColumn)
                                    $cellShape = $cell.Shape
                                    $textRange = $cellShape.TextFrame.TextRange
                                    $textRange.Text = [string]$rowValues[$columnOffset]
                                    Set-TextRangeReportFont $textRange
                                    if ($operation.PSObject.Properties.Name -contains 'font_colors') {
                                        $colorRows = @($operation.font_colors)
                                        if ($rowOffset -lt $colorRows.Count) {
                                            $colorValues = @($colorRows[$rowOffset])
                                            if ($columnOffset -lt $colorValues.Count) {
                                                $hexColor = [string]$colorValues[$columnOffset]
                                                if ($hexColor -match '^[0-9A-Fa-f]{6}$') {
                                                    $red = [Convert]::ToInt32($hexColor.Substring(0, 2), 16)
                                                    $green = [Convert]::ToInt32($hexColor.Substring(2, 2), 16)
                                                    $blue = [Convert]::ToInt32($hexColor.Substring(4, 2), 16)
                                                    $textRange.Font.Color.RGB = $red + ($green * 256) + ($blue * 65536)
                                                }
                                            }
                                        }
                                    }
                                }
                                finally {
                                    Release-ComObject $textRange
                                    Release-ComObject $cellShape
                                    Release-ComObject $cell
                                }
                            }
                        }
                    }
                    "paragraph" {
                        $textRange = $null
                        $paragraphRange = $null
                        try {
                            $textRange = $shape.TextFrame.TextRange
                            $paragraphRange = $textRange.Paragraphs([int]$operation.paragraph, 1)
                            $originalText = [string]$paragraphRange.Text
                            $replacement = [string]$operation.text
                            if ($originalText.EndsWith("`r")) {
                                $replacement = $replacement + "`r"
                            }
                            $paragraphRange.Text = $replacement
                            Set-TextRangeReportFont $paragraphRange
                        }
                        finally {
                            Release-ComObject $paragraphRange
                            Release-ComObject $textRange
                        }
                    }
                    "text" {
                        $textRange = $null
                        try {
                            $textRange = $shape.TextFrame.TextRange
                            $textRange.Text = [string]$operation.text
                            Set-TextRangeReportFont $textRange
                        }
                        finally {
                            Release-ComObject $textRange
                        }
                    }
                    "paragraph_list" {
                        $textRange = $null
                        try {
                            $textRange = $shape.TextFrame.TextRange
                            $paragraphs = @($operation.paragraphs | ForEach-Object { [string]$_ })
                            $textRange.Text = [string]::Join("`r", $paragraphs)
                            Set-TextRangeReportFont $textRange
                        }
                        finally {
                            Release-ComObject $textRange
                        }
                    }
                    default { throw "Unsupported operation kind: $($operation.kind)" }
                }
            }
            catch {
                throw "PowerPoint operation $operationIndex failed (slide $($operation.slide), shape '$($operation.shape)', kind '$($operation.kind)'): $($_.Exception.Message)"
            }
            finally {
                Release-ComObject $shape
                Release-ComObject $slide
            }
        }
        Set-PresentationReportFont $presentation
        $presentation.Save()
    }
    finally {
        if ($null -ne $presentation) {
            try { $presentation.Close() } catch { }
        }
        if ($null -ne $powerPoint -and $initialCount -eq 0) {
            try { $powerPoint.Quit() } catch { }
        }
        Release-ComObject $presentation
        Release-ComObject $powerPoint
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Invoke-ReplaceImage {
    $resolvedPresentation = Require-ExistingFile "PresentationPath" $PresentationPath
    $resolvedImage = Require-ExistingFile "ImagePath" $ImagePath
    Require-Value "SlideNumber" $SlideNumber
    Require-Value "TargetShapeName" $TargetShapeName
    $powerPoint = $null
    $presentation = $null
    $slide = $null
    $target = $null
    $newShape = $null
    $initialCount = 0
    try {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $initialCount = $powerPoint.Presentations.Count
        $presentation = Open-PowerPointPresentation $powerPoint $resolvedPresentation
        $slide = $presentation.Slides.Item($SlideNumber)
        $target = $slide.Shapes.Item($TargetShapeName)
        $left = [single]$target.Left
        $top = [single]$target.Top
        $width = [single]$target.Width
        $height = [single]$target.Height
        $target.Delete()
        $newShape = $slide.Shapes.AddPicture($resolvedImage, 0, -1, $left, $top, $width, $height)
        $newShape.LockAspectRatio = 0
        $newShape.Left = $left
        $newShape.Top = $top
        $newShape.Width = $width
        $newShape.Height = $height
        $newShape.Name = $TargetShapeName
        $presentation.Save()
    }
    finally {
        if ($null -ne $presentation) { try { $presentation.Close() } catch { } }
        if ($null -ne $powerPoint -and $initialCount -eq 0) { try { $powerPoint.Quit() } catch { } }
        Release-ComObject $newShape
        Release-ComObject $target
        Release-ComObject $slide
        Release-ComObject $presentation
        Release-ComObject $powerPoint
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Invoke-ReplaceExcelChart {
    $resolvedWorkbook = Require-ExistingFile "WorkbookPath" $WorkbookPath
    $resolvedPresentation = Require-ExistingFile "PresentationPath" $PresentationPath
    Require-Value "WorksheetName" $WorksheetName
    Require-Value "ChartNumber" $ChartNumber
    Require-Value "SlideNumber" $SlideNumber
    Require-Value "TargetShapeName" $TargetShapeName
    $temporaryPng = Join-Path ([System.IO.Path]::GetTempPath()) ("weekly_report_chart_" + [Guid]::NewGuid().ToString("N") + ".png")
    $excel = $null
    $workbookObject = $null
    $worksheet = $null
    $chartObjects = $null
    $chartObject = $null
    $chart = $null
    $excelInitialCount = 0
    try {
        $excel = New-Object -ComObject Excel.Application
        $excelInitialCount = $excel.Workbooks.Count
        $excel.DisplayAlerts = $false
        $workbookObject = $excel.Workbooks.Open($resolvedWorkbook, 0, $true)
        $excel.CalculateFullRebuild()
        $worksheet = $workbookObject.Worksheets.Item($WorksheetName)
        $chartObjects = $worksheet.ChartObjects()
        $chartObject = $chartObjects.Item($ChartNumber)
        $chart = $chartObject.Chart
        try { $chart.ChartArea.Font.Name = $ReportFont } catch { }
        try { if ($chart.HasTitle) { $chart.ChartTitle.Font.Name = $ReportFont } } catch { }
        try { if ($chart.HasLegend) { $chart.Legend.Font.Name = $ReportFont } } catch { }
        foreach ($axisType in @(1, 2)) {
            $axis = $null
            try {
                $axis = $chart.Axes($axisType, 1)
                $axis.TickLabels.Font.Name = $ReportFont
                if ($axis.HasTitle) { $axis.AxisTitle.Font.Name = $ReportFont }
            }
            catch { }
            finally { Release-ComObject $axis }
        }
        $seriesCollection = $null
        try {
            $seriesCollection = $chart.SeriesCollection()
            for ($seriesIndex = 1; $seriesIndex -le $seriesCollection.Count; $seriesIndex++) {
                $series = $null
                try {
                    $series = $seriesCollection.Item($seriesIndex)
                    if ($series.HasDataLabels) { $series.DataLabels().Font.Name = $ReportFont }
                }
                catch { }
                finally { Release-ComObject $series }
            }
        }
        finally { Release-ComObject $seriesCollection }
        $exported = $chart.Export($temporaryPng, "PNG")
        if (-not $exported -or -not (Test-Path -LiteralPath $temporaryPng)) {
            throw "Excel could not export chart $ChartNumber from worksheet '$WorksheetName'."
        }
    }
    finally {
        if ($null -ne $workbookObject) { try { $workbookObject.Close($false) } catch { } }
        if ($null -ne $excel -and $excelInitialCount -eq 0) { try { $excel.Quit() } catch { } }
        Release-ComObject $chart
        Release-ComObject $chartObject
        Release-ComObject $chartObjects
        Release-ComObject $worksheet
        Release-ComObject $workbookObject
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    try {
        $script:ImagePath = $temporaryPng
        Invoke-ReplaceImage
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPng) {
            Remove-Item -LiteralPath $temporaryPng -Force
        }
    }
}

function Invoke-ExportPdf {
    $resolvedPresentation = Require-ExistingFile "PresentationPath" $PresentationPath
    Require-Value "PdfPath" $PdfPath
    $resolvedPdf = [System.IO.Path]::GetFullPath($PdfPath)
    $pdfDirectory = Split-Path -Parent $resolvedPdf
    if (-not (Test-Path -LiteralPath $pdfDirectory)) {
        New-Item -ItemType Directory -Path $pdfDirectory -Force | Out-Null
    }
    $powerPoint = $null
    $presentation = $null
    $initialCount = 0
    try {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $initialCount = $powerPoint.Presentations.Count
        $presentation = Open-PowerPointPresentation $powerPoint $resolvedPresentation
        Set-PresentationReportFont $presentation
        $presentation.Save()
        # ppSaveAsPDF = 32. SaveAs is more reliable than ExportAsFixedFormat
        # through Windows PowerShell's late-bound COM adapter.
        $presentation.SaveAs($resolvedPdf, 32)
    }
    finally {
        if ($null -ne $presentation) { try { $presentation.Close() } catch { } }
        if ($null -ne $powerPoint -and $initialCount -eq 0) { try { $powerPoint.Quit() } catch { } }
        Release-ComObject $presentation
        Release-ComObject $powerPoint
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Invoke-ExtractExcelRange {
    $resolvedWorkbook = Require-ExistingFile "WorkbookPath" $WorkbookPath
    Require-Value "WorksheetName" $WorksheetName
    Require-Value "OutputPath" $OutputPath
    if ($StartRow -lt 1 -or $StartColumn -lt 1 -or $RowCount -lt 1 -or $ColumnCount -lt 1) {
        throw "StartRow, StartColumn, RowCount and ColumnCount must be positive integers."
    }
    $excel = $null
    $workbookObject = $null
    $worksheet = $null
    $initialCount = 0
    try {
        $excel = New-Object -ComObject Excel.Application
        $initialCount = $excel.Workbooks.Count
        $excel.DisplayAlerts = $false
        $workbookObject = $excel.Workbooks.Open($resolvedWorkbook, 0, $true)
        $excel.CalculateFullRebuild()
        $worksheet = $workbookObject.Worksheets.Item($WorksheetName)
        $matrix = New-Object System.Collections.ArrayList
        for ($rowOffset = 0; $rowOffset -lt $RowCount; $rowOffset++) {
            $rowValues = New-Object System.Collections.ArrayList
            for ($columnOffset = 0; $columnOffset -lt $ColumnCount; $columnOffset++) {
                $cell = $null
                try {
                    $cell = $worksheet.Cells.Item($StartRow + $rowOffset, $StartColumn + $columnOffset)
                    $value = if ($OutputMode -eq "raw") { $cell.Value2 } else { $cell.Text }
                    if ($null -eq $value) { $value = "" }
                    $cleaned = ([string]$value) -replace "[\r\n\t\x1e\x1f]", " "
                    [void]$rowValues.Add($cleaned)
                }
                finally {
                    Release-ComObject $cell
                }
            }
            [void]$matrix.Add(@($rowValues))
        }
        $outputDirectory = Split-Path -Parent ([System.IO.Path]::GetFullPath($OutputPath))
        if (-not (Test-Path -LiteralPath $outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
        }
        $json = ConvertTo-Json -InputObject @($matrix) -Depth 8
        [System.IO.File]::WriteAllText(
            [System.IO.Path]::GetFullPath($OutputPath),
            $json,
            (New-Object System.Text.UTF8Encoding($false))
        )
    }
    finally {
        if ($null -ne $workbookObject) { try { $workbookObject.Close($false) } catch { } }
        if ($null -ne $excel -and $initialCount -eq 0) { try { $excel.Quit() } catch { } }
        Release-ComObject $worksheet
        Release-ComObject $workbookObject
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

switch ($Action) {
    "CheckOffice" { Invoke-CheckOffice }
    "UpdatePresentation" { Invoke-UpdatePresentation }
    "ReplaceImage" { Invoke-ReplaceImage }
    "ReplaceExcelChart" { Invoke-ReplaceExcelChart }
    "ExportPdf" { Invoke-ExportPdf }
    "ExtractExcelRange" { Invoke-ExtractExcelRange }
}
