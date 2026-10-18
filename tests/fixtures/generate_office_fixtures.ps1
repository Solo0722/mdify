# Optional fixture generator. Requires desktop Microsoft Office; not used by CI.
# Creates only synthetic documents. Existing user documents are never edited.
$ErrorActionPreference = 'Stop'
$fixtureDir = $PSScriptRoot
$missing = [System.Type]::Missing

# Refuse to overwrite fixtures during an accidental rerun. Remove or relocate
# the three explicitly named fixtures before deliberately regenerating them.
foreach ($name in @('office.pptx', 'office.docx', 'office.xlsx')) {
    if (Test-Path -LiteralPath (Join-Path $fixtureDir $name)) {
        throw "Fixture already exists: $name"
    }
}

$ppt = $null
$deck = $null
$pptWasEmpty = $false
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pptWasEmpty = $ppt.Presentations.Count -eq 0
    $deck = $ppt.Presentations.Add(0)
    $overview = $deck.Slides.Add(1, 2)
    $overview.Shapes.Title.TextFrame.TextRange.Text = 'Overview'
    $overview.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "First bullet`rNested bullet`rLast bullet"
    $overview.Shapes.Placeholders.Item(2).TextFrame.TextRange.Paragraphs(2).IndentLevel = 2
    $overview.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Speaker note: synthetic fixture.'
    $box1 = $overview.Shapes.AddTextbox(1, 40, 320, 180, 40)
    $box1.TextFrame.TextRange.Text = 'Grouped left'
    $box2 = $overview.Shapes.AddTextbox(1, 240, 320, 180, 40)
    $box2.TextFrame.TextRange.Text = 'Grouped right'
    $names = @($box1.Name, $box2.Name)
    $null = $overview.Shapes.Range($names).Group()
    $results = $deck.Slides.Add(2, 11)
    $results.Shapes.Title.TextFrame.TextRange.Text = 'Results'
    $table = $results.Shapes.AddTable(2, 3, 40, 120, 600, 100).Table
    $table.Cell(1, 1).Shape.TextFrame.TextRange.Text = 'Merged header'
    $table.Cell(1, 1).Merge($table.Cell(1, 2))
    $table.Cell(1, 3).Shape.TextFrame.TextRange.Text = 'Third'
    $table.Cell(2, 1).Shape.TextFrame.TextRange.Text = 'A|B'
    $table.Cell(2, 2).Shape.TextFrame.TextRange.Text = 'C'
    $table.Cell(2, 3).Shape.TextFrame.TextRange.Text = 'D'
    $results.SlideShowTransition.Hidden = -1
    $results.MoveTo(1)
    $null = $deck.Slides.Add(3, 12)
    $deck.SaveAs((Join-Path $fixtureDir 'office.pptx'), 24)
    Write-Output 'Created office.pptx'
} finally {
    if ($null -ne $deck) { $deck.Close() }
    if ($null -ne $ppt) {
        if ($pptWasEmpty -and $ppt.Presentations.Count -eq 0) { $ppt.Quit() }
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
    }
}

# Metadata sanitization is also available as a standalone Python script.

$word = $null
$doc = $null
$wordWasEmpty = $false
try {
    $word = New-Object -ComObject Word.Application
    $wordWasEmpty = $word.Documents.Count -eq 0
    $doc = $word.Documents.Add($missing, $false, 0, $false)
    $doc.Content.Text = "Fixture heading`rFirst bullet`rSecond bullet`rPlain text`r"
    $doc.Paragraphs.Item(1).Range.Style = -2 # wdStyleHeading1
    $doc.Paragraphs.Item(2).Range.Style = -49 # wdStyleListBullet
    $doc.Paragraphs.Item(3).Range.Style = -49
    $range = $doc.Paragraphs.Item(4).Range
    $control = $doc.ContentControls.Add(1, $range)
    $control.Range.Text = 'Content control preserved'
    $doc.SaveAs2((Join-Path $fixtureDir 'office.docx'), 16)
    Write-Output 'Created office.docx'
} finally {
    if ($null -ne $doc) { $doc.Close(0) }
    if ($null -ne $word) {
        if ($wordWasEmpty -and $word.Documents.Count -eq 0) { $word.Quit() }
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}

$excel = $null
$book = $null
$excelWasEmpty = $false
try {
    $excel = New-Object -ComObject Excel.Application
    $excelWasEmpty = $excel.Workbooks.Count -eq 0
    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = 'Values'
    $sheet.Cells.Item(1, 1).Value2 = 'Date'
    $sheet.Cells.Item(1, 2).Value2 = 'Total'
    $sheet.Cells.Item(2, 1).Value2 = 45292
    $sheet.Cells.Item(2, 1).NumberFormat = 'yyyy-mm-dd'
    $sheet.Cells.Item(2, 2).Formula = '=2+3'
    $book.SaveAs((Join-Path $fixtureDir 'office.xlsx'), 51)
    Write-Output 'Created office.xlsx'
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) {
        if ($excelWasEmpty -and $excel.Workbooks.Count -eq 0) { $excel.Quit() }
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}
