# scripts/create_sample_presentation.ps1 - Creates a sample 5-slide educational presentation
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$sampleFile = Join-Path $projectRoot "sample_presentation.pptx"

Write-Host "Creating sample educational presentation at: $sampleFile" -ForegroundColor Cyan

$ppt = New-Object -ComObject PowerPoint.Application

try {
    $pres = $ppt.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoFalse)

    # ----------------------------------------------------
    # Slide 1: Title Slide
    # ----------------------------------------------------
    $s1 = $pres.Slides.Add(1, 1) # ppLayoutTitle
    $s1.Shapes.Title.TextFrame.TextRange.Text = "Introduction to Algorithms: Sorting & Searching"
    if ($s1.Shapes.Placeholders.Count -ge 2) {
        $s1.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "CS 101: Understanding How Computers Solve Problems Step-by-Step`nProf. Teaching Assistant"
    }

    # ----------------------------------------------------
    # Slide 2: Bullet Points
    # ----------------------------------------------------
    $s2 = $pres.Slides.Add(2, 2) # ppLayoutText
    $s2.Shapes.Title.TextFrame.TextRange.Text = "What is an Algorithm?"
    $tf2 = $s2.Shapes.Placeholders.Item(2).TextFrame
    $tf2.TextRange.Text = "A step-by-step sequence of instructions to solve a problem.`n" +
                          "Must be unambiguous: every step is crystal clear and executable.`n" +
                          "Must terminate: guaranteed to stop after a finite number of operations.`n" +
                          "Transforms input data into desired output results reliably."

    # ----------------------------------------------------
    # Slide 3: Comparison Table
    # ----------------------------------------------------
    $s3 = $pres.Slides.Add(3, 12) # ppLayoutBlank / title only
    $title3 = $s3.Shapes.AddTextbox(1, 50, 30, 860, 50)
    $title3.TextFrame.TextRange.Text = "Comparing Search Strategies: Linear vs. Binary"
    $title3.TextFrame.TextRange.Font.Size = 28
    $title3.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue

    # Add Table (3 rows, 4 columns)
    $tableShape = $s3.Shapes.AddTable(3, 4, 50, 120, 860, 200)
    $tbl = $tableShape.Table

    $headers = @("Algorithm", "Data Requirement", "Worst-Case Time", "Best Used For")
    for ($c = 1; $c -le 4; $c++) {
        $tbl.Cell(1, $c).Shape.TextFrame.TextRange.Text = $headers[$c - 1]
        $tbl.Cell(1, $c).Shape.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    }

    $r1 = @("Linear Search", "None (unsorted lists)", "O(N) - scans every element", "Small or disordered datasets")
    for ($c = 1; $c -le 4; $c++) {
        $tbl.Cell(2, $c).Shape.TextFrame.TextRange.Text = $r1[$c - 1]
    }

    $r2 = @("Binary Search", "Data must already be sorted", "O(log N) - halves space each step", "Large sorted databases or lists")
    for ($c = 1; $c -le 4; $c++) {
        $tbl.Cell(3, $c).Shape.TextFrame.TextRange.Text = $r2[$c - 1]
    }

    # ----------------------------------------------------
    # Slide 4: Diagram Shapes with Alt Text
    # ----------------------------------------------------
    $s4 = $pres.Slides.Add(4, 12)
    $title4 = $s4.Shapes.AddTextbox(1, 50, 30, 860, 50)
    $title4.TextFrame.TextRange.Text = "The Divide and Conquer Strategy"
    $title4.TextFrame.TextRange.Font.Size = 28
    $title4.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue

    # Box 1: Divide
    $box1 = $s4.Shapes.AddShape(1, 80, 150, 220, 120) # msoShapeRectangle
    $box1.TextFrame.TextRange.Text = "1. DIVIDE`nBreak into smaller subproblems"
    $box1.AlternativeText = "Divide step: splitting the input dataset into smaller, independent chunks that are easier to solve."
    $box1.Title = "Step 1: Divide"

    # Box 2: Conquer
    $box2 = $s4.Shapes.AddShape(1, 370, 150, 220, 120)
    $box2.TextFrame.TextRange.Text = "2. CONQUER`nSolve subproblems recursively"
    $box2.AlternativeText = "Conquer step: solving base cases directly and handling intermediate chunks with repeated logic."
    $box2.Title = "Step 2: Conquer"

    # Box 3: Combine
    $box3 = $s4.Shapes.AddShape(1, 660, 150, 220, 120)
    $box3.TextFrame.TextRange.Text = "3. COMBINE`nMerge sub-results into final answer"
    $box3.AlternativeText = "Combine step: seamlessly stitching individual partial solutions into the final complete result."
    $box3.Title = "Step 3: Combine"

    # ----------------------------------------------------
    # Slide 5: Abstract Concept (Time Complexity)
    # ----------------------------------------------------
    $s5 = $pres.Slides.Add(5, 2)
    $s5.Shapes.Title.TextFrame.TextRange.Text = "Why Time Complexity Matters in the Real World"
    $tf5 = $s5.Shapes.Placeholders.Item(2).TextFrame
    $tf5.TextRange.Text = "As data grows to millions or billions of items, efficiency dictates feasibility.`n" +
                          "Linear scan on 1 billion items: requires ~1,000,000,000 checks (seconds to minutes).`n" +
                          "Binary search on 1 billion items: requires ~30 checks (fractions of a microsecond)!`n" +
                          "Clever algorithms beat faster hardware every single time."

    if (Test-Path $sampleFile) {
        Remove-Item $sampleFile -Force
    }

    $pres.SaveAs($sampleFile)
    $pres.Close()
    Write-Host "Sample presentation created successfully with 5 diverse slides!" -ForegroundColor Green
}
finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
