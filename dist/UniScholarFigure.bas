Attribute VB_Name = "UniScholarFigure"
'==========================================================================
' UniScholarFigure 1.2 Clean-Room Clone — VBA Module
' Version: 1.2.2   Release: 2026-07-17
'
' Ribbon callbacks (called from customUI14.xml):
'   OnTrimPNG, OnGenTable, OnLayerStack, OnMatrixGrid,
'   OnTableImages, OnAddIcon, OnShowVersion
'
' Behaviors ported from the Python CLI clone (clean-room). No proprietary
' assets are used.
'==========================================================================
Option Explicit

Public Const UNISFIG_VERSION As String = "1.2.2"
Public Const UNISFIG_RELEASE As String = "2026-07-17"

' ---- Color palette for layer-stack and matrix ----
Private Const PAL1 As Long = &HC6864F   ' blue (BGR)
Private Const PAL2 As Long = &HDA5DA5   ' cyan
Private Const PAL3 As Long = &H68BD60   ' green
Private Const PAL4 As Long = &HB07CF1   ' pink
Private Const PAL5 As Long = &H3AA4FA   ' orange
Private Const PAL6 As Long = &HB276B2   ' purple
Private Const PAL7 As Long = &H3FCFDE   ' yellow

'==========================================================================
' Ribbon callbacks (one Sub per feature)
'==========================================================================

Public Sub OnTrimPNG(Optional control As IRibbonControl)
    TrimPNGFiles
End Sub

Public Sub OnGenTable(Optional control As IRibbonControl)
    GenTableFromInput
End Sub

Public Sub OnLayerStack(Optional control As IRibbonControl)
    DrawLayerStack
End Sub

Public Sub OnMatrixGrid(Optional control As IRibbonControl)
    DrawMatrixGrid
End Sub

Public Sub OnTableImages(Optional control As IRibbonControl)
    FillTableWithImages
End Sub

Public Sub OnAddIcon(Optional control As IRibbonControl)
    AddIconToSlide
End Sub

Public Sub OnShowVersion(Optional control As IRibbonControl)
    MsgBox "Uni-Scholar Figure" & vbCrLf & _
           "Version: " & UNISFIG_VERSION & vbCrLf & _
           "Release: " & UNISFIG_RELEASE, _
           vbInformation, "About Uni-Scholar Figure"
End Sub

'==========================================================================
' Bonus: Generate Demo — builds a 3-slide showcase exercising all features
'==========================================================================
Public Sub OnGenerateDemo(Optional control As IRibbonControl)
    Dim pres As Object
    Set pres = Application.Presentations.Add(msoTrue)

    ' Slide 1: Title + table + icons (Features 2 + 6)
    Dim s1 As Object
    Set s1 = pres.Slides.Add(1, 1)  ' 1 = ppLayoutTitle
    s1.Shapes(1).TextFrame.TextRange.Text = "Uni-Scholar Figure Demo"
    s1.Shapes(2).TextFrame.TextRange.Text = "v" & UNISFIG_VERSION & "  (" & UNISFIG_RELEASE & ")"

    ' Table on slide 1
    Dim tbl As Object
    Set tbl = s1.Shapes.AddTable(3, 4, 60, 250, 800, 90).Table
    tbl.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Method"
    tbl.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Precision"
    tbl.Cell(1, 3).Shape.TextFrame.TextRange.Text = "Recall"
    tbl.Cell(1, 4).Shape.TextFrame.TextRange.Text = "F1"
    tbl.Cell(2, 1).Shape.TextFrame.TextRange.Text = "Baseline"
    tbl.Cell(2, 2).Shape.TextFrame.TextRange.Text = "0.82"
    tbl.Cell(2, 3).Shape.TextFrame.TextRange.Text = "0.79"
    tbl.Cell(2, 4).Shape.TextFrame.TextRange.Text = "0.80"
    tbl.Cell(3, 1).Shape.TextFrame.TextRange.Text = "Ours"
    tbl.Cell(3, 2).Shape.TextFrame.TextRange.Text = "0.91"
    tbl.Cell(3, 3).Shape.TextFrame.TextRange.Text = "0.88"
    tbl.Cell(3, 4).Shape.TextFrame.TextRange.Text = "0.89"

    ' Sample icons
    Dim iconNames As Variant
    iconNames = Array("check", "info", "warning", "lightbulb", "search", "gear")
    Dim i As Long
    For i = LBound(iconNames) To UBound(iconNames)
        Dim shp As Object
        Set shp = s1.Shapes.AddShape(9, 720 + (i Mod 3) * 50, 380 + (i \ 3) * 50, 36, 36)
        shp.Fill.ForeColor.RGB = RGB(&H1F, &H4E, &H79)
        shp.Line.Fill.Visible = msoFalse
        shp.TextFrame.TextRange.Text = CStr(iconNames(i))
        shp.TextFrame.TextRange.Font.Size = 8
        shp.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = 3
    Next i

    ' Slide 2: Layer Stack (Feature 3)
    Dim s2 As Object
    Set s2 = pres.Slides.Add(2, 2)  ' 2 = ppLayoutText
    s2.Shapes(1).TextFrame.TextRange.Text = "Architecture (Feature 3)"
    Dim layers As String
    layers = "Presentation|Application|Business Logic|Data Access|Storage"
    Dim layerArr() As String
    layerArr = Split(layers, "|")
    Dim n As Long
    n = UBound(layerArr) + 1
    Dim boxW As Single, boxH As Single, gap As Single
    boxW = 468: boxH = 50: gap = 14
    Dim left As Single, top As Single
    left = (pres.PageSetup.SlideWidth - boxW) / 2
    top = 130
    Dim palette(0 To 6) As Long
    palette(0) = PAL1: palette(1) = PAL2: palette(2) = PAL3
    palette(3) = PAL4: palette(4) = PAL5: palette(5) = PAL6: palette(6) = PAL7
    For i = 0 To n - 1
        Dim shapeTop As Single
        shapeTop = top + i * (boxH + gap)
        Dim shp2 As Object
        Set shp2 = s2.Shapes.AddShape(5, left, shapeTop, boxW, boxH)
        shp2.Fill.ForeColor.RGB = palette(i Mod 7)
        shp2.Line.ForeColor.RGB = RGB(&H33, &H33, &H33)
        shp2.TextFrame.TextRange.Text = layerArr(i)
        shp2.TextFrame.TextRange.Font.Size = 16
        shp2.TextFrame.TextRange.Font.Bold = msoTrue
        shp2.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        shp2.TextFrame.TextRange.ParagraphFormat.Alignment = 3
        If i < n - 1 Then
            Dim arrow As Object
            Set arrow = s2.Shapes.AddShape(35, left + (boxW - 12) / 2, shapeTop + boxH, 12, 12)
            arrow.Fill.ForeColor.RGB = RGB(&H66, &H66, &H66)
            arrow.Line.Fill.Visible = msoFalse
        End If
    Next i

    ' Slide 3: Matrix Grid (Feature 4)
    Dim s3 As Object
    Set s3 = pres.Slides.Add(3, 2)
    s3.Shapes(1).TextFrame.TextRange.Text = "Matrix Offset (Feature 4)"
    Dim nRows As Long, nCols As Long
    nRows = 3: nCols = 5
    Dim cellSize As Single, sgap As Single
    cellSize = 60: sgap = 8
    Dim stride As Single
    stride = cellSize + sgap
    Dim offset As Single
    offset = 30
    Dim totalW As Single
    totalW = nCols * stride + Abs(offset)
    Dim startLeft As Single
    startLeft = (pres.PageSetup.SlideWidth - totalW) / 2
    Dim r As Long, c As Long
    For r = 0 To nRows - 1
        Dim dx As Single
        dx = IIf(r Mod 2 = 1, offset, 0)
        For c = 0 To nCols - 1
            Dim ms As Object
            Set ms = s3.Shapes.AddShape(5, _
                startLeft + dx + c * stride, _
                150 + r * stride, _
                cellSize, cellSize)
            ms.Fill.ForeColor.RGB = PAL1
            ms.Line.ForeColor.RGB = RGB(255, 255, 255)
            ms.TextFrame.TextRange.Text = r & "," & c
            ms.TextFrame.TextRange.Font.Size = 10
            ms.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
            ms.TextFrame.TextRange.ParagraphFormat.Alignment = 3
        Next c
    Next r

    MsgBox "Demo generated: 3 slides exercising Features 2, 3, 4, 6." & vbCrLf & _
           "Other features (PNG Trim, Table Images) need files — try them on your own." & vbCrLf & vbCrLf & _
           "Uni-Scholar Figure v" & UNISFIG_VERSION, vbInformation, "Demo"
End Sub

'==========================================================================
' Feature 1: PNG blank-margin auto-trim
'==========================================================================
Public Sub TrimPNGFiles()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select PNG files to trim"
        .Filters.Clear
        .Filters.Add "PNG Images", "*.png"
        .AllowMultiSelect = True
        If .Show <> -1 Then Exit Sub
    End With

    Dim i As Long, count As Long, failed As Long
    count = 0: failed = 0
    For i = 1 To fd.SelectedItems.count
        If TrimOnePNG(CStr(fd.SelectedItems(i))) Then
            count = count + 1
        Else
            failed = failed + 1
        End If
    Next i
    MsgBox "Trimmed " & count & " file(s)." & _
           IIf(failed > 0, " Failed: " & failed, ""), vbInformation
End Sub

' Returns True on success.
Private Function TrimOnePNG(ByVal filePath As String) As Boolean
#If Mac Then
    TrimOnePNG = TrimOnePNGMac(filePath)
    Exit Function
#Else
    TrimOnePNG = TrimOnePNGWin(filePath)
    Exit Function
#End If
End Function

'==========================================================================
' Mac implementation: shells out to Python3 + Pillow
'==========================================================================
#If Mac Then
Private Function TrimOnePNGMac(ByVal filePath As String) As Boolean
    On Error GoTo Fail
    ' Inline Python script — bbox detection + crop using Pillow
    Dim pyScript As String
    pyScript = "import sys,os" & vbCrLf & _
               "from PIL import Image" & vbCrLf & _
               "p=sys.argv[1]" & vbCrLf & _
               "im=Image.open(p)" & vbCrLf & _
               "if im.mode=='RGBA':" & vbCrLf & _
               "    bg=Image.new('RGBA',im.size,(255,255,255,255))" & vbCrLf & _
               "    im=Image.alpha_composite(bg,im).convert('RGB')" & vbCrLf & _
               "else:" & vbCrLf & _
               "    im=im.convert('RGB')" & vbCrLf & _
               "g=im.convert('L')" & vbCrLf & _
               "mask=g.point(lambda v:255 if v<245 else 0)" & vbCrLf & _
               "b=mask.getbbox()" & vbCrLf & _
               "if b is None: sys.exit(1)" & vbCrLf & _
               "pad=2" & vbCrLf & _
               "b=(max(0,b[0]-pad),max(0,b[1]-pad),min(im.size[0],b[2]+pad),min(im.size[1],b[3]+pad))" & vbCrLf & _
               "im.crop(b).save(p,'PNG',optimize=True)"

    Dim tmpPath As String
    tmpPath = "/tmp/unisfig_trim_" & Format(Now, "yyyymmddhhnnss") & ".py"
    Dim f As Integer
    f = FreeFile
    Open tmpPath For Output As #f
    Print #f, pyScript
    Close #f

    ' AppleScript: do shell script — captures stderr + exit code
    Dim script As String
    script = "do shell script ""python3 "" & quoted form of """ & tmpPath & """ & "" "" & quoted form of """ & filePath & """ 2>&1"""
    Dim errMsg As String
    On Error Resume Next
    errMsg = MacScript(script)
    If Err.Number <> 0 Or Len(errMsg) > 0 Then
        ' Check if it's a "no content" success or real failure
        If InStr(errMsg, "sys.exit(1)") > 0 Then
            TrimOnePNGMac = False
        ElseIf InStr(errMsg, "No module named PIL") > 0 Or InStr(errMsg, "pip3") > 0 Then
            MsgBox "PNG Trim on Mac requires Python3 + Pillow." & vbCrLf & vbCrLf & _
                   "Install once with:" & vbCrLf & _
                   "  brew install python" & vbCrLf & _
                   "  pip3 install Pillow", vbExclamation, "Setup Required"
            TrimOnePNGMac = False
        ElseIf Len(errMsg) = 0 Then
            TrimOnePNGMac = True
        Else
            TrimOnePNGMac = False
        End If
    Else
        TrimOnePNGMac = True
    End If
    On Error GoTo 0

    ' Cleanup
    On Error Resume Next
    MacScript "do shell script ""rm -f "" & quoted form of """ & tmpPath & """"
    Exit Function

Fail:
    TrimOnePNGMac = False
End Function
#End If

'==========================================================================
' Windows implementation: GDI+ flat API
'==========================================================================
#If Not Mac Then
Private Function TrimOnePNGWin(ByVal filePath As String) As Boolean
    On Error GoTo Fail

    Dim gdipToken As LongPtr
    If Not GdipInit(gdipToken) Then GoTo Fail

    Dim img As LongPtr
    If GdipCreateBitmapFromFile(StrPtr(filePath), img) <> 0 Then GoTo FailCleanup

    Dim imgW As Long, imgH As Long
    GdipGetImageWidth img, imgW
    GdipGetImageHeight img, imgH

    ' Scan for content bbox (pixels darker than threshold)
    Dim tol As Long
    tol = 245  ' L threshold below which = content
    Dim leftB As Long, topB As Long, rightB As Long, bottomB As Long
    leftB = -1: topB = -1: rightB = -1: bottomB = -1

    Dim x As Long, y As Long, pixel As Long
    Dim r As Long, g As Long, b As Long, lum As Long
    Dim foundContent As Boolean
    foundContent = False

    For y = 0 To imgH - 1
        For x = 0 To imgW - 1
            GdipBitmapGetPixel img, x, y, pixel
            ' pixel is ARGB; extract RGB
            b = pixel And &HFF&
            g = (pixel \ &H100&) And &HFF&
            r = (pixel \ &H10000) And &HFF&
            ' Luminance ~ 0.299R + 0.587G + 0.114B
            lum = (299 * r + 587 * g + 114 * b) \ 1000
            If lum < tol Then
                foundContent = True
                If leftB < 0 Or x < leftB Then leftB = x
                If rightB < 0 Or x > rightB Then rightB = x
                If topB < 0 Or y < topB Then topB = y
                If bottomB < 0 Or y > bottomB Then bottomB = y
            End If
        Next x
    Next y

    If Not foundContent Then
        GdipDisposeImage img
        GdipShutdown gdipToken
        TrimOnePNGWin = False
        Exit Function
    End If

    ' Add 2px padding
    leftB = leftB - 2: If leftB < 0 Then leftB = 0
    topB = topB - 2: If topB < 0 Then topB = 0
    rightB = rightB + 2: If rightB >= imgW Then rightB = imgW - 1
    bottomB = bottomB + 2: If bottomB >= imgH Then bottomB = imgH - 1

    Dim newW As Long, newH As Long
    newW = rightB - leftB + 1
    newH = bottomB - topB + 1

    ' Create new bitmap of cropped size
    Dim dest As LongPtr
    GdipCreateBitmapFromScan0 newW, newH, 0, &H26200A, 0, dest  ' PixelFormat32bppARGB
    Dim graphics As LongPtr
    GdipGetImageGraphicsContext dest, graphics
    GdipDrawImageRectI graphics, img, 0, 0, newW, newH, _
        leftB, topB, newW, newH, Unit.Pixel
    GdipDeleteGraphics graphics

    ' Save back to PNG
    Dim clsid As String
    ClsidFromString "{557CF406-1A04-11D3-9A73-0000F81EF32E}", clsid  ' PNG encoder
    Dim encoderParams(0 To 1) As Byte
    GdipSaveImageToFile dest, StrPtr(filePath), StrPtr(clsid), 0&

    GdipDisposeImage img
    GdipDisposeImage dest
    GdipShutdown gdipToken
    TrimOnePNGWin = True
    Exit Function

FailCleanup:
    GdipShutdown gdipToken
Fail:
    TrimOnePNGWin = False
End Function
#End If

'==========================================================================
' Feature 2: Generate Table
'==========================================================================
Public Sub GenTableFromInput()
    Dim raw As String
    raw = InputBox("Paste table data." & vbCrLf & _
                   "Supported: CSV (one row per line, comma-sep)," & vbCrLf & _
                   "           Markdown (| a | b | rows)," & vbCrLf & _
                   "           Plain (one cell per line).", _
                   "Generate Table", "Method,Precision,Recall,F1" & vbCrLf & _
                   "Baseline,0.82,0.79,0.80" & vbCrLf & _
                   "Ours,0.91,0.88,0.89")
    If Len(raw) = 0 Then Exit Sub

    Dim rows() As String
    rows = Split(raw, vbCrLf)
    If UBound(rows) < 0 Then Exit Sub

    Dim tableData() As Variant
    ReDim tableData(0 To UBound(rows))
    Dim nCols As Long, nRows As Long
    nCols = 0
    Dim i As Long
    For i = 0 To UBound(rows)
        Dim r As String
        r = Trim(rows(i))
        If Len(r) > 0 Then
            Dim cells() As String
            If Left(r, 1) = "|" Then
                ' Markdown — strip pipes, skip alignment row
                If IsAlignmentRow(r) Then
                    tableData(i) = Empty
                Else
                    Dim clean As String
                    clean = Mid(r, 2, Len(r) - 2)
                    cells = Split(clean, "|")
                    Dim j As Long
                    For j = 0 To UBound(cells)
                        cells(j) = Trim(cells(j))
                    Next j
                    tableData(i) = cells
                    If UBound(cells) + 1 > nCols Then nCols = UBound(cells) + 1
                End If
            ElseIf InStr(r, ",") > 0 Then
                cells = Split(r, ",")
                Dim k As Long
                For k = 0 To UBound(cells)
                    cells(k) = Trim(cells(k))
                Next k
                tableData(i) = cells
                If UBound(cells) + 1 > nCols Then nCols = UBound(cells) + 1
            Else
                ReDim cells(0 To 0)
                cells(0) = r
                tableData(i) = cells
                If 1 > nCols Then nCols = 1
            End If
        Else
            tableData(i) = Empty
        End If
    Next i

    ' Compact rows (skip empty)
    Dim compact() As Variant
    ReDim compact(0 To UBound(rows))
    Dim rc As Long
    rc = 0
    For i = 0 To UBound(rows)
        If Not IsEmpty(tableData(i)) Then
            compact(rc) = tableData(i)
            rc = rc + 1
        End If
    Next i
    nRows = rc
    If nRows = 0 Or nCols = 0 Then Exit Sub

    Dim slide As Object
    Set slide = ActiveWindow.View.Slide
    Dim tbl As Object
    Dim shp As Object
    Set shp = slide.Shapes.AddTable(nRows, nCols, 60, 100, 800, 30 * nRows)
    Set tbl = shp.Table

    For i = 0 To nRows - 1
        Dim rowCells As Variant
        rowCells = compact(i)
        Dim c As Long
        For c = 0 To nCols - 1
            If c <= UBound(rowCells) Then
                tbl.Cell(i + 1, c + 1).Shape.TextFrame.TextRange.Text = CStr(rowCells(c))
            Else
                tbl.Cell(i + 1, c + 1).Shape.TextFrame.TextRange.Text = ""
            End If
        Next c
    Next i

    MsgBox "Added " & nRows & "x" & nCols & " table.", vbInformation
End Sub

Private Function IsAlignmentRow(ByVal s As String) As Boolean
    Dim t As String
    t = Replace(s, "|", "")
    t = Replace(t, "-", "")
    t = Replace(t, ":", "")
    t = Replace(t, " ", "")
    IsAlignmentRow = (Len(t) = 0)
End Function

'==========================================================================
' Feature 3: Layer Stack
'==========================================================================
Public Sub DrawLayerStack()
    Dim raw As String
    raw = InputBox("Enter layer names, one per line." & vbCrLf & _
                   "(Top layer first, bottom layer last)", _
                   "Layer Stack", "Presentation" & vbCrLf & "Application" & vbCrLf & _
                   "Business Logic" & vbCrLf & "Data Access" & vbCrLf & "Storage")
    If Len(raw) = 0 Then Exit Sub

    Dim layers() As String
    layers = Split(raw, vbCrLf)
    Dim n As Long, i As Long
    n = 0
    For i = 0 To UBound(layers)
        If Len(Trim(layers(i))) > 0 Then n = n + 1
    Next i
    If n = 0 Then Exit Sub

    Dim cleanLayers() As String
    ReDim cleanLayers(0 To n - 1)
    Dim idx As Long
    idx = 0
    For i = 0 To UBound(layers)
        If Len(Trim(layers(i))) > 0 Then
            cleanLayers(idx) = Trim(layers(i))
            idx = idx + 1
        End If
    Next i

    Dim slide As Object
    Set slide = ActiveWindow.View.Slide

    Dim boxW As Single, boxH As Single, gap As Single
    boxW = 468  ' ~6.5in in points
    boxH = 58
    gap = 18

    Dim left As Single, top As Single
    left = (Application.ActivePresentation.PageSetup.SlideWidth - boxW) / 2
    top = 100

    Dim palette(0 To 6) As Long
    palette(0) = PAL1: palette(1) = PAL2: palette(2) = PAL3
    palette(3) = PAL4: palette(4) = PAL5: palette(5) = PAL6: palette(6) = PAL7

    For i = 0 To n - 1
        Dim shapeTop As Single
        shapeTop = top + i * (boxH + gap)
        Dim shp As Object
        Set shp = slide.Shapes.AddShape(5, left, shapeTop, boxW, boxH)  ' 5 = msoShapeRoundedRectangle
        shp.Fill.ForeColor.RGB = palette(i Mod 7)
        shp.Line.ForeColor.RGB = RGB(&H33, &H33, &H33)
        shp.TextFrame.TextRange.Text = cleanLayers(i)
        shp.TextFrame.TextRange.Font.Size = 18
        shp.TextFrame.TextRange.Font.Bold = msoTrue
        shp.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = 3  ' ppAlignCenter

        If i < n - 1 Then
            Dim arrowSize As Single
            arrowSize = 13
            Dim arrowLeft As Single
            arrowLeft = left + (boxW - arrowSize) / 2
            Dim arrowTop As Single
            arrowTop = shapeTop + boxH
            Dim arrow As Object
            Set arrow = slide.Shapes.AddShape(35, arrowLeft, arrowTop, arrowSize, arrowSize)  ' 35 = msoShapeDownArrow
            arrow.Fill.ForeColor.RGB = RGB(&H66, &H66, &H66)
            arrow.Line.Fill.Visible = msoFalse
        End If
    Next i

    MsgBox "Drew " & n & " layers.", vbInformation
End Sub

'==========================================================================
' Feature 4: Matrix Grid with Row Offset
'==========================================================================
Public Sub DrawMatrixGrid()
    Dim sRows As String, sCols As String, sOffset As String, sMode As String
    sRows = InputBox("Number of rows?", "Matrix Grid", "3")
    If Len(sRows) = 0 Then Exit Sub
    sCols = InputBox("Number of columns?", "Matrix Grid", "5")
    If Len(sCols) = 0 Then Exit Sub
    sOffset = InputBox("Row offset (points)?", "Matrix Grid", "36")
    If Len(sOffset) = 0 Then Exit Sub
    sMode = InputBox("Mode: alternate / progressive / none?", "Matrix Grid", "alternate")
    If Len(sMode) = 0 Then Exit Sub

    Dim nRows As Long, nCols As Long
    nRows = CLng(sRows): nCols = CLng(sCols)
    Dim offsetPt As Single
    offsetPt = CSng(sOffset)
    Dim mode As String
    mode = LCase(Trim(sMode))

    Dim slide As Object
    Set slide = ActiveWindow.View.Slide

    Dim cellSize As Single, gap As Single
    cellSize = 72: gap = 9
    Dim stride As Single
    stride = cellSize + gap

    Dim totalW As Single
    totalW = nCols * stride + Abs(offsetPt)
    Dim startLeft As Single
    startLeft = (Application.ActivePresentation.PageSetup.SlideWidth - totalW) / 2
    Dim top0 As Single
    top0 = 100

    Dim r As Long, c As Long
    For r = 0 To nRows - 1
        Dim dx As Single
        If mode = "alternate" Then
            dx = IIf(r Mod 2 = 1, offsetPt, 0)
        ElseIf mode = "progressive" Then
            dx = offsetPt * r
        Else
            dx = 0
        End If
        For c = 0 To nCols - 1
            Dim shp As Object
            Set shp = slide.Shapes.AddShape(5, _
                startLeft + dx + c * stride, _
                top0 + r * stride, _
                cellSize, cellSize)
            shp.Fill.ForeColor.RGB = PAL1
            shp.Line.ForeColor.RGB = RGB(255, 255, 255)
            shp.TextFrame.TextRange.Text = r & "," & c
            shp.TextFrame.TextRange.Font.Size = 11
            shp.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = 3
        Next c
    Next r

    MsgBox "Drew " & nRows & "x" & nCols & " matrix.", vbInformation
End Sub

'==========================================================================
' Feature 5: Table with images (preserving aspect ratio)
'==========================================================================
Public Sub FillTableWithImages()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select images"
        .Filters.Clear
        .Filters.Add "Images", "*.png;*.jpg;*.jpeg;*.gif;*.bmp"
        .AllowMultiSelect = True
        If .Show <> -1 Then Exit Sub
    End With

    Dim sRows As String, sCols As String
    sRows = InputBox("Grid rows?", "Table Images", "2")
    If Len(sRows) = 0 Then Exit Sub
    sCols = InputBox("Grid cols?", "Table Images", "3")
    If Len(sCols) = 0 Then Exit Sub
    Dim nRows As Long, nCols As Long
    nRows = CLng(sRows): nCols = CLng(sCols)

    Dim slide As Object
    Set slide = ActiveWindow.View.Slide

    Dim cellW As Single, cellH As Single
    cellW = 144: cellH = 144  ' 2in x 2in
    Dim gap As Single
    gap = 3.6
    Dim left0 As Single, top0 As Single
    left0 = 60: top0 = 100

    Dim totalW As Single, totalH As Single
    totalW = nCols * (cellW + gap)
    totalH = nRows * (cellH + gap)

    Dim shp As Object
    Set shp = slide.Shapes.AddTable(nRows, nCols, left0, top0, totalW, totalH)
    Dim tbl As Object
    Set tbl = shp.Table
    For c = 0 To nCols - 1
        tbl.Columns(c + 1).Width = cellW
    Next c
    For r = 0 To nRows - 1
        tbl.Rows(r + 1).Height = cellH
    Next r

    Dim fileIdx As Long
    fileIdx = 1
    Dim r As Long, c As Long
    For r = 0 To nRows - 1
        For c = 0 To nCols - 1
            If fileIdx <= fd.SelectedItems.count Then
                Dim imgPath As String
                imgPath = CStr(fd.SelectedItems(fileIdx))

                ' Determine image dimensions for aspect fit
                Dim picW As Single, picH As Single
                picW = 0: picH = 0
                GetImageSize imgPath, picW, picH

                Dim newW As Single, newH As Single
                If picW > 0 And picH > 0 Then
                    If picW / picH > cellW / cellH Then
                        newW = cellW
                        newH = cellW / (picW / picH)
                    Else
                        newH = cellH
                        newW = cellH * (picW / picH)
                    End If
                Else
                    newW = cellW: newH = cellH
                End If

                ' Place picture absolutely (covers the cell)
                Dim absX As Single, absY As Single
                absX = left0 + c * (cellW + gap) + (cellW - newW) / 2
                absY = top0 + r * (cellH + gap) + (cellH - newH) / 2
                slide.Shapes.AddPicture imgPath, msoFalse, msoTrue, _
                    absX, absY, newW, newH

                ' Cell margin to 0 for predictable math
                With tbl.Cell(r + 1, c + 1)
                    .MarginLeft = 0
                    .MarginTop = 0
                    .MarginRight = 0
                    .MarginBottom = 0
                End With
            End If
            fileIdx = fileIdx + 1
        Next c
    Next r

    MsgBox "Placed " & fd.SelectedItems.count & " images.", vbInformation
End Sub

Private Sub GetImageSize(ByVal filePath As String, ByRef w As Single, ByRef h As Single)
    On Error Resume Next
    Dim shp As Object
    ' Create temp shape to measure
    Set shp = ActivePresentation.Slides(1).Shapes.AddPicture(filePath, msoFalse, msoFalse, 0, 0)
    If Not shp Is Nothing Then
        w = shp.Width
        h = shp.Height
        shp.Delete
    End If
End Sub

'==========================================================================
' Feature 6: Add Icon
'==========================================================================
Public Sub AddIconToSlide()
    Dim icons As String
    icons = "check|cross|arrow-right|arrow-down|arrow-up|arrow-left|info|warning|lightbulb|search|gear|doc|book|chart|code|brain|atom|dna|microscope|beaker|flask|graph|database|server|cloud|lock|key|flag|star|heart|thumbsup|link|mail|phone|calendar|clock|user|team|settings|trash|edit|save|download|upload|copy|paste|filter|sort|expand|collapse|play|pause|stop|forward|backward"
    Dim name As String
    name = InputBox("Icon name. Available:" & vbCrLf & Replace(icons, "|", ", "), _
                    "Add Icon", "warning")
    If Len(name) = 0 Then Exit Sub

    Dim valid As Boolean
    valid = False
    Dim item As Variant
    For Each item In Split(icons, "|")
        If LCase(name) = item Then valid = True: Exit For
    Next
    If Not valid Then
        MsgBox "Unknown icon: " & name, vbExclamation
        Exit Sub
    End If

    Dim slide As Object
    Set slide = ActiveWindow.View.Slide

    Dim shp As Object
    Set shp = slide.Shapes.AddShape(9, 60, 60, 36, 36)  ' 9 = msoShapeOval
    shp.Fill.ForeColor.RGB = RGB(&H1F, &H4E, &H79)
    shp.Line.Fill.Visible = msoFalse
    shp.TextFrame.TextRange.Text = name
    shp.TextFrame.TextRange.Font.Size = 10
    shp.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    shp.TextFrame.TextRange.Font.Bold = msoTrue
    shp.TextFrame.TextRange.ParagraphFormat.Alignment = 3

    MsgBox "Added '" & name & "' icon.", vbInformation
End Sub

'==========================================================================
' GDI+ Flat API declarations (used by Feature 1 PNG trim, Windows only)
'==========================================================================
#If Not Mac Then
#If VBA7 Then
    Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (token As LongPtr, inputbuf As Any, Optional outputbuf As Any) As Long
    Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
    Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As LongPtr, ByVal bitmap As LongPtr) As Long
    Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" (ByVal image As LongPtr, Width As Long) As Long
    Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" (ByVal image As LongPtr, Height As Long) As Long
    Private Declare PtrSafe Function GdipBitmapGetPixel Lib "gdiplus" (ByVal bitmap As LongPtr, ByVal x As Long, ByVal y As Long, argb As Long) As Long
    Private Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal format As Long, ByVal scan0 As LongPtr, bitmap As LongPtr) As Long
    Private Declare PtrSafe Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As LongPtr, graphics As LongPtr) As Long
    Private Declare PtrSafe Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As LongPtr, ByVal image As LongPtr, ByVal dx As Long, ByVal dy As Long, ByVal dw As Long, ByVal dh As Long, ByVal sx As Long, ByVal sy As Long, ByVal sw As Long, ByVal sh As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As LongPtr = 0, Optional ByVal callback As LongPtr = 0, Optional ByVal callbackData As LongPtr = 0) As Long
    Private Declare PtrSafe Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As LongPtr) As Long
    Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long
    Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As LongPtr, ByVal filename As LongPtr, ByVal clsidEncoder As LongPtr, Optional ByVal encoderParams As LongPtr = 0) As Long
#Else
    Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As Any, Optional outputbuf As Any) As Long
    Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
    Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As Long, ByVal bitmap As Long) As Long
    Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As Long
    Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As Long
    Private Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal bitmap As Long, ByVal x As Long, ByVal y As Long, argb As Long) As Long
    Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal format As Long, ByVal scan0 As Long, bitmap As Long) As Long
    Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As Long, graphics As Long) As Long
    Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal dx As Long, ByVal dy As Long, ByVal dw As Long, ByVal dh As Long, ByVal sx As Long, ByVal sy As Long, ByVal sw As Long, ByVal sh As Long, ByVal srcUnit As Long) As Long
    Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
    Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
    Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As Long, ByVal filename As Long, ByVal clsidEncoder As Long, Optional ByVal encoderParams As Long = 0) As Long
#End If

Private Function GdipInit(ByRef token As LongPtr) As Boolean
    #If VBA7 Then
        Dim si(0 To 15) As Byte
        si(0) = 1  ' GdiplusStartupInput structure version = 1
        Dim tokenL As LongPtr
        If GdiplusStartup(tokenL, si(0)) = 0 Then
            token = tokenL
            GdipInit = True
        End If
    #Else
        Dim si(0 To 15) As Byte
        si(0) = 1
        Dim tokenL As Long
        If GdiplusStartup(tokenL, si(0)) = 0 Then
            token = tokenL
            GdipInit = True
        End If
    #End If
End Function

Private Declare PtrSafe Sub ClsidFromString Lib "ole32" (ByVal str As LongPtr, clsid As Any)
#End If
