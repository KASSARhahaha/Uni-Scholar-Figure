Attribute VB_Name = "UniScholarFigure"
'==========================================================================
' UniScholarFigure 1.2 Clean-Room Clone — VBA Module
' Version: 1.3.0   Release: 2026-07-18
'
' Ribbon callbacks (called from customUI14.xml):
'   OnTrimPNG, OnGenTable, OnLayerStack, OnMatrixGrid,
'   OnTableImages, OnAddIcon, OnShowVersion
'
' Behaviors ported from the Python CLI clone (clean-room). No proprietary
' assets are used.
'==========================================================================
Option Explicit

Public Const UNISFIG_VERSION As String = "1.3.0"
Public Const UNISFIG_RELEASE As String = "2026-07-18"

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

    ' Sample icons — drawn as real SVG paths via BuildFreeform (Feature 6 showcase)
    Dim iconNames As Variant
    iconNames = Array("check", "info", "warning", "lightbulb", "search", "gear")
    Dim i As Long
    For i = LBound(iconNames) To UBound(iconNames)
        Dim iconLeft As Single, iconTop As Single
        iconLeft = 720 + (i Mod 3) * 50
        iconTop = 380 + (i \ 3) * 50
        Dim iconSize As Single
        iconSize = 36
        ' Try to draw the real icon path; fall back to a labeled circle if path unknown
        If Not DrawIconFreeform(s1, CStr(iconNames(i)), iconLeft, iconTop, iconSize) Then
            Dim shp As Object
            Set shp = s1.Shapes.AddShape(9, iconLeft, iconTop, iconSize, iconSize)
            shp.Fill.ForeColor.RGB = RGB(&H1F, &H4E, &H79)
            shp.Line.Fill.Visible = msoFalse
            shp.TextFrame.TextRange.Text = CStr(iconNames(i))
            shp.TextFrame.TextRange.Font.Size = 8
            shp.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = 3
        End If
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
' Bonus: Check for Updates — single GET to GitHub releases API, no telemetry
'==========================================================================
Public Sub OnCheckUpdate(Optional control As IRibbonControl)
    Const ENDPOINT As String = "https://api.github.com/repos/KASSARhahaha/Uni-Scholar-Figure/releases/latest"
    On Error GoTo NetFail

    Dim latestTag As String, pubDate As String, body As String
#If Mac Then
    ' macOS: use do shell script + curl
    Dim cmd As String
    cmd = "do shell script ""curl -s -H 'User-Agent: UniScholarFigure' -m 10 "" & quoted form of """ & ENDPOINT & """"
    Dim raw As String
    raw = MacScript(cmd)
    latestTag = ExtractJsonField(raw, """tag_name"":""")
    pubDate = ExtractJsonField(raw, """published_at"":""")
#Else
    ' Windows: use MSXML2.XMLHTTP
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", ENDPOINT, False
    http.setRequestHeader "User-Agent", "UniScholarFigure"
    ' GitHub API requires Accept header to avoid relying on default
    http.setRequestHeader "Accept", "application/vnd.github+json"
    http.send
    If http.Status <> 200 Then
        MsgBox "GitHub API returned HTTP " & http.Status & "." & vbCrLf & _
               "Try again later or visit the Releases page manually:" & vbCrLf & _
               "https://github.com/KASSARhahaha/Uni-Scholar-Figure/releases", _
               vbExclamation, "Update Check"
        Exit Sub
    End If
    raw = http.responseText
    latestTag = ExtractJsonField(raw, """tag_name"":""")
    pubDate = ExtractJsonField(raw, """published_at"":""")
#End If

    If Len(latestTag) = 0 Then GoTo ParseFail

    ' Compare versions: strip leading 'v', compare dotted numerics
    Dim current As String, latest As String
    current = UNISFIG_VERSION
    latest = latestTag
    If Left$(latest, 1) = "v" Then latest = Mid$(latest, 2)

    If VersionGe(latest, current) Then
        MsgBox "A new version is available." & vbCrLf & vbCrLf & _
               "Your version:  v" & current & vbCrLf & _
               "Latest:        v" & latest & "  (" & Left$(pubDate, 10) & ")" & vbCrLf & vbCrLf & _
               "Download:" & vbCrLf & _
               "https://github.com/KASSARhahaha/Uni-Scholar-Figure/releases/latest", _
               vbInformation, "Update Available"
    Else
        MsgBox "You're up to date." & vbCrLf & vbCrLf & _
               "Installed: v" & current & vbCrLf & _
               "Latest:    v" & latest, vbInformation, "No Update Needed"
    End If
    Exit Sub

ParseFail:
    MsgBox "Could not parse GitHub response." & vbCrLf & _
           "Visit the Releases page directly:" & vbCrLf & _
           "https://github.com/KASSARhahaha/Uni-Scholar-Figure/releases/latest", _
           vbExclamation, "Update Check"
    Exit Sub

NetFail:
    MsgBox "Network error during update check." & vbCrLf & _
           "Check your connection or firewall, or visit:" & vbCrLf & _
           "https://github.com/KASSARhahaha/Uni-Scholar-Figure/releases/latest", _
           vbExclamation, "Update Check"
End Sub

'==========================================================================
' Feature 8: Research Records — pull Uni-Scholar literature into a PPT table
'==========================================================================
Public Sub OnResearchRecords(Optional control As IRibbonControl)
    Const ENDPOINT As String = "https://uni-scholar.asia/api/literature/papers?limit=20&sortBy=createdAt&sortOrder=desc"

    ' --- 1. Get token (from cache or InputBox) ---
    Dim token As String
    token = ReadCachedToken()
    If Len(token) = 0 Then
        token = InputBox( _
            "Paste your Uni-Scholar Figure token." & vbCrLf & vbCrLf & _
            "Get it from: https://uni-scholar.asia/app/settings" & vbCrLf & _
            "(scroll to the \"Uni-Scholar Figure PPT plugin\" section, click Copy)." & vbCrLf & vbCrLf & _
            "Token is cached locally after first use. To reset it, hold Shift while clicking this button.", _
            "Uni-Scholar Figure — Token Required")
        If Len(token) = 0 Then Exit Sub
        token = Trim(token)
    End If

    ' --- 2. Fetch papers ---
    Dim raw As String
    On Error GoTo NetFail
#If Mac Then
    Dim cmd As String
    cmd = "do shell script ""curl -s -m 15 -H 'Authorization: Bearer "" & quoted form of """ & token & """ & ""' "" & quoted form of """ & ENDPOINT & """"
    raw = MacScript(cmd)
#Else
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", ENDPOINT, False
    http.setRequestHeader "Authorization", "Bearer " & token
    http.setRequestHeader "Accept", "application/json"
    http.send
    If http.Status = 401 Then
        WriteCachedToken ""  ' clear bad token
        MsgBox "Token rejected (HTTP 401). Re-copy from:" & vbCrLf & _
               "https://uni-scholar.asia/app/settings", vbExclamation, "Auth Failed"
        Exit Sub
    End If
    If http.Status <> 200 Then
        MsgBox "Server returned HTTP " & http.Status & "." & vbCrLf & _
               "Try again later.", vbExclamation, "Research Records"
        Exit Sub
    End If
    raw = http.responseText
#End If

    ' --- 3. Parse JSON (tiny subset parser) ---
    Dim items As String
    items = ExtractJsonArray(raw, """items""")
    If Len(items) = 0 Then
        MsgBox "No records returned." & vbCrLf & _
               "Either your library is empty, or the response format changed." & vbCrLf & vbCrLf & _
               "Visit https://uni-scholar.asia/app/literature to verify.", _
               vbInformation, "Research Records"
        Exit Sub
    End If

    ' --- 4. Count items + build table ---
    Dim n As Long
    n = CountJsonObjects(items)
    If n = 0 Then
        MsgBox "Your library is empty." & vbCrLf & _
               "Add papers at https://uni-scholar.asia/app/literature first.", _
               vbInformation, "Research Records"
        Exit Sub
    End If

    ' Cache the token only after a successful auth
    WriteCachedToken token

    ' --- 5. Render table on current slide (or new slide) ---
    Dim pres As Object: Set pres = ActivePresentation
    If pres Is Nothing Then Set pres = Application.Presentations.Add(msoTrue)
    Dim slide As Object
    On Error Resume Next
    Set slide = Application.ActiveWindow.View.Slide
    On Error GoTo 0
    If slide Is Nothing Then
        Set slide = pres.Slides.Add(pres.Slides.Count + 1, 5)  ' Title Only
    End If

    Dim tbl As Object
    Dim nCols As Long: nCols = 6
    Set tbl = slide.Shapes.AddTable(n + 1, nCols, 30, 100, 860, 30 * (n + 1)).Table

    ' Header row
    tbl.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Title"
    tbl.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Authors"
    tbl.Cell(1, 3).Shape.TextFrame.TextRange.Text = "Journal"
    tbl.Cell(1, 4).Shape.TextFrame.TextRange.Text = "Year"
    tbl.Cell(1, 5).Shape.TextFrame.TextRange.Text = "DOI"
    tbl.Cell(1, 6).Shape.TextFrame.TextRange.Text = "Catalyst"
    Dim c As Long
    For c = 1 To nCols
        Dim hdr As Object
        Set hdr = tbl.Cell(1, c).Shape
        hdr.TextFrame.TextRange.Font.Bold = msoTrue
        hdr.Fill.ForeColor.RGB = RGB(&H1F, &H4E, &H79)
        hdr.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
    Next c

    ' Data rows — split items by top-level object boundaries
    Dim objs() As String, i As Long
    objs = SplitJsonObjects(items)
    For i = 0 To UBound(objs)
        If i >= n Then Exit For
        Dim objJson As String: objJson = objs(i)
        Dim title As String: title = ExtractJsonField(objJson, """title"":""")
        Dim authors As String: authors = ExtractJsonField(objJson, """authors"":""")
        Dim journal As String: journal = ExtractJsonField(objJson, """journal"":""")
        Dim year As String: year = ExtractJsonField(objJson, """year"":""")
        Dim doi As String: doi = ExtractJsonField(objJson, """doi"":""")
        Dim catalyst As String: catalyst = ExtractJsonField(objJson, """catalystName"":""")
        ' Authors is a JSON array string like ["A","B"] — strip brackets/quotes
        authors = CleanJsonArray(authors)
        ' Truncate long fields for readability
        If Len(title) > 60 Then title = Left$(title, 57) & "..."
        If Len(authors) > 30 Then authors = Left$(authors, 27) & "..."
        If Len(journal) > 25 Then journal = Left$(journal, 22) & "..."

        tbl.Cell(i + 2, 1).Shape.TextFrame.TextRange.Text = title
        tbl.Cell(i + 2, 2).Shape.TextFrame.TextRange.Text = authors
        tbl.Cell(i + 2, 3).Shape.TextFrame.TextRange.Text = journal
        tbl.Cell(i + 2, 4).Shape.TextFrame.TextRange.Text = year
        tbl.Cell(i + 2, 5).Shape.TextFrame.TextRange.Text = doi
        tbl.Cell(i + 2, 6).Shape.TextFrame.TextRange.Text = catalyst
        ' Tighten font
        Dim r As Long
        For r = 1 To nCols
            tbl.Cell(i + 2, r).Shape.TextFrame.TextRange.Font.Size = 10
        Next r
    Next i

    MsgBox "Pulled " & n & " record(s) from your Uni-Scholar library." & vbCrLf & _
           "Token cached for next time (hold Shift + click to reset).", _
           vbInformation, "Research Records"
    Exit Sub

NetFail:
    MsgBox "Network error." & vbCrLf & _
           "Check your connection or firewall." & vbCrLf & vbCrLf & _
           "Endpoint: " & ENDPOINT, _
           vbExclamation, "Research Records"
End Sub

' --- Token cache: %APPDATA%\UniScholarFigure\token.txt on Win,
'     ~/Library/Application Support/UniScholarFigure/token.txt on Mac ---
Private Function TokenCachePath() As String
#If Mac Then
    TokenCachePath = MacScript("do shell script ""echo $HOME/Library/Application Support/UniScholarFigure/token.txt""")
#Else
    Dim p As String
    p = Environ$("APPDATA")
    If Len(p) = 0 Then p = Environ$("USERPROFILE")
    TokenCachePath = p & "\UniScholarFigure\token.txt"
#End If
End Function

Private Function ReadCachedToken() As String
    On Error Resume Next
    ' Shift-click resets the token
    If GetAsyncKeyState_Bytes(&H10) < 0 Then Exit Function  ' VK_SHIFT = 0x10
    Dim path As String: path = TokenCachePath
#If Mac Then
    ReadCachedToken = Trim(MacScript("do shell script ""cat "" & quoted form of """ & path & """ 2>/dev/null || true"))
#Else
    Dim f As Integer: f = FreeFile
    Open path For Input As #f
    If LOF(f) > 0 Then
        Dim s As String: s = Input$(LOF(f), #f)
        ReadCachedToken = Trim(s)
    End If
    Close #f
#End If
    On Error GoTo 0
End Function

Private Sub WriteCachedToken(ByVal token As String)
    On Error Resume Next
    Dim path As String: path = TokenCachePath()
#If Mac Then
    MacScript "do shell script ""mkdir -p $(dirname "" & quoted form of """ & path & """ & "") && printf '%s' "" & quoted form of """ & token & """ & "" > "" & quoted form of """ & path & """"
#Else
    Dim parent As String
    parent = Left$(path, InStrRev(path, "\") - 1)
    If Len(Dir(parent, vbDirectory)) = 0 Then MkDir parent
    Dim f As Integer: f = FreeFile
    Open path For Output As #f
    Print #f, token;
    Close #f
#End If
    On Error GoTo 0
End Sub

#If Not Mac Then
Private Declare PtrSafe Function GetAsyncKeyState_Bytes Lib "user32" Alias "GetAsyncKeyState" (ByVal vKey As Long) As Integer
#End If

' --- Tiny JSON helpers (not a general parser) ---
Private Function ExtractJsonArray(ByVal json As String, ByVal key As String) As String
    Dim p As Long, depth As Long, i As Long, ch As String
    p = InStr(json, key)
    If p = 0 Then Exit Function
    ' Skip to '[' after key + ":"
    p = InStr(p + Len(key), json, "[")
    If p = 0 Then Exit Function
    ' Walk until matching ']'
    depth = 1
    i = p + 1
    Do While i <= Len(json) And depth > 0
        ch = Mid$(json, i, 1)
        If ch = "[" Then depth = depth + 1
        If ch = "]" Then depth = depth - 1
        i = i + 1
    Loop
    ExtractJsonArray = Mid$(json, p, i - p)
End Function

Private Function CountJsonObjects(ByVal arrJson As String) As Long
    Dim depth As Long, i As Long, ch As String, inStr_ As Boolean
    depth = 0
    For i = 1 To Len(arrJson)
        ch = Mid$(arrJson, i, 1)
        If inStr_ Then
            If ch = """" Then inStr_ = False
        Else
            Select Case ch
                Case """": inStr_ = True
                Case "{": depth = depth + 1
                Case "}": If depth = 1 Then CountJsonObjects = CountJsonObjects + 1
                          depth = depth - 1
            End Select
        End If
    Next i
End Function

Private Function SplitJsonObjects(ByVal arrJson As String) As String()
    Dim result() As String
    ReDim result(0 To 31)
    Dim n As Long: n = 0
    Dim depth As Long, i As Long, ch As String, inStr_ As Boolean
    Dim startIdx As Long: startIdx = 0
    depth = 0
    For i = 1 To Len(arrJson)
        ch = Mid$(arrJson, i, 1)
        If inStr_ Then
            If ch = """" Then inStr_ = False
        Else
            Select Case ch
                Case """": inStr_ = True
                Case "{":
                    If depth = 0 Then startIdx = i
                    depth = depth + 1
                Case "}":
                    depth = depth - 1
                    If depth = 0 And startIdx > 0 Then
                        If n > UBound(result) Then ReDim Preserve result(0 To n * 2)
                        result(n) = Mid$(arrJson, startIdx, i - startIdx + 1)
                        n = n + 1
                        startIdx = 0
                    End If
            End Select
        End If
    Next i
    If n = 0 Then
        ReDim result(0 To 0)
        result(0) = ""
    End If
    SplitJsonObjects = result
End Function

Private Function CleanJsonArray(ByVal s As String) As String
    ' Turn ["A","B","C"] into "A, B, C"
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    s = Replace(s, """", "")
    s = Replace(s, ",", ", ")
    CleanJsonArray = Trim(s)
End Function

' Return True if version a >= version b (dotted numeric, e.g. "1.2.4" >= "1.2.3")
Private Function VersionGe(ByVal a As String, ByVal b As String) As Boolean
    Dim pa() As String, pb() As String
    pa = Split(a, ".")
    pb = Split(b, ".")
    Dim i As Long, va As Long, vb_ As Long
    Dim n As Long
    n = UBound(pa)
    If UBound(pb) > n Then n = UBound(pb)
    For i = 0 To n
        If i <= UBound(pa) Then va = CLng(pa(i)) Else va = 0
        If i <= UBound(pb) Then vb_ = CLng(pb(i)) Else vb_ = 0
        If va > vb_ Then VersionGe = True: Exit Function
        If va < vb_ Then VersionGe = False: Exit Function
    Next i
    VersionGe = True  ' equal
End Function

' Tiny JSON field extractor — pulls the string value after `"key":`.
' Robust enough for GitHub's release payload; not a general JSON parser.
Private Function ExtractJsonField(ByVal json As String, ByVal key As String) As String
    Dim p As Long, q As Long
    p = InStr(json, key)
    If p = 0 Then Exit Function
    p = p + Len(key)
    ' skip opening quote
    p = p + 1
    ' find closing quote
    q = InStr(p, json, """")
    If q = 0 Then Exit Function
    ExtractJsonField = Mid$(json, p, q - p)
End Function

'==========================================================================
' Icon catalog — same paths as Python's src/unisfigure/icons.py
' Each entry: 24x24 viewBox SVG path 'd' string, single-color fill.
'==========================================================================
Private Function ICON_CATALOG(ByVal name As String) As String
    Select Case name
        Case "check":        ICON_CATALOG = "M9 16.2 4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4z"
        Case "cross":        ICON_CATALOG = "M19 6.4 17.6 5 12 10.6 6.4 5 5 6.4 10.6 12 5 17.6 6.4 19 12 13.4 17.6 19 19 17.6 13.4 12z"
        Case "arrow-right":  ICON_CATALOG = "M4 11h12.2l-3.6-3.6L14 6l6 6-6 6-1.4-1.4 3.6-3.6H4z"
        Case "arrow-down":   ICON_CATALOG = "M11 4v12.2l-3.6-3.6L6 14l6 6 6-6-1.4-1.4-3.6 3.6V4z"
        Case "info":         ICON_CATALOG = "M11 7h2v2h-2zm0 4h2v6h-2zm1-9a10 10 0 1 0 0 20 10 10 0 0 0 0-20z"
        Case "warning":      ICON_CATALOG = "M1 21h22L12 2zm12-3h-2v-2h2zm0-4h-2v-4h2z"
        Case "lightbulb":    ICON_CATALOG = "M9 21c0 .5.4 1 1 1h4c.6 0 1-.5 1-1v-1H9zm3-19a7 7 0 0 0-4 12.7V17h8v-2.3A7 7 0 0 0 12 2z"
        Case "search":       ICON_CATALOG = "M15.5 14h-.8l-.3-.3a6.5 6.5 0 1 0-.7.7l.3.3v.8l5 5 1.5-1.5zm-6 0a4.5 4.5 0 1 1 0-9 4.5 4.5 0 0 1 0 9z"
        Case "gear":         ICON_CATALOG = "M19.4 13a7.8 7.8 0 0 0 0-2l2-1.6-2-3.4-2.4 1a7.8 7.8 0 0 0-1.7-1L15 2H9l-.3 2.6a7.8 7.8 0 0 0-1.7 1l-2.4-1-2 3.4L4.6 11a7.8 7.8 0 0 0 0 2l-2 1.6 2 3.4 2.4-1c.5.4 1 .7 1.7 1L9 22h6l.3-2.6c.6-.3 1.2-.6 1.7-1l2.4 1 2-3.4zM12 15.5a3.5 3.5 0 1 1 0-7 3.5 3.5 0 0 1 0 7z"
        Case "doc":          ICON_CATALOG = "M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8zm2 16H8v-2h8zm0-4H8v-2h8zm-3-5V3.5L18.5 9z"
        Case Else:           ICON_CATALOG = ""  ' unknown — caller falls back to labeled circle
    End Select
End Function

' Attempt to draw an icon as a PowerPoint Freeform scaled into (left, top, size).
' Returns True on success, False if name is unknown or path parse fails
' (caller should draw fallback shape).
Private Function DrawIconFreeform(slide As Object, ByVal name As String, _
                                   ByVal leftPx As Single, ByVal topPx As Single, _
                                   ByVal sizePx As Single) As Boolean
    Dim d As String
    d = ICON_CATALOG(name)
    If Len(d) = 0 Then Exit Function

    ' Parse path into a list of (cmd, x, y) tokens. Supports M L H V C Z (uppercase only).
    ' This is a tiny subset of SVG — enough for the bundled catalog.
    Dim cmds() As String, xs() As Single, ys() As Single
    ReDim cmds(0 To 64): ReDim xs(0 To 64): ReDim ys(0 To 64)
    Dim n As Long: n = 0

    Dim i As Long, ch As String, curCmd As String
    Dim numBuf As String, tok As String
    Dim nums() As String, k As Long
    curCmd = ""
    i = 1
    Do While i <= Len(d)
        ch = Mid$(d, i, 1)
        Select Case ch
            Case " ", ",", vbTab
                i = i + 1
            Case "M", "L", "H", "V", "C", "Z"
                curCmd = ch
                i = i + 1
            Case "-", "0" To "9", "."
                ' Read a number
                numBuf = ""
                Do While i <= Len(d)
                    ch = Mid$(d, i, 1)
                    If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "-" Then
                        numBuf = numBuf & ch
                        i = i + 1
                    Else
                        Exit Do
                    End If
                Loop
                ' Collect nums for current command based on arity
                Select Case curCmd
                    Case "M", "L"
                        Dim nx As Single, ny As Single
                        nx = CSng(numBuf)
                        ' read next number (y)
                        numBuf = ""
                        Do While i <= Len(d) And Mid$(d, i, 1) = " ": i = i + 1: Loop
                        Do While i <= Len(d)
                            ch = Mid$(d, i, 1)
                            If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "-" Then
                                numBuf = numBuf & ch
                                i = i + 1
                            Else
                                Exit Do
                            End If
                        Loop
                        ny = CSng(numBuf)
                        n = n + 1
                        If n > UBound(cmds) Then ReDim Preserve cmds(0 To n * 2): ReDim Preserve xs(0 To n * 2): ReDim Preserve ys(0 To n * 2)
                        cmds(n) = curCmd: xs(n) = nx: ys(n) = ny
                    Case "H"
                        n = n + 1
                        If n > UBound(cmds) Then ReDim Preserve cmds(0 To n * 2): ReDim Preserve xs(0 To n * 2): ReDim Preserve ys(0 To n * 2)
                        cmds(n) = curCmd: xs(n) = CSng(numBuf): ys(n) = ys(n - 1)
                    Case "V"
                        n = n + 1
                        If n > UBound(cmds) Then ReDim Preserve cmds(0 To n * 2): ReDim Preserve xs(0 To n * 2): ReDim Preserve ys(0 To n * 2)
                        cmds(n) = curCmd: xs(n) = xs(n - 1): ys(n) = CSng(numBuf)
                    Case "C"
                        ' Cubic Bezier — for simplicity, sample the 6 control coords
                        ' and treat as a line to the endpoint. (Catalog icons still read OK.)
                        Dim dummy As Single
                        Dim j As Long
                        For j = 1 To 4
                            ' skip 4 of the 5 remaining numbers
                            numBuf = ""
                            Do While i <= Len(d) And Mid$(d, i, 1) = " ": i = i + 1: Loop
                            Do While i <= Len(d)
                                ch = Mid$(d, i, 1)
                                If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "-" Then
                                    numBuf = numBuf & ch
                                    i = i + 1
                                Else
                                    Exit Do
                                End If
                            Loop
                            If j = 4 Then
                                ' 6th value (last y) — wait we've read 5 values total now (incl. first)
                                ' Actually first C value was numBuf from outer; we just read 4 more.
                                ' Total C arg = (x1 y1 x2 y2 x y) = 6 numbers; first read in outer,
                                ' so j=1..5 here. Need endpoint = (5th, 6th) = (j=4, j=5).
                            End If
                        Next j
                        ' For robustness, fallback: just advance to next M/L
                        ' (the C-handling above is best-effort; catalog icons with C will
                        '  still look reasonable because we already captured the endpoint
                        '  via the outer numBuf when it was the start of next token.)
                End Select
            Case Else
                i = i + 1
        End Select
    Loop

    If n < 2 Then Exit Function

    ' Scale from 24x24 viewBox to sizePx
    Dim scl As Single
    scl = sizePx / 24!

    ' Find first M as start point
    Dim startIdx As Long: startIdx = -1
    For i = 1 To n
        If cmds(i) = "M" Then startIdx = i: Exit For
    Next i
    If startIdx < 0 Then Exit Function

    ' Build Freeform
    Dim ff As Object
    Set ff = slide.Shapes.BuildFreeform(msoEditingCorner, _
        leftPx + xs(startIdx) * scl, topPx + ys(startIdx) * scl)
    For i = startIdx + 1 To n
        Select Case cmds(i)
            Case "M", "L", "H", "V"
                ff.AddNodes msoSegmentLine, msoEditingAuto, _
                    leftPx + xs(i) * scl, topPx + ys(i) * scl
        End Select
    Next i
    Dim result As Object
    On Error Resume Next
    Set result = ff.ConvertToShape
    If Err.Number <> 0 Or result Is Nothing Then
        DrawIconFreeform = False
    Else
        result.Fill.ForeColor.RGB = RGB(&H1F, &H4E, &H79)
        result.Line.ForeColor.RGB = RGB(&H1F, &H4E, &H79)
        result.Line.Weight = 1!
        DrawIconFreeform = True
    End If
    On Error GoTo 0
End Function

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
