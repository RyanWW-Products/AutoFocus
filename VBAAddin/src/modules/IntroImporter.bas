'
Option Explicit

'======================  IntroDataImporter.bas  ======================
' This module imports pre-stimulation data from an Excel workbook
' and builds a series of chart or table slides in the active
' PowerPoint presentation.  It normalizes text (fixes curly quotes
' and other encoding issues), maps P-codes to long labels and enforces
' that order on charts, and stashes juror-response payloads for
' later retrieval by the SwapView add-in.  Table rows are sorted
' alphabetically by juror name when used.

' Build PRE-STIM slides (CHARTS by default) using ModuleSlideHelpers.
' This version:
'  • Normalizes text so curly quotes/dashes and UTF-8/1252 mojibake render correctly.
'  • Builds CHART slides for all questions; TABLE builders kept as optional.
'  • Maps P+, P, D+, D ? long labels and enforces that order on charts.
'  • Stashes Juror<TAB>Response payload so SwapView can rebuild either view.
'  • Sorts TABLE rows A-Z by juror name when used.

' Excel layout / interop constants
Private Const FIRST_QUESTION_COL As Long = 32   ' AF
Private Const COL_FIRST          As Long = 21   ' U (first name)
Private Const COL_LAST           As Long = 22   ' V (last name)
Private Const xlToLeft           As Long = -4159
Private Const xlUp               As Long = -4162

'========================
' Entry point
'========================
Public Sub ImportIntroData(control As IRibbonControl)
    ' Ensure AutoFocusMaster is loaded only once
    Dim presTmp As Presentation
    Set presTmp = ActivePresentation
    Dim autoFocusPresent As Boolean: autoFocusPresent = False
    ' Check existing designs for a master name containing "AutoFocusMaster"
    Dim d As Design
    For Each d In presTmp.Designs
        Dim smn As String
        On Error Resume Next
        smn = d.SlideMaster.name
        On Error GoTo 0
        If Len(smn) > 0 Then
            If InStr(1, smn, "AutoFocusMaster", vbTextCompare) > 0 Then
                autoFocusPresent = True
                Exit For
            End If
        End If
    Next d
    If Not autoFocusPresent Then
        ' Only load the AutoFocusMaster design if it hasn't been loaded yet
        EnsureAutoFocusMaster presTmp
    End If

    Dim fd As FileDialog, pathXLS As String
    Dim xlApp As Object, wb As Object, ws As Object
    Dim pres As Presentation
    Dim lastCol As Long, c As Long
    Dim startCol As Long, startColInput As String, tmpCol As Long
    Dim colType As String, qText As String

    ' Pick Excel file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Select Pre-Stim Excel File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xlsm;*.xls"
        If .show = 0 Then Exit Sub
        pathXLS = .SelectedItems(1)
    End With

    ' Open workbook (hidden Excel)
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set wb = xlApp.Workbooks.Open(pathXLS, ReadOnly:=True)
    Set ws = wb.Worksheets(1)

    ' Active deck
    Set pres = ActivePresentation

    ' Respondent roster slides (disabled per new requirements)
    ' The respondent roster has been removed; skip adding roster slides

    ' Title slide
    CopyTemplateSlide TPL_TITLE_PRESTIM, pres, "PRE-STIM Questionnaire"

    ' Ask user which column to start at (letter like AF or number); default to AF
    startCol = FIRST_QUESTION_COL
    startColInput = InputBox("Enter starting column (letter like AF or number):", _
                             "Pre-Stim Start Column", "AF")
    startColInput = Trim$(startColInput)
    If Len(startColInput) > 0 Then
        If IsNumeric(startColInput) Then
            If CLng(startColInput) >= 1 Then startCol = CLng(startColInput)
        Else
            On Error Resume Next
            tmpCol = ws.Range(UCase$(startColInput) & "1").Column
            If Err.Number = 0 And tmpCol > 0 Then startCol = tmpCol
            Err.Clear
            On Error GoTo 0
        End If
    End If
    
    ' Ask user for first-name and last-name columns
    Dim colFirstInput As String, colLastInput As String
    Dim colFirst As Long, colLast As Long
    colFirst = COL_FIRST
    colLast = COL_LAST
    colFirstInput = InputBox("Enter first name column (letter like U or number 21):", _
                             "First Name Column", "U")
    colFirstInput = Trim$(colFirstInput)
    If Len(colFirstInput) > 0 Then
        If IsNumeric(colFirstInput) Then
            If CLng(colFirstInput) >= 1 Then colFirst = CLng(colFirstInput)
        Else
            On Error Resume Next
            tmpCol = ws.Range(UCase$(colFirstInput) & "1").Column
            If Err.Number = 0 And tmpCol > 0 Then colFirst = tmpCol
            Err.Clear
            On Error GoTo 0
        End If
    End If
    colLastInput = InputBox("Enter last name column (letter like V or number 22):", _
                            "Last Name Column", "V")
    colLastInput = Trim$(colLastInput)
    If Len(colLastInput) > 0 Then
        If IsNumeric(colLastInput) Then
            If CLng(colLastInput) >= 1 Then colLast = CLng(colLastInput)
        Else
            On Error Resume Next
            tmpCol = ws.Range(UCase$(colLastInput) & "1").Column
            If Err.Number = 0 And tmpCol > 0 Then colLast = tmpCol
            Err.Clear
            On Error GoTo 0
        End If
    End If

    ' Iterate questions
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    For c = startCol To lastCol
        colType = UCase$(Trim$(SafeTxt(ws.Cells(1, c).Text)))
        If colType = "" Then colType = GuessColumnType(ws, c)
        If colType <> "TABLE" And colType <> "CHART" Then GoTo NextCol

        qText = ReadQuestionTitle(ws, c)
        If UCase$(Trim$(qText)) = "TABLE" Or UCase$(Trim$(qText)) = "CHART" Then GoTo NextCol
        If qText = "" Then GoTo NextCol

        ' Build based on Row 1 type (TABLE or CHART)
        If colType = "TABLE" Then
            BuildIntroTableSlides pres, ws, c, qText, colFirst, colLast
        Else
            BuildIntroChartSlide pres, ws, c, qText, colFirst, colLast
        End If
NextCol:
    Next c

    ' After inserting all slides, remove duplicate AutoFocusMaster designs
    Call RemoveExtraAutoFocusMasters(pres)

    wb.Close False
    xlApp.Quit
    MsgBox "Pre-stim slides imported!", vbInformation
End Sub

'==============================================================
' OPTIONAL TABLE BUILDERS (kept for parity; A-Z juror sort)
'==============================================================
Private Sub BuildIntroTableSlides( _
            pres As Presentation, _
            ws As Object, _
            dataCol As Long, _
            qText As String, _
            Optional colFirst As Long = 21, _
            Optional colLast As Long = 22)

    Dim lastRow As Long, r As Long, payload As String
    lastRow = ws.Cells(ws.rows.Count, dataCol).End(xlUp).Row

    For r = 3 To lastRow
        Dim resp As String
        resp = CleanResp(SafeTxt(CategoryString(ws.Cells(r, dataCol))))
        If resp <> "" Then
            payload = payload & resp & vbLf
        End If
    Next r

    If Right$(payload, 1) = vbLf Then
        payload = Left$(payload, Len(payload) - 1)
    End If

    Dim sld As Slide
    Set sld = CopyTemplateSlide(TPL_TABLE, pres, SafeTxt(qText))
    On Error Resume Next
        sld.Shapes("TargetTitle").TextFrame.TextRange.Text = SafeTxt(qText)
        Debug.Print "[Intro][TABLE][first] set:", SafeTxt(qText), _
                    " got:", sld.Shapes("TargetTitle").TextFrame.TextRange.Text
    On Error GoTo 0
    On Error Resume Next
        sld.Shapes("SampleSize").TextFrame.TextRange.Text = _
            "N=" & IIf(payload = "", 0, UBound(Split(payload, vbLf)) + 1)
    On Error GoTo 0
    StashPayload sld, payload, "TABLE"

    BuildIntroTablePages sld, ws, dataCol, colFirst, colLast
End Sub

Private Sub BuildIntroTablePages( _
        firstSld As Slide, _
        ws As Object, _
        dataCol As Long, _
        Optional colFirst As Long = 21, _
        Optional colLast As Long = 22)

    ' Dynamically paginate tables so that long responses do not cause the table to run off the slide.
    Const MAX_ROWS As Long = MAX_ROWS_PER_TABLE
    Const BOTTOM_MARGIN As Single = 0.75 * 72

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, dataCol).End(xlUp).Row
    If lastRow < 3 Then Exit Sub

    ' Collect all response rows and their original row numbers
    Dim rowList() As Long, rowCnt As Long, r As Long
    ReDim rowList(1 To lastRow - 2)
    For r = 3 To lastRow
        Dim cellText As String
        cellText = Trim$(CategoryString(ws.Cells(r, dataCol)))
        If cellText <> "" Then
            rowCnt = rowCnt + 1
            rowList(rowCnt) = r
        End If
    Next r
    If rowCnt = 0 Then Exit Sub
    ReDim Preserve rowList(1 To rowCnt)

    ' Sort jurors alphabetically
    Dim i As Long, j As Long, tmpRow As Long
    For i = 1 To rowCnt - 1
        For j = i + 1 To rowCnt
            If LCase$(JurorLabel(ws, rowList(i), colFirst, colLast)) > LCase$(JurorLabel(ws, rowList(j), colFirst, colLast)) Then
                tmpRow = rowList(i)
                rowList(i) = rowList(j)
                rowList(j) = tmpRow
            End If
        Next j
    Next i

    On Error Resume Next
      firstSld.Shapes("SampleSize").TextFrame.TextRange.Text = "N=" & rowCnt
    On Error GoTo 0

    Dim pres As Presentation
    Set pres = firstSld.Parent
    Dim slideBottom As Single: slideBottom = pres.PageSetup.SlideHeight - BOTTOM_MARGIN

    Dim sld As Slide: Set sld = firstSld
    Dim tblShp As Shape: Set tblShp = sld.Shapes("TargetTable")
    Dim tbl As Table: Set tbl = tblShp.Table
    ClearTableData tbl

    Dim rowCount As Long: rowCount = 0
    Dim payload As String: payload = ""

    For i = 1 To rowCnt
        r = rowList(i)
        Dim juror As String: juror = JurorLabel(ws, r, colFirst, colLast)
        Dim resp As String: resp = CleanResp(SafeTxt(CategoryString(ws.Cells(r, dataCol))))
        ' Insert a new row just above the stub row
        tbl.rows.Add rowCount + 2
        tbl.cell(rowCount + 2, 1).Shape.TextFrame.TextRange.Text = juror
        tbl.cell(rowCount + 2, 2).Shape.TextFrame.TextRange.Text = resp
        payload = payload & juror & vbTab & resp & vbCrLf
        rowCount = rowCount + 1
        ' Check if we've exceeded the maximum rows or overflowed the slide height
        If rowCount >= MAX_ROWS _
           Or (tblShp.Top + tblShp.Height + tbl.rows(1).Height) > slideBottom Then
            ' Remove the newly inserted row and roll back counters
            tbl.rows(rowCount + 1).Delete
            rowCount = rowCount - 1
            ' Remove the last appended juror/response from payload (length = juror + tab + resp + CRLF)
            If Len(payload) >= (Len(juror) + Len(resp) + 4) Then
                payload = Left$(payload, Len(payload) - (Len(juror) + Len(resp) + 4))
            End If
            ' Remove stub row, stash payload and start a new slide
            If tbl.rows.Count > 2 Then tbl.rows(tbl.rows.Count).Delete
            If Right$(payload, 2) = vbCrLf Then payload = Left$(payload, Len(payload) - 2)
            StashPayload sld, payload, "TABLE"
            ' Create new slide
            Set sld = InsertTemplateAfter(TPL_TABLE, sld)
            sld.Shapes("TargetTitle").TextFrame.TextRange.Text = _
                firstSld.Shapes("TargetTitle").TextFrame.TextRange.Text
            Debug.Print "[Intro][TABLE][cont] got:", _
                sld.Shapes("TargetTitle").TextFrame.TextRange.Text
            On Error Resume Next
              sld.Shapes("SampleSize").TextFrame.TextRange.Text = "N=" & rowCnt
            On Error GoTo 0
            Set tblShp = sld.Shapes("TargetTable")
            Set tbl = tblShp.Table
            ClearTableData tbl
            rowCount = 0
            payload = ""
            ' Re-add this row on the new slide
            tbl.rows.Add rowCount + 2
            tbl.cell(rowCount + 2, 1).Shape.TextFrame.TextRange.Text = juror
            tbl.cell(rowCount + 2, 2).Shape.TextFrame.TextRange.Text = resp
            payload = payload & juror & vbTab & resp & vbCrLf
            rowCount = 1
        End If
    Next i
    ' Final slide cleanup
    If tbl.rows.Count > 2 Then tbl.rows(tbl.rows.Count).Delete
    If Right$(payload, 2) = vbCrLf Then payload = Left$(payload, Len(payload) - 2)
    StashPayload sld, payload, "TABLE"
End Sub

'==============================================================
' CHARTS (default for every question)
'==============================================================
Private Sub BuildIntroChartSlide( _
            pres As Presentation, _
            ws As Object, _
            dataCol As Long, _
            qText As String, _
            Optional colFirst As Long = 21, _
            Optional colLast As Long = 22)

    Dim freq As Object:        Set freq = CreateObject("Scripting.Dictionary")
    Dim originalMap As Object: Set originalMap = CreateObject("Scripting.Dictionary")
    Dim payload As String
    Dim total As Long

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.rows.Count, dataCol).End(xlUp).Row

    For r = 3 To lastRow
        Dim resp As String
        resp = CleanResp(SafeTxt(CategoryString(ws.Cells(r, dataCol))))
        If resp <> "" Then
            Dim juror As String
            juror = JurorLabel(ws, r, colFirst, colLast)

            ' Map P-codes to long labels for keying; keep original in payload map
            Dim keyUp As String
            keyUp = UCase$(MapPCodeToLong(resp))

            If Not freq.Exists(keyUp) Then
                freq.Add keyUp, 1
                originalMap.Add keyUp, resp
            Else
                freq(keyUp) = freq(keyUp) + 1
            End If

            payload = payload & juror & vbTab & resp & vbCrLf
            total = total + 1
        End If
    Next r

    If total = 0 Then Exit Sub

    ' Groups for padding/sorting
    ' Groups for padding/sorting: pulled from shared helper module
    Dim pLongGrp As Variant: pLongGrp = PLongGroup()
    Dim likelyGrp As Variant: likelyGrp = LikelyGroup()
    Dim highGrp As Variant: highGrp = HighGroup()
    Dim agreeGrp As Variant: agreeGrp = AgreeGroup()
    Dim keysSorted As Variant
    ' First check for P-coded long labels (P+, P, D+, D): these override all other ordering
    If DictContainsAny(freq, pLongGrp) Then
        PadMissing freq, originalMap, pLongGrp
        keysSorted = PresentKeysInOrder(freq, pLongGrp)
    Else
        ' Pad built-in scales if present
        If DictContainsAny(freq, likelyGrp) Then PadMissing freq, originalMap, likelyGrp
        If DictContainsAny(freq, highGrp) Then PadMissing freq, originalMap, highGrp
        If DictContainsAny(freq, agreeGrp) Then PadMissing freq, originalMap, agreeGrp
        If freq.Exists("YES") Or freq.Exists("NO") Then
            PadMissing freq, originalMap, Array("Yes", "No")
        End If
        ' Try to match any user-defined answer group and enforce its order
        Dim matched As Boolean
        matched = False
        Dim gName As Variant, gArr As Variant, v As Variant
        Dim buf() As String
        Dim added As Boolean
        Dim ku As String
        On Error Resume Next
        Call InitializeAnswerGroups
        On Error GoTo 0
        If Not AnswerGroups Is Nothing Then
            For Each gName In AnswerGroups.keys
                Select Case UCase$(CStr(gName))
                Case "LIKELY", "HIGH", "AGREE", "PLONG", "PSHORT"
                        ' skip built-in groups
                    Case Else
                        gArr = AnswerGroups(gName)
                        ' count how many items from this custom group exist
                        Dim matchCount As Long
                        matchCount = 0
                        For Each v In gArr
                            If freq.Exists(UCase$(CStr(v))) Then
                                matchCount = matchCount + 1
                            End If
                        Next v
                        ' determine whether this group should apply
                        Dim groupSize As Long
                        groupSize = UBound(gArr) - LBound(gArr) + 1
                        Dim shouldApply As Boolean
                        If groupSize <= 2 Then
                            ' for two-item groups like Yes/No, require at least one match
                            shouldApply = (matchCount >= 1)
                        Else
                            ' for larger groups, require at least two matches
                            shouldApply = (matchCount >= 2)
                        End If
                        If shouldApply Then
                            ' Pad missing items from this group
                            PadMissing freq, originalMap, gArr
                            ' Build sorted keys based on custom group order
                            ReDim buf(0)
                            added = False
                            For Each v In gArr
                                ku = UCase$(CStr(v))
                                If freq.Exists(ku) Then
                                    If Not added Then
                                        buf(0) = ku
                                        added = True
                                    Else
                                        ReDim Preserve buf(UBound(buf) + 1)
                                        buf(UBound(buf)) = ku
                                    End If
                                End If
                            Next v
                            If added Then
                                ' Reverse the collected keys so that the insertion loop (which iterates in reverse) will
                                ' output the group in the correct order (e.g. All The Time, Sometimes, Rarely, Never)
                                Dim rev() As String
                                Dim idx As Long
                                ReDim rev(0 To UBound(buf))
                                For idx = 0 To UBound(buf)
                                    rev(idx) = buf(UBound(buf) - idx)
                                Next idx
                                keysSorted = rev
                                matched = True
                                Exit For
                            End If
                        End If
                End Select
            Next gName
        End If
        ' Fallback to generic sorter if no user-defined group matched
        If Not matched Then
            keysSorted = SortKeysCustom(freq.keys)
        End If
    End If

    Dim sld As Slide
    Set sld = CopyTemplateSlide(TPL_CHART, pres, SafeTxt(qText))
    On Error Resume Next
        sld.Shapes("TargetTitle").TextFrame.TextRange.Text = SafeTxt(qText)
        Debug.Print "[Intro][CHART][first] set:", SafeTxt(qText), _
                    " got:", sld.Shapes("TargetTitle").TextFrame.TextRange.Text
    On Error GoTo 0
    On Error Resume Next
        sld.Shapes("SampleSize").TextFrame.TextRange.Text = "N=" & total
    On Error GoTo 0

    Dim ch As Chart
    Set ch = FindChartOnSlide(sld)

    Dim wsC As Object
    Set wsC = GetChartDataWorksheet(ch)
    If wsC Is Nothing Then Exit Sub

    wsC.Cells.Clear
    wsC.Range("A1").value = "Category"
    wsC.Range("B1").value = "Pct"

    ' Fill categories and percentages. Numeric/mostly-numeric sets should be written forward
    ' so categories appear left-to-right ascending; otherwise keep reverse (A->Z top-to-bottom).
    Dim i As Long, destRow As Long
    Dim catLabel As String
    destRow = 2
    If ShouldWriteForward(keysSorted) Then
        For i = LBound(keysSorted) To UBound(keysSorted)
            ' Write category as text formula (prefix with apostrophe) to force text interpretation
            catLabel = ProperDisplayLabel( _
                keysSorted(i), originalMap, likelyGrp, highGrp, agreeGrp, pLongGrp)
            wsC.Cells(destRow, 1).value = "'" & catLabel
            wsC.Cells(destRow, 2).value = freq(keysSorted(i)) / total
            destRow = destRow + 1
        Next i
    Else
        For i = UBound(keysSorted) To LBound(keysSorted) Step -1
            ' Write category as text formula (prefix with apostrophe) to force text interpretation
            catLabel = ProperDisplayLabel( _
                keysSorted(i), originalMap, likelyGrp, highGrp, agreeGrp, pLongGrp)
            wsC.Cells(destRow, 1).value = "'" & catLabel
            wsC.Cells(destRow, 2).value = freq(keysSorted(i)) / total
            destRow = destRow + 1
        Next i
    End If

    ' --- fix: resize the embedded table and rebuild the single series ---
    Dim lastRowWritten As Long
    lastRowWritten = wsC.Cells(wsC.rows.Count, 1).End(xlUp).Row
    On Error Resume Next
    wsC.ListObjects(1).Resize wsC.Range(wsC.Cells(1, 1), wsC.Cells(lastRowWritten, 2))
    On Error GoTo 0

    ' Rebuild the chart series by specifying address strings rather than Range objects (avoids type mismatch)
    Dim wsNm As String
    wsNm = "'" & wsC.name & "'"
    Dim addrX As String, addrY As String
    addrX = "=" & wsNm & "!$A$2:$A$" & lastRowWritten
    addrY = "=" & wsNm & "!$B$2:$B$" & lastRowWritten

    With ch
        ' Remove existing series
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        Dim srs As Object
        Set srs = .SeriesCollection.NewSeries
        srs.XValues = addrX
        srs.Values = addrY
        On Error Resume Next
        ' Use header as series name if available
        srs.name = wsC.Cells(1, 2).value
        On Error GoTo 0
        .Refresh
        ' Force category axis to category scale (not date/time scale) to prevent bucketing
        On Error Resume Next
        .Axes(1).CategoryType = 1  ' xlCategoryScale
        On Error GoTo 0
    End With
    ' --- end fix ---

    wsC.Parent.Close False

    If Len(payload) >= 2 Then
        payload = Left$(payload, Len(payload) - 2)
    End If
    ' Final guard: ensure slide title reflects Row 2 text
    On Error Resume Next
        sld.Shapes("TargetTitle").TextFrame.TextRange.Text = SafeTxt(qText)
        Debug.Print "[Intro][CHART][final-guard] set:", SafeTxt(qText), _
                    " got:", sld.Shapes("TargetTitle").TextFrame.TextRange.Text
    On Error GoTo 0
    StashPayload sld, payload, "CHART"
End Sub

'==============================================================
' Respondent roster slides (sanitized text)
'==============================================================
Public Sub BuildRespondentSlides( _
        pres As Presentation, _
        ws As Object, _
        ByRef insertIdx As Long)

    Const MAX_TABLE_BOTTOM As Single = 6.5 * 72

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, "U").End(xlUp).Row
    If lastRow < 3 Then Exit Sub

    Dim m() As Variant
    m = Array("U", "V", "Q", "W", "Y", "Z", "AA", "AB", "AD", "AE")

    Dim tplPath As String
    tplPath = Replace(TPL_RESPONDENT, "%APPDATA%", Environ$("APPDATA"))
    If Dir(tplPath) = "" Then
        MsgBox "Respondent template not found:" & vbCrLf & tplPath, vbExclamation
        Exit Sub
    End If

    Dim rowIdx As Long: rowIdx = 3
    Do While rowIdx <= lastRow And Trim$(ws.Cells(rowIdx, "U").Text) <> ""
        pres.Slides.InsertFromFile tplPath, insertIdx, 1, 1

        Dim currSld As Slide
        Set currSld = pres.Slides(insertIdx + 1)
        insertIdx = insertIdx + 1

        Dim tblShp As Shape, tbl As Table
        On Error Resume Next
            Set tblShp = currSld.Shapes("TargetTable")
            Set tbl = tblShp.Table
        On Error GoTo 0

        If tbl Is Nothing Then
            MsgBox "TargetTable not found in respondent template.", vbExclamation
            Exit Do
        End If

        Dim rOut As Long: rOut = 2
        Do While rowIdx <= lastRow And Trim$(ws.Cells(rowIdx, "U").Text) <> ""
            If rOut > 2 Then tbl.rows.Add

            Dim c As Long, val As String
            For c = 0 To UBound(m)
                val = SafeTxt(CStr(ws.Cells(rowIdx, m(c)).value))
                tbl.cell(rOut, c + 1).Shape.TextFrame.TextRange.Text = val
            Next c

            If tblShp.Top + tblShp.Height > MAX_TABLE_BOTTOM Then
                Exit Do
            End If

            rOut = rOut + 1
            rowIdx = rowIdx + 1
        Loop
    Loop
End Sub

'==============================================================
' Local helpers (encoding-safe text + labels)
'==============================================================

' Build "First L." (or last name if no first); sanitized
Private Function JurorLabel(ws As Object, r As Long, Optional colFirst As Long = 21, Optional colLast As Long = 22) As String
    Dim fn As String: fn = SafeTxt(ws.Cells(r, colFirst).Text)
    Dim ln As String: ln = SafeTxt(ws.Cells(r, colLast).Text)

    If fn <> "" Then
        If ln <> "" Then
            JurorLabel = fn & " " & Left$(ln, 1) & "."
        Else
            JurorLabel = fn
        End If
    Else
        JurorLabel = ln
    End If
End Function

' Strong sanitizer: fix UTF-8/1252 mojibake + normalize punctuation
Private Function SafeTxt(ByVal s As String) As String
    s = NzStr(s)

    ' Common UTF-8 decoded as 1252 sequences ? proper Unicode
    s = Replace(s, "Ã¢â‚¬â„¢", ChrW(&H2019)) ' ’
    s = Replace(s, "Ã¢â‚¬Ëœ", ChrW(&H2018)) ' ‘
    s = Replace(s, "Ã¢â‚¬Å“", ChrW(&H201C)) ' “
    s = Replace(s, "Ã¢â‚¬Â", ChrW(&H201D)) ' ”
    s = Replace(s, "Ã¢â‚¬â€œ", ChrW(&H2013)) ' –
    s = Replace(s, "Ã¢â‚¬â€?", ChrW(&H2014)) ' —
    s = Replace(s, "Ã¢â‚¬Â¦", ChrW(&H2026)) ' …

    ' Single-byte misreads sometimes seen
    s = Replace(s, "â€™", ChrW(&H2019))
    s = Replace(s, "â€˜", ChrW(&H2018))
    s = Replace(s, "â€œ", ChrW(&H201C))
    s = Replace(s, "â€", ChrW(&H201D))
    s = Replace(s, "â€“", ChrW(&H2013))
    s = Replace(s, "â€”", ChrW(&H2014))
    s = Replace(s, "â€¦", ChrW(&H2026))

    ' Normalize curly quotes/apostrophes to plain ASCII
    s = Replace(s, ChrW(&H2018), "'")
    s = Replace(s, ChrW(&H2019), "'")
    s = Replace(s, ChrW(&H201C), Chr$(34))
    s = Replace(s, ChrW(&H201D), Chr$(34))

    ' Normalize NBSP to space
    s = Replace(s, ChrW(&HA0), " ")

    ' Trim / collapse CR-LF oddities
    s = Replace$(s, vbCr, vbLf)
    Do While InStr(s, vbLf & vbLf) > 0
        s = Replace$(s, vbLf & vbLf, vbLf)
    Loop
    If Right$(s, 1) = vbLf Then s = Left$(s, Len(s) - 1)

    SafeTxt = Trim$(s)
End Function

Private Function NzStr(ByVal s As String) As String
    If LenB(s) = 0 Then
        NzStr = ""
    Else
        NzStr = CStr(s)
    End If
End Function

' Robustly read the visible question title for a column.
' Starts at row 2 and scans downward to find the first non-empty cell
' that is not the type marker (TABLE/CHART).
Private Function ReadQuestionTitle(ByVal ws As Object, ByVal c As Long) As String
    Dim r As Long, t As String
    For r = 2 To 6
        t = SafeTxt(ws.Cells(r, c).Text)
        If Len(t) > 0 Then
            If UCase$(Trim$(t)) <> "TABLE" And UCase$(Trim$(t)) <> "CHART" Then
                ReadQuestionTitle = t
                Exit Function
            End If
        End If
    Next r
    ReadQuestionTitle = SafeTxt(ws.Cells(2, c).Text)
End Function

' Map P-codes to long labels; leave others as-is
Private Function MapPCodeToLong(ByVal resp As String) As String
    Select Case UCase$(Trim$(resp))
        Case "P+": MapPCodeToLong = "Strongly Favor the Plaintiff"
        Case "P":  MapPCodeToLong = "Somewhat Favor the Plaintiff"
        Case "D":  MapPCodeToLong = "Somewhat Favor the Defendant"
        Case "D+": MapPCodeToLong = "Strongly Favor the Defendant"
        Case Else: MapPCodeToLong = resp
    End Select
End Function

' Present only keys that exist in d, but in given order (returns UPPER keys)
Private Function PresentKeysInOrder(ByVal d As Object, ByVal order As Variant) As Variant
    Dim out() As String, v As Variant
    ReDim out(0)
    For Each v In order
        If d.Exists(UCase$(CStr(v))) Then
            If out(0) = "" Then
                out(0) = UCase$(CStr(v))
            Else
                ReDim Preserve out(UBound(out) + 1)
                out(UBound(out)) = UCase$(CStr(v))
            End If
        End If
    Next v
    PresentKeysInOrder = out
End Function

' Any element of grp present in dictionary keys?
Private Function DictContainsAny(ByVal d As Object, ByVal grp As Variant) As Boolean
    Dim v As Variant
    For Each v In grp
        If d.Exists(UCase$(CStr(v))) Then
            DictContainsAny = True
            Exit Function
        End If
    Next v
End Function

' Display label selection (multi-line loops to avoid "Next without For")
Private Function ProperDisplayLabel( _
        ByVal keyUp As String, _
        ByVal mapOrig As Object, _
        ByVal likelyGrp As Variant, _
        ByVal highGrp As Variant, _
        ByVal agreeGrp As Variant, _
        ByVal pLongGrp As Variant) As String

    Dim v As Variant

    For Each v In pLongGrp
        If UCase$(v) = keyUp Then
            ProperDisplayLabel = v
            Exit Function
        End If
    Next v

    For Each v In likelyGrp
        If UCase$(v) = keyUp Then
            ProperDisplayLabel = v
            Exit Function
        End If
    Next v

    For Each v In highGrp
        If UCase$(v) = keyUp Then
            ProperDisplayLabel = v
            Exit Function
        End If
    Next v

    For Each v In agreeGrp
        If UCase$(v) = keyUp Then
            ProperDisplayLabel = v
            Exit Function
        End If
    Next v

    If IsNumeric(keyUp) Then
        ProperDisplayLabel = keyUp
        Exit Function
    End If

    If Not mapOrig Is Nothing Then
        If mapOrig.Exists(keyUp) Then
            ProperDisplayLabel = SafeTxt(StrConv(LCase$(mapOrig(keyUp)), vbProperCase))
            Exit Function
        End If
    End If

    ProperDisplayLabel = SafeTxt(StrConv(LCase$(keyUp), vbProperCase))
End Function

' Convert a worksheet cell value to plain text string (1:1 display).
' - Uses .Text to get the formatted display from Excel.
' - Falls back to string conversion only if needed.
' - No date/time conversion or reformatting applied.
Private Function CategoryString(ByVal cellObj As Object) As String
    If cellObj Is Nothing Then Exit Function
    Dim txt As String
    txt = CStr(cellObj.Text)
    If Len(txt) > 0 Then
        CategoryString = txt
    Else
        CategoryString = CStr(cellObj.Value2)
    End If
End Function