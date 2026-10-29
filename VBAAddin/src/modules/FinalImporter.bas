Attribute VB_Name = "FinalImporter"
''
Option Explicit

'==================  Final Data Importer v2.bas  ==================
' Imports Post-Case (Final) data, builds TABLE or CHART slides,
' paginates long tables, stashes RawPayload on every slide for SwapView,
' and sorts tables alphabetically by juror name.

'-------- constants ------------------------------------------------
Public Const xlToLeft As Long = -4159   ' Excel.XlDirection.xlToLeft
Public Const xlUp     As Long = -4162
Private Const FIRST_DATA_COL As Long = 23     ' Column W (first data column)

' If you keep template constants in ModuleSlideHelpers, reuse those:
'   TPL_TABLE, TPL_CHART, TPL_TITLE_POSTCASE
' Otherwise, define here:
'Private Const TPL_TITLE_POSTCASE As String = _
'  "%APPDATA%\Microsoft\AddIns\TQ FG Slides\TQ FG TITLE - POST CASE.pptx"

'==================================================================
' Entry point
'==================================================================
Public Sub ImportFinalData(control As IRibbonControl)
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

    '�� choose Excel file ��
    Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Select Post-Case Excel File"
        .Filters.Clear: .Filters.Add "Excel", "*.xlsx;*.xlsm;*.xls"
        If .show = 0 Then Exit Sub
    End With
    Dim xlFile As String: xlFile = fd.SelectedItems(1)

    '�� open Excel (hidden) ��
    Dim xlApp As Object: Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Dim wb As Object: Set wb = xlApp.Workbooks.Open(xlFile, ReadOnly:=True)
    Dim ws As Object: Set ws = wb.Worksheets(1)

    '╒═ prompt for first & last name columns ═╕
    Dim colFirstInput As String: colFirstInput = InputBox("First name column?", "Column for First Names", "U")
    If colFirstInput = "" Then GoTo CleanUp
    Dim colFirst As Long: colFirst = ColRefToNum(colFirstInput)
    If colFirst <= 0 Then
        MsgBox "Invalid first name column.", vbExclamation
        GoTo CleanUp
    End If

    Dim colLastInput As String: colLastInput = InputBox("Last name column?", "Column for Last Names", "V")
    If colLastInput = "" Then GoTo CleanUp
    Dim colLast As Long: colLast = ColRefToNum(colLastInput)
    If colLast <= 0 Then
        MsgBox "Invalid last name column.", vbExclamation
        GoTo CleanUp
    End If

    '╒═ current presentation ═╕
    Dim pres As Presentation: Set pres = ActivePresentation

    '╒═ Post-case title slide ═╕
    Call CopyTemplateSlide(TPL_TITLE_POSTCASE, pres, "POST-CASE Questionnaire")

    '╒═ iterate question columns ═╕
    Dim lastCol As Long: lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = FIRST_DATA_COL To lastCol
        Dim colType As String: colType = Trim$(UCase$(ws.Cells(1, c).Text))
        If colType = "" Then colType = GuessColumnType(ws, c)
        If colType <> "TABLE" And colType <> "CHART" Then GoTo NextCol

        Dim qText As String: qText = ws.Cells(2, c).Text
        If qText = "" Then GoTo NextCol

        If colType = "TABLE" Then
            BuildTableSlides pres, ws, c, qText, colFirst, colLast
        Else
            BuildChartSlide pres, ws, c, qText, colFirst, colLast
        End If
NextCol:
    Next c

    ' After inserting all slides, remove duplicate AutoFocusMaster designs
    Call RemoveExtraAutoFocusMasters(pres)

CleanUp:
    wb.Close False: xlApp.Quit
    MsgBox "Post-case slides imported!", vbInformation
End Sub

'==================================================================
' BuildTableSlides � import phase w/ pagination & alphabetical sort
'==================================================================
Public Sub BuildTableSlides( _
        pres As Presentation, _
        ws As Object, _
        dataCol As Long, _
        qText As String, _
        Optional colFirst As Long = 21, _
        Optional colLast As Long = 22)

    qText = CleanSymbols(qText)

    ' Reduce the number of rows per slide to leave extra space at the bottom
    Const MAX_ROWS_PER_SLIDE As Long = 14
    Const SLIDE_MARGIN_PT   As Single = 0.75 * 72

    Dim bottomLimit As Single
    bottomLimit = pres.PageSetup.SlideHeight - SLIDE_MARGIN_PT

    ' 1) collect non-empty responses as (Juror, Resp) pairs
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, dataCol).End(xlUp).Row
    If lastRow < 3 Then Exit Sub

    Dim pairs() As String
    Dim totalCount As Long, r As Long
    ReDim pairs(0)
    For r = 3 To lastRow
        Dim resp As String: resp = CleanSymbols(Trim$(ws.Cells(r, dataCol).Text))
        If resp <> "" Then
            Dim fName As String: fName = CleanSymbols(Trim$(ws.Cells(r, colFirst).Text))
            Dim lName As String: lName = CleanSymbols(Trim$(ws.Cells(r, colLast).Text))
            Dim juror As String
            If fName <> "" Then juror = fName & " " & Left$(lName, 1) & "." Else juror = lName

            If totalCount = 0 Then
                pairs(0) = juror & vbTab & resp
            Else
                ReDim Preserve pairs(totalCount)
                pairs(UBound(pairs)) = juror & vbTab & resp
            End If
            totalCount = totalCount + 1
        End If
    Next r
    If totalCount = 0 Then Exit Sub

    ' 2) A-Z sort by Juror
    Call SortPairsByJuror(pairs)

    ' 3) first slide
    Dim sld As Slide: Set sld = CopyTemplateSlide(TPL_TABLE, pres, qText)
    On Error Resume Next
      sld.Shapes("SampleSize").TextFrame.TextRange.Text = "N=" & totalCount
    On Error GoTo 0

    Dim tblShp As Shape: Set tblShp = sld.Shapes("TargetTable")
    Dim tbl As Table: Set tbl = tblShp.Table
    ClearTableData tbl

    Dim rowCount As Long, payload As String
    rowCount = 0: payload = ""

    ' 4) write rows dynamically, starting a new slide when the table would overflow
    Dim idx As Long
    For idx = LBound(pairs) To UBound(pairs)
        Dim parts() As String
        parts = Split(pairs(idx), vbTab, 2)
        ' Insert a new row just above the stub row (rowCount+2)
        tbl.rows.Add rowCount + 2
        tbl.cell(rowCount + 2, 1).Shape.TextFrame.TextRange.Text = parts(0)
        tbl.cell(rowCount + 2, 2).Shape.TextFrame.TextRange.Text = parts(1)
        payload = payload & parts(0) & vbTab & parts(1) & vbCrLf
        rowCount = rowCount + 1
        ' Overflow check: too many rows or table height exceeds available space
        If rowCount >= MAX_ROWS_PER_SLIDE _
           Or (tblShp.Top + tblShp.Height + tbl.rows(1).Height) > bottomLimit Then
            ' Remove the inserted row to roll back
            tbl.rows(rowCount + 1).Delete
            rowCount = rowCount - 1
            ' Trim the recently appended payload (juror\tresponse\r\n)
            If Len(payload) >= (Len(parts(0)) + Len(parts(1)) + 2 + 2) Then
                payload = Left$(payload, Len(payload) - (Len(parts(0)) + Len(parts(1)) + 4))
            End If
            ' Remove stub and stash the current slide
            If tbl.rows.Count > 2 Then tbl.rows(tbl.rows.Count).Delete
            If Right$(payload, 2) = vbCrLf Then payload = Left$(payload, Len(payload) - 2)
            StashPayload sld, payload, "TABLE"
            ' Start new slide
            rowCount = 0
            payload = ""
            Set sld = CopyTemplateSlide(TPL_TABLE, pres, qText)
            On Error Resume Next
              sld.Shapes("SampleSize").TextFrame.TextRange.Text = "N=" & totalCount
            On Error GoTo 0
            Set tblShp = sld.Shapes("TargetTable")
            Set tbl = tblShp.Table
            ClearTableData tbl
            ' Insert current row into the new slide
            tbl.rows.Add rowCount + 2
            tbl.cell(rowCount + 2, 1).Shape.TextFrame.TextRange.Text = parts(0)
            tbl.cell(rowCount + 2, 2).Shape.TextFrame.TextRange.Text = parts(1)
            payload = payload & parts(0) & vbTab & parts(1) & vbCrLf
            rowCount = 1
        End If
    Next idx
    ' Final slide cleanup and stash
    If tbl.rows.Count > 2 Then tbl.rows(tbl.rows.Count).Delete
    If Right$(payload, 2) = vbCrLf Then payload = Left$(payload, Len(payload) - 2)
    StashPayload sld, payload, "TABLE"
End Sub

'==================================================================
' BuildChartSlide � chart import w/ P-code mapping & correct order
'==================================================================
Public Sub BuildChartSlide( _
        pres As Presentation, _
        ws As Object, _
        dataCol As Long, _
        qText As String, _
        Optional colFirst As Long = 21, _
        Optional colLast As Long = 22)

    qText = CleanSymbols(qText)

    ' tally + remember original-case responses (+ payload for SwapView)
    Dim freq As Object:        Set freq = CreateObject("Scripting.Dictionary")
    Dim originalMap As Object: Set originalMap = CreateObject("Scripting.Dictionary")
    Dim payload As String, total As Long
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, dataCol).End(xlUp).Row

    Dim r As Long
    For r = 3 To lastRow
        Dim respRaw As String: respRaw = CleanSymbols(Trim$(ws.Cells(r, dataCol).Text))
        If respRaw = "" Then GoTo skipR

        ' juror as First-name L.
        Dim fRaw As String: fRaw = CleanSymbols(Trim$(ws.Cells(r, colFirst).Text))
        Dim lRaw As String: lRaw = CleanSymbols(Trim$(ws.Cells(r, colLast).Text))
        Dim juror As String
        If fRaw <> "" Then juror = fRaw & " " & Left$(lRaw, 1) & "." Else juror = lRaw

        ' P-code mapping to long label for the KEY
        Dim mapped As String: mapped = MapPCode(respRaw)          ' possibly unchanged
        Dim keyUp As String: keyUp = UCase$(mapped)

        If Not freq.Exists(keyUp) Then
            freq.Add keyUp, 1
            originalMap.Add keyUp, mapped
        Else
            freq(keyUp) = freq(keyUp) + 1
        End If

        payload = payload & juror & vbTab & respRaw & vbCrLf       ' keep original response in payload
        total = total + 1
skipR:
    Next r
    If total = 0 Then Exit Sub

    ' group arrays for padding/sorting when NOT P-coded
    Dim likelyGrp As Variant: likelyGrp = LikelyGroup()
    Dim highGrp As Variant: highGrp = HighGroup()
    Dim agreeGrp As Variant: agreeGrp = AgreeGroup()
    Dim pGrp As Variant: pGrp = PLongGroup()
    ' detect P-coded group (based on mapped, long labels)
    Dim hasP As Boolean: hasP = ContainsAnyKeys(freq, pGrp)

    ' build the slide
    Dim sld As Slide: Set sld = CopyTemplateSlide(TPL_CHART, pres, qText)
    On Error Resume Next
      sld.Shapes("SampleSize").TextFrame.TextRange.Text = "N=" & total
    On Error GoTo 0

    Dim ch As Chart: Set ch = FindChartOnSlide(sld)
    Dim wsC As Object: Set wsC = GetChartDataWorksheet(ch)
    If wsC Is Nothing Then Exit Sub
    wsC.Cells.Clear
    wsC.Range("A1").value = "Category"
    wsC.Range("B1").value = "Pct"

    Dim keysSorted As Variant
    ' Declare shared temporary variables once to avoid duplicates inside If/Else blocks
    Dim buf() As String
    Dim v As Variant
    If hasP Then
        ' write in enforced P order, but only those present
        ' Use the shared buf() declared earlier
        ReDim buf(0)
        For Each v In pGrp
            If freq.Exists(UCase$(v)) Then
                If buf(0) = "" Then
                    buf(0) = UCase$(v)
                Else
                    ReDim Preserve buf(UBound(buf) + 1)
                    buf(UBound(buf)) = UCase$(v)
                End If
            End If
        Next v
        keysSorted = buf
    Else
        ' regular scales: pad missing buckets for known groups
        If ContainsAnyKeys(freq, likelyGrp) Then PadMissing freq, originalMap, likelyGrp
        If ContainsAnyKeys(freq, highGrp) Then PadMissing freq, originalMap, highGrp
        If ContainsAnyKeys(freq, agreeGrp) Then PadMissing freq, originalMap, agreeGrp
        If freq.Exists("YES") Or freq.Exists("NO") Then _
            PadMissing freq, originalMap, Array("Yes", "No")
        ' Try to match any user-defined group (e.g., Frequency1) and build a
        ' sorted list based on that group's defined order.  Skip built-in groups.
        Dim matched As Boolean
        matched = False
        Dim gName As Variant, gArr As Variant
        ' Use the shared buf() and v variables declared earlier
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
                        ' count how many items from this group exist in the freq dictionary
                        Dim matchCount As Long
                        matchCount = 0
                        For Each v In gArr
                            If freq.Exists(UCase$(CStr(v))) Then
                                matchCount = matchCount + 1
                            End If
                        Next v
                        ' determine if this group should apply based on number of matches
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
                            ' pad any missing items from this custom group
                            PadMissing freq, originalMap, gArr
                            ' build an ordered buffer of keys according to gArr
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
                                keysSorted = buf
                                matched = True
                                Exit For
                            End If
                        End If
                End Select
            Next gName
        End If
        ' If no custom group matched, fall back to the generic sorter
        If Not matched Then
            keysSorted = SortKeysCustom(freq.keys)
        End If
    End If

    ' write chart rows in original order so categories appear in input order
    Dim rOut As Long: rOut = 2
    Dim i As Long
    For i = LBound(keysSorted) To UBound(keysSorted)
        Dim k As Variant
        k = keysSorted(i)
        Dim labelText As String
        If hasP Then
            ' label is exactly the long P label we mapped to
            labelText = ToDisplayLabelFromPKey(CStr(k))
        Else
            labelText = ProperLabelFromKey(CStr(k), originalMap, likelyGrp, highGrp, agreeGrp)
        End If
        wsC.Cells(rOut, 1).value = labelText
        wsC.Cells(rOut, 2).value = freq(CStr(k)) / total
        rOut = rOut + 1
    Next i

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
        srs.name = wsC.Cells(1, 2).value
        On Error GoTo 0
        .Refresh
    End With
    ' --- end fix ---

    wsC.Parent.Close False
    StashPayload sld, payload, "CHART"
End Sub

'==================================================================
' Helpers (local to this module)
'==================================================================

' Map P-codes ? long labels (leave others unchanged)
Private Function MapPCode(ByVal resp As String) As String
    Select Case UCase$(Trim$(resp))
        Case "P+": MapPCode = "Strongly Favor the Plaintiff"
        Case "P":  MapPCode = "Somewhat Favor the Plaintiff"
        Case "D":  MapPCode = "Somewhat Favor the Defendant"
        Case "D+": MapPCode = "Strongly Favor the Defendant"
        Case Else: MapPCode = resp
    End Select
End Function

' When we�re in the P-path, we already have uppercased keys of the long labels
Private Function ToDisplayLabelFromPKey(ByVal keyUp As String) As String
    ' just sentence case those exact long labels
    ToDisplayLabelFromPKey = StrConv(LCase$(keyUp), vbProperCase)
    ' but keep �the� lowercase
    ToDisplayLabelFromPKey = Replace(ToDisplayLabelFromPKey, " The ", " the ")
End Function

Private Function ProperLabelFromKey( _
    ByVal keyUp As String, _
    ByVal mapOrig As Object, _
    ByVal likelyGrp As Variant, _
    ByVal highGrp As Variant, _
    ByVal agreeGrp As Variant) As String

    Dim v As Variant

    ' Likely scale
    For Each v In likelyGrp
        If UCase$(CStr(v)) = keyUp Then
            ProperLabelFromKey = CStr(v)
            Exit Function
        End If
    Next v

    ' High/Low scale
    For Each v In highGrp
        If UCase$(CStr(v)) = keyUp Then
            ProperLabelFromKey = CStr(v)
            Exit Function
        End If
    Next v

    ' Agree/Disagree scale
    For Each v In agreeGrp
        If UCase$(CStr(v)) = keyUp Then
            ProperLabelFromKey = CStr(v)
            Exit Function
        End If
    Next v

    ' Numeric scale
    If IsNumeric(keyUp) Then
        ProperLabelFromKey = keyUp
        Exit Function
    End If

    ' Fallback to original-case value, then proper-case
    If Not mapOrig Is Nothing Then
        If mapOrig.Exists(keyUp) Then
            ProperLabelFromKey = StrConv(LCase$(mapOrig(keyUp)), vbProperCase)
            Exit Function
        End If
    End If

    ' Ultimate fallback � proper-case the key itself
    ProperLabelFromKey = StrConv(LCase$(keyUp), vbProperCase)
End Function

' True if any element of grp exists in dictionary keys (case-insensitive)
Private Function ContainsAnyKeys(d As Object, grp As Variant) As Boolean
    Dim v
    For Each v In grp
        If d.Exists(UCase$(CStr(v))) Then ContainsAnyKeys = True: Exit Function
    Next v
End Function

' A-Z sort by Juror on "Juror<tab>Resp" pairs
Private Sub SortPairsByJuror(ByRef pairs() As String)
    Dim i As Long, j As Long
    For i = LBound(pairs) To UBound(pairs) - 1
        For j = i + 1 To UBound(pairs)
            Dim a As String: a = LCase$(Split(pairs(i), vbTab, 2)(0))
            Dim b As String: b = LCase$(Split(pairs(j), vbTab, 2)(0))
            If a > b Then
                Dim tmp As String: tmp = pairs(i)
                pairs(i) = pairs(j)
                pairs(j) = tmp
            End If
        Next j
    Next i
End Sub

' Convert column reference (letter like "U" or number like "21") to column number
Private Function ColRefToNum(ByVal s As String) As Long
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If IsNumeric(s) Then ColRefToNum = CLng(s): Exit Function
    s = UCase$(s)
    Dim i As Long, n As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "A" Or Mid$(s, i, 1) > "Z" Then ColRefToNum = 0: Exit Function
        n = n * 26 + (Asc(Mid$(s, i, 1)) - 64)
    Next
    ColRefToNum = n
End Function



