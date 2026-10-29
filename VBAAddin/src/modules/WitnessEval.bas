Option Explicit

'========================  WitnessEvalUpdated.bas  ========================
' This module imports witness CSV/XLS files and builds charts and tables
' summarizing juror responses.  It has been updated to:
'  � Use the witness evaluation title slide (TPL_TITLE_WITNESS_EVAL) instead of the post-case title.
'  � Write category rows in reverse order so categories appear A�Z from top to bottom.
'  � Define ExtractFirstLastName to parse first/last names properly (resolving missing function error).
'--------------------------------------------------------------------------

Private Const xlToLeft As Long = -4159
Private Const xlUp     As Long = -4162

' Chart templates
Private Const TPL_BAR_X3_WIT As String = _
    "%APPDATA%\Microsoft\AddIns\TQ FG Slides\TQ FG Barchart x3.pptx"
Private Const TPL_TITLE_WITNESS_EVAL As String = _
    "%APPDATA%\Microsoft\AddIns\TQ FG Slides\TQ FG TITLE - WITNESS EVAL.pptx"

'=============================  ENTRY  =============================
Public Sub ImportWitness3x_FromColumn(control As IRibbonControl)
    ' Ensure AutoFocusMaster is available
    EnsureAutoFocusMaster ActivePresentation

    Dim files As Collection: Set files = PickManyFiles("Select witness CSV/XLS(X) files")
    If files Is Nothing Or files.Count = 0 Then Exit Sub

    ' Open all witnesses
    Dim W As New Collection, i As Long
    For i = 1 To files.Count
        Dim d As Object: Set d = OpenWitnessDict(CStr(files(i)))
        If Not d Is Nothing Then W.Add d
    Next
    If W.Count = 0 Then GoTo Tidy

    ' Ask where to start (column letter or number)
    Dim ws1 As Object: Set ws1 = W(1)("ws")
    Dim guessC As Long: guessC = GuessLikertStartCol(ws1)
    Dim defTxt As String: defTxt = ColNumToLetter(IIf(guessC > 0, guessC, 1))
    Dim ans As String
    ans = InputBox( _
        "Enter the Excel column (letter or number) where the first" & vbCrLf & _
        "numeric-prefixed question begins (e.g., W or 23).", _
        "Witness Eval � Start Column", defTxt)
    If Trim$(ans) = "" Then GoTo Tidy

    Dim startCol As Long: startCol = ColRefToNum(ans)
    If startCol <= 0 Then
        MsgBox "Couldn't parse that column. Please try again.", vbExclamation
        GoTo Tidy
    End If
    
    ' Ask user for first-name and last-name columns
    Dim colFirstInput As String, colLastInput As String
    Dim colFirst As Long, colLast As Long
    Dim tmpCol As Long
    colFirst = 21
    colLast = 22
    colFirstInput = InputBox("Enter first name column (letter like U or number 21):", _
                             "First Name Column", "U")
    colFirstInput = Trim$(colFirstInput)
    If Len(colFirstInput) > 0 Then
        If IsNumeric(colFirstInput) Then
            If CLng(colFirstInput) >= 1 Then colFirst = CLng(colFirstInput)
        Else
            On Error Resume Next
            tmpCol = ws1.Range(UCase$(colFirstInput) & "1").Column
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
            tmpCol = ws1.Range(UCase$(colLastInput) & "1").Column
            If Err.Number = 0 And tmpCol > 0 Then colLast = tmpCol
            Err.Clear
            On Error GoTo 0
        End If
    End If

    ' Insert witness evaluation title slide
    Call CopyTemplateSlide(TPL_TITLE_WITNESS_EVAL, ActivePresentation, "Witness Evaluations")

    ' Walk columns on the first witness
    Dim lastC As Long: lastC = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    Dim c As Long, qText As String, inTables As Boolean

    For c = startCol To lastC
        qText = CleanSymbols(CStr(ws1.Cells(1, c).Text))
        If qText = "" Then GoTo NextC

        ' stay in CHART mode while any witness has 1-/2-/3-/4- prefix;
        ' once a column doesn�t, switch to TABLES for the rest.
        If Not inTables Then
            If AnyWitnessHasLikertPrefix(W, c) Then
                BuildQuestionAs3x_ByCol ActivePresentation, W, c, qText, colFirst, colLast
            Else
                inTables = True
                BuildOpenEndedTables_ByCol ActivePresentation, W, c, qText, colFirst, colLast
            End If
        Else
            BuildOpenEndedTables_ByCol ActivePresentation, W, c, qText, colFirst, colLast
        End If
NextC:
    Next c

Tidy:
    ' Close Excel hosts
    On Error Resume Next
    For i = 1 To W.Count
        If Not W(i)("wb") Is Nothing Then W(i)("wb").Close False
        If Not W(i)("xApp") Is Nothing Then W(i)("xApp").Quit
    Next
    On Error GoTo 0

    MsgBox "Witness evaluation slides imported.", vbInformation
End Sub

'=============================  FILE I/O  ==============================
Private Function PickManyFiles(prompt As String) As Collection
    On Error GoTo fallback
    Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = prompt
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "CSV or Excel", "*.csv;*.xlsx;*.xlsm;*.xls"
        If .show = 0 Then Exit Function
        Dim c As New Collection, i As Long
        For i = 1 To .SelectedItems.Count: c.Add .SelectedItems(i): Next
        Set PickManyFiles = c
    End With
    Exit Function
fallback:
    Dim p As String: p = InputBox("Path to a CSV/XLSX (blank = cancel):", prompt, "")
    If Len(p) = 0 Then Exit Function
    Dim c2 As New Collection: c2.Add p: Set PickManyFiles = c2
End Function

Private Function OpenWitnessDict(pathX As String) As Object
    On Error GoTo fail
    Dim xApp As Object: Set xApp = CreateObject("Excel.Application"): xApp.Visible = False
    Dim wb As Object: Set wb = xApp.Workbooks.Open(pathX, ReadOnly:=True)
    Dim ws As Object: Set ws = wb.Worksheets(1)

    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("Path") = pathX
    d("Name") = CleanSymbols(StripLeadingIndex(GetFileStem(pathX)))
    Set d("xApp") = xApp: Set d("wb") = wb: Set d("ws") = ws
    Set OpenWitnessDict = d
    Exit Function
fail:
    Set OpenWitnessDict = Nothing
End Function

Private Function GetFileStem(fullPath As String) As String
    Dim nm As String: nm = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
    If InStr(nm, ".") > 0 Then nm = Left$(nm, InStrRev(nm, ".") - 1)
    GetFileStem = nm
End Function

Private Function StripLeadingIndex(ByVal s As String) As String
    ' e.g., "W3-Rodrigo Ferreira" -> "Rodrigo Ferreira"
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    With re: .Pattern = "^\s*[A-Za-z]?\d+[\-\._ ]+": .IgnoreCase = True: End With
    If re.Test(s) Then StripLeadingIndex = re.Replace(s, "") Else StripLeadingIndex = s
End Function

'==========================  COLUMN HELPERS  ==========================

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

Private Function ColNumToLetter(colNum As Long) As String
    Dim n As Long: n = colNum
    Do While n > 0
        ColNumToLetter = Chr$(((n - 1) Mod 26) + 65) & ColNumToLetter
        n = (n - 1) \ 26
    Loop
End Function

Private Function GuessLikertStartCol(ByVal ws As Object) As Long
    Dim lastC As Long: lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastC
        If ColumnHasLikertPrefix(ws, c) Then GuessLikertStartCol = c: Exit Function
    Next
End Function

Private Function AnyWitnessHasLikertPrefix(ByVal W As Collection, ByVal c As Long) As Boolean
    Dim i As Long, ws As Object
    For i = 1 To W.Count
        Set ws = W(i)("ws")
        If ColumnHasLikertPrefix(ws, c) Then AnyWitnessHasLikertPrefix = True: Exit Function
    Next
End Function

Private Function ColumnHasLikertPrefix(ByVal ws As Object, ByVal c As Long) As Boolean
    On Error GoTo nope
    Dim lastR As Long: lastR = ws.Cells(ws.rows.Count, c).End(xlUp).Row
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    With re: .Pattern = "^\s*[1-4][\.\-\)]\s*": .IgnoreCase = True: End With
    Dim r As Long, v As String
    For r = 2 To lastR
        v = Trim$(CStr(ws.Cells(r, c).Text))
        If v <> "" Then If re.Test(v) Then ColumnHasLikertPrefix = True: Exit Function
    Next
nope:
End Function

'==========================  BUILD � CHARTS  ==========================

Private Sub BuildQuestionAs3x_ByCol(pres As Presentation, ByVal W As Collection, _
                                    ByVal col As Long, ByVal qText As String, _
                                    Optional colFirst As Long = 21, _
                                    Optional colLast As Long = 22)
    If W.Count = 0 Then Exit Sub

    Dim startAt As Long: startAt = 1
    Do While startAt <= W.Count
        Dim take As Long: take = IIf(startAt + 2 <= W.Count, 3, W.Count - startAt + 1)

        ' --- union of categories + per-witness tallies ---
        Dim unionU As Object: Set unionU = CreateObject("Scripting.Dictionary")
        Dim dicts() As Object, totals() As Long
        ReDim dicts(1 To take): ReDim totals(1 To take)

        Dim j As Long, pos As Long
        For j = 1 To take
            pos = startAt + j - 1
            Set dicts(j) = CreateObject("Scripting.Dictionary")
            totals(j) = TallyOneByCol(W(pos)("ws"), col, dicts(j))
            Dim k As Variant
            For Each k In dicts(j).keys
                If Not unionU.Exists(CStr(k)) Then unionU.Add CStr(k), True
            Next k
        Next

        ' --- category order + pad zeros ---
        Dim orderedU As Variant
        orderedU = DetectScaleFromKeys(unionU.keys)
        If IsEmpty(orderedU) Then orderedU = OrderKeysWitness(unionU.keys)
        PadAllMissing dicts, orderedU

        ' --- spawn template and write workbook ---
        Dim tplPath As String: tplPath = Replace(TPL_BAR_X3_WIT, "%APPDATA%", Environ$("APPDATA"))
        If Dir$(tplPath) = "" Then Exit Sub

        Dim tpl As Presentation: Set tpl = Presentations.Open(tplPath, msoFalse)
        tpl.Slides(1).Copy: tpl.Close

        pres.Slides.Paste
        Dim sld As Slide: Set sld = pres.Slides(pres.Slides.Count)
        sld.Shapes("TargetTitle").TextFrame.TextRange.Text = CleanSymbols(qText)

        Dim ch As Chart: Set ch = FindChartOnSlide(sld)
        Dim wsC As Object: Set wsC = GetChartDataWorksheet(ch)
        wsC.Cells.Clear

        ' Header row and series names (First + Last)
        wsC.Cells(1, 1).value = "Category"
        Dim displayNames() As String: ReDim displayNames(1 To take)
        Dim cCol As Long: cCol = 2
        For j = 1 To take
            pos = startAt + j - 1
            displayNames(j) = ExtractFirstLastName(W(pos)("Name"))
            wsC.Cells(1, cCol).value = displayNames(j)
            cCol = cCol + 1
        Next

        ' Rows (category labels + % values)
        Dim r As Long: r = 2
        Dim ii As Long
        Dim keyUp As String
        Dim v As Long
        If ShouldWriteForward(orderedU) Then
            For ii = LBound(orderedU) To UBound(orderedU)
                keyUp = CStr(orderedU(ii))
                wsC.Cells(r, 1).value = ProperWitnessLabel(keyUp)
                cCol = 2
                For j = 1 To take
                    v = 0
                    If dicts(j).Exists(keyUp) Then v = CLng(dicts(j)(keyUp))
                    wsC.Cells(r, cCol).value = IIf(totals(j) > 0, v / totals(j), 0)
                    cCol = cCol + 1
                Next
                r = r + 1
            Next ii
        Else
            For ii = UBound(orderedU) To LBound(orderedU) Step -1
                keyUp = CStr(orderedU(ii))
                wsC.Cells(r, 1).value = ProperWitnessLabel(keyUp)
                cCol = 2
                For j = 1 To take
                    v = 0
                    If dicts(j).Exists(keyUp) Then v = CLng(dicts(j)(keyUp))
                    wsC.Cells(r, cCol).value = IIf(totals(j) > 0, v / totals(j), 0)
                    cCol = cCol + 1
                Next
                r = r + 1
            Next ii
        End If

        ' --- ensure the embedded table fits the data (marquee fix) ---
        Dim lastRow As Long: lastRow = r - 1
        Dim lastCol As Long: lastCol = 1 + take

        On Error Resume Next
        wsC.ListObjects(1).Resize wsC.Range(wsC.Cells(1, 1), wsC.Cells(lastRow, lastCol))
        On Error GoTo 0

        ' --- rebuild series so the chart points at the range we just wrote ---
        Dim wsNm As String
        wsNm = "'" & wsC.name & "'"          ' quote sheet name in case it has spaces

        With ch
            Do While .SeriesCollection.Count > 0
                .SeriesCollection(1).Delete
            Loop
            Dim addrX As String, addrY As String
            ' A-column contains categories (row 2..lastRow)
            addrX = "=" & wsNm & "!$A$2:$A$" & lastRow
            For j = 1 To take
                Dim s As Series
                Set s = .SeriesCollection.NewSeries
                s.name = displayNames(j)
                ' B/C/D... contain series values starting at row 2
                addrY = "=" & wsNm & "!$" & ColNumToLetter(1 + j) & "$2:$" & ColNumToLetter(1 + j) & "$" & lastRow
                s.XValues = addrX
                s.Values = addrY
            Next j
        End With

        wsC.Parent.Close False
        ch.Refresh

        ' --- set chart legend and Data1/2/3 header badges, hide unused groups ---
        On Error Resume Next
        For j = 1 To take
            ch.SeriesCollection(j).name = displayNames(j)
        Next j
        On Error GoTo 0
        UpdateDataHeaderBoxes sld, displayNames, take   ' hides Data2/Data3 groups when unused

        ' payload for SwapView
        Dim payload As String
        payload = BuildSeriesPayload(orderedU, dicts, totals, displayNames)
        StashPayload sld, payload, "CHART"

        startAt = startAt + take
    Loop
End Sub

' Update Data1/2/3Text and Data1/2/3Box; hide whole Data# group when unused
Private Sub UpdateDataHeaderBoxes(ByVal sld As Slide, _
                                  ByRef displayNames() As String, _
                                  ByVal take As Long)
    Dim i As Long, t As String
    For i = 1 To 3
        If i <= take And i >= LBound(displayNames) And i <= UBound(displayNames) Then
            t = displayNames(i)
            SetGroupVisible sld, "Data" & i, True
            SetNamedVisible sld, "Data" & i & "Text", True, "Data" & i
            SetNamedVisible sld, "Data" & i & "Box", True, "Data" & i
        Else
            t = ""
            SetGroupVisible sld, "Data" & i, False
            SetNamedVisible sld, "Data" & i & "Text", False, "Data" & i
            SetNamedVisible sld, "Data" & i & "Box", False, "Data" & i
        End If
        SetNamedText sld, "Data" & i & "Text", t, "Data" & i
    Next i
End Sub

Private Sub SetGroupVisible(ByVal sld As Slide, ByVal groupName As String, ByVal show As Boolean)
    On Error Resume Next
    Dim grp As Shape: Set grp = sld.Shapes(groupName)
    If Not grp Is Nothing Then grp.Visible = IIf(show, msoTrue, msoFalse)
    On Error GoTo 0
End Sub

' Looks for a shape named childName at slide root; if not found, looks inside group groupName.
Private Sub SetNamedText(ByVal sld As Slide, ByVal childName As String, _
                         ByVal value As String, ByVal groupName As String)
    Dim shp As Shape
    ' Try slide root
    On Error Resume Next
    Set shp = sld.Shapes(childName)
    On Error GoTo 0
    If Not shp Is Nothing Then
        shp.TextFrame.TextRange.Text = value
        Exit Sub
    End If
    ' Try inside the group
    Dim grp As Shape
    On Error Resume Next
    Set grp = sld.Shapes(groupName)
    On Error GoTo 0
    If Not grp Is Nothing And grp.Type = msoGroup Then
        Dim i As Long
        For i = 1 To grp.GroupItems.Count
            If StrComp(grp.GroupItems(i).name, childName, vbTextCompare) = 0 Then
                grp.GroupItems(i).TextFrame.TextRange.Text = value
                Exit Sub
            End If
        Next i
    End If
End Sub

' Tally one column; dictionary keys are UPPER labels. Returns N.
Private Function TallyOneByCol(ByVal ws As Object, ByVal c As Long, ByVal d As Object) As Long
    On Error GoTo done
    Dim lastR As Long: lastR = ws.Cells(ws.rows.Count, c).End(xlUp).Row
    Dim r As Long, resp As String, ku As String, n As Long
    For r = 2 To lastR
        resp = CleanResp(CleanSymbols(CStr(ws.Cells(r, c).Text)))
        If resp <> "" Then
            ku = UCase$(resp)
            If Not d.Exists(ku) Then d.Add ku, 1 Else d(ku) = d(ku) + 1
            n = n + 1
        End If
    Next
done:
    TallyOneByCol = n
End Function

Private Function BuildSeriesPayload(ByVal orderedU As Variant, _
                                    ByRef dicts() As Object, _
                                    ByRef totals() As Long, _
                                    ByRef fullNames() As String) As String
    Dim buf As String, ii As Long, j As Long
    Dim keyUp As String
    Dim v As Long
    If ShouldWriteForward(orderedU) Then
        For ii = LBound(orderedU) To UBound(orderedU)
            keyUp = CStr(orderedU(ii))
            buf = buf & ProperWitnessLabel(keyUp)
            For j = LBound(fullNames) To UBound(fullNames)
                v = 0
                If dicts(j).Exists(keyUp) Then v = CLng(dicts(j)(keyUp))
                buf = buf & vbTab & IIf(totals(j) > 0, Round(v / totals(j), 3), 0)
            Next
            buf = buf & vbCrLf
        Next
    Else
        For ii = UBound(orderedU) To LBound(orderedU) Step -1
            keyUp = CStr(orderedU(ii))
            buf = buf & ProperWitnessLabel(keyUp)
            For j = LBound(fullNames) To UBound(fullNames)
                v = 0
                If dicts(j).Exists(keyUp) Then v = CLng(dicts(j)(keyUp))
                buf = buf & vbTab & IIf(totals(j) > 0, Round(v / totals(j), 3), 0)
            Next
            buf = buf & vbCrLf
        Next
    End If
    If Len(buf) >= 2 And Right$(buf, 2) = vbCrLf Then buf = Left$(buf, Len(buf) - 2)
    BuildSeriesPayload = buf
End Function

'========================  BUILD � TABLES  ========================

Private Sub BuildOpenEndedTables_ByCol(pres As Presentation, ByVal W As Collection, _
                                       ByVal col As Long, ByVal qText As String, _
                                       Optional colFirst As Long = 21, _
                                       Optional colLast As Long = 22)
    Dim i As Long
    For i = 1 To W.Count
        BuildWitnessTable_ByCol pres, W(i), col, qText, colFirst, colLast
    Next
End Sub

Private Sub BuildWitnessTable_ByCol(pres As Presentation, ByVal d As Object, _
                                    ByVal col As Long, ByVal qText As String, _
                                    Optional colFirst As Long = 21, _
                                    Optional colLast As Long = 22)
    Dim ws As Object: Set ws = d("ws")
    Dim lastR As Long: lastR = ws.Cells(ws.rows.Count, col).End(xlUp).Row
    If lastR < 2 Then Exit Sub
    ' Collect responses
    Dim tmp() As Variant, n As Long, r As Long
    ReDim tmp(1 To lastR - 1, 1 To 2)
    For r = 2 To lastR
        Dim resp As String: resp = CleanSymbols(CStr(ws.Cells(r, col).Text))
        If Trim$(resp) <> "" Then
            n = n + 1
            tmp(n, 1) = MakeJuror(ws, r, colFirst, colLast)
            tmp(n, 2) = resp
        End If
    Next
    If n = 0 Then Exit Sub
    ReDim Preserve tmp(1 To n, 1 To 2)
    SortPairsByJuror tmp
    ' NEW: split question into title (left of dash) + bar text (right of dash)
    Dim baseTitle As String, nameFromQ As String, haveName As Boolean
    haveName = ParseTitleAndBarText(qText, baseTitle, nameFromQ)
    ' Fallback if no trailing name exists
    If Not haveName Then
        nameFromQ = ExtractFirstLastName(d("Name"))
    End If
    ' First page
    Dim sld As Slide
    Set sld = CopyTemplateSlide(TPL_TABLE, pres, CleanSymbols(baseTitle))
    On Error Resume Next
        sld.Shapes("TargetBarText").TextFrame.TextRange.Text = nameFromQ
        sld.Shapes("SampleSize").TextFrame.TextRange.Text = "N=" & n
    On Error GoTo 0
    Dim tblShp As Shape: Set tblShp = sld.Shapes("TargetTable")
    Dim tbl As Table:    Set tbl = tblShp.Table
    ClearTableData tbl
    ' Reduce number of rows per page to avoid overlapping logo (approx 20% reduction)
    Const MAX_ROWS As Long = 14
    Dim bottom As Single: bottom = pres.PageSetup.SlideHeight - 0.75 * 72
    Dim rowCount As Long, payload As String
    ' Process sorted responses in ascending chunks of MAX_ROWS
    Dim pageStart As Long: pageStart = 1
    Do While pageStart <= n
        Dim pageEnd As Long
        pageEnd = pageStart + MAX_ROWS - 1
        If pageEnd > n Then pageEnd = n
        ' For pages after the first, spawn a new slide
        If pageStart > 1 Then
            tbl.rows(tbl.rows.Count).Delete
            StashPayload sld, RTrimCRLF(payload), "TABLE"
            payload = "": rowCount = 0
            Set sld = CopyTemplateSlide(TPL_TABLE, pres, CleanSymbols(baseTitle))
            On Error Resume Next
                sld.Shapes("TargetBarText").TextFrame.TextRange.Text = nameFromQ
                sld.Shapes("SampleSize").TextFrame.TextRange.Text = "N=" & n
            On Error GoTo 0
            Set tblShp = sld.Shapes("TargetTable")
            Set tbl = tblShp.Table
            ClearTableData tbl
        End If
        ' Insert rows in reverse within this page so they appear A?Z top-to-bottom
        Dim ri As Long
        For ri = pageEnd To pageStart Step -1
            tbl.rows.Add 2
            tbl.cell(2, 1).Shape.TextFrame.TextRange.Text = tmp(ri, 1)
            tbl.cell(2, 2).Shape.TextFrame.TextRange.Text = tmp(ri, 2)
            payload = payload & tmp(ri, 1) & vbTab & tmp(ri, 2) & vbCrLf
            rowCount = rowCount + 1
        Next ri
        pageStart = pageEnd + 1
    Loop
    tbl.rows(tbl.rows.Count).Delete
    StashPayload sld, RTrimCRLF(payload), "TABLE"
End Sub

Private Function FindHeaderLike(ws As Object, ByVal candidates As Variant, ByVal fallbackCol As Long) As Long
    Dim lastC As Long: lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, h As String, i As Long
    For c = 1 To lastC
        h = LCase$(Trim$(CStr(ws.Cells(1, c).Text)))
        For i = LBound(candidates) To UBound(candidates)
            If InStr(h, CStr(candidates(i))) > 0 Then FindHeaderLike = c: Exit Function
        Next
    Next
    FindHeaderLike = fallbackCol
End Function

Private Function MakeJuror(ws As Object, ByVal r As Long, ByVal colFirst As Long, ByVal colLast As Long) As String
    Dim f As String: f = CleanSymbols(Trim$(CStr(ws.Cells(r, colFirst).Text)))
    Dim l As String: l = CleanSymbols(Trim$(CStr(ws.Cells(r, colLast).Text)))
    If f <> "" Then
        MakeJuror = f & IIf(l <> "", " " & Left$(l, 1) & ".", "")
    Else
        MakeJuror = l
    End If
End Function

Private Sub SortPairsByJuror(ByRef pairs() As Variant)
    On Error GoTo done
    Dim i As Long, j As Long, a As Variant, b As Variant
    For i = LBound(pairs, 1) To UBound(pairs, 1) - 1
        For j = i + 1 To UBound(pairs, 1)
            If StrComp(CStr(pairs(i, 1)), CStr(pairs(j, 1)), vbTextCompare) > 0 Then
                a = pairs(i, 1): b = pairs(i, 2)
                pairs(i, 1) = pairs(j, 1): pairs(i, 2) = pairs(j, 2)
                pairs(j, 1) = a: pairs(j, 2) = b
            End If
        Next
    Next
done:
End Sub

Private Function RTrimCRLF(ByVal s As String) As String
    If Len(s) >= 2 And Right$(s, 2) = vbCrLf Then RTrimCRLF = Left$(s, Len(s) - 2) Else RTrimCRLF = s
End Function

'====================  CATEGORY ORDERING & LABELS  ====================

Private Function CanonicalGroups() As Variant
    CanonicalGroups = Array( _
        Array("Very Negative", "Somewhat Negative", "Somewhat Positive", "Very Positive"), _
        Array("Very Unbelievable", "Somewhat Unbelievable", "Somewhat Believable", "Very Believable"), _
        Array("Very Evasive", "Somewhat Evasive", "A Little Bit Evasive", "Not At All Evasive"), _
        Array("Not At All Honest", "A Little Bit Honest", "Somewhat Honest", "Very Honest"), _
        Array("Very Unlikable", "Somewhat Unlikable", "Somewhat Likable", "Very Likable"), _
        Array("Mostly Helped The Plaintiff", "Somewhat Helped The Plaintiff", "Somewhat Helped The Defendant", "Mostly Helped The Defendant"), _
        Array("Very Unlikely", "Unlikely", "Likely", "Very Likely"), _
        Array("Way Too Low", "Too Low", "About Right", "Too High", "Way Too High"), _
        Array("Strongly Disagree", "Somewhat Disagree", "Somewhat Agree", "Strongly Agree"), _
        Array("Yes", "No") _
    )
End Function

Private Function DetectScaleFromKeys(ByVal keysU As Variant) As Variant
    If IsEmpty(keysU) Then Exit Function
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(keysU) To UBound(keysU)
        d(CStr(keysU(i))) = True
    Next i
    Dim groups As Variant
    groups = CanonicalGroups()
    Dim g As Long, j As Long, k As Long
    For g = LBound(groups) To UBound(groups)
        For j = LBound(groups(g)) To UBound(groups(g))
            If d.Exists(UCase$(CStr(groups(g)(j)))) Then
                Dim out() As String
                ReDim out(LBound(groups(g)) To UBound(groups(g)))
                For k = LBound(groups(g)) To UBound(groups(g))
                    out(k) = UCase$(CStr(groups(g)(k)))
                Next k
                DetectScaleFromKeys = out
                Exit Function
            End If
        Next j
    Next g
End Function

Private Sub PadAllMissing(ByRef dicts() As Object, ByVal orderedU As Variant)
    Dim j As Long, ii As Long
    For j = LBound(dicts) To UBound(dicts)
        For ii = LBound(orderedU) To UBound(orderedU)
            If Not dicts(j).Exists(CStr(orderedU(ii))) Then dicts(j)(CStr(orderedU(ii))) = 0
        Next
    Next
End Sub

Private Function OrderKeysWitness(ByVal keys As Variant) As Variant
    If IsEmpty(keys) Then
        OrderKeysWitness = Array("N/A")
        Exit Function
    End If
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        d(UCase$(CStr(keys(i)))) = True
    Next i
    If d.Count = 0 Then
        OrderKeysWitness = Array("N/A")
        Exit Function
    End If
    ' numeric?
    Dim allNum As Boolean
    allNum = True
    Dim k As Variant
    For Each k In d.keys
        If Not IsNumeric(k) Then
            allNum = False
            Exit For
        End If
    Next k
    If allNum Then
        Dim nums() As Double, j As Long
        ReDim nums(0 To d.Count - 1)
        j = 0
        For Each k In d.keys
            nums(j) = CDbl(k)
            j = j + 1
        Next k
        QuickSortD nums, 0, UBound(nums)
        Dim outN() As String
        ReDim outN(0 To UBound(nums))
        For j = LBound(nums) To UBound(nums)
            outN(j) = CStr(nums(j))
        Next j
        OrderKeysWitness = outN
        Exit Function
    End If
    ' alphabetic fallback
    Dim out() As String
    ReDim out(0 To d.Count - 1)
    i = 0
    For Each k In d.keys
        out(i) = CStr(k)
        i = i + 1
    Next k
    BubbleSort out
    OrderKeysWitness = out
End Function
Private Function ProperWitnessLabel(ByVal keyUp As String) As String
    Dim groups As Variant
    groups = CanonicalGroups()

    Dim i As Long, j As Long
    For i = LBound(groups) To UBound(groups)
        For j = LBound(groups(i)) To UBound(groups(i))
            If UCase$(CStr(groups(i)(j))) = UCase$(keyUp) Then
                ProperWitnessLabel = CStr(groups(i)(j))
                Exit Function
            End If
        Next j
    Next i

    If IsNumeric(keyUp) Then
        ProperWitnessLabel = keyUp
    Else
        ProperWitnessLabel = StrConv(LCase$(keyUp), vbProperCase)
    End If
End Function


'==========================  SORT HELPERS  ==========================

Private Sub QuickSortD(a() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Double, t As Double
    i = lo: j = hi: p = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < p: i = i + 1: Loop
        Do While a(j) > p: j = j - 1: Loop
        If i <= j Then
            t = a(i): a(i) = a(j): a(j) = t: i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortD a, lo, j
    If i < hi Then QuickSortD a, i, hi
End Sub

Private Sub BubbleSort(a() As String)
    Dim i As Long, j As Long, t As String
    For i = LBound(a) To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If a(i) > a(j) Then
                t = a(i): a(i) = a(j): a(j) = t
            End If
        Next
    Next
End Sub

'==========================  NAME HELPERS  ==========================

Private Function ExtractLastName(ByVal fullName As String) As String
    Dim s As String: s = Trim$(fullName)
    If s = "" Then Exit Function
    s = Replace$(s, "  ", " ")
    Dim parts() As String: parts = Split(s, " ")
    Dim skip As Object: Set skip = CreateObject("Scripting.Dictionary")
    skip.CompareMode = 1
    skip("JR") = True: skip("JR.") = True
    skip("SR") = True: skip("SR.") = True
    skip("II") = True: skip("III") = True: skip("IV") = True: skip("V") = True
    Dim i As Long, tok As String
    For i = UBound(parts) To LBound(parts) Step -1
        tok = UCase$(Replace$(parts(i), ".", ""))
        If Not skip.Exists(tok) Then
            ExtractLastName = parts(i)
            Exit Function
        End If
    Next i
    ExtractLastName = parts(UBound(parts))
End Function

' ExtractFirstLastName: return "First Last" (skipping honorifics and suffixes)
Private Function ExtractFirstLastName(ByVal fullName As String) As String
    Dim s As String: s = Trim$(CleanSymbols(fullName))
    If s = "" Then Exit Function
    ' compress spaces
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    Dim parts() As String: parts = Split(s, " ")
    ' skip honorifics
    Dim skipFront As Object: Set skipFront = CreateObject("Scripting.Dictionary")
    skipFront.CompareMode = 1
    skipFront("MR") = True: skipFront("MR.") = True
    skipFront("MRS") = True: skipFront("MRS.") = True
    skipFront("MS") = True: skipFront("MS.") = True
    skipFront("MISS") = True
    skipFront("DR") = True: skipFront("DR.") = True
    skipFront("PROF") = True: skipFront("PROF.") = True
    Dim firstTok As String, lastTok As String
    Dim i As Long, tok As String
    For i = LBound(parts) To UBound(parts)
        tok = UCase$(Replace$(parts(i), ".", ""))
        If Not skipFront.Exists(tok) Then
            firstTok = parts(i)
            Exit For
        End If
    Next i
    lastTok = ExtractLastName(s)
    If firstTok = "" Then
        ExtractFirstLastName = lastTok
    ElseIf lastTok = "" Or StrComp(firstTok, lastTok, vbTextCompare) = 0 Then
        ExtractFirstLastName = firstTok
    Else
        ExtractFirstLastName = firstTok & " " & lastTok
    End If
End Function

Private Sub SetNamedVisible(ByVal sld As Slide, ByVal childName As String, _
                            ByVal show As Boolean, Optional ByVal groupName As String = "")
    Dim shp As Shape, grp As Shape, i As Long
    ' Try slide root
    On Error Resume Next
    Set shp = sld.Shapes(childName)
    On Error GoTo 0
    If Not shp Is Nothing Then
        On Error Resume Next
        shp.Visible = IIf(show, msoTrue, msoFalse)
        On Error GoTo 0
        Exit Sub
    End If
    ' Try inside the specified group
    If Len(groupName) > 0 Then
        On Error Resume Next
        Set grp = sld.Shapes(groupName)
        On Error GoTo 0
        If Not grp Is Nothing And grp.Type = msoGroup Then
            For i = 1 To grp.GroupItems.Count
                If StrComp(grp.GroupItems(i).name, childName, vbTextCompare) = 0 Then
                    On Error Resume Next
                    grp.GroupItems(i).Visible = IIf(show, msoTrue, msoFalse)
                    On Error GoTo 0
                    Exit Sub
                End If
            Next i
        End If
    End If
End Sub

Private Function ParseTitleAndBarText(ByVal qText As String, _
                                      ByRef baseTitle As String, _
                                      ByRef barText As String) As Boolean
    Dim s As String, sn As String, p As Long
    s = CleanSymbols(qText)
    sn = Replace(Replace(Replace(s, "�", "-"), "�", "-"), vbTab, " ")
    Do While InStr(sn, "  ") > 0: sn = Replace(sn, "  ", " "): Loop
    p = InStrRev(sn, " - ")
    If p > 0 Then
        baseTitle = Trim$(Left$(sn, p - 1))
        barText = Trim$(Mid$(sn, p + 3))
    Else
        p = InStrRev(sn, "-")
        If p > 0 Then
            baseTitle = Trim$(Left$(sn, p - 1))
            barText = Trim$(Mid$(sn, p + 1))
        End If
    End If
    If Len(barText) > 0 And Len(baseTitle) > 0 Then
        ParseTitleAndBarText = True
    Else
        baseTitle = s
        barText = ""
    End If
End Function

