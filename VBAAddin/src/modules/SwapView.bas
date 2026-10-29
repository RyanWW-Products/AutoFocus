Attribute VB_Name = "SwapView"
''
Option Explicit

' Swap between TABLE and CHART views on the active slide.
Public Sub SwapView(control As IRibbonControl)
    Dim curSld As Slide
    Set curSld = ActiveWindow.View.Slide
    If curSld Is Nothing Then Exit Sub
    Dim oldMode As String
    oldMode = curSld.Tags("ViewMode")
    If oldMode = "" Then oldMode = "TABLE"
    ' 1) Gather ALL payload for this question
    Dim rawAll As String
    If oldMode = "TABLE" Then
        rawAll = GatherPayload(curSld, "TABLE")
        RemoveContinuationSlides curSld
    Else
        rawAll = curSld.Shapes("RawPayload").TextFrame.TextRange.Text
        RemoveChartPages curSld
    End If
    ' 2) Split into String() array
    Dim lines() As String
    lines = PayloadLines(rawAll)
    ' 3) Clear all shapes except Title + RawPayload
    Dim i As Long
    For i = curSld.Shapes.Count To 1 Step -1
        With curSld.Shapes(i)
            If .name <> "TargetTitle" And .name <> "RawPayload" Then .Delete
        End With
    Next
    ' 4) Insert the opposite template (just to get placeholders)
    Dim tmpSld As Slide
    Dim newMode As String, tplPath As String
    If oldMode = "TABLE" Then
        tplPath = TPL_CHART: newMode = "CHART"
    Else
        tplPath = TPL_TABLE: newMode = "TABLE"
    End If
    Set tmpSld = InsertTemplateAfter(tplPath, curSld)
    ' copy all shapes over
    Dim shp As Shape
    For Each shp In tmpSld.Shapes
        If shp.name <> "TargetTitle" Then
            shp.Copy
            curSld.Shapes.Paste(1).name = shp.name
        End If
    Next
    tmpSld.Delete
    ' 5) Rebuild the new view
    If newMode = "CHART" Then
        BuildChartFromRaw curSld, lines
        StashPayload curSld, rawAll, "CHART"
    Else
        RebuildTablePages curSld, lines
    End If
    ' 6) Tag for next flip
    curSld.Tags.Add "ViewMode", newMode
End Sub

' Rebuild the TABLE view from an array of lines.  Sorts names alphabetically.
Public Sub RebuildTablePages(firstSld As Slide, ByRef lines() As String)
    Const MAX_ROWS_PER_SLIDE As Long = 18
    Const SLIDE_MARGIN_PT   As Single = 0.75 * 72
    ' sort lines by juror name
    lines = SortLinesByName(lines)
    Dim bottomLimit  As Single
    bottomLimit = firstSld.Parent.PageSetup.SlideHeight - SLIDE_MARGIN_PT
    Dim totalLines   As Long
    totalLines = UBound(lines) - LBound(lines) + 1
    If totalLines <= 0 Then Exit Sub
    Dim idx     As Long: idx = LBound(lines)
    Dim prevSld As Slide: Set prevSld = firstSld
    Do While idx <= UBound(lines)
        Dim sld      As Slide
        Dim tblShp   As Shape
        Dim tbl      As Table
        Dim rowCount As Long
        Dim payload  As String
        ' 1) reuse firstSld or insert a continuation
        If idx = LBound(lines) Then
            Set sld = firstSld
        Else
            Set sld = InsertTemplateAfter(TPL_TABLE, prevSld)
            On Error Resume Next
            sld.Shapes("TargetTitle").TextFrame.TextRange.Text = _
              firstSld.Shapes("TargetTitle").TextFrame.TextRange.Text
            On Error GoTo 0
        End If
        ' now set N= on every slide
        On Error Resume Next
        sld.Shapes("SampleSize").TextFrame.TextRange.Text = "N=" & totalLines
        On Error GoTo 0
        ' 2) clear stub and start filling…
        Set tblShp = sld.Shapes("TargetTable")
        Set tbl = tblShp.Table
        ClearTableData tbl
        rowCount = 0
        payload = ""
        ' 3) add rows until overflow…
        Do While idx <= UBound(lines)
            Dim parts() As String
            parts = Split(lines(idx), vbTab, 2)
                ' Insert each row at the end of the current data so that
                ' juror names remain in ascending order top-to-bottom.  The
                ' header is row 1 and a blank stub row is kept at the bottom
                ' (rowCount+2) to measure height.  We therefore insert a new
                ' row just before the stub row using (rowCount + 2) and write
                ' into that row.
                tbl.rows.Add rowCount + 2
                tbl.cell(rowCount + 2, 1).Shape.TextFrame.TextRange.Text = Trim$(parts(0))
                If UBound(parts) >= 1 Then
                    tbl.cell(rowCount + 2, 2).Shape.TextFrame.TextRange.Text = Trim$(parts(1))
                Else
                    tbl.cell(rowCount + 2, 2).Shape.TextFrame.TextRange.Text = ""
                End If
            payload = payload & parts(0) & vbTab & IIf(UBound(parts) >= 1, parts(1), "") & vbCrLf
            rowCount = rowCount + 1
            If rowCount >= MAX_ROWS_PER_SLIDE _
               Or (tblShp.Top + tblShp.Height + tbl.rows(1).Height) > bottomLimit Then
                Exit Do
            End If
            idx = idx + 1
        Loop
        ' 4) delete stub & stash payload
        If tbl.rows.Count > 2 Then tbl.rows(tbl.rows.Count).Delete
        If Len(payload) >= 2 Then
            StashPayload sld, Left$(payload, Len(payload) - 2), "TABLE"
        Else
            StashPayload sld, payload, "TABLE"
        End If
        ' 5) next page
        Set prevSld = sld
        idx = idx + 1
    Loop
End Sub

' BuildChartFromRaw – SwapView version, honours P codes and uses ProperLabel
Public Sub BuildChartFromRaw( _
    sld As Slide, _
    lines() As String)
    ' tally & remember original responses
    Dim freq        As Object: Set freq = CreateObject("Scripting.Dictionary")
    Dim originalMap As Object: Set originalMap = CreateObject("Scripting.Dictionary")
    Dim total       As Long
    Dim i           As Long
    For i = LBound(lines) To UBound(lines)
        Dim parts() As String: parts = Split(lines(i), vbTab, 2)
        If UBound(parts) < 1 Then GoTo nx
        Dim respRaw  As String: respRaw = CleanSymbols(Trim$(parts(1)))
        If respRaw = "" Then GoTo nx
        Dim respKey  As String: respKey = UCase$(respRaw)
        If Not freq.Exists(respKey) Then
            freq.Add respKey, 1
            originalMap.Add respKey, respRaw
        Else
            freq(respKey) = freq(respKey) + 1
        End If
        total = total + 1
nx:
    Next
    If total = 0 Then Exit Sub
    ' pad missing buckets for P codes and standard scales
    Dim pdGrp As Variant: pdGrp = Array("P+", "P", "D+", "D")
    Dim likelyGrp As Variant: likelyGrp = Array("Very Unlikely", "Unlikely", "Likely", "Very Likely")
    Dim highGrp As Variant: highGrp = Array("Way Too Low", "Too Low", "About Right", "Too High", "Way Too High")
    Dim agreeGrp As Variant: agreeGrp = Array("Strongly Disagree", "Somewhat Disagree", "Somewhat Agree", "Strongly Agree")
    If ContainsAny(freq, pdGrp) Then PadMissing freq, originalMap, pdGrp
    If ContainsAny(freq, likelyGrp) Then PadMissing freq, originalMap, likelyGrp
    If ContainsAny(freq, highGrp) Then PadMissing freq, originalMap, highGrp
    If ContainsAny(freq, agreeGrp) Then PadMissing freq, originalMap, agreeGrp
    If freq.Exists("YES") Or freq.Exists("NO") Then PadMissing freq, originalMap, Array("Yes", "No")
    ' sort keys using custom groups when applicable
    Dim keysSorted As Variant
    Dim matched As Boolean
    matched = False
    Dim gName As Variant, gArr As Variant
    Dim buf() As String
    Dim added As Boolean
    Dim ku As String
    Dim v As Variant
    ' attempt to find a user-defined answer group that applies
    On Error Resume Next
    Call InitializeAnswerGroups
    On Error GoTo 0
    If Not AnswerGroups Is Nothing Then
        For Each gName In AnswerGroups.keys
            Select Case UCase$(CStr(gName))
                Case "LIKELY", "HIGH", "AGREE", "PLONG", "PSHORT"
                    ' skip built-in groups handled earlier
                Case Else
                    gArr = AnswerGroups(gName)
                    ' count how many items from this group exist in the frequency dictionary
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
    If Not matched Then
        ' fallback: generic sorting
        Dim keysUpper As Variant: keysUpper = freq.keys
        keysSorted = SortKeysCustom(keysUpper)
    End If
    ' reuse chart
    Dim ch As Chart: Set ch = sld.Shapes("TargetChart").Chart
    ch.ChartData.Activate
    Dim wb2 As Object: Set wb2 = ch.ChartData.Workbook
    Dim wsC As Object: Set wsC = wb2.Worksheets(1)
    wsC.Cells.Clear
    wsC.Range("A1").value = "Category"
    wsC.Range("B1").value = "Pct"
    Dim j As Long
    For i = LBound(keysSorted) To UBound(keysSorted)
        wsC.Cells(i + 2, 1).value = ProperLabel(keysSorted(i), originalMap)
        wsC.Cells(i + 2, 2).value = freq(keysSorted(i)) / total
    Next
    wb2.Close False
    On Error Resume Next
      sld.Shapes("SampleSize").TextFrame.TextRange.Text = "N=" & total
    On Error GoTo 0
End Sub

' Normalize and split payload into lines array (duplicate of helper)
Private Function PayloadLines(raw As String) As String()
    Dim tmp As String           ' <- add this!
    tmp = Replace$(raw, vbCr, vbLf)
    Do While InStr(tmp, vbLf & vbLf)
        tmp = Replace$(tmp, vbLf & vbLf, vbLf)
    Loop
    If Right$(tmp, 1) = vbLf Then tmp = Left$(tmp, Len(tmp) - 1)
    PayloadLines = Split(tmp, vbLf)
End Function

' Gather payload across continuation slides of the same type
Public Function GatherPayload(cur As Slide, mode As String) As String
    Dim pres    As Presentation: Set pres = cur.Parent
    Dim title   As String:       title = cur.Shapes("TargetTitle").TextFrame.TextRange.Text
    Dim buf     As String
    Dim s       As Slide
    Dim i       As Long
    For i = cur.SlideIndex To pres.Slides.Count
        Set s = pres.Slides(i)
        If s.Shapes("TargetTitle").TextFrame.TextRange.Text <> title Then Exit For
        If s.Tags("ViewMode") = mode Then
            buf = buf & s.Shapes("RawPayload").TextFrame.TextRange.Text & vbCrLf
        End If
    Next
    If Right$(buf, 2) = vbCrLf Then buf = Left$(buf, Len(buf) - 2)
    GatherPayload = buf
End Function

Private Sub RemoveContinuationSlides(cur As Slide)
    Dim pres     As Presentation: Set pres = cur.Parent
    Dim title    As String:       title = cur.Shapes("TargetTitle").TextFrame.TextRange.Text
    Dim i        As Long:         i = cur.SlideIndex + 1
    Do While i <= pres.Slides.Count
        With pres.Slides(i)
            If .Tags("ViewMode") = "TABLE" _
               And .Shapes("TargetTitle").TextFrame.TextRange.Text = title Then
                .Delete
            Else
                Exit Do
            End If
        End With
    Loop
End Sub

Private Sub RemoveChartPages(cur As Slide)
    Dim pres     As Presentation: Set pres = cur.Parent
    Dim title    As String:       title = cur.Shapes("TargetTitle").TextFrame.TextRange.Text
    Dim i        As Long:         i = cur.SlideIndex + 1
    Do While i <= pres.Slides.Count
        With pres.Slides(i)
            If .Tags("ViewMode") = "CHART" _
               And .Shapes("TargetTitle").TextFrame.TextRange.Text = title Then
                .Delete
            Else
                Exit Do
            End If
        End With
    Loop
End Sub

Private Function InsertTemplateAfter( _
    ByVal tplPath As String, _
    ByVal afterSld As Slide _
) As Slide
    Dim pres     As Presentation: Set pres = afterSld.Parent
    Dim fullPath As String:       fullPath = Replace$(tplPath, "%APPDATA%", Environ$("APPDATA"))
    If Dir$(fullPath) = "" Then MsgBox "Missing template: " & fullPath, vbExclamation: Exit Function
    pres.Slides.InsertFromFile fullPath, afterSld.SlideIndex, 1, 1
    Set InsertTemplateAfter = pres.Slides(afterSld.SlideIndex + 1)
End Function

' Sort an array (for completeness)
Public Sub SortArray(arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
        Next
    Next
End Sub

' BuildTableFromRaw – rebuild a paginated TABLE view from lines array.  Sorts names first.
Public Sub BuildTableFromRaw( _
        tblSlide As Slide, _
        ByRef lines() As String)
    ' Sort lines by name
    lines = SortLinesByName(lines)
    Const MAX_ROWS      As Long = 18                 ' real rows per slide
    Const BOTTOM_MARGIN As Single = 0.75 * 72        ' ¾" bottom margin
    Dim pres        As Presentation: Set pres = tblSlide.Parent
    Dim slideBottom As Single:        slideBottom = pres.PageSetup.SlideHeight - BOTTOM_MARGIN
    Dim tblShp As Shape: Set tblShp = tblSlide.Shapes("TargetTable")
    Dim tbl    As Table: Set tbl = tblShp.Table
    ' start with header + stub
    ClearTableData tbl
    Dim rowCount As Long
    Dim payload As String
    Dim parts() As String
    Dim i As Long
    rowCount = 0: payload = ""
    For i = LBound(lines) To UBound(lines)
        parts = Split(lines(i), vbTab, 2)
        If UBound(parts) < 1 Then GoTo NextLine
        ' overflow check
        If rowCount >= MAX_ROWS _
           Or (tblShp.Top + tblShp.Height + tbl.rows(1).Height) > slideBottom Then
            tbl.rows(2).Delete
            StashPayload tblSlide, payload, "TABLE"
            Set tblSlide = InsertTemplateAfter(TPL_TABLE, tblSlide)
            Set tblShp = tblSlide.Shapes("TargetTable")
            Set tbl = tblShp.Table
            ClearTableData tbl
            rowCount = 0
            payload = ""
        End If
        tbl.rows.Add 2
        tbl.cell(2, 1).Shape.TextFrame.TextRange.Text = Trim$(parts(0))
        tbl.cell(2, 2).Shape.TextFrame.TextRange.Text = Trim$(parts(1))
        payload = payload & parts(0) & vbTab & parts(1) & vbCrLf
        rowCount = rowCount + 1
NextLine:
    Next i
    tbl.rows(2).Delete
    StashPayload tblSlide, payload, "TABLE"
End Sub


