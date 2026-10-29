Attribute VB_Name = "CombinedImporter"
'==================  FinalDataImporterv2.bas  ==================
' Build the 3-series comparison charts:  Plaintiff / Defense / Rebuttal (post-case)
' This version resets counts per question, pads and orders P+/P/D+/D
' categories, and converts them to descriptive labels on the chart.
Option Explicit
Const xlToLeft      As Long = -4159
Const xlUp          As Long = -4162
Const FIRST_DATA_COL As Long = 5
Const TPL_BAR_X3     As String = _
        "%APPDATA%\Microsoft\AddIns\TQ FG Slides\TQ FG Barchart x3.pptx"

Public Sub ImportCombinedData()
' Ensure AutoFocusMaster is available
EnsureAutoFocusMaster ActivePresentation

    Dim t0 As Single: t0 = Timer
    Debug.Print String(60, "=")
    Debug.Print ">>> ImportCombinedData  " & Format$(Now, "hh:nn:ss")
    ' pick Excel files
    Dim fP$, fd$, fR$
    fP = PickXLS("Select **Plaintiff** file"): If fP = "" Then Exit Sub
    fd = PickXLS("Select **Defense** file"):  If fd = "" Then Exit Sub
    fR = PickXLS("Select **Rebuttal** file"): If fR = "" Then Exit Sub
    ' open workbooks
    Dim xApp As Object: Set xApp = CreateObject("Excel.Application")
    xApp.Visible = False
    Dim wbP As Object: Set wbP = xApp.Workbooks.Open(fP, ReadOnly:=True)
    Dim wbD As Object: Set wbD = xApp.Workbooks.Open(fd, ReadOnly:=True)
    Dim wbR As Object: Set wbR = xApp.Workbooks.Open(fR, ReadOnly:=True)
    Dim wsP As Object: Set wsP = wbP.Worksheets(1)
    Dim wsD As Object: Set wsD = wbD.Worksheets(1)
    Dim wsR As Object: Set wsR = wbR.Worksheets(1)
    ' set up PowerPoint
    Dim pres As Presentation:         Set pres = ActivePresentation
    Dim insertAt As Long:             insertAt = pres.Slides.Count + 1
    Dim slidesAdded As Long
    Dim lastCol As Long
    lastCol = wsP.Cells(2, wsP.Columns.Count).End(xlToLeft).Column
    ' groups for padding
    Dim likelyGrp, highGrp, agreeGrp, pdGrp
    likelyGrp = Array("Very Unlikely", "Unlikely", "Likely", "Very Likely")
    highGrp = Array("Way Too Low", "Too Low", "About Right", _
                      "Too High", "Way Too High")
    agreeGrp = Array("Strongly Disagree", "Somewhat Disagree", _
                      "Somewhat Agree", "Strongly Agree")
    pdGrp = Array("P+", "P", "D+", "D")
    ' iterate columns
    Dim col As Long
    For col = FIRST_DATA_COL To lastCol
        Dim hdr As String: hdr = Trim$(UCase$(wsP.Cells(1, col).Text))
        If hdr <> "CHART" Then GoTo NextCol
        Dim qText As String: qText = wsP.Cells(2, col).Text
        If qText = "" Then GoTo NextCol
        Dim dP As Object, dD As Object, dR As Object
        Set dP = CreateObject("Scripting.Dictionary")
        Set dD = CreateObject("Scripting.Dictionary")
        Set dR = CreateObject("Scripting.Dictionary")
        Dim nP As Long, nD As Long, nR As Long
        nP = 0: nD = 0: nR = 0
        TallyColumn wsP, col, dP, nP
        TallyColumn wsD, col, dD, nD
        TallyColumn wsR, col, dR, nR
        If nP + nD + nR = 0 Then GoTo NextCol
        ' pad P codes first
        If ContainsAny(dP, pdGrp) Or ContainsAny(dD, pdGrp) Or ContainsAny(dR, pdGrp) Then
            PadMissing dP, Nothing, pdGrp
            PadMissing dD, Nothing, pdGrp
            PadMissing dR, Nothing, pdGrp
        End If
        ' other scales
        If ContainsAny(dP, likelyGrp) Or ContainsAny(dD, likelyGrp) _
           Or ContainsAny(dR, likelyGrp) Then
            PadMissing dP, Nothing, likelyGrp
            PadMissing dD, Nothing, likelyGrp
            PadMissing dR, Nothing, likelyGrp
        End If
        If ContainsAny(dP, highGrp) Or ContainsAny(dD, highGrp) _
           Or ContainsAny(dR, highGrp) Then
            PadMissing dP, Nothing, highGrp
            PadMissing dD, Nothing, highGrp
            PadMissing dR, Nothing, highGrp
        End If
        If ContainsAny(dP, agreeGrp) Or ContainsAny(dD, agreeGrp) _
           Or ContainsAny(dR, agreeGrp) Then
            PadMissing dP, Nothing, agreeGrp
            PadMissing dD, Nothing, agreeGrp
            PadMissing dR, Nothing, agreeGrp
        End If
        If dP.Exists("YES") Or dD.Exists("YES") Or dR.Exists("YES") _
           Or dP.Exists("NO") Or dD.Exists("NO") Or dR.Exists("NO") Then
            PadMissing dP, Nothing, Array("Yes", "No")
            PadMissing dD, Nothing, Array("Yes", "No")
            PadMissing dR, Nothing, Array("Yes", "No")
        End If
        ' union + alphabetical sort
        Dim union As Object: Set union = CreateObject("Scripting.Dictionary")
        AddKeys union, dP: AddKeys union, dD: AddKeys union, dR
        Dim keys() As Variant: keys = union.keys
        QuickSort keys
        ' reorder P codes to desired order
        Dim ordered() As Variant
        Dim iKey As Long, idx As Long, pCode As Variant
        ReDim ordered(0 To UBound(keys))
        idx = 0
        For Each pCode In pdGrp
            For iKey = LBound(keys) To UBound(keys)
                If UCase$(CStr(keys(iKey))) = pCode Then
                    ordered(idx) = keys(iKey)
                    idx = idx + 1
                End If
            Next iKey
        Next pCode
        Dim isP As Boolean
        For iKey = LBound(keys) To UBound(keys)
            isP = False
            For Each pCode In pdGrp
                If UCase$(CStr(keys(iKey))) = pCode Then isP = True: Exit For
            Next pCode
            If Not isP Then
                ordered(idx) = keys(iKey)
                idx = idx + 1
            End If
        Next iKey
        keys = ordered
        ' build slide
        Dim tplPath As String: tplPath = Replace$(TPL_BAR_X3, _
                                            "%APPDATA%", Environ$("APPDATA"))
        If Dir$(tplPath) = "" Then GoTo NextCol
        Dim tpl As Presentation: Set tpl = _
            Presentations.Open(tplPath, msoFalse)
        tpl.Slides(1).Copy: tpl.Close
        pres.Slides.Paste Index:=insertAt
        Dim sld As Slide:     Set sld = pres.Slides(insertAt)
        insertAt = insertAt + 1: slidesAdded = slidesAdded + 1
        sld.Shapes("TargetTitle").TextFrame.TextRange.Text = qText
        Dim ch As Chart: Set ch = FindChartOnSlide(sld)
        Dim wsC As Object: Set wsC = GetChartDataWorksheet(ch)
        wsC.Cells.Clear
        wsC.Range("A1").Resize(1, 4).value = _
            Array("Category", "Plaintiff (%)", "Defense (%)", "Rebuttal (%)")
        Dim i As Long, r As Long
        Dim payload As String: payload = ""
        r = 2
        For i = LBound(keys) To UBound(keys)
            Dim keyVal As String: keyVal = CStr(keys(i))
            wsC.Cells(r, 1).value = ProperLabel(keyVal, Nothing)
            wsC.Cells(r, 2).value = dP(keyVal) / nP
            wsC.Cells(r, 3).value = dD(keyVal) / nD
            wsC.Cells(r, 4).value = dR(keyVal) / nR
            payload = payload & keyVal & vbTab & _
                      Round(dP(keyVal) / nP, 3) & vbTab & _
                      Round(dD(keyVal) / nD, 3) & vbTab & _
                      Round(dR(keyVal) / nR, 3) & vbCrLf
            r = r + 1
        Next i
        ch.Refresh
        wsC.Parent.Close False
        On Error Resume Next
            sld.Shapes("SampleSize").TextFrame.TextRange.Text = "N=" & nP + nD + nR
        On Error GoTo 0
        StashPayload sld, payload, "CHART"
NextCol:
    Next col
    ' tidy
    wbP.Close False: wbD.Close False: wbR.Close False
    xApp.Quit
    Debug.Print "<< finished – added " & slidesAdded & " slide(s)"
    Debug.Print "elapsed " & Format$(Timer - t0, "0.0") & " s"
    Debug.Print String(60, "=") & vbLf
    MsgBox slidesAdded & " comparison slide(s) imported!", vbInformation
End Sub

'========================  helpers  ===============================
Private Function PickXLS(prompt As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = prompt
        .Filters.Clear: .Filters.Add "Excel", "*.xlsx;*.xlsm;*.xls"
        If .show = 0 Then Exit Function
        PickXLS = .SelectedItems(1)
    End With
End Function
Private Sub TallyColumn(ws As Object, c As Long, d As Object, ByRef n As Long)
    Dim lastRow As Long, r As Long, raw As String, resp As String
    lastRow = ws.Cells(ws.rows.Count, c).End(xlUp).Row
    For r = 3 To lastRow
        raw = ws.Cells(r, c).Text
        resp = CleanResp(raw)
        If resp <> "" Then
            If Not d.Exists(resp) Then d.Add resp, 1 Else d(resp) = d(resp) + 1
            n = n + 1
        End If
    Next r
End Sub
Private Sub AddKeys(dst As Object, src As Object)
    Dim k
    For Each k In src.keys
        If Not dst.Exists(k) Then dst.Add k, 1
    Next k
End Sub
Private Sub AddZeros(d As Object, keys As Variant)
    Dim k
    For Each k In keys
        If Not d.Exists(k) Then d.Add k, 0
    Next k
End Sub
Private Sub QuickSort(a As Variant, Optional lo As Long, Optional hi As Long)
    If Not IsArray(a) Then Exit Sub
    If hi = 0 Then hi = UBound(a)
    Dim i As Long, j As Long, p As String, t As String
    i = lo: j = hi: p = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < p: i = i + 1: Loop
        Do While a(j) > p: j = j - 1: Loop
        If i <= j Then t = a(i): a(i) = a(j): a(j) = t: i = i + 1: j = j - 1
    Loop
    If lo < j Then QuickSort a, lo, j
    If i < hi Then QuickSort a, i, hi
End Sub
Private Function DictTotal(d As Object) As Long
    Dim v
    For Each v In d.items
        DictTotal = DictTotal + CLng(v)
    Next v
End Function
