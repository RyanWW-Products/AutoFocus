'=============================================
' Module: ModuleSlideHelpers.bas (Updated)
' A consolidated library of shared Subs/Functions.
' Updated to support P+/P/D+/D mapping, ordering, and alphabetical
' sorting of juror tables.
' Includes payload viewer utilities and helper routines to manage
' duplicate masters.
'=============================================
Option Explicit

'- TEMPLATE PATHS (adjust if you've renamed any constants) -
Public Const TPL_TABLE         As String = "%APPDATA%\Microsoft\AddIns\TQ FG Slides\TQ FG Juror Response.pptx"
Public Const TPL_CHART         As String = "%APPDATA%\Microsoft\AddIns\TQ FG Slides\TQ FG Barchart Percentage.pptx"
Public Const TPL_TITLE_PRESTIM  As String = "%APPDATA%\Microsoft\AddIns\TQ FG Slides\TQ FG TITLE - PRESTIM.pptx"
Public Const TPL_TITLE_POSTCASE As String = "%APPDATA%\Microsoft\AddIns\TQ FG Slides\TQ FG TITLE - POST CASE.pptx"
Public Const TPL_TITLE_WITNESS_EVAL As String = "%APPDATA%\Microsoft\AddIns\TQ FG Slides\TQ FG TITLE - WITNESS EVAL.pptx"
Public Const TPL_RESPONDENT    As String = "%APPDATA%\Microsoft\AddIns\TQ FG Slides\TQ FG Respondent.pptx"
Public Const TPL_SLIDEMASTER As String = _
    "%APPDATA%\Microsoft\AddIns\TQ FG Slides\TQ FG Slidemaster.pptx"
Public Const LAYOUT_AUTOF_MASTER As String = "AutoFocusMaster"

' Adjusted to reduce table height by ~20%, leaving more space at the bottom for the logo
Public Const MAX_ROWS_PER_TABLE As Long = 14    ' real rows that fit on one slide

' Normalize common mis-encoded punctuation and UTF-8 sequences
Public Function CleanSymbols(ByVal txt As String) As String
    Dim s As String: s = txt

    ' Fix curly apostrophes and quotes mis-decoded from UTF-8 sequences (e.g. ’, “).
    s = Replace(s, "’", "'")
    s = Replace(s, "‘", "'")
    s = Replace(s, "“", Chr$(34))
    s = Replace(s, "”", Chr$(34))

    ' Normalize Unicode curly quotes to ASCII equivalents
    s = Replace(s, ChrW(&H2018), "'")
    s = Replace(s, ChrW(&H2019), "'")
    s = Replace(s, ChrW(&H201C), Chr$(34))
    s = Replace(s, ChrW(&H201D), Chr$(34))

    ' Fix common dash mis-encodings (en dash, em dash)
    s = Replace(s, "–", "-")
    s = Replace(s, "—", "-")
    s = Replace(s, ChrW(&H2013), "-")
    s = Replace(s, ChrW(&H2014), "-")

    ' Fix ellipsis
    s = Replace(s, "…", "...")
    s = Replace(s, ChrW(&H2026), "...")

    ' Additional historical UTF-8 mojibake sequences
    s = Replace(s, "â€™", "'")
    s = Replace(s, "â€˜", "'")
    s = Replace(s, "â€œ", Chr$(34))
    s = Replace(s, "â€�", Chr$(34))
    s = Replace(s, "â€“", "-")
    s = Replace(s, "â€—", "-")
    s = Replace(s, "â€¦", "...")
    ' Remove stray accent bytes
    s = Replace(s, "�", "")

    CleanSymbols = s
End Function

'
'-----------------------------
' Strip numeric/letter prefixes like "1-", "A)", "3."
'-----------------------------
Public Function CleanResp(raw As String) As String
    If Len(raw) = 0 Then Exit Function
    raw = CleanSymbols(raw)
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = "^\s*(\d+|[A-Za-z])[\)\.\-]+\s*"
        .IgnoreCase = True
        .Global = False
    End With
    CleanResp = Trim$(re.Replace(raw, ""))
End Function

'
'-----------------------------
' Store a hidden textbox called "RawPayload" on a slide
'-----------------------------
Public Sub StashPayload(sld As Slide, txt As String, viewMode As String)
    On Error Resume Next:    sld.Shapes("RawPayload").Delete:    On Error GoTo 0
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=0, Top:=0, Width:=5, Height:=5 _
    )
    shp.name = "RawPayload": shp.Visible = msoFalse
    If Right$(txt, 2) = vbCrLf Then txt = Left$(txt, Len(txt) - 2)
    shp.TextFrame.TextRange.Text = txt
    sld.Tags.Add "ViewMode", viewMode
End Sub

'
'-----------------------------
' Remove all but header+stub from a TABLE, then blank the stub
'-----------------------------
Public Sub ClearTableData(tbl As Table)
    Do While tbl.rows.Count > 2
        tbl.rows(tbl.rows.Count).Delete
    Loop
    tbl.cell(2, 1).Shape.TextFrame.TextRange.Text = ""
    tbl.cell(2, 2).Shape.TextFrame.TextRange.Text = ""
End Sub

'
'-----------------------------
' Normalize CR/LF and split into a zero-based String() array
'-----------------------------
Public Function PayloadLines(raw As String) As String()
    Dim tmp As String: tmp = Replace$(raw, vbCr, vbLf)
    Do While InStr(tmp, vbLf & vbLf)
        tmp = Replace$(tmp, vbLf & vbLf, vbLf)
    Loop
    If Right$(tmp, 1) = vbLf Then tmp = Left$(tmp, Len(tmp) - 1)
    PayloadLines = Split(tmp, vbLf)
End Function

'
'-----------------------------
' Insert a single slide from a template just after afterSld
'-----------------------------
Public Function InsertTemplateAfter(tplPath As String, afterSld As Slide) As Slide
    Dim pres As Presentation: Set pres = afterSld.Parent
    Dim fullPath As String: fullPath = Replace$(tplPath, "%APPDATA%", Environ$("APPDATA"))
    If Dir$(fullPath) = "" Then
        MsgBox "Missing template: " & fullPath, vbExclamation
        Exit Function
    End If
    pres.Slides.InsertFromFile fullPath, afterSld.SlideIndex, 1, 1
    Set InsertTemplateAfter = pres.Slides(afterSld.SlideIndex + 1)
End Function

'
'-----------------------------
' Copy one slide from a template to the end of the deck, optionally set its title
'-----------------------------
Public Function CopyTemplateSlide( _
    tplPath As String, _
    pres As Presentation, _
    Optional titleText As String = "" _
) As Slide
    Dim p As String: p = Replace(tplPath, "%APPDATA%", Environ$("APPDATA"))
    If Dir$(p) = "" Then
        MsgBox "Missing template: " & p, vbExclamation
        Exit Function
    End If
    pres.Slides.InsertFromFile p, pres.Slides.Count, 1, 1
    Dim sld As Slide: Set sld = pres.Slides(pres.Slides.Count)
    If titleText <> "" Then
        On Error Resume Next
          sld.Shapes("TargetTitle").TextFrame.TextRange.Text = titleText
        On Error GoTo 0
    End If

    ' Re-attach new slide to the existing AutoFocusMaster design if present
    On Error Resume Next
    Dim idx As Long: idx = GetDesignIndexByMasterName(pres, LAYOUT_AUTOF_MASTER)
    If idx > 0 Then sld.Design = pres.Designs(idx)
    On Error GoTo 0

    ' Prune any duplicate AutoFocusMaster designs (clean up after insertion)
    PruneDuplicateMasters pres, LAYOUT_AUTOF_MASTER

    Set CopyTemplateSlide = sld
End Function

Private Function GetDesignIndexByMasterName(ByVal pres As Presentation, ByVal masterName As String) As Long
    Dim i As Long
    For i = 1 To pres.Designs.Count
        Dim smn As String
        On Error Resume Next
        smn = pres.Designs(i).SlideMaster.name
        On Error GoTo 0
        If InStr(1, smn, masterName, vbTextCompare) > 0 Then
            GetDesignIndexByMasterName = i
            Exit Function
        End If
    Next
    GetDesignIndexByMasterName = 0
End Function

Public Sub PruneDuplicateMasters(ByVal pres As Presentation, ByVal masterName As String)
    Dim i As Long
    For i = pres.Designs.Count To 1 Step -1
        Dim d As Design
        Set d = pres.Designs(i)
        Dim smn As String
        On Error Resume Next
        smn = d.SlideMaster.name
        On Error GoTo 0
        If Len(smn) > 0 Then
            If InStr(1, smn, masterName, vbTextCompare) > 0 And i <> GetDesignIndexByMasterName(pres, masterName) Then
                ' Check if any slide uses this design
                Dim used As Boolean: used = False
                Dim s As Slide
                For Each s In pres.Slides
                    If s.Design Is d Then
                        used = True
                        Exit For
                    End If
                Next
                If Not used Then
                    On Error Resume Next
                    d.Delete
                    On Error GoTo 0
                End If
            End If
        End If
    Next i
End Sub

'
'-----------------------------
' Remove any continuation slides tagged "TABLE" with the same title
'-----------------------------
Public Sub RemoveContinuationSlides(cur As Slide)
    Dim pres As Presentation: Set pres = cur.Parent
    Dim title As String: title = cur.Shapes("TargetTitle").TextFrame.TextRange.Text
    Dim i As Long: i = cur.SlideIndex + 1
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

'
'-----------------------------
' Remove any continuation slides tagged "CHART" with the same title
'-----------------------------
Public Sub RemoveChartPages(cur As Slide)
    Dim pres As Presentation: Set pres = cur.Parent
    Dim title As String: title = cur.Shapes("TargetTitle").TextFrame.TextRange.Text
    Dim i As Long: i = cur.SlideIndex + 1
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

'
'-----------------------------
' Find & return the Chart object on a slide (even if grouped)
'-----------------------------
Public Function FindChartOnSlide(sld As Slide) As Chart
    Dim sh As Shape
    For Each sh In sld.Shapes
        If sh.HasChart Then
            Set FindChartOnSlide = sh.Chart: Exit Function
        ElseIf sh.Type = msoGroup Then
            Dim g As Shape
            For Each g In sh.GroupItems
                If g.HasChart Then
                    Set FindChartOnSlide = g.Chart: Exit Function
                End If
            Next
        End If
    Next
End Function

'
'-----------------------------
' Simple bubble-sort for a 1-D Variant array
'-----------------------------
Public Sub SortArray(arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next
    Next
End Sub

'
'---------------------------------------------
' SortKeysCustom - returns an array of UPPERCASE keys
' in your exact custom order (falling back to A-Z)
'  Updated to recognise P+, P, D+, and D as a special group ordered
'  Strongly Favor Plaintiff ? Somewhat Favor Plaintiff ? Strongly Favor Defendant ? Somewhat Favor Defendant.
'---------------------------------------------
' Helper: strip dashes, slashes, and colons from a string for numeric detection
Private Function StripDateChars(ByVal s As String) As String
    s = Replace(s, "-", "")
    s = Replace(s, "/", "")
    s = Replace(s, ":", "")
    StripDateChars = s
End Function

' Determine whether a set of keys should be written forward (ascending)
' Returns True for: all-numeric keys (after stripping date chars),
' or mostly-numeric sets where 1-2 non-numeric items remain but start with digits.
Public Function ShouldWriteForward(ByVal keys As Variant) As Boolean
    On Error GoTo done
    If Not IsArray(keys) Then Exit Function
    Dim k As Variant
    Dim numericCount As Long: numericCount = 0
    Dim nonNumericCount As Long: nonNumericCount = 0
    Dim startsLikeNum As Boolean: startsLikeNum = True
    Dim total As Long: total = 0
    For Each k In keys
        total = total + 1
        Dim ks As String: ks = CStr(k)
        If IsNumeric(StripDateChars(ks)) Then
            numericCount = numericCount + 1
        Else
            nonNumericCount = nonNumericCount + 1
            If Len(ks) = 0 Or Not IsNumeric(Left$(ks, 1)) Then
                startsLikeNum = False
            End If
        End If
    Next
    ' All numeric -> forward
    If numericCount = total And total > 0 Then
        ShouldWriteForward = True
        Exit Function
    End If
    ' Mostly numeric (all-but-1 or all-but-2) and non-numeric items start with digits
    If startsLikeNum And (nonNumericCount = 1 Or nonNumericCount = 2) And numericCount >= 1 Then
        ShouldWriteForward = True
        Exit Function
    End If
done:
    ShouldWriteForward = False
End Function

Public Function SortKeysCustom(arr As Variant) As Variant
    ' Build a quick lookup of the keys that are present
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, k As Variant
    For i = LBound(arr) To UBound(arr)
        dict(arr(i)) = True
    Next
    '----- 1) YES / NO special case -----
    If dict.Exists("YES") And dict.Exists("NO") Then
        SortKeysCustom = Array("YES", "NO")
        Exit Function
    End If
    '----- 1a) P/D favor codes special case -----
    Dim pdGrp As Variant: pdGrp = Array("P+", "P", "D+", "D")
    Dim outKeys() As String: ReDim outKeys(0)
    Dim found As Boolean: found = False
    ' append any P/D codes in order
    Call AppendIfPresent(dict, pdGrp, outKeys, found)
    If found Then
        SortKeysCustom = ToUpperArray(outKeys)
        Exit Function
    End If
    '----- 2) Numeric scale? -----
    Dim allNum As Boolean: allNum = True
    For Each k In dict.keys
        If Not IsNumeric(StripDateChars(CStr(k))) Then allNum = False: Exit For
    Next
    If allNum Then
        Dim nums() As Double: ReDim nums(0 To dict.Count - 1)
        i = 0
        For Each k In dict.keys: nums(i) = CDbl(StripDateChars(CStr(k))): i = i + 1: Next
        Call SortArray(nums)                         'bubble-sort numeric
        Dim outN() As String: ReDim outN(0 To UBound(nums))
        For i = LBound(nums) To UBound(nums): outN(i) = CStr(nums(i)): Next
        SortKeysCustom = outN
        Exit Function
    End If
    '----- 2a) Mostly numeric with 1-2 non-numeric prepended items? -----
    ' Count numeric and non-numeric keys (stripping dashes/slashes for detection)
    Dim numericCount As Long: numericCount = 0
    Dim nonNumericCount As Long: nonNumericCount = 0
    Dim hasNonNumericPrepended As Boolean: hasNonNumericPrepended = True
    For Each k In dict.keys
        If IsNumeric(StripDateChars(CStr(k))) Then
            numericCount = numericCount + 1
        Else
            nonNumericCount = nonNumericCount + 1
            ' Check if non-numeric starts with a digit
            If Len(CStr(k)) = 0 Or Not IsNumeric(Left$(CStr(k), 1)) Then
                hasNonNumericPrepended = False
            End If
        End If
    Next
    ' Apply this rule only if: all-but-1 or all-but-2 are numeric, and non-numeric items start with digits
    If hasNonNumericPrepended And (nonNumericCount = 1 Or nonNumericCount = 2) And numericCount >= 1 Then
        ' Sort numeric items first
        Dim numItems() As Double: ReDim numItems(0 To numericCount - 1)
        Dim nonNumItems() As String: ReDim nonNumItems(0 To nonNumericCount - 1)
        Dim numIdx As Long: numIdx = 0
        Dim nonNumIdx As Long: nonNumIdx = 0
        For Each k In dict.keys
            If IsNumeric(StripDateChars(CStr(k))) Then
                numItems(numIdx) = CDbl(StripDateChars(CStr(k)))
                numIdx = numIdx + 1
            Else
                nonNumItems(nonNumIdx) = CStr(k)
                nonNumIdx = nonNumIdx + 1
            End If
        Next
        Call SortArray(numItems)
        ' Combine: sorted numbers first, then non-numeric items
        Dim combined() As String: ReDim combined(0 To numericCount + nonNumericCount - 1)
        For i = 0 To numericCount - 1
            combined(i) = CStr(numItems(i))
        Next
        For i = 0 To nonNumericCount - 1
            combined(numericCount + i) = nonNumItems(i)
        Next
        SortKeysCustom = combined
        Exit Function
    End If
    '----- 3) Four-point Likely group (partial OK) -----
    Dim likelyGrp As Variant: likelyGrp = LikelyGroup()
    Erase outKeys: ReDim outKeys(0): found = False
    Call AppendIfPresent(dict, likelyGrp, outKeys, found)
    If found Then SortKeysCustom = ToUpperArray(outKeys): Exit Function
    '----- 4) Five-point High/Low group (partial OK) -----
    Dim highGrp As Variant: highGrp = HighGroup()
    Erase outKeys: ReDim outKeys(0): found = False
    Call AppendIfPresent(dict, highGrp, outKeys, found)
    If found Then SortKeysCustom = ToUpperArray(outKeys): Exit Function
    '----- 5) Four-point Agree group (partial OK) -----
    Dim agreeGrp As Variant: agreeGrp = AgreeGroup()
    Erase outKeys: ReDim outKeys(0): found = False
    Call AppendIfPresent(dict, agreeGrp, outKeys, found)
    If found Then SortKeysCustom = ToUpperArray(outKeys): Exit Function
    '----- 6) Custom dynamic answer groups from the AnswerGroups dictionary -----
    ' At this point, the keys are neither YES/NO, P/D codes, numeric, nor any
    ' of the built-in Likely/High/Agree scales.  We next check the global
    ' AnswerGroups dictionary (managed by GroupArrays) for any user-defined
    ' groups.  If enough keys from a user-defined group are present, we
    ' reorder all keys so that matching group items come first (in the
    ' group-defined order) and any remaining keys follow in their original
    ' appearance order.  Built-in groups like Likely, High, Agree, PLong,
    ' PShort, and YesNo are skipped here because they were handled
    ' earlier or via the dedicated functions above.
    On Error Resume Next
    Call InitializeAnswerGroups
    Dim grpName As Variant
    Dim gArr As Variant
    Dim matchCount As Long
    Dim groupSize As Long
    Dim shouldApply As Boolean
    Dim ku As String
    Dim idx As Long
    Dim tmp() As String
    Dim added As Boolean
    If Not AnswerGroups Is Nothing Then
        For Each grpName In AnswerGroups.keys
            Select Case UCase$(CStr(grpName))
                Case "LIKELY", "HIGH", "AGREE", "PLONG", "PSHORT", "YESNO"
                    ' skip built-in groups
                Case Else
                    gArr = AnswerGroups(grpName)
                    ' Count how many items from this group exist in the current key set
                    matchCount = 0
                    For Each k In gArr
                        If dict.Exists(UCase$(CStr(k))) Then
                            matchCount = matchCount + 1
                        End If
                    Next k
                    groupSize = UBound(gArr) - LBound(gArr) + 1
                    ' Determine if this group should apply based on the number of matches
                    If groupSize <= 2 Then
                        ' For two-item groups (e.g., Yes/No), require at least one match
                        shouldApply = (matchCount >= 1)
                    Else
                        ' For larger groups, require at least two matches
                        shouldApply = (matchCount >= 2)
                    End If
                    If shouldApply Then
                        ' Build a reordered array: start with matching group items (in group order)
                        added = False
                        Erase tmp
                        Dim g As Variant
                        For Each g In gArr
                            ku = UCase$(CStr(g))
                            If dict.Exists(ku) Then
                                If Not added Then
                                    ReDim tmp(0)
                                    tmp(0) = ku
                                    added = True
                                Else
                                    ReDim Preserve tmp(UBound(tmp) + 1)
                                    tmp(UBound(tmp)) = ku
                                End If
                            End If
                        Next g
                        ' Append any remaining keys that were not part of the group in their original order
                       
                        For i = LBound(arr) To UBound(arr)
                            ku = UCase$(CStr(arr(i)))
                            Dim present As Boolean
                            present = False
                            If added Then
                                For idx = LBound(tmp) To UBound(tmp)
                                    If tmp(idx) = ku Then
                                        present = True
                                        Exit For
                                    End If
                                Next idx
                            End If
                            If Not present Then
                                If Not added Then
                                    ReDim tmp(0)
                                    tmp(0) = ku
                                    added = True
                                Else
                                    ReDim Preserve tmp(UBound(tmp) + 1)
                                    tmp(UBound(tmp)) = ku
                                End If
                            End If
                        Next i
                        SortKeysCustom = tmp
                        Exit Function
                    End If
            End Select
        Next grpName
    End If
    On Error GoTo 0

    '----- 7) Fallback: preserve original order of appearance -----
    ' When no known or custom groups apply, keep the order of keys as provided
    SortKeysCustom = arr
End Function

Private Function ContainsAll(d As Object, vArray As Variant) As Boolean
    Dim i As Long
    For i = LBound(vArray) To UBound(vArray)
        If Not d.Exists(UCase$(vArray(i))) Then Exit Function
    Next
    ContainsAll = True
End Function

Private Function ToUpperArray(vArray As Variant) As Variant
    Dim i As Long, out() As Variant
    ReDim out(LBound(vArray) To UBound(vArray))
    For i = LBound(vArray) To UBound(vArray)
        out(i) = UCase$(vArray(i))
    Next
    ToUpperArray = out
End Function

Private Sub AppendIfPresent(d As Object, grp As Variant, _
                             ByRef buf() As String, ByRef something As Boolean)
    Dim g
    For Each g In grp
        If d.Exists(UCase$(g)) Then
            If Not something Then
                ReDim buf(0)
            Else
                ReDim Preserve buf(UBound(buf) + 1)
            End If
            buf(UBound(buf)) = g
            something = True
        End If
    Next g
End Sub

'====================================================================
' Ensure every element of grp() exists in freq / origMap;
' add with a zero count if missing.
'====================================================================
Public Sub PadMissing(ByRef freq As Object, _
                      ByRef origMap As Object, _
                      grp As Variant)
    Dim v
    For Each v In grp
        Dim keyUp As String: keyUp = UCase$(v)
        If Not freq.Exists(keyUp) Then
            freq.Add keyUp, 0
            If Not origMap Is Nothing Then origMap.Add keyUp, v
        End If
    Next v
End Sub

Public Function ContainsAny(d As Object, grp As Variant) As Boolean
    Dim v
    For Each v In grp
        If d.Exists(UCase$(v)) Then ContainsAny = True: Exit Function
    Next v
End Function

'--------------------------------------------
' ProperLabel - converts an UPPER-KEY back to a display label.
' Updated to recognise P codes and map them to long form labels.
'--------------------------------------------
Public Function ProperLabel(ByVal keyUp As Variant, ByVal mapOrig As Object) As String
    Dim v
    ' normalise
    If IsNull(keyUp) Then ProperLabel = "": Exit Function
    keyUp = UCase$(CStr(keyUp))
    '- handle P/D favour codes -
    Select Case keyUp
        Case "P+": ProperLabel = "Strongly Favor the Plaintiff": Exit Function
        Case "P": ProperLabel = "Somewhat Favor the Plaintiff": Exit Function
        Case "D+": ProperLabel = "Strongly Favor the Defendant": Exit Function
        Case "D": ProperLabel = "Somewhat Favor the Defendant": Exit Function
    End Select
    '- predefined groups using shared helpers (must match SortKeysCustom)
    Dim likelyGrp As Variant, highGrp As Variant, agreeGrp As Variant
    likelyGrp = LikelyGroup()
    highGrp = HighGroup()
    agreeGrp = AgreeGroup()
    ' 1) Likely scale
    For Each v In likelyGrp
        If UCase$(CStr(v)) = keyUp Then
            ProperLabel = v: Exit Function
        End If
    Next v
    ' 2) High/Low scale
    For Each v In highGrp
        If UCase$(CStr(v)) = keyUp Then
            ProperLabel = v: Exit Function
        End If
    Next v
    ' 3) Agree/Disagree scale
    For Each v In agreeGrp
        If UCase$(CStr(v)) = keyUp Then
            ProperLabel = v: Exit Function
        End If
    Next v
    ' 4) Numeric scale
    If IsNumeric(keyUp) Then ProperLabel = keyUp: Exit Function
    ' 5) Fallback to original-case value, then proper-case
    If Not mapOrig Is Nothing Then
        If mapOrig.Exists(keyUp) Then
            ProperLabel = StrConv(LCase$(mapOrig(keyUp)), vbProperCase)
            Exit Function
        End If
    End If
    ' 6) Ultimate fallback - proper-case the key itself
    ProperLabel = StrConv(LCase$(keyUp), vbProperCase)
End Function

'====================================================================
' BuildIntroChartSlide - fills CHART template with padded scales.
' There are two overloads: this version is called from Module helpers and
' accepts a slide instead of a presentation.  It honours P codes.
'====================================================================
Public Sub BuildIntroChartSlide( _
            sld As Slide, _
            ws As Object, _
            dataCol As Long)
    '-- tally responses into dictionaries --
    Dim freq As Object:    Set freq = CreateObject("Scripting.Dictionary")
    Dim origMap As Object: Set origMap = CreateObject("Scripting.Dictionary")
    Dim payload As String, total As Long
    Dim r As Long, lastRow As Long, resp As String, keyUp As String
    lastRow = ws.Cells(ws.rows.Count, dataCol).End(xlUp).Row
    For r = 3 To lastRow
        resp = CleanResp(ws.Cells(r, dataCol).Text)
        If resp <> "" Then
            keyUp = UCase$(resp)
            If Not freq.Exists(keyUp) Then
                freq.Add keyUp, 1
                origMap.Add keyUp, resp
            Else
                freq(keyUp) = freq(keyUp) + 1
            End If
            payload = payload & resp & vbCrLf
            total = total + 1
        End If
    Next r
    If total = 0 Then Exit Sub
    '-- pad missing buckets (Agree / Likely / High-Low / Yes-No / P codes) --
    Dim pdGrp As Variant, likelyGrp As Variant, highGrp As Variant, agreeGrp As Variant
    pdGrp = PShortGroup()
    likelyGrp = LikelyGroup()
    highGrp = HighGroup()
    agreeGrp = AgreeGroup()
    If ContainsAny(freq, pdGrp) Then PadMissing freq, origMap, pdGrp
    If ContainsAny(freq, likelyGrp) Then PadMissing freq, origMap, likelyGrp
    If ContainsAny(freq, highGrp) Then PadMissing freq, origMap, highGrp
    If ContainsAny(freq, agreeGrp) Then PadMissing freq, origMap, agreeGrp
    If freq.Exists("YES") Or freq.Exists("NO") Then _
            PadMissing freq, origMap, Array("Yes", "No")
    ' -- sort once, using padded list. Prefer any user-defined group order over built-ins --
    Dim keysSorted As Variant
    Dim matched As Boolean
    matched = False
    Dim gName As Variant, gArr As Variant
    Dim v As Variant
    Dim buf() As String
    Dim added As Boolean
    Dim ku As String
    On Error Resume Next
    ' Initialize answer groups; ignore errors if not present
    Call InitializeAnswerGroups
    On Error GoTo 0
    If Not AnswerGroups Is Nothing Then
        For Each gName In AnswerGroups.keys
            Select Case UCase$(CStr(gName))
                Case "LIKELY", "HIGH", "AGREE", "PLONG", "PSHORT"
                    ' skip built-in groups handled above
                Case Else
                    gArr = AnswerGroups(gName)
                    ' Count how many items from this group exist in the frequency dictionary
                    Dim matchCount As Long
                    matchCount = 0
                    For Each v In gArr
                        If freq.Exists(UCase$(CStr(v))) Then
                            matchCount = matchCount + 1
                        End If
                    Next v
                    ' Determine if this group should apply: require at least two matches for groups with
                    ' more than two items; one match suffices for two-item groups like Yes/No
                    Dim groupSize As Long
                    groupSize = UBound(gArr) - LBound(gArr) + 1
                    Dim shouldApply As Boolean
                    If groupSize <= 2 Then
                        shouldApply = (matchCount >= 1)
                    Else
                        shouldApply = (matchCount >= 2)
                    End If
                    If shouldApply Then
                        ' Pad missing buckets for this custom group
                        PadMissing freq, origMap, gArr
                        ' Build the sorted key list according to this group's order
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
    ' Fallback to generic sorter if no custom group matched
    If Not matched Then
        keysSorted = SortKeysCustom(freq.keys)
    End If
    '-- update N= label --
    On Error Resume Next
      sld.Shapes("SampleSize").TextFrame.TextRange.Text = "N=" & total
    On Error GoTo 0
    '-- open embedded workbook --
    Dim ch As Chart: Set ch = sld.Shapes("TargetChart").Chart
    Dim wsC As Object: Set wsC = GetChartDataWorksheet(ch)
    If wsC Is Nothing Then Exit Sub
    wsC.Cells.Clear
    wsC.Range("A1").value = "Category"
    wsC.Range("B1").value = "Pct"
    '-- proper-case labels & write rows in original order --
    Dim i As Long, dest As Long: dest = 2
    For i = LBound(keysSorted) To UBound(keysSorted)
        wsC.Cells(dest, 1).value = ProperLabel(keysSorted(i), origMap)
        wsC.Cells(dest, 2).value = freq(keysSorted(i)) / total
        dest = dest + 1
    Next i
    wsC.Parent.Close False
    StashPayload sld, Left$(payload, Len(payload) - 2), "CHART"
End Sub

'====================================================================
' Safely return the worksheet inside a chart's data workbook.
'====================================================================
Public Function GetChartDataWorksheet(ch As Chart) As Object
    Const MAX_RETRY As Long = 20            ' 20 � 50 ms ? 1 s
    Dim wb As Object, i As Long
    For i = 1 To MAX_RETRY
        On Error Resume Next
            ch.ChartData.Activate           ' may raise RPC_E_CALL_REJECTED
            Set wb = ch.ChartData.Workbook  ' Nothing if previous call failed
        On Error GoTo 0
        If Not wb Is Nothing Then
            Set GetChartDataWorksheet = wb.Worksheets(1)
            Exit Function
        End If
        DoEvents                             ' yield to UI
    Next i
    MsgBox "Couldn't get chart data after several retries.", vbExclamation
End Function

Public Function GuessColumnType(ws As Object, c As Long) As String
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim longTxt As Long, lastR As Long: lastR = ws.Cells(ws.rows.Count, c).End(xlUp).Row
    Dim r As Long
    For r = 3 To lastR
        Dim v As String: v = Trim$(ws.Cells(r, c).Text)
        If v <> "" Then
            If Len(v) > 20 Then longTxt = longTxt + 1
            If Not dict.Exists(v) Then dict.Add v, 1
            If dict.Count > 6 And longTxt >= 3 Then Exit For
        End If
    Next r
    GuessColumnType = IIf(dict.Count > 6 Or longTxt >= 3, "TABLE", "CHART")
End Function

Public Function ForceTable(colNum As Long) As Boolean
    Select Case colNum
        Case 32, 33, 36          ' AF, AG, AJ ...
            ForceTable = True
        Case Else
            ForceTable = False
    End Select
End Function

'====================================================================
' Additional helpers for sorting juror names in RawPayload arrays
'====================================================================
Public Function SortLinesByName(lines() As String) As String()
    ' Returns a new array with lines sorted alphabetically by the first field
    Dim sorted() As String
    If Not IsArray(lines) Then SortLinesByName = lines: Exit Function
    Dim i As Long, j As Long
    ReDim sorted(LBound(lines) To UBound(lines))
    ' Copy lines into sorted() array
    For i = LBound(lines) To UBound(lines)
        sorted(i) = lines(i)
    Next i
    Dim temp As String, nameI As String, nameJ As String
    For i = LBound(sorted) To UBound(sorted) - 1
        For j = i + 1 To UBound(sorted)
            nameI = GetLineJurorName(sorted(i))
            nameJ = GetLineJurorName(sorted(j))
            If StrComp(nameI, nameJ, vbTextCompare) > 0 Then
                temp = sorted(i)
                sorted(i) = sorted(j)
                sorted(j) = temp
            End If
        Next j
    Next i
    SortLinesByName = sorted
End Function

Private Function GetLineJurorName(line As String) As String
    Dim parts() As String
    parts = Split(line, vbTab, 2)
    If UBound(parts) >= 0 Then
        GetLineJurorName = Trim$(parts(0))
    Else
        GetLineJurorName = ""
    End If
End Function

'============================================================
' BuildTableFromRaw - rebuild a paginated TABLE view from a
' 0-based String() payload containing "Juror<TAB>Response" lines.
' It sorts rows alphabetically by juror name and uses the same
' stub+continuation logic as your importer.
'============================================================
Public Sub BuildTableFromRaw( _
        tblSlide As Slide, _
        ByRef lines() As String)
    ' Sort the lines by juror name before paginating
    lines = SortLinesByName(lines)
    Const MAX_ROWS      As Long = MAX_ROWS_PER_TABLE        ' real rows per slide
    Const BOTTOM_MARGIN As Single = 0.75 * 72               ' �" bottom margin
    Dim pres        As Presentation: Set pres = tblSlide.Parent
    Dim slideBottom As Single:        slideBottom = pres.PageSetup.SlideHeight - BOTTOM_MARGIN
    Dim tblShp As Shape: Set tblShp = tblSlide.Shapes("TargetTable")
    Dim tbl    As Table: Set tbl = tblShp.Table
    '--- start with header + invisible stub row ---
    ClearTableData tbl
    Dim rowCount As Long    ' number of real rows on current slide
    Dim payload  As String  ' to re-stash RawPayload
    Dim parts()  As String
    Dim i        As Long
    rowCount = 0: payload = ""
    ' Process sorted lines in ascending chunks of MAX_ROWS
    Dim pageStart As Long: pageStart = LBound(lines)
    Do While pageStart <= UBound(lines)
        Dim pageEnd As Long
        pageEnd = pageStart + MAX_ROWS - 1
        If pageEnd > UBound(lines) Then pageEnd = UBound(lines)
        ' For pages after the first, finalize current slide and spawn a new one
        If pageStart > LBound(lines) Then
            tbl.rows(2).Delete
            StashPayload tblSlide, payload, "TABLE"
            payload = "": rowCount = 0
            Set tblSlide = CopyTemplateSlide( _
                             TPL_TABLE, _
                             pres, _
                             tblSlide.Shapes("TargetTitle").TextFrame.TextRange.Text)
            Set tblShp = tblSlide.Shapes("TargetTable")
            Set tbl = tblShp.Table
            ClearTableData tbl
        End If
        ' Insert rows in reverse within this page so they appear A?Z top-to-bottom
        Dim idx As Long
        For idx = pageEnd To pageStart Step -1
            parts = Split(lines(idx), vbTab, 2)
            If UBound(parts) >= 1 Then
                tbl.rows.Add 2
                tbl.cell(2, 1).Shape.TextFrame.TextRange.Text = Trim$(parts(0))
                tbl.cell(2, 2).Shape.TextFrame.TextRange.Text = Trim$(parts(1))
                payload = payload & parts(0) & vbTab & parts(1) & vbCrLf
                rowCount = rowCount + 1
            End If
        Next idx
        pageStart = pageEnd + 1
    Loop
    ' final slide: remove stub and stash one last time
    tbl.rows(2).Delete
    StashPayload tblSlide, payload, "TABLE"
End Sub

'============================================================
' Payload viewer - toggles a readable overlay of the slide's RawPayload
'  - If payload lines contain a TAB, shows a 2-column table (Juror | Response)
'  - Else shows a monospaced text panel
'  - Call again to hide the overlay
'============================================================
Public Sub ShowSlidePayload(Optional ByVal sld As Slide)
    If sld Is Nothing Then
        On Error Resume Next
        Set sld = ActiveWindow.View.Slide
        On Error GoTo 0
    End If
    If sld Is Nothing Then
        MsgBox "No active slide to show payload on.", vbExclamation
        Exit Sub
    End If

    ' Toggle off if viewer already present
    If HasPayloadViewer(sld) Then
        RemovePayloadViewer sld
        Exit Sub
    End If

    Dim raw As String
    raw = GetRawPayloadText(sld)
    If Len(raw) = 0 Then
        MsgBox "No RawPayload found on this slide.", vbInformation
        Exit Sub
    End If

    ' Normalize & split
    Dim lines() As String
    lines = PayloadLines(raw)

    ' Heuristic: if at least one line contains a TAB, treat as 2-column data
    Dim showAsTable As Boolean
    showAsTable = LooksTabular(lines)

    If showAsTable Then
        BuildPayloadViewerTable sld, lines
    Else
        BuildPayloadViewerText sld, Join(lines, vbCrLf)
    End If
End Sub
' Wrapper that satisfies Ribbon onAction signature
Public Sub ShowSlidePayload_Ribbon(control As IRibbonControl)
    ' Just call your utility; it will pick up the active slide
    ShowSlidePayload
End Sub


Private Function GetRawPayloadText(ByVal sld As Slide) As String
    On Error Resume Next
    Dim shp As Shape
    Set shp = sld.Shapes("RawPayload")
    On Error GoTo 0
    If Not shp Is Nothing Then
        GetRawPayloadText = shp.TextFrame.TextRange.Text
    Else
        GetRawPayloadText = ""
    End If
End Function

Private Function LooksTabular(ByRef lines() As String) As Boolean
    On Error GoTo done
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), vbTab) > 0 Then
            LooksTabular = True
            Exit Function
        End If
    Next i
done:
End Function

Public Sub RemovePayloadViewer(ByVal sld As Slide)
    On Error Resume Next
    sld.Shapes("PayloadViewerGroup").Delete
    On Error GoTo 0
End Sub

Private Function HasPayloadViewer(ByVal sld As Slide) As Boolean
    On Error Resume Next
    HasPayloadViewer = Not sld.Shapes("PayloadViewerGroup") Is Nothing
    On Error GoTo 0
End Function

Private Sub BuildPayloadViewerText(ByVal sld As Slide, ByVal textBody As String)
    ' Clean any prior viewer
    RemovePayloadViewer sld

    ' Create a group to make toggling easy
    Dim grp As Shape
    Set grp = sld.Shapes.AddShape(Type:=msoShapeRectangle, _
                                  Left:=36, Top:=72, Width:=sld.Parent.PageSetup.SlideWidth - 72, _
                                  Height:=sld.Parent.PageSetup.SlideHeight - 144)
    grp.name = "PayloadViewerGroup"
    grp.Fill.ForeColor.RGB = RGB(245, 245, 245)
    grp.Fill.Transparency = 0.06
    grp.line.ForeColor.RGB = RGB(180, 180, 180)

    Dim txt As Shape
    Set txt = sld.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                                    Left:=grp.Left + 12, Top:=grp.Top + 12, _
                                    Width:=grp.Width - 24, Height:=grp.Height - 24)
    txt.name = "PayloadViewerText"
    With txt.TextFrame
        .AutoSize = ppAutoSizeNone
        .MarginLeft = 6: .MarginRight = 6: .MarginTop = 4: .MarginBottom = 4
        .WordWrap = msoTrue
        .TextRange.Text = textBody
        With .TextRange.Font
            .name = "Consolas"
            .Size = 10
            .Color.RGB = RGB(30, 30, 30)
        End With
    End With

    ' Group them
    Dim arr(1 To 2) As Variant
    Set arr(1) = grp: Set arr(2) = txt
    sld.Shapes.Range(arr).Group.name = "PayloadViewerGroup"
End Sub

Private Sub BuildPayloadViewerTable(ByVal sld As Slide, ByRef lines() As String)
    ' Clean any prior viewer
    RemovePayloadViewer sld

    ' Table bounds
    Dim leftPt As Single, topPt As Single, wPt As Single, hPt As Single
    leftPt = 36: topPt = 72
    wPt = sld.Parent.PageSetup.SlideWidth - 72
    hPt = sld.Parent.PageSetup.SlideHeight - 144

    ' Outer backer
    Dim back As Shape
    Set back = sld.Shapes.AddShape(msoShapeRectangle, leftPt, topPt, wPt, hPt)
    back.Fill.ForeColor.RGB = RGB(245, 245, 245)
    back.Fill.Transparency = 0.06
    back.line.ForeColor.RGB = RGB(180, 180, 180)
    back.name = "PayloadViewerGroup" ' temporary

    ' Create a 2-column table: header + N rows
    Dim rows As Long: rows = (UBound(lines) - LBound(lines) + 1) + 1
    Dim tblShp As Shape
    Set tblShp = sld.Shapes.AddTable(NumRows:=rows, NumColumns:=2, _
                                     Left:=leftPt + 12, Top:=topPt + 12, _
                                     Width:=wPt - 24, Height:=hPt - 24)
    Dim tbl As Table: Set tbl = tblShp.Table
    ' Header
    tbl.cell(1, 1).Shape.TextFrame.TextRange.Text = "Juror"
    tbl.cell(1, 2).Shape.TextFrame.TextRange.Text = "Response"
    StyleViewerHeaderCell tbl.cell(1, 1)
    StyleViewerHeaderCell tbl.cell(1, 2)

    ' Rows
    Dim i As Long, parts() As String, r As Long
    r = 2
    For i = LBound(lines) To UBound(lines)
        parts = Split(lines(i), vbTab, 2)
        If UBound(parts) >= 1 Then
            tbl.cell(r, 1).Shape.TextFrame.TextRange.Text = parts(0)
            tbl.cell(r, 2).Shape.TextFrame.TextRange.Text = parts(1)
        Else
            tbl.cell(r, 1).Shape.TextFrame.TextRange.Text = lines(i)
            tbl.cell(r, 2).Shape.TextFrame.TextRange.Text = ""
        End If
        StyleViewerBodyCell tbl.cell(r, 1)
        StyleViewerBodyCell tbl.cell(r, 2)
        r = r + 1
    Next i

    ' Group backer + table as one togglable viewer
    Dim arr(1 To 2) As Variant
    Set arr(1) = back: Set arr(2) = tblShp
    sld.Shapes.Range(arr).Group.name = "PayloadViewerGroup"
End Sub

Private Sub StyleViewerHeaderCell(ByVal cell As cell)
    With cell.Shape.TextFrame.TextRange
        .Font.Bold = True
        .Font.name = "Segoe UI"
        .Font.Size = 12
        .ParagraphFormat.Alignment = ppAlignLeft
    End With
    cell.Shape.Fill.ForeColor.RGB = RGB(235, 235, 235)
    cell.Borders(ppBorderTop).ForeColor.RGB = RGB(200, 200, 200)
    cell.Borders(ppBorderBottom).ForeColor.RGB = RGB(200, 200, 200)
    cell.Borders(ppBorderLeft).ForeColor.RGB = RGB(200, 200, 200)
    cell.Borders(ppBorderRight).ForeColor.RGB = RGB(200, 200, 200)
End Sub

Private Sub StyleViewerBodyCell(ByVal cell As cell)
    With cell.Shape.TextFrame.TextRange
        .Font.Bold = False
        .Font.name = "Segoe UI"
        .Font.Size = 11
        .ParagraphFormat.Alignment = ppAlignLeft
    End With
    cell.Shape.Fill.Visible = msoFalse
    cell.Borders(ppBorderTop).ForeColor.RGB = RGB(220, 220, 220)
    cell.Borders(ppBorderBottom).ForeColor.RGB = RGB(220, 220, 220)
    cell.Borders(ppBorderLeft).ForeColor.RGB = RGB(220, 220, 220)
    cell.Borders(ppBorderRight).ForeColor.RGB = RGB(220, 220, 220)
End Sub

Public Function HasLayoutByName(ByVal pres As Presentation, ByVal layoutName As String) As Boolean
    Dim d As Design, cl As CustomLayout
    For Each d In pres.Designs
        For Each cl In d.SlideMaster.CustomLayouts
            ' Match on substring rather than exact name, so "AutoFocusMaster (2)" counts as found
            If InStr(1, cl.name, layoutName, vbTextCompare) > 0 Then
                HasLayoutByName = True
                Exit Function
            End If
        Next cl
    Next d
End Function

Private Sub LoadDesignIfMissing(ByVal pres As Presentation, ByVal fullPath As String, ByVal layoutName As String)
    On Error Resume Next
    pres.Designs.Load fullPath   ' adds the design to pres.Designs, does not apply
    On Error GoTo 0
End Sub

Private Function HasDesignWithMasterLike(ByVal pres As Presentation, _
                                         ByVal masterName As String) As Boolean
    Dim d As Design
    For Each d In pres.Designs
        On Error Resume Next
        Dim smn As String: smn = d.SlideMaster.name
        On Error GoTo 0
        If Len(smn) > 0 Then
            ' Catch "AutoFocusMaster", "1_AutoFocusMaster", etc.
            If InStr(1, smn, masterName, vbTextCompare) > 0 _
               Or InStr(1, Replace$(smn, "_", " "), masterName, vbTextCompare) > 0 Then
                HasDesignWithMasterLike = True
                Exit Function
            End If
        End If
    Next d
End Function

Public Sub EnsureAutoFocusMaster( _
    ByVal pres As Presentation, _
    Optional ByVal masterName As String = "AutoFocusMaster", _
    Optional ByVal masterFile As String = "")

    On Error GoTo CleanExit
    If pres Is Nothing Then Exit Sub

    ' Resolve source path
    Dim srcPath As String
    If Len(masterFile) > 0 Then
        srcPath = masterFile
    Else
        ' Uses your constant if present; else construct from %APPDATA%
        #If Mac Then
            srcPath = Replace("%APPDATA%\Microsoft\AddIns\TQ FG Slides\TQ FG Slidemaster.pptx", "%APPDATA%", Environ$("APPDATA"))
        #Else
            srcPath = Replace("%APPDATA%\Microsoft\AddIns\TQ FG Slides\TQ FG Slidemaster.pptx", "%APPDATA%", Environ$("APPDATA"))
        #End If
        If Len(Dir$(srcPath)) = 0 Then Exit Sub
    End If

    ' Quick fast-path: if we've marked this deck before AND still see the master, bail.
    If pres.Tags("AutoFocusMasterLoaded") = "1" Then
        If HasMasterLike(pres, masterName) Then Exit Sub
    End If

    ' If already present, tag and exit.
    If HasMasterLike(pres, masterName) Then
        pres.Tags.Add "AutoFocusMasterLoaded", "1"
        Exit Sub
    End If

    ' -------- Attempt 1: non-destructive load into Designs --------
    On Error Resume Next
    pres.Designs.Load srcPath      ' appends a design; does not re-template
    On Error GoTo 0

    If HasMasterLike(pres, masterName) Then
        pres.Tags.Add "AutoFocusMasterLoaded", "1"
        Exit Sub
    End If

    ' -------- Attempt 2 (fallback): insert one slide that uses the master, then delete it --------
    Dim src As Presentation
    On Error Resume Next
    Set src = Presentations.Open(srcPath, msoFalse, msoTrue, msoFalse)  ' ReadOnly, no window
    On Error GoTo 0
    If Not src Is Nothing Then
        Dim s As Slide, hitIdx As Long: hitIdx = -1

        ' Find a source slide whose master/group name contains masterName
        For Each s In src.Slides
            If Not s.CustomLayout Is Nothing Then
                Dim mName As String
                On Error Resume Next
                mName = s.CustomLayout.Parent.name  ' Parent is the Master
                On Error GoTo 0
                If Len(mName) > 0 Then
                    If InStr(1, mName, masterName, vbTextCompare) > 0 _
                    Or InStr(1, Replace$(mName, "_", " "), masterName, vbTextCompare) > 0 Then
                        hitIdx = s.SlideIndex
                        Exit For
                    End If
                End If
            End If
        Next s

        If hitIdx <> -1 Then
            Dim insertAfter As Long
            insertAfter = IIf(pres.Slides.Count > 0, pres.Slides.Count, 0)
            pres.Slides.InsertFromFile src.fullName, insertAfter, hitIdx, hitIdx
            ' Delete the imported slide; the design stays attached to the deck.
            pres.Slides(insertAfter + 1).Delete
        End If

        src.Close
    End If

    ' Final presence check; tag for future runs
    If HasMasterLike(pres, masterName) Then
        pres.Tags.Add "AutoFocusMasterLoaded", "1"
    End If

CleanExit:
    ' nothing
End Sub

' True if any Design's SlideMaster or any CustomLayout name contains masterName.
Private Function HasMasterLike(ByVal pres As Presentation, ByVal masterName As String) As Boolean
    Dim d As Design, cl As CustomLayout
    For Each d In pres.Designs
        On Error Resume Next
        Dim smn As String: smn = d.SlideMaster.name
        On Error GoTo 0
        If Len(smn) > 0 Then
            If InStr(1, smn, masterName, vbTextCompare) > 0 _
            Or InStr(1, Replace$(smn, "_", " "), masterName, vbTextCompare) > 0 Then
                HasMasterLike = True
                Exit Function
            End If
        End If
        For Each cl In d.SlideMaster.CustomLayouts
            If InStr(1, cl.name, masterName, vbTextCompare) > 0 Then
                HasMasterLike = True
                Exit Function
            End If
        Next cl
    Next d
End Function



