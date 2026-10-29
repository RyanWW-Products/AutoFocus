Attribute VB_Name = "GroupArrays"
Option Explicit
' Global dictionary of answer groups
Public AnswerGroups As Object

' File location used to persist custom answer groups across sessions.
' The file is saved into the same Add-Ins folder as the templates.  Each
' line has the format: GroupName:Item1|Item2|Item3...

 ' `ANSWER_GROUPS_FILE_NAME` and `AnswerGroupsFilePath()`
 ' To support dynamic persistence of answer groups and avoid constant-value compilation errors, we define a filename constant and a helper
 ' function that builds the full path to the answer group file.  If you rename or move the location
 ' for storing custom answer groups, update `ANSWER_GROUPS_FILE_NAME` accordingly.

Private Const ANSWER_GROUPS_FILE_NAME As String = "answer_groups.txt"

Private Function AnswerGroupsFilePath() As String
    AnswerGroupsFilePath = Environ$("APPDATA") & _
        "\Microsoft\AddIns\TQ FG Slides\" & ANSWER_GROUPS_FILE_NAME
End Function

Public Sub InitializeAnswerGroups()
    ' Ensure the AnswerGroups dictionary exists
    If AnswerGroups Is Nothing Then
        Set AnswerGroups = CreateObject("Scripting.Dictionary")
    End If
    ' Always reset to the built-in defaults
    AnswerGroups.RemoveAll
    AnswerGroups("Likely") = Array("Very Unlikely", "Unlikely", "Likely", "Very Likely")
    AnswerGroups("High") = Array("Way Too Low", "Too Low", "About Right", "Too High", "Way Too High")
    AnswerGroups("Agree") = Array("Strongly Disagree", "Somewhat Disagree", "Somewhat Agree", "Strongly Agree")
    AnswerGroups("PLong") = Array("Strongly Favor the Plaintiff", "Somewhat Favor the Plaintiff", "Somewhat Favor the Defendant", "Strongly Favor the Defendant")
    AnswerGroups("PShort") = Array("P+", "P", "D+", "D")
    AnswerGroups("YesNo") = Array("Yes", "No")
    ' Load any persisted custom groups from disk.  These will override or
    ' supplement the defaults.  If the file does not exist, nothing is loaded.
    LoadAnswerGroups
End Sub

' Retrieve a group by name; returns empty array if not found
Public Function GetAnswerGroup(ByVal groupName As String) As Variant
    InitializeAnswerGroups
    groupName = Trim$(groupName)
    If AnswerGroups.Exists(groupName) Then
        GetAnswerGroup = AnswerGroups(groupName)
    Else
        GetAnswerGroup = Array()
    End If
End Function

' Menu to manage answer groups: view, add, edit, remove
Public Sub ManageAnswerGroups(control As IRibbonControl)
    InitializeAnswerGroups
    Dim choice As String
    Do
        choice = InputBox("Answer Group Manager:" & vbCrLf & _
                         "1. View groups" & vbCrLf & _
                         "2. Add group" & vbCrLf & _
                         "3. Edit group" & vbCrLf & _
                         "4. Remove group" & vbCrLf & _
                         "0. Cancel" & vbCrLf & _
                         "Enter your choice (0-4):", "Manage Answer Groups")
        If choice = "" Or choice = "0" Then Exit Sub
        Select Case Trim$(choice)
            Case "1"
                Dim info As String
                info = "Current Answer Groups:" & vbCrLf
                Dim k As Variant
                For Each k In AnswerGroups.keys
                    ' Join group items with a comma and space for display
                    info = info & k & " -> " & Join(AnswerGroups(k), ", ") & vbCrLf
                Next k
                MsgBox info, vbInformation, "Answer Groups"
            Case "2"
                Dim newName As String, newItems As String
                newName = InputBox("Enter new group name (no spaces):", "Add Group")
                If newName = "" Then GoTo ContinueLoop
                If AnswerGroups.Exists(newName) Then
                    MsgBox "Group already exists: " & newName, vbExclamation
                    GoTo ContinueLoop
                End If
                newItems = InputBox("Enter items separated by commas:", "Add Group")
                If newItems = "" Then GoTo ContinueLoop
                Dim arr() As String, i As Long
                ' split items by comma
                arr = Split(newItems, ",")
                For i = 0 To UBound(arr)
                    arr(i) = Trim$(arr(i))
                Next i
                AnswerGroups(newName) = arr
                ' Persist changes to disk
                SaveAnswerGroups
                MsgBox "Group added: " & newName, vbInformation
            Case "3"
                Dim editName As String, editItems As String
                editName = InputBox("Enter the group name to edit:", "Edit Group")
                If editName = "" Then GoTo ContinueLoop
                If Not AnswerGroups.Exists(editName) Then
                    MsgBox "Group does not exist: " & editName, vbExclamation
                    GoTo ContinueLoop
                End If
                editItems = InputBox("Enter new items for " & editName & " separated by commas:", "Edit Group")
                If editItems = "" Then GoTo ContinueLoop
                Dim arr2() As String, j As Long
                ' split items by comma
                arr2 = Split(editItems, ",")
                For j = 0 To UBound(arr2)
                    arr2(j) = Trim$(arr2(j))
                Next j
                AnswerGroups(editName) = arr2
                ' Persist changes to disk
                SaveAnswerGroups
                MsgBox "Group updated: " & editName, vbInformation
            Case "4"
                Dim delName As String
                delName = InputBox("Enter the group name to remove:", "Remove Group")
                If delName = "" Then GoTo ContinueLoop
                If Not AnswerGroups.Exists(delName) Then
                    MsgBox "Group does not exist: " & delName, vbExclamation
                    GoTo ContinueLoop
                End If
                AnswerGroups.Remove delName
                ' Persist changes to disk
                SaveAnswerGroups
                MsgBox "Group removed: " & delName, vbInformation
        End Select
ContinueLoop:
    Loop
End Sub


' Shared group arrays to be used across all importers.

' Returns the four-point Likely group: Very Unlikely ? Unlikely ? Likely ? Very Likely
Public Function LikelyGroup() As Variant
    InitializeAnswerGroups
    If Not AnswerGroups Is Nothing Then
        If AnswerGroups.Exists("Likely") Then
            LikelyGroup = AnswerGroups("Likely")
            Exit Function
        End If
    End If
    LikelyGroup = Array("Very Unlikely", "Unlikely", "Likely", "Very Likely")
End Function

' Returns the five-point High/Low group: Way Too Low ? Too Low ? About Right ? Too High ? Way Too High
Public Function HighGroup() As Variant
    InitializeAnswerGroups
    If Not AnswerGroups Is Nothing Then
        If AnswerGroups.Exists("High") Then
            HighGroup = AnswerGroups("High")
            Exit Function
        End If
    End If
    HighGroup = Array("Way Too Low", "Too Low", "About Right", "Too High", "Way Too High")
End Function

' Returns the four-point Agree/Disagree group: Strongly Disagree ? Somewhat Disagree ? Somewhat Agree ? Strongly Agree
Public Function AgreeGroup() As Variant
    InitializeAnswerGroups
    If Not AnswerGroups Is Nothing Then
        If AnswerGroups.Exists("Agree") Then
            AgreeGroup = AnswerGroups("Agree")
            Exit Function
        End If
    End If
    AgreeGroup = Array("Strongly Disagree", "Somewhat Disagree", "Somewhat Agree", "Strongly Agree")
End Function

' Returns the four-point P-long group (long labels for P/D codes):
' Strongly Favor the Plaintiff ? Somewhat Favor the Plaintiff ? Somewhat Favor the Defendant ? Strongly Favor the Defendant
Public Function PLongGroup() As Variant
    InitializeAnswerGroups
    If Not AnswerGroups Is Nothing Then
        If AnswerGroups.Exists("PLong") Then
            PLongGroup = AnswerGroups("PLong")
            Exit Function
        End If
    End If
    PLongGroup = Array("Strongly Favor the Plaintiff", "Somewhat Favor the Plaintiff", "Somewhat Favor the Defendant", "Strongly Favor the Defendant")
End Function

' Returns the four-point P-short group (shorthand P/D codes): P+ ? P ? D+ ? D
Public Function PShortGroup() As Variant
    InitializeAnswerGroups
    If Not AnswerGroups Is Nothing Then
        If AnswerGroups.Exists("PShort") Then
            PShortGroup = AnswerGroups("PShort")
            Exit Function
        End If
    End If
    PShortGroup = Array("P+", "P", "D+", "D")
End Function

' Returns the Yes/No two-point group.
Public Function YesNoGroup() As Variant
    InitializeAnswerGroups
    If Not AnswerGroups Is Nothing Then
        If AnswerGroups.Exists("YesNo") Then
            YesNoGroup = AnswerGroups("YesNo")
            Exit Function
        End If
    End If
    YesNoGroup = Array("Yes", "No")
End Function

'----------------------------------------------------------------------
' Persist the AnswerGroups dictionary to a text file.  Each line in the
' file has the format: GroupName:Item1|Item2|Item3.  The file lives in
' the Add-Ins folder defined by ANSWER_GROUPS_FILE.
Private Sub SaveAnswerGroups()
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Ensure containing directory exists
    Dim folderPath As String
    ' Determine the file path using the AnswerGroupsFilePath() helper
    folderPath = fso.GetParentFolderName(AnswerGroupsFilePath())
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    Dim f As Integer
    f = FreeFile
    ' Write to the AnswerGroups file path
    Open AnswerGroupsFilePath() For Output As #f
    Dim k As Variant
    If Not AnswerGroups Is Nothing Then
        For Each k In AnswerGroups.keys
            Dim arr As Variant
            arr = AnswerGroups(k)
            Dim line As String
            line = CStr(k) & ":" & Join(arr, "|")
            Print #f, line
        Next k
    End If
    Close #f
    On Error GoTo 0
End Sub

'----------------------------------------------------------------------
' Load persisted AnswerGroups from the file defined by ANSWER_GROUPS_FILE.
' Existing entries in AnswerGroups will be overridden by loaded values.
Private Sub LoadAnswerGroups()
    On Error Resume Next
    ' If the file doesn't exist, nothing to load
    If Len(Dir$(AnswerGroupsFilePath())) = 0 Then Exit Sub
    Dim f As Integer
    f = FreeFile
    ' Open the persisted file for reading
    Open AnswerGroupsFilePath() For Input As #f
    Dim line As String
    Do While Not EOF(f)
        Line Input #f, line
        line = Trim$(line)
        If Len(line) > 0 Then
            Dim pos As Long
            pos = InStr(line, ":")
            If pos > 0 Then
                Dim name As String
                name = Trim$(Left$(line, pos - 1))
                Dim items As String
                items = Mid$(line, pos + 1)
                Dim arr() As String
                arr = Split(items, "|")
                Dim i As Long
                For i = 0 To UBound(arr)
                    arr(i) = Trim$(arr(i))
                Next i
                If Not AnswerGroups Is Nothing Then
                    AnswerGroups(name) = arr
                End If
            End If
        End If
    Loop
    Close #f
    On Error GoTo 0
End Sub

'----------------------------------------------------------------------
' RemoveExtraAutoFocusMasters
'
' Walk through the presentation's designs and ensure only one
' AutoFocusMaster design remains.  Any additional designs whose slide
' master name contains "AutoFocusMaster" will be removed, and slides
' using those designs will be reassigned to the base AutoFocusMaster
' design.  This prevents PowerPoint from adding a new duplicate
' master every time a slide is inserted from a template file.
Public Sub RemoveExtraAutoFocusMasters(ByVal pres As Presentation)
    If pres Is Nothing Then Exit Sub
    Dim baseIdx As Long: baseIdx = 0
    Dim i As Long
    Dim smn As String
    ' Find the first design whose SlideMaster name contains "AutoFocusMaster"
    For i = 1 To pres.Designs.Count
        smn = ""
        On Error Resume Next
        smn = pres.Designs(i).SlideMaster.name
        On Error GoTo 0
        If Len(smn) > 0 Then
            If InStr(1, smn, "AutoFocusMaster", vbTextCompare) > 0 Then
                baseIdx = i
                Exit For
            End If
        End If
    Next i
    ' If no AutoFocusMaster design exists, nothing to do
    If baseIdx = 0 Then Exit Sub
    ' Walk backwards through designs and remove any additional AutoFocusMaster
    For i = pres.Designs.Count To 1 Step -1
        smn = ""
        On Error Resume Next
        smn = pres.Designs(i).SlideMaster.name
        On Error GoTo 0
        If i <> baseIdx And Len(smn) > 0 Then
            If InStr(1, smn, "AutoFocusMaster", vbTextCompare) > 0 Then
                ' Reassign slides that currently use this duplicate design to the base design
                Dim sld As Slide
                For Each sld In pres.Slides
                    On Error Resume Next
                    If sld.Design.name = pres.Designs(i).name Then
                        sld.Design = pres.Designs(baseIdx)
                    End If
                    On Error GoTo 0
                Next sld
                ' Delete the duplicate design
                On Error Resume Next
                pres.Designs(i).Delete
                On Error GoTo 0
            End If
        End If
    Next i
End Sub



