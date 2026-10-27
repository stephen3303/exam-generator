Option Compare Database
Option Explicit

' Admin: Rules and lists (see Build_pgAdmRules). On the left the overlap
' rule (tblSettings OverlapMaxPct and OverlapWindowExams) and the questions
' a new topic starts with (DefaultQuestionCount). Moving a slider lists the
' topics that would be short of approved questions under the new values
' (modAdmin.TopicsShortText) before anything is saved; Save writes all three.
'
' On the right the difficulty levels and classification levels. Add asks
' for the new entry; More, or a right-click on a row, renames it, moves it,
' turns a difficulty level off or on, or deletes an entry nothing uses yet
' (one in use can be turned off or renamed instead). Classification levels
' are listed lowest first: their rank decides an exam's overall marking.

Private Const MSG_TITLE As String = "Rules and lists"

Private m_dirty As Boolean      ' the rule has changes not saved yet
Private m_data As Boolean       ' a data file is linked

Private Sub Form_Load()
    Dim c As Variant
    On Error GoTo Err_Handler
    m_data = DataLinked()
    If m_data Then
        Me.sfDiff.Form.RecordSource = DiffSql()
        Me.sfClass.Form.RecordSource = ClassSql()
        LoadRule
    Else
        ShowAlert Me, "alShort", "This copy was built without the data file, so there are no settings to show.", "warning", _
            "No data file linked"
        For Each c In Array("btnDiffAdd", "btnDiffMore", "btnClassAdd", "btnClassMore")
            SetEnabled Me.Controls(CStr(c)), False
        Next
        ShowState
    End If
    Exit Sub
Err_Handler:
    HandleError "pgAdmRules.Form_Load"
End Sub

' frmMain asks before it swaps or closes this page.
Public Function CanLeave() As Boolean
    On Error GoTo Err_Handler
    If Not m_dirty Then CanLeave = True: Exit Function
    Select Case KitMsgBox("Save your changes to the overlap rule?", vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbYes: CanLeave = SaveRule()
        Case vbNo: CanLeave = True
    End Select
    Exit Function
Err_Handler:
    HandleError "pgAdmRules.CanLeave"
End Function

' Ctrl+S saves the rule (modKitKeys).
Public Function Shortcut(ByVal action As String) As Boolean
    On Error GoTo Err_Handler
    If action = SHORTCUT_SAVE Then
        If m_dirty Then SaveRule
        Shortcut = True
    End If
    Exit Function
Err_Handler:
    HandleError "pgAdmRules.Shortcut", action
    Shortcut = True
End Function

' Left and Right on a slider move it a step (the slider's steps are
' buttons, which take no arrow keys of their own).
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim nm As String
    On Error Resume Next                    ' no control has the focus
    nm = Me.ActiveControl.Name
    On Error GoTo Err_Handler
    If KeyCode <> vbKeyLeft And KeyCode <> vbKeyRight Then Exit Sub
    If Left$(nm, 8) = "sldPct_b" Or Left$(nm, 11) = "sldWindow_b" Then
        SliderStep Me, Left$(nm, InStr(nm, "_b") - 1), IIf(KeyCode = vbKeyLeft, -1, 1)
        KeyCode = 0
    End If
    Exit Sub
Err_Handler:
    HandleError "pgAdmRules.Form_KeyDown", nm
End Sub

' A slider moved (its onChange).
Public Sub RuleChanged()
    On Error GoTo Err_Handler
    If Not m_data Then Exit Sub
    Edited
    ShowShort
    Exit Sub
Err_Handler:
    HandleError "pgAdmRules.RuleChanged"
End Sub

Private Sub txtDefault_Change()
    On Error GoTo Err_Handler
    Edited
    Exit Sub
Err_Handler:
    HandleError "pgAdmRules.txtDefault_Change"
End Sub

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    SaveRule
    Exit Sub
Err_Handler:
    HandleError "pgAdmRules.btnSave_Click"
End Sub

Private Sub btnUndo_Click()
    On Error GoTo Err_Handler
    LoadRule
    Exit Sub
Err_Handler:
    HandleError "pgAdmRules.btnUndo_Click"
End Sub

' ---------------------------------------------------------------------------
' The overlap rule
' ---------------------------------------------------------------------------

Private Sub LoadRule()
    SetSlider Me, "sldPct", SettingLong(SET_OVERLAP_PCT, 30)
    SetSlider Me, "sldWindow", SettingLong(SET_OVERLAP_WINDOW, 10)
    Me.txtDefault = DefaultQuestionCount()
    m_dirty = False
    ShowShort
    ShowState
End Sub

Private Sub Edited()
    If m_dirty Or Not m_data Then Exit Sub
    m_dirty = True
    ShowState
End Sub

Private Sub ShowState()
    SetEnabled Me.btnSave, m_dirty
    SetEnabled Me.btnUndo, m_dirty
End Sub

' The topics short of approved questions under the sliders' values.
Private Sub ShowShort()
    Dim s As String, n As Long
    s = TopicsShortText(SliderValue(Me, "sldPct"), SliderValue(Me, "sldWindow"), n, 5)
    If n = 0 Then
        ShowAlert Me, "alShort", "Every active topic has enough approved questions for exams under these values.", _
            "success", "No topics short"
    Else
        ShowAlert Me, "alShort", s, "warning", n & IIf(n = 1, " topic is", " topics are") & " short of approved questions"
    End If
End Sub

' Save the three settings. False when the default count isn't a whole
' number from 1 to 50 (the admin is told).
Private Function SaveRule() As Boolean
    Dim t As String
    t = Trim$(FieldText(Me.txtDefault))
    If Not IsWhole(t, 1, 50) Then
        KitMsgBox "Questions per topic to start with must be a whole number from 1 to 50.", vbExclamation, MSG_TITLE
        Me.txtDefault.SetFocus
        Exit Function
    End If
    SetSetting SET_OVERLAP_PCT, CStr(SliderValue(Me, "sldPct"))
    SetSetting SET_OVERLAP_WINDOW, CStr(SliderValue(Me, "sldWindow"))
    SetSetting SET_DEFAULT_COUNT, CStr(CLng(t))
    Me.txtDefault = CLng(t)
    m_dirty = False
    ShowState
    KitToast "Saved. New exams follow the new rule.", "success"
    SaveRule = True
End Function

' True when t is a whole number from lo to hi.
Private Function IsWhole(ByVal t As String, ByVal lo As Long, ByVal hi As Long) As Boolean
    If Len(t) = 0 Or Len(t) > 4 Then Exit Function
    If t Like "*[!0-9]*" Then Exit Function
    IsWhole = (CLng(t) >= lo And CLng(t) <= hi)
End Function

' A field's text: what is typed so far while it has the focus, else its value.
Private Function FieldText(c As Control) As String
    Dim act As String
    On Error Resume Next
    act = Me.ActiveControl.Name
    On Error GoTo 0
    If act = c.Name Then FieldText = Nz(c.Text, "") Else FieldText = Nz(c.Value, "")
End Function

' ---------------------------------------------------------------------------
' The lists
' ---------------------------------------------------------------------------

Private Sub btnDiffAdd_Click()
    Dim v As Variant, problem As String
    On Error GoTo Err_Handler
    v = KitInputBox("Name of the new difficulty level:", "Difficulty levels")
    If IsNull(v) Then Exit Sub
    problem = NameProblem(v, 30, "tblDifficultyLevels", "LevelName", "difficulty level")
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Sub
    KitExecute "INSERT INTO tblDifficultyLevels (LevelName, SortOrder, IsActive) VALUES (" & Q(Trim$(v)) & ", " & _
        NextOrder("tblDifficultyLevels", "SortOrder") & ", True)"
    ListChanged Me.sfDiff, DLookup("DifficultyID", "tblDifficultyLevels", "LevelName=" & Q(Trim$(v)))
    KitToast "Level added.", "success"
    Exit Sub
Err_Handler:
    HandleError "pgAdmRules.btnDiffAdd_Click"
End Sub

Private Sub btnClassAdd_Click()
    Dim m As Variant, pm As Variant, problem As String
    On Error GoTo Err_Handler
    m = KitInputBox("Marking of the new classification level (as an exam's banner shows it):", "Classification levels")
    If IsNull(m) Then Exit Sub
    problem = NameProblem(m, 30, "tblClassLevels", "Marking", "classification level")
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Sub
    pm = KitInputBox("Its portion mark (as it shows before a question, without the brackets):", "Classification levels")
    If IsNull(pm) Then Exit Sub
    problem = NameProblem(pm, 10, "tblClassLevels", "PortionMark", "portion mark")
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Sub
    KitExecute "INSERT INTO tblClassLevels (Marking, PortionMark, LevelRank) VALUES (" & Q(Trim$(m)) & ", " & _
        Q(Trim$(pm)) & ", " & NextOrder("tblClassLevels", "LevelRank") & ")"
    ListChanged Me.sfClass, DLookup("ClassLevelID", "tblClassLevels", "Marking=" & Q(Trim$(m)))
    KitToast "Level added at the top. Move it down if it ranks lower.", "success"
    Exit Sub
Err_Handler:
    HandleError "pgAdmRules.btnClassAdd_Click"
End Sub

Private Sub btnDiffMore_Click()
    On Error GoTo Err_Handler
    DiffActions False
    Exit Sub
Err_Handler:
    HandleError "pgAdmRules.btnDiffMore_Click"
End Sub

Private Sub btnClassMore_Click()
    On Error GoTo Err_Handler
    ClassActions False
    Exit Sub
Err_Handler:
    HandleError "pgAdmRules.btnClassMore_Click"
End Sub

' Row click in a list (subGrid): the row is selected; nothing else opens.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
End Sub

' Right-click on a row (subGrid): its actions at the pointer.
Public Sub GridRightClick(ByVal grid As String)
    On Error GoTo Err_Handler
    If grid = "sfDiff" Then DiffActions True Else ClassActions True
    Exit Sub
Err_Handler:
    HandleError "pgAdmRules.GridRightClick", grid
End Sub

' The selected difficulty level's actions.
Private Sub DiffActions(ByVal atPointer As Boolean)
    Dim id As Long, nm As String, isOn As Boolean, used As Long, k As String, v As Variant, problem As String
    Dim a(5) As String
    id = Nz(Me.sfDiff.Form.txtSel, 0)
    If id = 0 Then KitToast "Pick a level in the list first.", "info": Exit Sub
    nm = Nz(DLookup("LevelName", "tblDifficultyLevels", "DifficultyID=" & id), "")
    isOn = Nz(DLookup("IsActive", "tblDifficultyLevels", "DifficultyID=" & id), False)
    a(0) = "ren|Rename|E70F"
    a(1) = "up|Move up|E70E"
    a(2) = "down|Move down|E70D"
    a(3) = "onoff|" & IIf(isOn, "Turn off", "Turn on") & "|E7E8"
    a(4) = "-"
    a(5) = "del|!Delete|E74D"
    If atPointer Then k = KitMenuAt(a(0), a(1), a(2), a(3), a(4), a(5)) Else k = KitMenu(Me.btnDiffMore, a(0), a(1), a(2), a(3), a(4), a(5))
    Select Case k
        Case "ren"
            v = KitInputBox("New name for """ & nm & """:", "Difficulty levels", nm)
            If IsNull(v) Then Exit Sub
            problem = NameProblem(v, 30, "tblDifficultyLevels", "LevelName", "difficulty level", "DifficultyID=" & id)
            If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Sub
            KitExecute "UPDATE tblDifficultyLevels SET LevelName=" & Q(Trim$(v)) & " WHERE DifficultyID=" & id
        Case "up", "down"
            If Not MoveRow("tblDifficultyLevels", "DifficultyID", "SortOrder", id, IIf(k = "up", -1, 1)) Then Exit Sub
        Case "onoff"
            KitExecute "UPDATE tblDifficultyLevels SET IsActive=" & IIf(isOn, "False", "True") & " WHERE DifficultyID=" & id
            KitToast IIf(isOn, "Turned off. Questions that have it keep it; new ones can't pick it.", "Turned on."), "success"
        Case "del"
            used = DCount("*", "tblQuestions", "DifficultyID=" & id)
            If used > 0 Then
                KitMsgBox """" & nm & """ is on " & used & IIf(used = 1, " question", " questions") & ", so it can't be deleted. " & _
                    "Turn it off instead: those questions keep it, and new ones can't pick it.", vbInformation, MSG_TITLE
                Exit Sub
            End If
            If Not KitConfirmTyped("Delete the difficulty level """ & nm & """?", "No question uses it. This can't be undone.", _
                PURGE_WORD, "Delete", "Difficulty levels") Then Exit Sub
            KitExecute "DELETE FROM tblDifficultyLevels WHERE DifficultyID=" & id
            ListChanged Me.sfDiff, Null
            KitToast "Level deleted.", "destructive"
            Exit Sub
        Case Else
            Exit Sub
    End Select
    ListChanged Me.sfDiff, id
End Sub

' The selected classification level's actions.
Private Sub ClassActions(ByVal atPointer As Boolean)
    Dim id As Long, nm As String, pm As String, used As Long, k As String, v As Variant, problem As String
    Dim a(5) As String
    id = Nz(Me.sfClass.Form.txtSel, 0)
    If id = 0 Then KitToast "Pick a level in the list first.", "info": Exit Sub
    nm = Nz(DLookup("Marking", "tblClassLevels", "ClassLevelID=" & id), "")
    pm = Nz(DLookup("PortionMark", "tblClassLevels", "ClassLevelID=" & id), "")
    a(0) = "ren|Change marking|E70F"
    a(1) = "pm|Change portion mark|E70F"
    a(2) = "up|Move up (ranks lower)|E70E"
    a(3) = "down|Move down (ranks higher)|E70D"
    a(4) = "-"
    a(5) = "del|!Delete|E74D"
    If atPointer Then k = KitMenuAt(a(0), a(1), a(2), a(3), a(4), a(5)) Else k = KitMenu(Me.btnClassMore, a(0), a(1), a(2), a(3), a(4), a(5))
    Select Case k
        Case "ren"
            v = KitInputBox("New marking for """ & nm & """:", "Classification levels", nm)
            If IsNull(v) Then Exit Sub
            problem = NameProblem(v, 30, "tblClassLevels", "Marking", "classification level", "ClassLevelID=" & id)
            If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Sub
            KitExecute "UPDATE tblClassLevels SET Marking=" & Q(Trim$(v)) & " WHERE ClassLevelID=" & id
        Case "pm"
            v = KitInputBox("New portion mark for """ & nm & """:", "Classification levels", pm)
            If IsNull(v) Then Exit Sub
            problem = NameProblem(v, 10, "tblClassLevels", "PortionMark", "portion mark", "ClassLevelID=" & id)
            If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Sub
            KitExecute "UPDATE tblClassLevels SET PortionMark=" & Q(Trim$(v)) & " WHERE ClassLevelID=" & id
        Case "up", "down"
            If Not MoveRow("tblClassLevels", "ClassLevelID", "LevelRank", id, IIf(k = "up", -1, 1)) Then Exit Sub
        Case "del"
            used = DCount("*", "tblQuestions", "ClassLevelID=" & id) + DCount("*", "tblQuestionParts", "ClassLevelID=" & id) + _
                DCount("*", "tblExams", "OverallClassLevelID=" & id)
            If used > 0 Then
                KitMsgBox """" & nm & """ is used by questions or exams, so it can't be deleted. You can still change " & _
                    "its marking.", vbInformation, MSG_TITLE
                Exit Sub
            End If
            If Not KitConfirmTyped("Delete the classification level """ & nm & """?", _
                "No question, part or exam uses it. This can't be undone.", PURGE_WORD, "Delete", "Classification levels") Then Exit Sub
            KitExecute "DELETE FROM tblClassLevels WHERE ClassLevelID=" & id
            ListChanged Me.sfClass, Null
            KitToast "Level deleted.", "destructive"
            Exit Sub
        Case Else
            Exit Sub
    End Select
    ListChanged Me.sfClass, id
End Sub

' Reload a list and select id in it (Null: none).
Private Sub ListChanged(sf As Control, ByVal id As Variant)
    Dim g As Form
    Set g = sf.Form
    g.Requery
    g.txtSel = Null
    If Not IsNull(id) Then
        With g.RecordsetClone
            .FindFirst "id=" & id
            If Not .NoMatch Then g.Bookmark = .Bookmark: g.txtSel = id
        End With
    End If
    g.Recalc
End Sub

Private Function DiffSql() As String
    DiffSql = "SELECT DifficultyID AS id, LevelName, IIf(IsActive,'Active','Off') AS state, SortOrder " & _
        "FROM tblDifficultyLevels ORDER BY SortOrder, DifficultyID"
End Function

Private Function ClassSql() As String
    ClassSql = "SELECT ClassLevelID AS id, Marking, PortionMark, LevelRank FROM tblClassLevels ORDER BY LevelRank, ClassLevelID"
End Function
