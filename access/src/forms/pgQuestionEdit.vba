Option Compare Database
Option Explicit

' Question editor (see Build_pgQuestionEdit). Opened from the question bank
' with TempVars("EditQuestionID") (0 for a new question) and
' TempVars("EditMode") ("newversion" to start a new version of an approved
' question). Unbound: the question's own fields are on the page and its
' parts, answer options and references in the working copy (modQuestions),
' so nothing reaches the shared data file until Save, which checks every
' rule and writes it all in one transaction (SaveQuestion).
'
' A draft or rejected question can be changed; one in review, approved or
' retired opens read-only. An approved one is changed with More > Edit as
' new version, whose first Save retires the approved version. Submit for
' approval saves, checks the submission rules too, and sends it to the DH.
'
' Someone else's save while the editor is open shows as a newer audit row
' than the one the editor loaded (m_stamp); Save asks before overwriting it.

Private Const MSG_TITLE As String = "Question editor"

Private m_id As Long            ' the question being changed (0: new, or a new version not saved yet)
Private m_baseId As Long        ' the approved question a new version replaces (0 otherwise)
Private m_key As Long           ' its QuestionKey (0 for a new question)
Private m_ver As Long           ' its version number
Private m_status As String      ' its status name
Private m_readOnly As Boolean   ' it can't be changed (in review, approved, retired), or m_noRight
Private m_noRight As Boolean    ' this person may not change it (not its author, nor an SME on its topic)
Private m_stamp As Long         ' the newest audit row of the question when it was loaded
Private m_snap As String        ' the page's own fields when loaded or saved (IsDirty)
Private m_workDirty As Boolean  ' parts, options or references changed since
Private m_type As String        ' the type switch's key, to put it back
Private m_loading As Boolean

Public Property Get HelpTopic() As String
    HelpTopic = "qedit"
End Property

Private Sub Form_Load()
    Dim qid As Long, mode As String
    On Error GoTo Err_Handler
    If Not DataLinked() Then
        Me.lblSubtitle.Caption = "This copy was built without the data file, so there are no questions to open."
        m_readOnly = True
        ApplyReadOnly
        SetEnabled Me.btnMore, False
        Exit Sub
    End If
    EnsureWorkTables
    Me.sfParts.Form.RecordSource = WorkPartsSql()
    Me.sfOptions.Form.RecordSource = WorkOptionsSql()
    Me.sfRefs.Form.RecordSource = WorkRefsSql()
    qid = Nz(TempVars("EditQuestionID"), 0)
    mode = Nz(TempVars("EditMode"), "")
    LoadQuestion qid, mode = "newversion"
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.Form_Load", qid
End Sub

' frmMain asks before it swaps or closes this page: unsaved changes are
' saved, dropped or kept, as the user says.
Public Function CanLeave() As Boolean
    On Error GoTo Err_Handler
    If m_readOnly Or Not IsDirty() Then CanLeave = True: Exit Function
    Select Case KitMsgBox("Save your changes to " & Called() & " before you go?" & vbCrLf & vbCrLf & _
            "Yes saves it as a draft. No drops the changes.", vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbYes: CanLeave = SaveNow(False)
        Case vbNo: CanLeave = True
    End Select
    Exit Function
Err_Handler:
    HandleError "pgQuestionEdit.CanLeave"
End Function

' Ctrl+S saves, Ctrl+N starts a new question (modKitKeys).
Public Function Shortcut(ByVal action As String) As Boolean
    On Error GoTo Err_Handler
    Select Case action
        Case SHORTCUT_SAVE
            SaveClicked
            Shortcut = True
        Case SHORTCUT_NEW
            If CanLeave() Then LoadQuestion 0, False
            Shortcut = True
    End Select
    Exit Function
Err_Handler:
    HandleError "pgQuestionEdit.Shortcut", action
    Shortcut = True
End Function

Private Sub txtTopic_KeyDown(KeyCode As Integer, Shift As Integer)
    KitPickKey Me.txtTopic, KeyCode
End Sub

' ---------------------------------------------------------------------------
' Header buttons
' ---------------------------------------------------------------------------

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    SaveClicked
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnSave_Click", m_id
End Sub

Private Sub btnSubmit_Click()
    Dim problem As String, who As String
    On Error GoTo Err_Handler
    If m_readOnly Then Exit Sub
    If KitMsgBox("Submit " & Called() & " for approval?" & vbCrLf & vbCrLf & "It goes to the department head, then the XO " & _
            "and the senior member. You can't change it while it is in review; if it is rejected, it comes back to you.", _
            vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then Exit Sub
    If Not SaveNow(True) Then Exit Sub
    problem = SubmitQuestion(m_id)
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Sub
    who = RoleHolders("DH")
    LoadQuestion m_id, False
    If Len(who) = 0 Then
        KitMsgBox "Submitted. Nobody holds the department head role yet, so it waits until the admin gives someone " & _
            "that role on the People and routing page.", vbInformation, MSG_TITLE
    Else
        KitToast "Submitted to " & who & ".", "success"
    End If
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnSubmit_Click", m_id
End Sub

Private Sub btnBack_Click()
    On Error GoTo Err_Handler
    TempVars("BankQuestionID") = IIf(m_id <> 0, m_id, m_baseId)
    NavTo "questions"
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnBack_Click"
End Sub

' More: undo, a new question, a new version, history, delete a draft.
Private Sub btnMore_Click()
    Dim canNew As Boolean, canDel As Boolean
    On Error GoTo Err_Handler
    If m_id <> 0 Then canNew = (Len(NewVersionProblem(m_id)) = 0)
    If m_id <> 0 Then canDel = (Len(DeleteProblem(m_id)) = 0)
    If canNew Or canDel Then
        If Not CanWriteQuestion(m_id) Then canNew = False: canDel = False
    End If
    Select Case KitMenu(Me.btnMore, IIf(IsDirty() And Not m_readOnly, "", "~") & "undo|Undo changes|E7A7", _
            IIf(canNew, "", "~") & "newver|Edit as new version|E70F", "new|New question|E710", _
            IIf(m_key <> 0, "", "~") & "hist|History|E81C", "-", _
            IIf(canDel, "", "~") & "del|!Delete draft|E74D")
        Case "undo": UndoChanges
        Case "newver": LoadQuestion m_id, True
        Case "new": If CanLeave() Then LoadQuestion 0, False
        Case "hist": RunDialog "dlgHistory", "key=" & m_key
        Case "del": DeleteThis
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnMore_Click", m_id
End Sub

' ---------------------------------------------------------------------------
' The question's own fields
' ---------------------------------------------------------------------------

Private Sub cboDiff_AfterUpdate()
    Changed
End Sub

Private Sub cboMark_AfterUpdate()
    Changed
End Sub

Private Sub txtText_AfterUpdate()
    Changed
End Sub

Private Sub txtRefAns_AfterUpdate()
    Changed
End Sub

' Points: shown as 4 or 1.5 when a number, and the parts' total checked.
Private Sub txtPoints_AfterUpdate()
    On Error GoTo Err_Handler
    If IsNumeric(Me.txtPoints) Then Me.txtPoints = PointsText(Me.txtPoints)
    ShowPoints
    Changed
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.txtPoints_AfterUpdate"
End Sub

' The type switch moved (its onChange). True/false and short answer
' questions have no answer options: offer to remove them, or switch back.
Public Sub TypeChanged()
    Dim n As Long, k As String
    On Error GoTo Err_Handler
    k = SegmentValue(Me, "segType")
    n = DCount("*", "wq_option")
    If k <> "mc" And n > 0 Then
        If KitMsgBox("A " & LCase$(TypeOfKey(k)) & " question has no answer options. Remove its " & n & _
                IIf(n = 1, " option?", " options?"), vbYesNo + vbQuestion + vbDefaultButton2, MSG_TITLE) <> vbYes Then
            SetSegment Me, "segType", m_type
            Exit Sub
        End If
        CurrentDb.Execute "DELETE FROM wq_option", dbFailOnError
        m_workDirty = True
    End If
    m_type = k
    RefreshOptions 0
    Changed
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.TypeChanged", k
End Sub

' A field changed: the subtitle says there are unsaved changes.
Private Sub Changed()
    If m_loading Then Exit Sub
    ShowState
End Sub

' ---------------------------------------------------------------------------
' Tables
' ---------------------------------------------------------------------------

' A row clicked (subGrid): select it; a double-click changes it.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    ShowRowState
    If Not dbl Or m_readOnly Then Exit Sub
    Select Case grid
        Case "sfParts": EditPart
        Case "sfOptions": EditOption
        Case "sfRefs": EditRef
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.GridClick", grid
End Sub

' A right-click on a row: its actions.
Public Sub GridRightClick(ByVal grid As String)
    On Error GoTo Err_Handler
    If m_readOnly Then Exit Sub
    Select Case grid
        Case "sfParts"
            Select Case KitMenuAt("edit|Edit part|E70F", "up|Move up|E70E", "down|Move down|E70D", "-", "rem|!Remove part|E74D")
                Case "edit": EditPart
                Case "up": MovePart -1
                Case "down": MovePart 1
                Case "rem": RemovePart
            End Select
        Case "sfOptions"
            Select Case KitMenuAt("edit|Edit option|E70F", "ok|Mark correct|E73E", "up|Move up|E70E", "down|Move down|E70D", _
                    "-", "rem|!Remove option|E74D")
                Case "edit": EditOption
                Case "ok": MarkCorrect
                Case "up": MoveOption -1
                Case "down": MoveOption 1
                Case "rem": RemoveOption
            End Select
        Case "sfRefs"
            Select Case KitMenuAt("edit|Change paragraph|E70F", "-", "rem|!Stop citing|E74D")
                Case "edit": EditRef
                Case "rem": RemoveRef
            End Select
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.GridRightClick", grid
End Sub

Private Sub btnAddPart_Click()
    On Error GoTo Err_Handler
    AddPart
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnAddPart_Click"
End Sub

Private Sub btnEditPart_Click()
    On Error GoTo Err_Handler
    EditPart
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnEditPart_Click"
End Sub

Private Sub btnUpPart_Click()
    On Error GoTo Err_Handler
    MovePart -1
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnUpPart_Click"
End Sub

Private Sub btnDownPart_Click()
    On Error GoTo Err_Handler
    MovePart 1
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnDownPart_Click"
End Sub

Private Sub btnRemovePart_Click()
    On Error GoTo Err_Handler
    RemovePart
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnRemovePart_Click"
End Sub

Private Sub btnAddOption_Click()
    On Error GoTo Err_Handler
    AddOption
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnAddOption_Click"
End Sub

Private Sub btnEditOption_Click()
    On Error GoTo Err_Handler
    EditOption
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnEditOption_Click"
End Sub

Private Sub btnUpOption_Click()
    On Error GoTo Err_Handler
    MoveOption -1
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnUpOption_Click"
End Sub

Private Sub btnDownOption_Click()
    On Error GoTo Err_Handler
    MoveOption 1
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnDownOption_Click"
End Sub

Private Sub btnRemoveOption_Click()
    On Error GoTo Err_Handler
    RemoveOption
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnRemoveOption_Click"
End Sub

Private Sub btnCorrectOption_Click()
    On Error GoTo Err_Handler
    MarkCorrect
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnCorrectOption_Click"
End Sub

Private Sub btnAddRef_Click()
    On Error GoTo Err_Handler
    If m_readOnly Then Exit Sub
    If RunDialog("dlgReference", "mode=cite") Then
        m_workDirty = True
        RefreshRefs Nz(TempVars("WorkRefID"), 0)
        ShowState
    End If
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnAddRef_Click"
End Sub

Private Sub btnEditRef_Click()
    On Error GoTo Err_Handler
    EditRef
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnEditRef_Click"
End Sub

Private Sub btnRemoveRef_Click()
    On Error GoTo Err_Handler
    RemoveRef
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnRemoveRef_Click"
End Sub

Private Sub btnCopyHistory_Click()
    On Error GoTo Err_Handler
    If m_key = 0 Then Exit Sub
    If KitCopyText(HistoryText(m_key)) Then KitToast "History copied.", "success"
    Exit Sub
Err_Handler:
    HandleError "pgQuestionEdit.btnCopyHistory_Click", m_key
End Sub

' ---------------------------------------------------------------------------
' Parts
' ---------------------------------------------------------------------------

Private Sub AddPart()
    Dim n As Long, pid As Long, nOpts As Long
    If m_readOnly Then Exit Sub
    n = WorkPartCount()
    If n >= 26 Then KitMsgBox "A question can have at most 26 parts (A to Z).", vbInformation, MSG_TITLE: Exit Sub
    If Not RunDialog("dlgPart", "total=" & Nz(Me.txtPoints, "") & ";mark=" & Nz(Me.cboMark, "")) Then Exit Sub
    pid = Nz(TempVars("PartID"), 0)
    m_workDirty = True
    ' The first part: the question's own answer options belong to a part now.
    nOpts = DCount("*", "wq_option", "part_id Is Null")
    If n = 0 And nOpts > 0 Then
        If KitMsgBox("The question has " & nOpts & IIf(nOpts = 1, " answer option", " answer options") & " of its own. Once it has " & _
                "parts, each option belongs to a part. Move " & IIf(nOpts = 1, "it", "them") & " to part A?", _
                vbYesNo + vbQuestion, MSG_TITLE) = vbYes Then
            CurrentDb.Execute "UPDATE wq_option SET part_id=" & pid & " WHERE part_id Is Null", dbFailOnError
        End If
    End If
    RelabelParts
    RefreshParts pid
    RefreshOptions 0
    ShowState
End Sub

Private Sub EditPart()
    Dim pid As Long
    pid = SelId(Me.sfParts)
    If m_readOnly Or pid = 0 Then Exit Sub
    If Not RunDialog("dlgPart", "id=" & pid & ";total=" & Nz(Me.txtPoints, "")) Then Exit Sub
    m_workDirty = True
    RefreshParts pid
    ShowState
End Sub

' Move the selected part up (-1) or down (1); the parts are lettered again.
Private Sub MovePart(ByVal delta As Long)
    Dim pid As Long
    pid = SelId(Me.sfParts)
    If m_readOnly Or pid = 0 Then Exit Sub
    If Not MoveRow("wq_part", "id", "SortOrder", pid, delta) Then Exit Sub
    RelabelParts
    m_workDirty = True
    RefreshParts pid
    RefreshOptions 0
    ShowState
End Sub

' Remove the selected part and its answer options; the rest are lettered again.
Private Sub RemovePart()
    Dim pid As Long, lbl As String, nOpts As Long
    pid = SelId(Me.sfParts)
    If m_readOnly Or pid = 0 Then Exit Sub
    lbl = PartLabelOf(pid)
    nOpts = DCount("*", "wq_option", "part_id=" & pid)
    If KitMsgBox("Remove part " & lbl & "?" & IIf(nOpts > 0, " Its " & nOpts & IIf(nOpts = 1, " answer option goes", " answer options go") & _
            " with it.", "") & " The parts after it are lettered again.", vbYesNo + vbQuestion + vbDefaultButton2, MSG_TITLE) <> vbYes Then Exit Sub
    CurrentDb.Execute "DELETE FROM wq_option WHERE part_id=" & pid, dbFailOnError
    CurrentDb.Execute "DELETE FROM wq_part WHERE id=" & pid, dbFailOnError
    RelabelParts
    RenumberOptions
    m_workDirty = True
    RefreshParts 0
    RefreshOptions 0
    ShowState
End Sub

' ---------------------------------------------------------------------------
' Answer options
' ---------------------------------------------------------------------------

' Add an option to the question, or to the part picked beside the buttons.
Private Sub AddOption()
    Dim part As Variant, v As Variant, t As Variant, grp As String, n As Long, id As Long
    If m_readOnly Then Exit Sub
    If SegmentValue(Me, "segType") <> "mc" Then
        KitMsgBox "Only multiple choice questions have answer options. Pick Multiple choice under Type first.", vbInformation, MSG_TITLE
        Exit Sub
    End If
    part = Null
    If WorkPartCount() > 0 Then
        part = Me.cboOptPart
        If IsNull(part) Then
            KitMsgBox "Pick the part the option is for in the box beside the buttons.", vbInformation, MSG_TITLE
            Me.cboOptPart.SetFocus
            Exit Sub
        End If
    End If
    grp = GroupWhere(part)
    n = DCount("*", "wq_option", grp)
    If n >= 26 Then KitMsgBox "A question or part can have at most 26 options (A to Z).", vbInformation, MSG_TITLE: Exit Sub
    v = KitInputBox("Text of option " & OptionLetter(n + 1) & IIf(IsNull(part), "", " of part " & PartLabelOf(part)) & ":", "Add option")
    If IsNull(v) Then Exit Sub
    t = OptionTextOK(v)
    If IsNull(t) Then Exit Sub
    id = AddOptionRow(part, t, n + 1)
    m_workDirty = True
    RefreshOptions id
    ShowState
End Sub

Private Sub EditOption()
    Dim id As Long, v As Variant, t As Variant, rs As DAO.Recordset, cap As String
    id = SelId(Me.sfOptions)
    If m_readOnly Or id = 0 Then Exit Sub
    Set rs = CurrentDb.OpenRecordset("SELECT part_id, OptionText, SortOrder FROM wq_option WHERE id=" & id, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Sub
    cap = "Option " & OptionLetter(rs!SortOrder) & IIf(IsNull(rs!part_id), "", " of part " & PartLabelOf(rs!part_id))
    v = KitInputBox(cap & ":", "Change option", Nz(rs!OptionText, ""))
    rs.Close
    If IsNull(v) Then Exit Sub
    t = OptionTextOK(v)
    If IsNull(t) Then Exit Sub
    KitExecute "UPDATE wq_option SET OptionText=" & Q(t) & " WHERE id=" & id
    m_workDirty = True
    RefreshOptions id
    ShowState
End Sub

' Make the selected option its group's correct answer; on the correct one,
' take the mark off.
Private Sub MarkCorrect()
    Dim id As Long, part As Variant, isOk As Boolean
    id = SelId(Me.sfOptions)
    If m_readOnly Or id = 0 Then Exit Sub
    part = DLookup("part_id", "wq_option", "id=" & id)
    isOk = Nz(DLookup("IsCorrect", "wq_option", "id=" & id), False)
    If isOk Then
        KitExecute "UPDATE wq_option SET IsCorrect=False WHERE id=" & id
    Else
        KitExecute "UPDATE wq_option SET IsCorrect=(id=" & id & ") WHERE " & GroupWhere(part)
    End If
    m_workDirty = True
    RefreshOptions id
    ShowState
End Sub

Private Sub MoveOption(ByVal delta As Long)
    Dim id As Long
    id = SelId(Me.sfOptions)
    If m_readOnly Or id = 0 Then Exit Sub
    If Not MoveRow("wq_option", "id", "SortOrder", id, delta, GroupWhere(DLookup("part_id", "wq_option", "id=" & id))) Then Exit Sub
    m_workDirty = True
    RefreshOptions id
    ShowState
End Sub

Private Sub RemoveOption()
    Dim id As Long
    id = SelId(Me.sfOptions)
    If m_readOnly Or id = 0 Then Exit Sub
    If KitMsgBox("Remove this option? The options after it are lettered again.", vbYesNo + vbQuestion + vbDefaultButton2, _
            MSG_TITLE) <> vbYes Then Exit Sub
    CurrentDb.Execute "DELETE FROM wq_option WHERE id=" & id, dbFailOnError
    RenumberOptions
    m_workDirty = True
    RefreshOptions 0
    ShowState
End Sub

' The working copy's options in part's group (Null: the question's own).
Private Function GroupWhere(ByVal part As Variant) As String
    If IsNull(part) Then GroupWhere = "part_id Is Null" Else GroupWhere = "part_id=" & CLng(part)
End Function

' An option's text as saved, or Null (after saying why) when there is none
' or too much.
Private Function OptionTextOK(ByVal v As Variant) As Variant
    OptionTextOK = CleanText(v)
    If IsNull(OptionTextOK) Then
        KitMsgBox "An option needs its text.", vbInformation, MSG_TITLE
    ElseIf Len(OptionTextOK) > MAX_TEXT Then
        KitMsgBox "An option can be at most " & MAX_TEXT & " characters; this one is " & Len(OptionTextOK) & ".", vbInformation, MSG_TITLE
        OptionTextOK = Null
    End If
End Function

Private Function AddOptionRow(ByVal part As Variant, ByVal t As String, ByVal n As Long) As Long
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM wq_option WHERE False", dbOpenDynaset)
    rs.AddNew
    rs!part_id = part
    rs!OptionText = t
    rs!IsCorrect = False
    rs!SortOrder = n
    rs.Update
    rs.Bookmark = rs.LastModified
    AddOptionRow = rs!id
    rs.Close
End Function

' ---------------------------------------------------------------------------
' References
' ---------------------------------------------------------------------------

Private Sub EditRef()
    Dim id As Long
    id = SelId(Me.sfRefs)
    If m_readOnly Or id = 0 Then Exit Sub
    If Not RunDialog("dlgReference", "mode=cite;id=" & id) Then Exit Sub
    m_workDirty = True
    RefreshRefs id
    ShowState
End Sub

Private Sub RemoveRef()
    Dim id As Long
    id = SelId(Me.sfRefs)
    If m_readOnly Or id = 0 Then Exit Sub
    If KitMsgBox("Stop citing " & RefText(Nz(DLookup("RefID", "wq_ref", "id=" & id), 0)) & "? The reference stays in the list.", _
            vbYesNo + vbQuestion + vbDefaultButton2, MSG_TITLE) <> vbYes Then Exit Sub
    CurrentDb.Execute "DELETE FROM wq_ref WHERE id=" & id, dbFailOnError
    m_workDirty = True
    RefreshRefs 0
    ShowState
End Sub

' ---------------------------------------------------------------------------
' Load, save, delete
' ---------------------------------------------------------------------------

' Load question qid (0: a new question) into the page and the working
' copy. asNewVersion starts a new version of approved question qid: the
' same fields, not saved until Save.
Private Sub LoadQuestion(ByVal qid As Long, ByVal asNewVersion As Boolean)
    Dim rs As DAO.Recordset, st As String
    m_loading = True
    ClearWork
    m_id = 0: m_baseId = 0: m_key = 0: m_ver = 1: m_status = QS_DRAFT
    Me.txtTopicKey = Null
    SetSegment Me, "segType", "mc"
    Me.cboDiff = Null: Me.txtPoints = Null: Me.cboMark = Null: Me.txtText = Null: Me.txtRefAns = Null
    If qid <> 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblQuestions WHERE QuestionID=" & qid, dbOpenSnapshot)
        If rs.EOF Then
            rs.Close
            m_loading = False
            KitMsgBox "That question isn't in the bank any more, so a new one is open instead.", vbInformation, MSG_TITLE
            LoadQuestion 0, False
            Exit Sub
        End If
        st = QStatusName(Nz(rs!StatusID, 0))
        If asNewVersion Then
            If Len(NewVersionProblem(qid)) > 0 Then
                KitMsgBox NewVersionProblem(qid) & " So it opens as it is.", vbInformation, MSG_TITLE
                asNewVersion = False
            End If
        End If
        If asNewVersion And Not CanWriteQuestion(qid) Then
            KitMsgBox "Only the question's author or an SME on its topic can start a new version of it, so it opens as it is.", _
                vbInformation, MSG_TITLE
            asNewVersion = False
        End If
        m_key = Nz(rs!QuestionKey, 0)
        If asNewVersion Then
            m_baseId = qid
            m_ver = Nz(DMax("VersionNo", "tblQuestions", "QuestionKey=" & m_key), 0) + 1
        Else
            m_id = qid
            m_ver = Nz(rs!VersionNo, 1)
            m_status = st
        End If
        Me.txtTopicKey = rs!TopicID
        SetSegment Me, "segType", IIf(Len(TypeKey(QTypeName(Nz(rs!QuestionTypeID, 0)))) > 0, TypeKey(QTypeName(Nz(rs!QuestionTypeID, 0))), "mc")
        Me.cboDiff = rs!DifficultyID
        Me.txtPoints = PointsText(rs!TotalPoints)
        Me.cboMark = rs!ClassLevelID
        Me.txtText = rs!QuestionText
        Me.txtRefAns = rs!ReferenceAnswer
        rs.Close
        LoadWork qid
    End If
    ShowTopic
    m_type = SegmentValue(Me, "segType")
    ' Read only while it can't be changed, or when this person may not change
    ' it (not its author, nor an SME on its topic). Save checks again.
    m_noRight = False
    If m_id <> 0 And StatusEditable(m_status) Then m_noRight = Not CanWriteQuestion(m_id)
    m_readOnly = (m_id <> 0 And Not StatusEditable(m_status)) Or m_noRight
    m_stamp = LastQuestionAudit(m_key)
    m_workDirty = False
    TempVars("EditQuestionID") = IIf(m_id <> 0, m_id, m_baseId)
    TempVars("EditMode") = IIf(m_baseId <> 0, "newversion", "")
    RefreshParts 0
    RefreshOptions 0
    RefreshRefs 0
    RefreshHistory
    ApplyReadOnly
    m_snap = Snapshot()
    m_loading = False
    ShowState
    If m_id <> 0 Then KitRecent QuestionRef(m_key, m_ver) & " " & Left$(Replace(Nz(Me.txtText, ""), vbCrLf, " "), 40), _
        "OpenQuestion(" & m_id & ")", "E82D"
End Sub

' The topic's name in the picker; a topic turned off since is shown with (off).
Private Sub ShowTopic()
    KitPickShow Me.txtTopic
    ' The picker lists only active topics this person is an SME on; any other
    ' topic is named here, with (off) when it is turned off.
    If IsNull(Me.txtTopic) And Not IsNull(Me.txtTopicKey) Then
        Me.txtTopic = Nz(DLookup("TopicName", "tblTopics", "TopicID=" & Me.txtTopicKey), "?") & _
            IIf(Nz(DLookup("IsActive", "tblTopics", "TopicID=" & Me.txtTopicKey), False), "", " (off)")
    End If
End Sub

' Undo: load the question again as it was saved (or as it is approved,
' for a new version).
Private Sub UndoChanges()
    If KitMsgBox("Drop your changes to " & Called() & "?", vbYesNo + vbQuestion + vbDefaultButton2, MSG_TITLE) <> vbYes Then Exit Sub
    If m_baseId <> 0 Then LoadQuestion m_baseId, True Else LoadQuestion m_id, False
End Sub

Private Sub SaveClicked()
    Dim wasNew As Boolean, oldVer As Long
    If m_noRight Then
        KitMsgBox "Only its author or an SME on its topic can change " & Called() & ".", vbInformation, MSG_TITLE
        Exit Sub
    End If
    If m_readOnly Then
        KitMsgBox Called() & " is " & LCase$(m_status) & ", so it can't be changed." & _
            IIf(m_status = QS_APPROVED, " To change it, use More > Edit as new version.", ""), vbInformation, MSG_TITLE
        Exit Sub
    End If
    If m_id <> 0 And Not IsDirty() Then KitToast "Nothing to save.", "info": Exit Sub
    wasNew = (m_baseId <> 0)
    oldVer = m_ver - 1
    If Not SaveNow(False) Then Exit Sub
    If wasNew Then
        KitToast "Saved as version " & m_ver & ". Version " & oldVer & " stays in use until this one is approved.", "success"
    Else
        KitToast "Saved as a draft.", "success"
    End If
End Sub

' Check the rules and save. forSubmit checks the submission rules too, and
' doesn't save a question that has no changes. False when it wasn't saved
' (the user was told why).
Private Function SaveNow(ByVal forSubmit As Boolean) As Boolean
    Dim problems As String, typeName As String, newId As Long, problem As String, last As Long
    If m_readOnly Then Exit Function
    CommitTyping
    typeName = TypeOfKey(SegmentValue(Me, "segType"))
    problems = QuestionProblems(Me.txtTopicKey, typeName, Me.cboDiff, Me.txtPoints, Me.cboMark, Me.txtText, forSubmit)
    If Len(problems) > 0 Then
        KitMsgBox Called() & " can't be " & IIf(forSubmit, "submitted", "saved") & " yet:" & vbCrLf & vbCrLf & _
            ChrW(&H2022) & " " & Replace(problems, vbCrLf, vbCrLf & ChrW(&H2022) & " "), vbExclamation, MSG_TITLE
        GoToProblem Split(problems, vbCrLf)(0)
        Exit Function
    End If
    If forSubmit And m_id <> 0 And Not IsDirty() Then SaveNow = True: Exit Function
    ' Someone else saved since this was loaded.
    last = LastQuestionAudit(m_key)
    If last <> m_stamp Then
        If KitMsgBox("Someone else changed this question while you had it open (" & AuditWho(last) & ")." & vbCrLf & vbCrLf & _
                "Save yours over theirs? No keeps theirs; Undo changes in More then shows it.", _
                vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then Exit Function
    End If
    problem = SaveQuestion(m_id, m_baseId, CLng(Me.txtTopicKey), typeName, CLng(Me.cboDiff), CDbl(Me.txtPoints), Me.cboMark, _
        Me.txtText, Me.txtRefAns, newId)
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Function
    LoadQuestion newId, False
    SaveNow = True
End Function

' Ctrl+S keeps the focus in the box being typed in, whose value isn't
' updated yet: take what is typed.
Private Sub CommitTyping()
    Dim c As Control
    On Error Resume Next                    ' no control has the focus
    Set c = Me.ActiveControl
    If c Is Nothing Then Exit Sub
    If c.ControlType = acTextBox And Not c.Locked Then
        If Nz(c.Text, "") <> Nz(c.Value, "") Then c.Value = IIf(Len(c.Text) = 0, Null, c.Text)
    End If
End Sub

' Show where the first problem is: its field, or its tab.
Private Sub GoToProblem(ByVal line As String)
    On Error Resume Next                    ' a control that can't take the focus is left
    Select Case True
        Case line Like "Topic*": Me.txtTopic.SetFocus
        Case line Like "Type*": Me.segType_b0.SetFocus
        Case line Like "Difficulty*": Me.cboDiff.SetFocus
        Case line Like "Points*": Me.txtPoints.SetFocus
        Case line Like "Portion marking*": Me.cboMark.SetFocus
        Case line Like "Question text*": Me.tabQ = 0: Me.txtText.SetFocus
        Case line Like "Part*": Me.tabQ = 1
        Case line Like "Answer options*": Me.tabQ = 2
        Case line Like "References*": Me.tabQ = 3
    End Select
End Sub

' Delete this draft (More), and go back to the bank.
Private Sub DeleteThis()
    Dim problem As String, shown As Long, id As Long
    problem = DeleteProblem(m_id)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: Exit Sub
    If Not KitConfirmTyped("Delete " & Called() & "?", "It was never submitted and is on no exam. Its parts, options and " & _
            "citations go with it" & IIf(m_ver > 1, "; the approved version stays in use", "") & _
            ". This can't be undone.", PURGE_WORD, "Delete", MSG_TITLE) Then Exit Sub
    id = m_id
    shown = DeleteDraft(id)
    ' Nothing left to save: the page can be left.
    m_id = 0: m_readOnly = True: m_workDirty = False: m_snap = Snapshot()
    KitToast IIf(shown <> 0, "Draft deleted. The approved version is still in use.", "Draft deleted."), "destructive"
    TempVars("BankQuestionID") = shown
    NavTo "questions"
End Sub

' ---------------------------------------------------------------------------
' Showing the question
' ---------------------------------------------------------------------------

' "Q12 v2", or "the new question".
Private Function Called() As String
    If m_key = 0 Then Called = "the new question" Else Called = QuestionRef(m_key, m_ver)
End Function

' The page's own fields as one string, to tell whether they changed.
Private Function Snapshot() As String
    Snapshot = Nz(Me.txtTopicKey, "") & "|" & SegmentValue(Me, "segType") & "|" & Nz(Me.cboDiff, "") & "|" & _
        Nz(Me.txtPoints, "") & "|" & Nz(Me.cboMark, "") & "|" & Nz(Me.txtText, "") & "|" & Nz(Me.txtRefAns, "")
End Function

Private Function IsDirty() As Boolean
    IsDirty = m_workDirty Or (Snapshot() <> m_snap)
End Function

' The subtitle: which question and version, where it stands, and whether
' there are unsaved changes.
Private Sub ShowState()
    Dim s As String, stepCode As String, why As String, tip As String
    If m_key = 0 Then
        s = "New question, not saved yet."
    ElseIf m_baseId <> 0 Then
        s = QuestionRef(m_key, m_ver) & ", a new version of v" & m_ver - 1 & ", not saved yet. v" & m_ver - 1 & _
            " stays in use until this one is approved."
    Else
        s = QuestionRef(m_key, m_ver) & ", " & m_status & "."
        Select Case m_status
            Case QS_REVIEW
                stepCode = Nz(DLookup("RoleCode", "tblRoles", "RoleID=" & Nz(DLookup("CurrentStepRoleID", "tblQuestions", _
                    "QuestionID=" & m_id), 0)), "")
                s = s & IIf(Len(stepCode) > 0, " Waiting on the " & stepCode & ".", "") & " It can't be changed while it is in review."
            Case QS_APPROVED
                s = s & " To change it, use More > Edit as new version."
            Case QS_RETIRED
                s = s & " A retired version is kept for the record and can't be changed."
            Case QS_REJECTED
                why = RejectionText(m_id)
                tip = why
                If Len(why) > 110 Then why = Left$(why, 107) & "..."    ' the whole reason is in the tip and in History
                s = s & IIf(Len(why) > 0, " " & why, " See History for why.") & " Change it and submit it again."
        End Select
        If m_noRight Then s = s & " Only its author or an SME on its topic can change it."
    End If
    If Not m_readOnly And IsDirty() Then s = s & "  Unsaved changes."
    Me.lblSubtitle.Caption = s
    Me.lblSubtitle.ControlTipText = tip
End Sub

' Lock everything that changes the question while it can't be changed.
Private Sub ApplyReadOnly()
    Dim nm As Variant, ro As Boolean
    ro = m_readOnly
    For Each nm In Array("txtTopic", "segType_val", "cboDiff", "txtPoints", "txtText", "txtRefAns")
        Me.Controls(nm).Locked = ro
    Next
    Me.cboMark.Locked = ro Or WorkPartCount() > 0
    For Each nm In Array("btnSave", "btnSubmit", "btnAddPart", "btnAddOption", "btnAddRef", "cboOptPart")
        SetEnabled Me.Controls(nm), Not ro
    Next
    ShowRowState
End Sub

' The row buttons: on while a row is selected and the question can be changed.
Private Sub ShowRowState()
    Dim ro As Boolean, p As Boolean, o As Boolean, r As Boolean, isOk As Boolean
    ro = m_readOnly
    p = SelId(Me.sfParts) <> 0 And Not ro
    o = SelId(Me.sfOptions) <> 0 And Not ro
    r = SelId(Me.sfRefs) <> 0 And Not ro
    SetEnabled Me.btnEditPart, p: SetEnabled Me.btnUpPart, p: SetEnabled Me.btnDownPart, p: SetEnabled Me.btnRemovePart, p
    SetEnabled Me.btnEditOption, o: SetEnabled Me.btnUpOption, o: SetEnabled Me.btnDownOption, o
    SetEnabled Me.btnRemoveOption, o: SetEnabled Me.btnCorrectOption, o
    If o Then isOk = Nz(DLookup("IsCorrect", "wq_option", "id=" & SelId(Me.sfOptions)), False)
    Me.btnCorrectOption.Caption = IIf(isOk, "Not correct", "Mark correct")
    SetEnabled Me.btnEditRef, r: SetEnabled Me.btnRemoveRef, r
End Sub

' The parts table (pid selected), the marking rolled up from them, the
' points line and the tab's count.
Private Sub RefreshParts(ByVal pid As Long)
    Dim n As Long
    Me.sfParts.Form.Requery
    GridSelect Me.sfParts, pid
    n = WorkPartCount()
    If n > 0 Then
        Me.cboMark = RolledUpMarking()
        Me.cboMark.ControlTipText = "Set from the parts: the highest part's marking"
    Else
        Me.cboMark.ControlTipText = "The question's portion marking"
    End If
    Me.cboMark.Locked = m_readOnly Or n > 0
    Me.tbParts.Caption = "Parts" & IIf(n > 0, " (" & n & ")", "")
    ShowPoints
    ShowRowState
End Sub

' The line over the parts: how their points add up against the question's.
Private Sub ShowPoints()
    Dim n As Long, sum As Double, total As Variant
    n = WorkPartCount()
    If n = 0 Then
        ShowAlert Me, "alPoints", "No parts. A question with parts (A, B...) gives each its own text, points and portion " & _
            "marking; their points must add up to the question's.", "info"
        Exit Sub
    End If
    sum = WorkPointsSum()
    total = Me.txtPoints
    If Len(PointsProblem(total, "Points")) > 0 Then
        ShowAlert Me, "alPoints", "The parts give " & PointsText(sum) & " points. Set the question's points (above) to match.", "warning"
    ElseIf Abs(sum - CDbl(total)) < 0.001 Then
        ShowAlert Me, "alPoints", "The parts give " & PointsText(sum) & " points, the question's total. The question's marking " & _
            "is the highest of theirs.", "success"
    Else
        ShowAlert Me, "alPoints", "The parts give " & PointsText(sum) & " points, but the question is worth " & PointsText(total) & _
            ". They must match before it can be saved.", "warning"
    End If
End Sub

' The options table (id selected), the part picker beside it, the line
' over it and the tab's count.
Private Sub RefreshOptions(ByVal id As Long)
    Dim n As Long, nParts As Long, keep As Variant
    Me.sfOptions.Form.Requery
    GridSelect Me.sfOptions, id
    nParts = WorkPartCount()
    keep = Me.cboOptPart
    Me.cboOptPart.RowSource = "SELECT id, PartLabel FROM wq_part ORDER BY SortOrder"
    If nParts = 0 Then
        Me.cboOptPart = Null
    ElseIf IsNull(keep) Then
        Me.cboOptPart = DLookup("id", "wq_part", "SortOrder=1")
    ElseIf DCount("*", "wq_part", "id=" & keep) = 0 Then
        Me.cboOptPart = DLookup("id", "wq_part", "SortOrder=1")
    Else
        Me.cboOptPart = keep
    End If
    SetVisible Me.cboOptPart, nParts > 0
    SetVisible Me.lblOptPart, nParts > 0
    n = DCount("*", "wq_option")
    Me.tbOptions.Caption = "Answer options" & IIf(n > 0, " (" & n & ")", "")
    ShowOptionsLine
    ShowRowState
End Sub

' The line over the options: what the submission rules still want of them.
Private Sub ShowOptionsLine()
    Dim lines As Variant, ln As Variant, first As String, more As Long
    If SegmentValue(Me, "segType") <> "mc" Then
        ShowAlert Me, "alOptions", "Only multiple choice questions have answer options. Put a true/false or short answer " & _
            "question's answer on the Reference answer tab.", "info"
        Exit Sub
    End If
    lines = Split(QuestionProblems(Me.txtTopicKey, QT_MC, Me.cboDiff, Me.txtPoints, Me.cboMark, Me.txtText, True), vbCrLf)
    For Each ln In lines
        If Left$(ln, 16) = "Answer options: " Then
            If Len(first) = 0 Then first = Mid$(ln, 17) Else more = more + 1
        End If
    Next
    If Len(first) = 0 Then
        ShowAlert Me, "alOptions", IIf(WorkPartCount() > 0, "Each part has", "The question has") & _
            " its options and one correct answer.", "success"
    Else
        ShowAlert Me, "alOptions", first & IIf(more > 0, " (and " & more & " more to fix before it is submitted)", ""), "warning"
    End If
End Sub

' The references table (id selected), its line and the tab's count.
Private Sub RefreshRefs(ByVal id As Long)
    Dim n As Long
    Me.sfRefs.Form.Requery
    GridSelect Me.sfRefs, id
    n = DCount("*", "wq_ref")
    If n = 0 Then
        ShowAlert Me, "alRefs", "No references yet. Cite at least one before the question is submitted.", "warning"
    Else
        ShowAlert Me, "alRefs", n & IIf(n = 1, " reference", " references") & " cited. The paragraph says where in it the " & _
            "answer is.", "info"
    End If
    Me.tbRefs.Caption = "References" & IIf(n > 0, " (" & n & ")", "")
    ShowRowState
End Sub

' The History tab: every version's audit rows, newest first.
Private Sub RefreshHistory()
    Dim n As Long
    If m_key = 0 Then
        SetVisible Me.sfHistory, False
        SetEnabled Me.btnCopyHistory, False
        Me.lblHistoryEmpty.Caption = "No history yet. It starts when the question is first saved."
        Exit Sub
    End If
    Me.sfHistory.Form.RecordSource = HistorySql(m_key)
    n = Me.sfHistory.Form.RecordsetClone.RecordCount
    SetVisible Me.sfHistory, True
    SetEnabled Me.btnCopyHistory, n > 0
    Me.lblHistoryEmpty.Caption = "Every version of Q" & m_key & ": who wrote, changed, submitted, approved and retired it."
End Sub

' ---------------------------------------------------------------------------

' The id of the selected row of a table, 0 for none.
Private Function SelId(sf As Control) As Long
    SelId = Nz(sf.Form.txtSel, 0)
End Function
