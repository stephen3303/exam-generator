Option Compare Database
Option Explicit

' Approvals (see Build_pgApprovals). Two tabs; each tab's controls are
' listed in a hidden label, grpQ and grpX.
'
' Questions: the questions waiting at a step of the chain whose role the
' signed-in person holds (modApprovals.ApprovalQueueSql), the longest
' waiting first, filtered by the search box and the column funnels. The
' picked one shows on the right: where it is in the chain, who wrote and
' submitted it, and the whole question as submitted. Approve sends it on (at
' the last step it goes into the bank, retiring the version it replaces);
' Reject opens dlgReject for the reason and sends it back to its author.
' Both are checked again in the data file when clicked (RequireApprover).
'
' Exams: the exams waiting at a step whose role the person holds
' (modExamApprovals.ExamQueueSql), laid out the same way, with the whole
' exam, answers included, on the right. Approve sends it on; the SM's final
' approval makes its Word document (modExamDoc). Reject opens
' dlgRejectExam, which sends one topic or all of them back to their SMEs.

Private Const MSG_TITLE As String = "Approvals"

Private m_tab As String         ' the open tab: q or x
Private m_search As String
Private m_qid As Long           ' the question shown on the right (0: none)
Private m_xsearch As String
Private m_xid As Long           ' the exam shown on the right (0: none)

Public Property Get HelpTopic() As String
    HelpTopic = "approvals"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    m_tab = "q"
    SelectTab Me, "q"
    SetCount Me.tab_x, 0
    If Not DataLinked() Then
        ShowGroup
        ShowEmpty Me, "emptyList", True, "No data file linked", "This copy was built without the data file, so nothing waits for approval."
        SetVisible Me.emptyList_b, False
        SetVisible Me.sfList, False
        SetVisible Me.sfXList, False
        ShowQuestion 0
        Exit Sub
    End If
    RefreshList 0
    RefreshXList 0
    ShowGroup
    Exit Sub
Err_Handler:
    HandleError "pgApprovals.Form_Load"
End Sub

' ---------------------------------------------------------------------------
' Tabs and search (the kit's builders call these)
' ---------------------------------------------------------------------------

Public Function TabPick(ByVal key As String)
    On Error GoTo Err_Handler
    If key = m_tab Then Exit Function
    m_tab = key
    SelectTab Me, key
    ShowGroup
    Exit Function
Err_Handler:
    HandleError "pgApprovals.TabPick", key
End Function

Public Sub SearchFor(ByVal text As String)
    On Error GoTo Err_Handler
    m_search = Trim$(text)
    RefreshList m_qid
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "pgApprovals.SearchFor", text
End Sub

Public Sub SearchExams(ByVal text As String)
    On Error GoTo Err_Handler
    m_xsearch = Trim$(text)
    RefreshXList m_xid
    Me.txtXSearch.SetFocus
    Me.txtXSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "pgApprovals.SearchExams", text
End Sub

' A column funnel changed the rows shown (GridHeaders' onFilter).
Public Sub ListFiltered(ByVal sfName As String)
    On Error GoTo Err_Handler
    If sfName = "sfXList" Then AfterXFilter Else AfterFilter
    Exit Sub
Err_Handler:
    HandleError "pgApprovals.ListFiltered", sfName
End Sub

' ---------------------------------------------------------------------------
' Buttons and rows
' ---------------------------------------------------------------------------

Private Sub btnApprove_Click()
    On Error GoTo Err_Handler
    ApproveClicked
    Exit Sub
Err_Handler:
    HandleError "pgApprovals.btnApprove_Click", m_qid
    RefreshList m_qid                       ' a refused approval: show where it stands now
End Sub

Private Sub btnReject_Click()
    On Error GoTo Err_Handler
    RejectClicked
    Exit Sub
Err_Handler:
    HandleError "pgApprovals.btnReject_Click", m_qid
    RefreshList m_qid
End Sub

Private Sub btnHistory_Click()
    On Error GoTo Err_Handler
    If m_qid <> 0 Then RunDialog "dlgHistory", "key=" & Nz(DLookup("QuestionKey", "tblQuestions", "QuestionID=" & m_qid), 0)
    Exit Sub
Err_Handler:
    HandleError "pgApprovals.btnHistory_Click", m_qid
End Sub

Private Sub btnXApprove_Click()
    On Error GoTo Err_Handler
    XApproveClicked
    Exit Sub
Err_Handler:
    HandleError "pgApprovals.btnXApprove_Click", m_xid
    RefreshXList m_xid
End Sub

Private Sub btnXReject_Click()
    On Error GoTo Err_Handler
    XRejectClicked
    Exit Sub
Err_Handler:
    HandleError "pgApprovals.btnXReject_Click", m_xid
    RefreshXList m_xid
End Sub

Private Sub btnXHistory_Click()
    On Error GoTo Err_Handler
    If m_xid <> 0 Then RunDialog "dlgHistory", "exam=" & m_xid
    Exit Sub
Err_Handler:
    HandleError "pgApprovals.btnXHistory_Click", m_xid
End Sub

' The Exams tab's empty state button: clear its search and column filters.
Private Sub emptyXList_b_Click()
    On Error GoTo Err_Handler
    ClearGridFilters Me.sfXList
    SetSearch Me, "txtXSearch", ""
    Exit Sub
Err_Handler:
    HandleError "pgApprovals.emptyXList_b_Click"
End Sub

' The empty state's button: clear the search and the column filters.
Private Sub emptyList_b_Click()
    On Error GoTo Err_Handler
    ClearGridFilters Me.sfList
    SetSearch Me, "txtSearch", ""
    Exit Sub
Err_Handler:
    HandleError "pgApprovals.emptyList_b_Click"
End Sub

' A row clicked: show it on the right.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    Dim id As Long
    On Error GoTo Err_Handler
    If grid = "sfXList" Then
        id = Nz(Me.sfXList.Form.txtSel, 0)
        ShowExam id
        Exit Sub
    End If
    If grid <> "sfList" Then Exit Sub
    id = Nz(Me.sfList.Form.txtSel, 0)
    ShowQuestion id
    Exit Sub
Err_Handler:
    HandleError "pgApprovals.GridClick", grid & " " & id
End Sub

' A right-click on a row: its actions.
Public Sub GridRightClick(ByVal grid As String)
    Dim ok As Boolean, k As String
    On Error GoTo Err_Handler
    If grid = "sfXList" Then
        If m_xid = 0 Then Exit Sub
        ok = (Len(ExamApproveProblem(m_xid)) = 0)
        k = KitMenuAt(IIf(ok, "", "~") & "approve|Approve|E8FB", IIf(ok, "", "~") & "reject|!Reject|E711", "-", _
            "hist|History|E81C", "exams|Show on the Exams page|E8A5")
        Select Case k
            Case "approve": XApproveClicked
            Case "reject": XRejectClicked
            Case "hist": btnXHistory_Click
            Case "exams"
                TempVars("ExamsExamID") = m_xid
                NavTo "exams"
        End Select
        Exit Sub
    End If
    If grid <> "sfList" Or m_qid = 0 Then Exit Sub
    ok = CanApprove(m_qid)
    k = KitMenuAt(IIf(ok, "", "~") & "approve|Approve|E8FB", IIf(ok, "", "~") & "reject|!Reject|E711", "-", _
        "hist|History|E81C", "bank|Show in the question bank|E82D")
    Select Case k
        Case "approve": ApproveClicked
        Case "reject": RejectClicked
        Case "hist": btnHistory_Click
        Case "bank"
            TempVars("BankQuestionID") = m_qid
            NavTo "questions"
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgApprovals.GridRightClick", grid
    RefreshList m_qid
End Sub

' Approve the shown question at the person's step. The last step's approval
' of a new version retires the version in use, so that one asks first.
Private Sub ApproveClicked()
    Dim problem As String, done As String, key As Long, ver As Long, oldVer As Variant
    If m_qid = 0 Then Exit Sub
    problem = ApproveProblem(m_qid)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: RefreshList m_qid: Exit Sub
    If NextStepRoleID(Nz(DLookup("CurrentStepRoleID", "tblQuestions", "QuestionID=" & m_qid), 0)) = 0 Then
        key = Nz(DLookup("QuestionKey", "tblQuestions", "QuestionID=" & m_qid), 0)
        ver = Nz(DLookup("VersionNo", "tblQuestions", "QuestionID=" & m_qid), 0)
        oldVer = DLookup("VersionNo", "tblQuestions", "QuestionKey=" & key & " AND IsCurrent AND QuestionID<>" & m_qid)
        If Not IsNull(oldVer) Then
            If KitMsgBox("Approve " & QuestionRef(key, ver) & "? It replaces v" & oldVer & " in the bank: v" & oldVer & _
                    " is retired and won't be drawn for new exams. Exams that already have v" & oldVer & " keep it.", _
                    vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then Exit Sub
        End If
    End If
    problem = ApproveQuestion(m_qid, done)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: RefreshList m_qid: Exit Sub
    KitToast done, "success"
    RefreshList m_qid                       ' still listed when the next step is also theirs
End Sub

' Send the shown question back to its author, with the reason dlgReject asks for.
Private Sub RejectClicked()
    Dim problem As String
    If m_qid = 0 Then Exit Sub
    problem = ApproveProblem(m_qid)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: RefreshList m_qid: Exit Sub
    If RunDialog("dlgReject", "id=" & m_qid) Then
        KitToast "Sent back to its author with your reason.", "info"
    End If
    RefreshList m_qid                       ' gone from the list, or as it stands now
End Sub

' Approve the shown exam at the person's step. The final approval asks
' first, then makes the exam's Word document; the approval stands if that
' fails, and Download a copy on the Exams page tries again.
Private Sub XApproveClicked()
    Dim problem As String, done As String, isFinal As Boolean, note As String, mk As String
    If m_xid = 0 Then Exit Sub
    problem = ExamApproveProblem(m_xid)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: RefreshXList m_xid: Exit Sub
    If NextStepRoleID(Nz(DLookup("CurrentStepRoleID", "tblExams", "ExamID=" & m_xid), 0)) = 0 Then
        mk = LevelMarking(OverallLevelID(m_xid))
        If KitMsgBox("Give " & ExamRef(m_xid, DLookup("VersionNo", "tblExams", "ExamID=" & m_xid)) & " final approval? " & _
                IIf(Len(mk) > 0, "It is marked " & mk & " overall, its highest portion marking. ", "") & _
                "Its Word document is made now, and anyone can download a copy from the Exams page.", _
                vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then Exit Sub
    End If
    problem = ApproveExam(m_xid, done, isFinal)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: RefreshXList m_xid: Exit Sub
    If isFinal Then
        problem = MakeExamDocumentNow(m_xid, note)
        If Len(problem) > 0 Then
            KitMsgBox done & vbCrLf & vbCrLf & problem, vbExclamation, MSG_TITLE
        Else
            KitToast done & IIf(Len(note) > 0, " " & note, ""), "success"
        End If
    Else
        KitToast done, "success"
    End If
    RefreshXList m_xid                      ' still listed when the next step is also theirs
End Sub

' Send the shown exam back to its SMEs, with the topics and the reason
' dlgReject asks for.
Private Sub XRejectClicked()
    Dim problem As String
    If m_xid = 0 Then Exit Sub
    problem = ExamApproveProblem(m_xid)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: RefreshXList m_xid: Exit Sub
    If RunDialog("dlgRejectExam", "id=" & m_xid) Then
        KitToast "Sent back to its SMEs with your reason. They find it under My reviews.", "info"
    End If
    RefreshXList m_xid
End Sub

' ---------------------------------------------------------------------------
' The list and the detail
' ---------------------------------------------------------------------------

' Load the questions waiting for the person, for the search, and show
' openId (0, or not listed: the first).
Private Sub RefreshList(ByVal openId As Long)
    If Not DataLinked() Then Exit Sub
    Me.sfList.Form.RecordSource = ApprovalQueueSql(m_search)
    SetGridFilter Me.sfList, ""             ' the column funnels, on the new rows
    GridHighlight Me.sfList, m_search
    SetCount Me.tab_q, ApprovalCount()
    m_qid = openId
    AfterFilter
End Sub

' After the rows shown changed: the empty state, and the shown question
' kept when still listed (else the first).
Private Sub AfterFilter()
    Dim n As Long, any As Boolean
    n = Me.sfList.Form.RecordsetClone.RecordCount
    any = ApprovalCount() > 0
    If m_tab = "q" Then
        If Not any Then
            ShowEmpty Me, "emptyList", n = 0, "Nothing waiting for you", "Questions sent to your step of the chain show here."
            SetVisible Me.emptyList_b, False
        Else
            ShowEmpty Me, "emptyList", n = 0, "No questions match", "No question waiting for you fits the search and column filters."
            Me.emptyList_b.Caption = "Show all"
        End If
        SetVisible Me.sfList, n > 0
    End If
    If MarkRow(Me.sfList, m_qid) Then
        ShowQuestion m_qid
    ElseIf IsNull(SelectFirstRow(Me.sfList)) Then
        ShowQuestion 0
    End If
End Sub

' Show question qid on the right (0: the hint instead).
Private Sub ShowQuestion(ByVal qid As Long)
    Dim rs As DAO.Recordset, isOpen As Boolean, nm As Variant, c As Control, s As String, problem As String
    Dim stepNo As Long, author As String, last As Long, oldVer As Variant
    m_qid = qid
    isOpen = (qid <> 0)
    If isOpen Then
        Set rs = CurrentDb.OpenRecordset("SELECT q.QuestionKey, q.VersionNo, q.ReplacesQuestionID, q.CurrentStepRoleID, " & _
            "q.AuthorPersonID, tp.TopicName FROM tblQuestions AS q LEFT JOIN tblTopics AS tp ON q.TopicID=tp.TopicID " & _
            "WHERE q.QuestionID=" & qid, dbOpenSnapshot)
        If rs.EOF Then
            isOpen = False: m_qid = 0
        Else
            Me.lblQRef.Caption = QuestionRef(rs!QuestionKey, rs!VersionNo) & IIf(IsNull(rs!TopicName), "", ", " & rs!TopicName)
            author = PersonName(rs!AuthorPersonID)
            Me.txtAuthor = IIf(Len(author) > 0, author, Null)
            last = Nz(DMax("AuditID", "tblAuditTrail", "ItemType='" & ITEM_QUESTION & "' AND ItemID=" & qid & _
                " AND ActionType=" & Q(AU_SUBMITTED)), 0)
            Me.lblQBy.Caption = "Written by " & IIf(Len(author) > 0, author, "someone no longer listed") & _
                IIf(last <> 0, ". Submitted by " & AuditWho(last), "") & "."
            stepNo = StepNumber(Nz(rs!CurrentStepRoleID, 0))
            problem = ApproveProblem(qid)
            If Len(problem) > 0 Then s = problem Else s = "Waiting on you as the " & RoleCodeOf(rs!CurrentStepRoleID) & "."
            If Nz(rs!ReplacesQuestionID, 0) <> 0 Then
                oldVer = DLookup("VersionNo", "tblQuestions", "QuestionID=" & rs!ReplacesQuestionID)
                If Not IsNull(oldVer) Then s = s & " A new version of v" & oldVer & ", which stays in use until this one's last approval."
            End If
            Me.lblQState.Caption = s
            Me.txtPreview = QuestionPreview(qid)
            SetEnabled Me.btnApprove, Len(problem) = 0
            SetEnabled Me.btnReject, Len(problem) = 0
        End If
        rs.Close
    End If
    isOpen = isOpen And m_tab = "q"
    For Each nm In Array("lblQRef", "lblQBy", "lblQState", "capPreview", "txtPreview", "btnApprove", "btnReject", "btnHistory")
        SetVisible Me.Controls(nm), isOpen
    Next
    SetVisible Me.lblDetailEmpty, Not isOpen And m_tab = "q"
    ' The stepper and the avatar paint their own parts, so they aren't in grpQ.
    For Each c In Me.Controls
        If c.Name Like "stpQ*" Or c.Name Like "avAuthor*" Then SetVisible c, isOpen
    Next
    If isOpen Then SetStep Me, "stpQ", IIf(stepNo < 1, 1, stepNo)
End Sub

' Load the exams waiting for the person, for the search, and show openId
' (0, or not listed: the first).
Private Sub RefreshXList(ByVal openId As Long)
    If Not DataLinked() Then Exit Sub
    Me.sfXList.Form.RecordSource = ExamQueueSql(m_xsearch)
    SetGridFilter Me.sfXList, ""
    GridHighlight Me.sfXList, m_xsearch
    SetCount Me.tab_x, ExamApprovalCount()
    m_xid = openId
    AfterXFilter
End Sub

' After the exams shown changed: the empty state, and the shown exam kept
' when still listed (else the first).
Private Sub AfterXFilter()
    Dim n As Long
    n = Me.sfXList.Form.RecordsetClone.RecordCount
    If m_tab = "x" Then
        If ExamApprovalCount() = 0 Then
            ShowEmpty Me, "emptyXList", n = 0, "Nothing waiting for you", _
                "Exams sent to your step of the chain show here once their SMEs sign off every topic."
            SetVisible Me.emptyXList_b, False
        Else
            ShowEmpty Me, "emptyXList", n = 0, "No exams match", "No exam waiting for you fits the search and column filters."
            Me.emptyXList_b.Caption = "Show all"
        End If
        SetVisible Me.sfXList, n > 0
    End If
    If MarkRow(Me.sfXList, m_xid) Then
        ShowExam m_xid
    ElseIf IsNull(SelectFirstRow(Me.sfXList)) Then
        ShowExam 0
    Else
        ShowExam Nz(Me.sfXList.Form.txtSel, 0)
    End If
End Sub

' Show exam examId on the right (0: the hint instead).
Private Sub ShowExam(ByVal examId As Long)
    Dim rs As DAO.Recordset, isOpen As Boolean, nm As Variant, c As Control, s As String, problem As String
    Dim nQ As Long, nT As Long, mk As String, by As String
    m_xid = examId
    isOpen = (examId <> 0)
    If isOpen Then
        Set rs = CurrentDb.OpenRecordset("SELECT e.VersionNo, e.RecipientName, e.CreatedByPersonID, e.CreatedOn, e.CurrentStepRoleID, " & _
            "tm.TemplateName FROM tblExams AS e LEFT JOIN tblTemplates AS tm ON e.TemplateID=tm.TemplateID " & _
            "WHERE e.ExamID=" & examId, dbOpenSnapshot)
        If rs.EOF Then
            isOpen = False: m_xid = 0
        Else
            Me.lblXRef.Caption = ExamRef(examId, rs!VersionNo) & " for " & rs!RecipientName
            nQ = DCount("*", "tblExamQuestions", "ExamID=" & examId)
            nT = DCount("*", "tblExamTopics", "ExamID=" & examId)
            by = PersonName(rs!CreatedByPersonID)
            Me.lblXBy.Caption = "From the template " & Nz(rs!TemplateName, "(deleted)") & ", generated by " & _
                IIf(Len(by) > 0, by, "someone no longer listed") & " on " & Format$(rs!CreatedOn, "d mmm yyyy") & ". " & _
                nQ & IIf(nQ = 1, " question", " questions") & " in " & nT & IIf(nT = 1, " topic", " topics") & "."
            problem = ExamApproveProblem(examId)
            If Len(problem) > 0 Then s = problem Else s = "Waiting on you as the " & RoleCodeOf(rs!CurrentStepRoleID) & "."
            mk = LevelMarking(OverallLevelID(examId))
            If Len(mk) > 0 Then s = s & " Its highest portion marking is " & mk & "."
            Me.lblXState.Caption = s
            Me.txtXPreview = ExamPreview(examId)
            SetEnabled Me.btnXApprove, Len(problem) = 0
            SetEnabled Me.btnXReject, Len(problem) = 0
        End If
        rs.Close
    End If
    isOpen = isOpen And m_tab = "x"
    For Each nm In Array("lblXRef", "lblXBy", "lblXState", "capXPreview", "txtXPreview", "btnXApprove", "btnXReject", "btnXHistory")
        SetVisible Me.Controls(nm), isOpen
    Next
    SetVisible Me.lblXDetailEmpty, Not isOpen And m_tab = "x"
    For Each c In Me.Controls
        If c.Name Like "stpX*" Then SetVisible c, isOpen
    Next
    If isOpen Then SetStep Me, "stpX", ExamStepNumber(examId)
End Sub

' ---------------------------------------------------------------------------

' Show the open tab's controls and hide the other's.
Private Sub ShowGroup()
    Dim nm As Variant, onQ As Boolean
    onQ = (m_tab = "q")
    For Each nm In Split(Me.grpQ.Caption, ";")
        If Len(nm) > 0 Then SetVisible Me.Controls(nm), onQ
    Next
    For Each nm In Split(Me.grpX.Caption, ";")
        If Len(nm) > 0 Then SetVisible Me.Controls(nm), Not onQ
    Next
    If onQ Then
        ShowEmpty Me, "emptyXList", False
        ShowExam m_xid
        If DataLinked() Then AfterFilter Else ShowQuestion 0
    Else
        ShowEmpty Me, "emptyList", False
        ShowQuestion m_qid
        If DataLinked() Then
            AfterXFilter
        Else
            ShowEmpty Me, "emptyXList", True, "No data file linked", "This copy was built without the data file, so nothing waits for approval."
            SetVisible Me.emptyXList_b, False
            SetVisible Me.sfXList, False
            ShowExam 0
        End If
    End If
End Sub

' Select row id of a table. False when it isn't listed.
Private Function MarkRow(sf As Control, ByVal id As Long) As Boolean
    Dim g As Form
    Set g = sf.Form
    g.txtSel = Null
    If id <> 0 Then
        With g.RecordsetClone
            .FindFirst "id=" & id
            If Not .NoMatch Then
                g.Bookmark = .Bookmark
                g.txtSel = id
                MarkRow = True
            End If
        End With
    End If
    g.Recalc
End Function
