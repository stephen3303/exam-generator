Option Compare Database
Option Explicit

' Approvals (see Build_pgApprovals). Two tabs; the Questions tab's controls
' are listed in the hidden label grpQ.
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
' Exams: a placeholder until build phase 8.

Private Const MSG_TITLE As String = "Approvals"

Private m_tab As String         ' the open tab: q or x
Private m_search As String
Private m_qid As Long           ' the question shown on the right (0: none)

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
        ShowQuestion 0
        Exit Sub
    End If
    RefreshList 0
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

' A column funnel changed the questions shown (GridHeaders' onFilter).
Public Sub ListFiltered(ByVal sfName As String)
    On Error GoTo Err_Handler
    AfterFilter
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
    If GridSelect(Me.sfList, m_qid) Then
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

' ---------------------------------------------------------------------------

' Show the open tab's controls and hide the other's.
Private Sub ShowGroup()
    Dim nm As Variant, onQ As Boolean
    onQ = (m_tab = "q")
    For Each nm In Split(Me.grpQ.Caption, ";")
        If Len(nm) > 0 Then SetVisible Me.Controls(nm), onQ
    Next
    ShowEmpty Me, "emptyExams", Not onQ
    If onQ Then
        If DataLinked() Then AfterFilter Else ShowQuestion 0
    Else
        ShowEmpty Me, "emptyList", False
        ShowQuestion m_qid
    End If
End Sub
