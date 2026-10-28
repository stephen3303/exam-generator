Option Compare Database
Option Explicit

' My reviews (see Build_pgReview). The signed-in SME's review tasks, one per
' exam topic (modExams.ReviewTasksSql): Open, oldest first, or Done, newest
' first (signed off by them, or Closed by another SME's sign-off), filtered
' by the search box and the column funnels. The picked one shows on the
' right: the exam and where the topic stands, the questions drawn for it
' (Swapped, or Given lately when they were on one of the exams generated
' before this one), and
' the picked question in full. Swap opens dlgSwap; Sign off completes the
' topic for every SME on it. Both are checked again in the data file when
' clicked (RequireReviewer). The Dashboard can ask for a task to be shown
' first through TempVars("ReviewTaskID").

Private Const MSG_TITLE As String = "My reviews"

Private m_done As Boolean       ' the Done chip is on
Private m_search As String
Private m_task As Long          ' the task shown on the right (0: none)
Private m_et As Long            ' its exam topic
Private m_keys As String        ' RecentKeysList for its exam
Private m_stamp As Long         ' its change number when its questions were loaded (TopicChangeNo)

Public Property Get HelpTopic() As String
    HelpTopic = "review"
End Property

Private Sub Form_Load()
    Dim openId As Long
    On Error GoTo Err_Handler
    ChipOn Me, "chip", 0
    If Not DataLinked() Then
        ShowEmpty Me, "emptyList", True, "No data file linked", "This copy was built without the data file, so there is nothing to review."
        SetVisible Me.emptyList_b, False
        SetVisible Me.sfList, False
        ShowTask 0
        Exit Sub
    End If
    SendSignedOffExams                      ' an exam whose topics are all signed off goes on to the DH
    openId = Nz(TempVars("ReviewTaskID"), 0)
    TempVars("ReviewTaskID") = Null
    RefreshList openId
    Exit Sub
Err_Handler:
    HandleError "pgReview.Form_Load"
End Sub

' ---------------------------------------------------------------------------
' Chips and search (the kit's builders call these)
' ---------------------------------------------------------------------------

Public Function ChipPick(ByVal i As Long)
    On Error GoTo Err_Handler
    If (i = 1) = m_done Then Exit Function
    m_done = (i = 1)
    ChipOn Me, "chip", i
    RefreshList 0
    Exit Function
Err_Handler:
    HandleError "pgReview.ChipPick", i
End Function

Public Sub SearchFor(ByVal text As String)
    On Error GoTo Err_Handler
    m_search = Trim$(text)
    RefreshList m_task
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "pgReview.SearchFor", text
End Sub

' A column funnel changed the rows shown (GridHeaders' onFilter).
Public Sub ListFiltered(ByVal sfName As String)
    On Error GoTo Err_Handler
    AfterFilter
    Exit Sub
Err_Handler:
    HandleError "pgReview.ListFiltered", sfName
End Sub

' ---------------------------------------------------------------------------
' Buttons and rows
' ---------------------------------------------------------------------------

Private Sub btnSignOff_Click()
    On Error GoTo Err_Handler
    SignOffClicked
    Exit Sub
Err_Handler:
    HandleError "pgReview.btnSignOff_Click", m_et
    RefreshList m_task                      ' a refused sign-off: show where it stands now
End Sub

Private Sub btnSwap_Click()
    On Error GoTo Err_Handler
    SwapClicked
    Exit Sub
Err_Handler:
    HandleError "pgReview.btnSwap_Click", m_et
    RefreshList m_task
End Sub

' The empty state's button: clear the search and the column filters.
Private Sub emptyList_b_Click()
    On Error GoTo Err_Handler
    ClearGridFilters Me.sfList
    SetSearch Me, "txtSearch", ""
    Exit Sub
Err_Handler:
    HandleError "pgReview.emptyList_b_Click"
End Sub

' A row clicked: a task shows on the right; a question shows in full under
' the questions, and a double-click swaps it.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    Select Case grid
        Case "sfList"
            ShowTask Nz(Me.sfList.Form.txtSel, 0)
        Case "sfQs"
            ShowPicked
            If dbl And Me.btnSwap.Enabled Then SwapClicked
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgReview.GridClick", grid
    If grid = "sfQs" Then RefreshList m_task
End Sub

' A right-click on a row: its actions.
Public Sub GridRightClick(ByVal grid As String)
    Dim ok As Boolean, k As String, eq As Long
    On Error GoTo Err_Handler
    If m_et = 0 Then Exit Sub
    ok = Me.btnSignOff.Enabled
    Select Case grid
        Case "sfList"
            k = KitMenuAt(IIf(ok, "", "~") & "signoff|Sign off|E73E")
        Case "sfQs"
            ShowPicked
            k = KitMenuAt(IIf(ok, "", "~") & "swap|Swap...|E8AB", "-", "bank|Show in the question bank|E82D")
    End Select
    Select Case k
        Case "signoff": SignOffClicked
        Case "swap": SwapClicked
        Case "bank"
            eq = Nz(Me.sfQs.Form.txtSel, 0)
            If eq <> 0 Then
                TempVars("BankQuestionID") = Nz(DLookup("QuestionID", "tblExamQuestions", "ExamQuestionID=" & eq), 0)
                NavTo "questions"
            End If
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgReview.GridRightClick", grid
    RefreshList m_task
End Sub

' Sign off the shown topic, after asking: it can't be undone, and it closes
' the topic's other SMEs' reviews.
Private Sub SignOffClicked()
    Dim problem As String, done As String, n As Long
    If m_et = 0 Then Exit Sub
    problem = ReviewProblem(m_et)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: RefreshList m_task: Exit Sub
    n = Me.sfQs.Form.RecordsetClone.RecordCount
    If KitMsgBox("Sign off " & ExamTopicName(m_et) & "? Its " & n & IIf(n = 1, " question is", " questions are") & _
            " then settled for this exam, and the topic's other SMEs' reviews close. It can't be undone.", _
            vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then Exit Sub
    problem = SignOffTopic(m_et, m_stamp, done)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: RefreshList m_task: Exit Sub
    KitToast WithNoticeNote(done), "success"
    RefreshList 0                           ' it moves to Done
End Sub

' Swap the picked question for another of the topic (dlgSwap).
Private Sub SwapClicked()
    Dim eq As Long, problem As String
    If m_et = 0 Then Exit Sub
    eq = Nz(Me.sfQs.Form.txtSel, 0)
    If eq = 0 Then
        KitMsgBox "Pick the question to swap out in the list of questions drawn.", vbInformation, MSG_TITLE
        Exit Sub
    End If
    problem = ReviewProblem(m_et)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: RefreshList m_task: Exit Sub
    If RunDialog("dlgSwap", "eq=" & eq & ";stamp=" & m_stamp) Then KitToast Nz(TempVars("SwapDone"), "Swapped."), "success"
    ShowTask m_task                         ' the topic as it stands now, the same row picked
    If GridSelect(Me.sfQs, eq) Then ShowPicked Else SelectFirstRow Me.sfQs: ShowPicked
End Sub

' ---------------------------------------------------------------------------
' The list and the detail
' ---------------------------------------------------------------------------

' Load the person's tasks for the chip and the search, and show openId (0,
' or not listed: the first).
Private Sub RefreshList(ByVal openId As Long)
    If Not DataLinked() Then Exit Sub
    Me.sfList.Form.RecordSource = ReviewTasksSql(m_done, m_search)
    SetGridFilter Me.sfList, ""             ' the column funnels, on the new rows
    GridHighlight Me.sfList, m_search
    StyleChip Me.Controls("chip_0"), Not m_done, ReviewCount(False)
    StyleChip Me.Controls("chip_1"), m_done, ReviewCount(True)
    m_task = openId
    AfterFilter
End Sub

' After the rows shown changed: the empty state, and the shown task kept
' when still listed (else the first).
Private Sub AfterFilter()
    Dim n As Long
    n = Me.sfList.Form.RecordsetClone.RecordCount
    If ReviewCount(m_done) = 0 Then
        If m_done Then
            ShowEmpty Me, "emptyList", n = 0, "Nothing done yet", "The topics you or another SME signed off show here."
        Else
            ShowEmpty Me, "emptyList", n = 0, "Nothing to review", "When the DH generates an exam, each topic you're an SME of shows here."
        End If
        SetVisible Me.emptyList_b, False
    Else
        ShowEmpty Me, "emptyList", n = 0, "No topics match", "No review here fits the search and column filters."
        Me.emptyList_b.Caption = "Show all"
    End If
    SetVisible Me.sfList, n > 0
    If GridSelect(Me.sfList, m_task) Then
        ShowTask m_task
    ElseIf IsNull(SelectFirstRow(Me.sfList)) Then
        ShowTask 0
    Else
        ShowTask Nz(Me.sfList.Form.txtSel, 0)
    End If
End Sub

' Show task taskId on the right (0: the hint instead).
Private Sub ShowTask(ByVal taskId As Long)
    Dim rs As DAO.Recordset, isOpen As Boolean, c As Control, problem As String, s As String, n As Long, given As Long, replaced As Long
    Dim rj As String, cnt As String
    m_task = taskId: m_et = 0
    isOpen = (taskId <> 0)
    If isOpen Then
        Set rs = CurrentDb.OpenRecordset("SELECT tk.ExamTopicID, tk.TaskStatus, e.ExamID, e.RecipientName, e.CreatedByPersonID, " & _
            "e.CreatedOn, tm.TemplateName, tp.TopicName, et.QuestionCount " & _
            "FROM (((tblTasks AS tk INNER JOIN tblExamTopics AS et ON tk.ExamTopicID=et.ExamTopicID) " & _
            "INNER JOIN tblExams AS e ON et.ExamID=e.ExamID) INNER JOIN tblTopics AS tp ON et.TopicID=tp.TopicID) " & _
            "LEFT JOIN tblTemplates AS tm ON e.TemplateID=tm.TemplateID " & _
            "WHERE tk.TaskID=" & taskId & " AND tk.AssignedPersonID=" & SignedInPersonID(), dbOpenSnapshot)
        If rs.EOF Then
            isOpen = False: m_task = 0
        Else
            m_et = rs!ExamTopicID
            m_stamp = TopicChangeNo(m_et)   ' before the questions: a change after this one is caught
            m_keys = RecentKeysList(rs!ExamID)
            n = rs!QuestionCount
            Me.lblExam.Caption = rs!TopicName & " for " & rs!RecipientName
            Me.lblExamBy.Caption = "From the template " & Nz(rs!TemplateName, "(deleted)") & ", generated by " & _
                SignerName(rs!CreatedByPersonID) & " on " & Format$(rs!CreatedOn, "d mmm yyyy") & ". " & _
                n & IIf(n = 1, " question", " questions") & " from this topic."
            Me.sfQs.Form.RecordSource = ExamTopicQsSql(m_et, m_keys)
            problem = ReviewProblem(m_et)
            If Len(problem) = 0 Then
                given = GivenLatelyCount(m_et, m_keys)
                replaced = ReplacedCount(m_et)
                s = "Check each question and swap any that aren't right, then sign off. One sign-off completes the topic for all its SMEs."
                If replaced > 0 Then s = s & " " & replaced & " of them " & IIf(replaced = 1, "was", "were") & " replaced or retired " & _
                    "since the exam was drawn; swap " & IIf(replaced = 1, "it", "them") & " out (Swap lists the newer version when there is one)."
                If given > 0 Then s = s & " " & given & " of them " & IIf(given = 1, "was", "were") & " on an exam generated before this one; the " & _
                    "overlap rule allows " & RepeatsAllowed(n) & "."
                rj = ExamRejectionText(rs!ExamID)    ' sent back by an approver: why, first, then the counts
                If Len(rj) > 0 Then
                    cnt = ""
                    If replaced > 0 Then cnt = replaced & " replaced since the draw."
                    If given > 0 Then cnt = cnt & IIf(Len(cnt) > 0, " ", "") & given & " given lately; the rule allows " & RepeatsAllowed(n) & "."
                    If Len(rj) + Len(cnt) > 220 Then rj = Left$(rj, 165 - Len(cnt)) & "... (the rest is in its History on the Exams page)"
                    ShowAlert Me, "alReview", rj & IIf(Len(cnt) > 0, " " & cnt, ""), "warning", "Sent back"
                Else
                    ShowAlert Me, "alReview", s, "info"
                End If
            ElseIf rs!TaskStatus = TS_OPEN Then
                ShowAlert Me, "alReview", problem, "warning"
            Else
                ShowAlert Me, "alReview", problem, "success"
            End If
            SetEnabled Me.btnSignOff, Len(problem) = 0
            SetEnabled Me.btnSwap, Len(problem) = 0
        End If
        rs.Close
    End If
    For Each c In Me.Controls
        Select Case True
            Case c.Name Like "hdrQs*", c.Name = "sfQs", c.Name Like "lblExam*", c.Name Like "cap*", c.Name = "txtPreview", _
                    c.Name = "btnSignOff", c.Name = "btnSwap"
                SetVisible c, isOpen
        End Select
    Next
    If Not isOpen Then HideAlert Me, "alReview"
    SetVisible Me.lblDetailEmpty, Not isOpen
    If isOpen Then
        SelectFirstRow Me.sfQs
        ShowPicked
    End If
End Sub

' The picked question of the topic, in full, under the questions.
Private Sub ShowPicked()
    Dim eq As Long, qid As Long
    eq = Nz(Me.sfQs.Form.txtSel, 0)
    If eq <> 0 Then qid = Nz(DLookup("QuestionID", "tblExamQuestions", "ExamQuestionID=" & eq), 0)
    If qid <> 0 Then
        Me.txtPreview = QuestionPreview(qid)
    Else
        Me.txtPreview = "Pick a question above to read it in full."
    End If
End Sub
