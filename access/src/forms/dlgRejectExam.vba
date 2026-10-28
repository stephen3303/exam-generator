Option Compare Database
Option Explicit

' Send an exam back (Build_dlgRejectExam), from the Approvals page's Exams
' tab. OpenArgs "id=<ExamID>;step=<the step's RoleID the page showed>". The
' approver picks one topic or every topic, sees whose reviews reopen
' (ExamReturnsToText) and gives the reason, which is required. Send back
' calls modExamApprovals.RejectExam: those topics' sign-offs are cleared,
' their SMEs get new reviews and the exam is Rejected, its version one
' higher. RunDialog sees DialogResult "ok". Cancel, or Esc, changes
' nothing.

Private Const MSG_TITLE As String = "Send back"

Private m_id As Long
Private m_step As Long          ' the step the page showed: sends back from that step only

Public Property Get HelpTopic() As String
    HelpTopic = "approvals.exams"
End Property

Private Sub Form_Load()
    Dim ref As String
    On Error GoTo Err_Handler
    m_id = ArgLong("id")
    m_step = ArgLong("step")
    ref = ExamRef(m_id, DLookup("VersionNo", "tblExams", "ExamID=" & m_id))
    Me.Caption = "Send back " & ref
    Me.lblHead.Caption = "Send " & ref & " back to its SMEs?"
    Me.lblSub.Caption = "The topics you send back lose their sign-off and their SMEs review them again. Once every topic is " & _
        "signed off again, the exam starts again at the DH."
    Me.cboTopic.RowSource = "SELECT 0 AS id, '(Every topic)' AS nm, 0 AS ord FROM tblExams WHERE ExamID=" & m_id & _
        " UNION SELECT et.ExamTopicID, tp.TopicName, 1 FROM tblExamTopics AS et INNER JOIN tblTopics AS tp ON et.TopicID=tp.TopicID " & _
        "WHERE et.ExamID=" & m_id & " ORDER BY ord, nm"
    Me.cboTopic = 0
    ShowWho
    CenterOnApp Me
    Me.txtReason.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgRejectExam.Form_Load", Nz(Me.OpenArgs, "")
End Sub

Private Sub cboTopic_AfterUpdate()
    On Error GoTo Err_Handler
    If IsNull(Me.cboTopic) Then Me.cboTopic = 0
    ShowWho
    Exit Sub
Err_Handler:
    HandleError "dlgRejectExam.cboTopic_AfterUpdate", m_id
End Sub

Private Sub btnReject_Click()
    Dim problem As String
    On Error GoTo Err_Handler
    problem = ReasonProblem(Me.txtReason, EXAM_REASON_WHO)
    If Len(problem) > 0 Then
        KitMsgBox problem, vbExclamation, MSG_TITLE
        Me.txtReason.SetFocus
        Exit Sub
    End If
    problem = RejectExam(m_id, Me.txtReason, Nz(Me.cboTopic, 0), m_step)
    If Len(problem) > 0 Then
        KitMsgBox problem, vbInformation, MSG_TITLE
        If Len(ExamApproveProblem(m_id, m_step)) > 0 Then DoCmd.Close acForm, Me.Name, acSaveNo   ' acted on first: nothing to fix here
        Exit Sub
    End If
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgRejectExam.btnReject_Click", m_id
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgRejectExam.btnCancel_Click"
End Sub

' Whose reviews reopen for the topic picked; Send back waits while a topic
' it would reopen has no SME.
Private Sub ShowWho()
    Me.txtTo = ExamReturnsToText(m_id, Nz(Me.cboTopic, 0))
    SetEnabled Me.btnReject, Len(ExamTopicsWithoutSme(m_id, Nz(Me.cboTopic, 0))) = 0
End Sub

' The number after name= in OpenArgs ("id=3"), or 0.
Private Function ArgLong(ByVal name As String) As Long
    Dim part As Variant, v As String
    For Each part In Split(Nz(Me.OpenArgs, ""), ";")
        If LCase$(Left$(part, Len(name) + 1)) = LCase$(name) & "=" Then v = Mid$(part, Len(name) + 2)
    Next
    If IsNumeric(v) Then ArgLong = CLng(v)
End Function
