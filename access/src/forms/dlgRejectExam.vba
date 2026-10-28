Option Compare Database
Option Explicit

' Send an exam back (Build_dlgRejectExam), from the Approvals page's Exams
' tab. OpenArgs "id=<ExamID>;step=<the step's RoleID the page showed>;
' ver=<the version it showed>". The
' approver picks one topic or every topic, sees whose reviews reopen
' (ExamReturnsToText) and gives the reason, which is required. Send back
' calls modExamApprovals.RejectExam: those topics' sign-offs are cleared,
' their SMEs get new reviews and the exam is Rejected, its version one
' higher. RunDialog sees DialogResult "ok". Cancel, or Esc, changes
' nothing.

Private Const MSG_TITLE As String = "Send back"

Private m_id As Long
Private m_step As Long          ' the step the page showed: sends back from that step only
Private m_ver As Long           ' the version the page showed: sends back that version only

Public Property Get HelpTopic() As String
    HelpTopic = "approvals.exams"
End Property

Private Sub Form_Load()
    Dim ref As String
    On Error GoTo Err_Handler
    m_id = OpenArgLong(Me, "id")
    m_step = OpenArgLong(Me, "step")
    m_ver = OpenArgLong(Me, "ver")
    ref = ExamRef(m_id, IIf(m_ver <> 0, m_ver, Nz(DLookup("VersionNo", "tblExams", "ExamID=" & m_id), 1)))
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
    problem = RejectExam(m_id, Me.txtReason, Nz(Me.cboTopic, 0), m_step, m_ver)
    If Len(problem) > 0 Then
        KitMsgBox problem, vbInformation, MSG_TITLE
        ' Acted on first (moved on, or sent back and resubmitted): nothing to fix here.
        If Len(ExamApproveProblem(m_id, m_step, m_ver)) > 0 Then
            DoCmd.Close acForm, Me.Name, acSaveNo
        End If
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
