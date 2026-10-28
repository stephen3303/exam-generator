Option Compare Database
Option Explicit

' Reject a question (Build_dlgReject), from the Approvals page. OpenArgs
' "id=<QuestionID>". Shows who it goes back to and asks the reason, which is
' required. Reject sends it back (modApprovals.RejectQuestion: Rejected, at
' no step, the reason in the audit row); RunDialog sees DialogResult "ok".
' Cancel, or Esc, changes nothing.

Private Const MSG_TITLE As String = "Reject"

Private m_qid As Long

Public Property Get HelpTopic() As String
    HelpTopic = "approvals.reject"
End Property

Private Sub Form_Load()
    Dim key As Variant, ver As Variant
    On Error GoTo Err_Handler
    m_qid = OpenArgLong(Me, "id")
    key = DLookup("QuestionKey", "tblQuestions", "QuestionID=" & m_qid)
    ver = DLookup("VersionNo", "tblQuestions", "QuestionID=" & m_qid)
    Me.Caption = "Reject " & QuestionRef(key, ver)
    Me.lblHead.Caption = "Reject " & QuestionRef(key, ver) & "?"
    Me.lblSub.Caption = "It goes back to its author as rejected, with your reason. Once changed and submitted again, it starts at the DH."
    Me.txtTo = ReturnsToText(m_qid)
    CenterOnApp Me
    Me.txtReason.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgReject.Form_Load", Nz(Me.OpenArgs, "")
End Sub

Private Sub btnReject_Click()
    Dim problem As String
    On Error GoTo Err_Handler
    problem = ReasonProblem(Me.txtReason)
    If Len(problem) > 0 Then
        KitMsgBox problem, vbExclamation, MSG_TITLE
        Me.txtReason.SetFocus
        Exit Sub
    End If
    problem = RejectQuestion(m_qid, Me.txtReason)
    If Len(problem) > 0 Then                ' someone acted on it first: nothing to fix here
        KitMsgBox problem, vbInformation, MSG_TITLE
        DoCmd.Close acForm, Me.Name, acSaveNo
        Exit Sub
    End If
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgReject.btnReject_Click", m_qid
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgReject.btnCancel_Click"
End Sub
