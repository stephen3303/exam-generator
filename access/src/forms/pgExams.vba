Option Compare Database
Option Explicit

' Exams (see Build_pgExams). The list of exams comes in build phase 8; until
' then the page holds the department head's New exam button, which opens
' dlgNewExam. Only a DH sees it, and the right is checked again in the data
' file when the exam is generated (modExams.RequireExamMaker).

Private Const MSG_TITLE As String = "Exams"

Public Property Get HelpTopic() As String
    HelpTopic = "exams"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    SetVisible Me.btnNew, HasRole("DH") Or Not DataLinked()
    SetEnabled Me.btnNew, DataLinked()
    Exit Sub
Err_Handler:
    HandleError "pgExams.Form_Load"
End Sub

Private Sub btnNew_Click()
    Dim problem As String
    On Error GoTo Err_Handler
    problem = ExamMakerProblem()
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: Exit Sub
    If RunDialog("dlgNewExam") Then KitToast Nz(TempVars("NewExamDone"), "Exam generated."), "success"
    Exit Sub
Err_Handler:
    HandleError "pgExams.btnNew_Click"
End Sub
