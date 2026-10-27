Option Compare Database
Option Explicit

' The open topic on pgAdmTopics (Build_subTopic), bound to tblTopics. The
' page saves a changed topic before it opens another; every save checks
' the name (filled in, unique) and the questions per exam (1 to 50) in
' Form_BeforeUpdate and is refused until they are right. Save and the
' Active switch save at once and show the change in the page's list.

Private Const REQUIRED As String = "txtName=Topic;txtCount=Questions per exam"

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim problem As String
    On Error GoTo Err_Handler
    If Not CheckRequired(Me, REQUIRED) Then Cancel = True: Exit Sub
    problem = TopicProblem()
    If Len(problem) > 0 Then
        KitRefuseSave
        KitMsgBox problem, vbExclamation, "Topics"
        Cancel = True
    End If
    Exit Sub
Err_Handler:
    Cancel = True
    HandleError "subTopic.Form_BeforeUpdate"
End Sub

' A save the page caused (another row, a search): its list catches up after.
Private Sub Form_AfterUpdate()
    On Error GoTo Err_Handler
    Me.Parent.TopicChanged
    Exit Sub
Err_Handler:
    HandleError "subTopic.Form_AfterUpdate"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    KitFormError Me, DataErr, Response
End Sub

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    SaveTopic
    Exit Sub
Err_Handler:
    HandleError "subTopic.btnSave_Click"
End Sub

' The Active switch (Toggle's onChange): save at once, like a settings switch.
Public Sub ActiveChanged()
    On Error GoTo Err_Handler
    If Not Me.Dirty Then Exit Sub
    If KitSaveRecord(Me) Then Me.Parent.TopicSaved Me!TopicID
    Exit Sub
Err_Handler:
    HandleError "subTopic.ActiveChanged"
End Sub

' Save the topic now (Save, Ctrl+S). False when it can't be saved yet.
Public Function SaveTopic() As Boolean
    ' Checked here as well as in Form_BeforeUpdate: from a click the focus
    ' can move to the first empty field, which it can't during an update.
    If Not CheckRequired(Me, REQUIRED) Then Exit Function
    If Not Me.Dirty Then SaveTopic = True: Exit Function
    If Not KitSaveRecord(Me) Then Exit Function
    Me.Parent.TopicSaved Me!TopicID
    KitToast "Topic saved.", "success"
    SaveTopic = True
End Function

' What is wrong with the topic as typed, or "".
Private Function TopicProblem() As String
    Dim n As Variant
    TopicProblem = NameProblem(Nz(Me.txtName, ""), 100, "tblTopics", "TopicName", "topic", "TopicID=" & Nz(Me!TopicID, 0))
    If Len(TopicProblem) > 0 Then Exit Function
    n = Me.txtCount
    If Not IsNumeric(n) Then
        TopicProblem = "Questions per exam must be a whole number from 1 to 50."
    ElseIf n < 1 Or n > 50 Or n <> Int(n) Then
        TopicProblem = "Questions per exam must be a whole number from 1 to 50."
    End If
End Function
