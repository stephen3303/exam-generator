Option Compare Database
Option Explicit

' The open template on pgAdmTemplates (Build_subTemplate), bound to
' tblTemplates. The page saves a changed template before it opens another;
' every save checks the name (filled in, unique) in Form_BeforeUpdate and
' is refused until it is right. Save and the Active switch save at once and
' show the change in the page's list.

Private Const REQUIRED As String = "txtName=Template"

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim problem As String
    On Error GoTo Err_Handler
    RequireExamAdmin                        ' the right is checked now, not as of sign-in
    If Not CheckRequired(Me, REQUIRED) Then Cancel = True: Exit Sub
    problem = NameProblem(Nz(Me.txtName, ""), 100, "tblTemplates", "TemplateName", "template", _
        "TemplateID=" & Nz(Me!TemplateID, 0))
    If Len(problem) > 0 Then
        KitRefuseSave
        KitMsgBox problem, vbExclamation, "Templates"
        Cancel = True
    End If
    Exit Sub
Err_Handler:
    Cancel = True
    HandleError "subTemplate.Form_BeforeUpdate"
End Sub

' A save the page caused (another row, a search): its list catches up after.
Private Sub Form_AfterUpdate()
    On Error GoTo Err_Handler
    Me.Parent.TemplateChanged
    Exit Sub
Err_Handler:
    HandleError "subTemplate.Form_AfterUpdate"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    KitFormError Me, DataErr, Response
End Sub

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    SaveTemplate
    Exit Sub
Err_Handler:
    HandleError "subTemplate.btnSave_Click"
End Sub

' The Active switch (Toggle's onChange): save at once, like a settings switch.
Public Sub ActiveChanged()
    On Error GoTo Err_Handler
    If Not Me.Dirty Then Exit Sub
    If KitSaveRecord(Me) Then Me.Parent.TemplateSaved Me!TemplateID
    Exit Sub
Err_Handler:
    HandleError "subTemplate.ActiveChanged"
End Sub

' Save the template now (Save, Ctrl+S, or before its topics change). False
' when it can't be saved yet.
Public Function SaveTemplate(Optional ByVal quiet As Boolean = False) As Boolean
    If Not CheckRequired(Me, REQUIRED) Then Exit Function
    If Not Me.Dirty Then SaveTemplate = True: Exit Function
    If Not KitSaveRecord(Me) Then Exit Function
    Me.Parent.TemplateSaved Me!TemplateID
    If Not quiet Then KitToast "Template saved.", "success"
    SaveTemplate = True
End Function
