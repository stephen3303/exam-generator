Option Compare Database
Option Explicit

' Export the question bank (Build_dlgBankExport), from the Question bank
' page's Export menu, for the admin. Every topic is listed with how many
' questions export would write from it, all ticked at first. Export asks
' where to save the workbook (KitSaveAs) and writes it (modBank.ExportBank)
' under the busy veil; RunDialog sees DialogResult "ok" and the page shows
' TempVars("BankExportNote") as a toast. Cancel, or Esc, writes nothing.
'
' Owner: the question bank's export and import. Depends on: modBank,
' modErrors; the kit's modKitFiles, modKitUI, modKitUtil.

Private Const MSG_TITLE As String = "Export the question bank"

Public Property Get HelpTopic() As String
    HelpTopic = "questions.bank"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    TempVars("BankExportNote") = Null
    Me.lstTopics.RowSource = BankTopicsSql()
    SelectAll True
    CenterOnApp Me
    Exit Sub
Err_Handler:
    HandleError "dlgBankExport.Form_Load"
End Sub

' Select all / Select none.
Public Function SelectAll(ByVal onOff As Boolean)
    Dim i As Long
    On Error GoTo Err_Handler
    For i = 0 To Me.lstTopics.ListCount - 1
        Me.lstTopics.Selected(i) = onOff
    Next
    ShowCount
    Exit Function
Err_Handler:
    HandleError "dlgBankExport.SelectAll", onOff
End Function

Private Sub lstTopics_Click()
    On Error GoTo Err_Handler
    ShowCount
    Exit Sub
Err_Handler:
    HandleError "dlgBankExport.lstTopics_Click"
End Sub

Private Sub btnExport_Click()
    Dim ids As String, dest As String, problem As String, n As Long, busy As Boolean
    Dim num As Long, desc As String, src As String
    On Error GoTo Err_Handler
    ids = PickedIds()
    If Len(ids) = 0 Then KitMsgBox "Tick at least one topic to export.", vbInformation, MSG_TITLE: Exit Sub
    dest = KitSaveAs("Export the question bank", "Question bank " & Format$(Date, "yyyy-mm-dd") & ".xlsx", _
        "Excel workbook|*.xlsx", "bankexport")
    If Len(dest) = 0 Then Exit Sub
    KitBusy "Exporting the question bank to Excel...": busy = True
    problem = ExportBank(dest, ids, n)
    KitBusyEnd: busy = False
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Sub
    TempVars("BankExportNote") = "Exported " & n & " " & Plural(n, "question") & " to " & Mid$(dest, InStrRev(dest, "\") + 1) & "."
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    ' KitBusyEnd resets Err, so keep the error to show it.
    num = Err.Number: desc = Err.Description: src = Err.Source
    If busy Then KitBusyEnd
    Err.Number = num: Err.Description = desc: Err.Source = src
    HandleError "dlgBankExport.btnExport_Click", dest
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgBankExport.btnCancel_Click"
End Sub

' The ticked topics' TopicIDs, separated by commas ("" for none).
Private Function PickedIds() As String
    Dim v As Variant, s As String
    For Each v In Me.lstTopics.ItemsSelected
        s = s & "," & Me.lstTopics.ItemData(v)
    Next
    If Len(s) > 0 Then PickedIds = Mid$(s, 2)
End Function

' How many questions and topics are ticked, and Export only with one.
Private Sub ShowCount()
    Dim ids As String, n As Long, k As Long
    ids = PickedIds()
    k = Me.lstTopics.ItemsSelected.Count
    If k > 0 Then n = BankExportCount(ids)
    If k = 0 Then
        Me.lblCount.Caption = "No topic ticked."
    Else
        Me.lblCount.Caption = n & " " & Plural(n, "question") & " in " & k & " " & Plural(k, "topic") & _
            ". Each question's version in use is written; an open new version stays in the bank."
    End If
    SetEnabled Me.btnExport, k > 0
End Sub
