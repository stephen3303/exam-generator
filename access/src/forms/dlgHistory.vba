Option Compare Database
Option Explicit

' A question's history (Build_dlgHistory). OpenArgs "key=<QuestionKey>":
' every version's audit rows, newest first (HistorySql): who wrote,
' changed, submitted, approved, rejected and retired it, and when. The
' circle is the first version's author. Copy puts it on the clipboard as
' text.

Private m_key As Long

Public Property Get HelpTopic() As String
    HelpTopic = "questions.history"
End Property

Private Sub Form_Load()
    Dim rs As DAO.Recordset, n As Long, nVer As Long, st As String
    On Error GoTo Err_Handler
    m_key = Val(Mid$(Nz(Me.OpenArgs, ""), InStr(Nz(Me.OpenArgs, "") & "=", "=") + 1))
    Me.Caption = "History of Q" & m_key
    Me.lblHead.Caption = "History of Q" & m_key
    Me.sfHist.Form.RecordSource = HistorySql(m_key)
    n = Me.sfHist.Form.RecordsetClone.RecordCount
    nVer = DCount("*", "tblQuestions", "QuestionKey=" & m_key)
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 q.VersionNo, q.StatusID, " & PersonNameSql() & " AS author " & _
        "FROM tblQuestions AS q LEFT JOIN " & PERSON_TABLE & " AS p ON q.AuthorPersonID=p.person_id " & _
        "WHERE q.QuestionKey=" & m_key & " ORDER BY q.VersionNo", dbOpenSnapshot)
    If Not rs.EOF Then Me.txtAuthor = rs!author
    rs.Close
    st = Nz(DLookup("StatusID", "tblQuestions", "QuestionKey=" & m_key & " AND IsCurrent"), 0)
    Me.lblSub.Caption = "Written by " & Nz(Me.txtAuthor, "someone no longer listed") & ". " & nVer & IIf(nVer = 1, " version", " versions") & _
        IIf(Val(st) <> 0, "; the current one is " & LCase$(QStatusName(Val(st))) & ".", ".")
    SetVisible Me.sfHist, n > 0
    SetEnabled Me.btnCopy, n > 0
    CenterOnApp Me
    Exit Sub
Err_Handler:
    HandleError "dlgHistory.Form_Load", m_key
End Sub

' A row clicked (subGrid): nothing more to show than the row itself.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
End Sub

Private Sub btnCopy_Click()
    On Error GoTo Err_Handler
    If KitCopyText(HistoryText(m_key)) Then KitToast "History copied.", "success"
    Exit Sub
Err_Handler:
    HandleError "dlgHistory.btnCopy_Click", m_key
End Sub

Private Sub btnClose_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgHistory.btnClose_Click"
End Sub
