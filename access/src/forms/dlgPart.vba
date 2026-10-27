Option Compare Database
Option Explicit

' A part of the question open in the editor (Build_dlgPart). OpenArgs:
' "total=<question's points>;mark=<ClassLevelID>" to add a part (the mark
' starts its portion marking), or "id=<wq_part id>;total=<points>" to
' change one. The label is given by the part's place (A, B...) and the
' editor letters the parts again after a move, so it can't be typed. OK
' checks the part's text, points and marking, writes the working copy
' (wq_part) and hands back the part in TempVars("PartID"); RunDialog sees
' DialogResult "ok". Nothing reaches the data file until the editor's Save.

Private Const MSG_TITLE As String = "Part"

Private m_id As Long            ' the part being changed (0 when adding)
Private m_total As Variant      ' the question's points, as typed in the editor

Public Property Get HelpTopic() As String
    HelpTopic = "qedit.part"
End Property

Private Sub Form_Load()
    Dim rs As DAO.Recordset, mark As Long
    On Error GoTo Err_Handler
    m_id = ArgLong("id")
    m_total = ArgText("total")
    If m_id <> 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM wq_part WHERE id=" & m_id, dbOpenSnapshot)
        If Not rs.EOF Then
            Me.txtLabel = rs!PartLabel
            Me.txtPartPoints = PointsText(rs!Points)
            Me.cboPartMark = rs!ClassLevelID
            Me.txtPartText = rs!PartText
            Me.txtPartAns = rs!ReferenceAnswer
        End If
        rs.Close
        Me.Caption = "Change part " & Me.txtLabel
        Me.lblHead.Caption = "Part " & Me.txtLabel
    Else
        Me.txtLabel = Chr$(65 + WorkPartCount())
        mark = ArgLong("mark")
        If mark <> 0 Then Me.cboPartMark = mark
        If Len(OthersPoints()) > 0 Then Me.txtPartPoints = OthersPoints()
        Me.Caption = "Add a part"
        Me.lblHead.Caption = "Add part " & Me.txtLabel
    End If
    Me.txtLabel.Locked = True
    Me.txtLabel.TabStop = False
    Me.txtLabel.ControlTipText = "Parts are lettered by their order; move a part to change its letter"
    ShowHint
    CenterOnApp Me
    Me.txtPartText.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgPart.Form_Load", Nz(Me.OpenArgs, "")
End Sub

Private Sub txtPartPoints_AfterUpdate()
    On Error GoTo Err_Handler
    If IsNumeric(Me.txtPartPoints) Then Me.txtPartPoints = PointsText(Me.txtPartPoints)
    ShowHint
    Exit Sub
Err_Handler:
    HandleError "dlgPart.txtPartPoints_AfterUpdate"
End Sub

Private Sub btnOK_Click()
    Dim s As String, p As String, t As Variant, ans As Variant, rs As DAO.Recordset
    On Error GoTo Err_Handler
    t = CleanText(Me.txtPartText)
    ans = CleanText(Me.txtPartAns)
    If IsNull(t) Then s = s & ChrW(&H2022) & " Text: write the part as it will print." & vbCrLf
    If Len(Nz(t, "")) > MAX_TEXT Then s = s & ChrW(&H2022) & " Text can be at most " & MAX_TEXT & " characters; it is " & Len(t) & "." & vbCrLf
    If Len(Nz(ans, "")) > MAX_TEXT Then s = s & ChrW(&H2022) & " The reference answer can be at most " & MAX_TEXT & " characters." & vbCrLf
    p = PointsProblem(Me.txtPartPoints, "Points")
    If Len(p) > 0 Then s = s & ChrW(&H2022) & " " & p & vbCrLf
    If IsNull(Me.cboPartMark) Then s = s & ChrW(&H2022) & " Portion marking: pick the part's marking." & vbCrLf
    If Len(s) > 0 Then
        KitMsgBox "The part can't be kept yet:" & vbCrLf & vbCrLf & s, vbExclamation, MSG_TITLE
        Exit Sub
    End If
    If m_id <> 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM wq_part WHERE id=" & m_id, dbOpenDynaset)
        If rs.EOF Then rs.Close: KitMsgBox "That part was removed.", vbInformation, MSG_TITLE: Exit Sub
        rs.Edit
    Else
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM wq_part WHERE False", dbOpenDynaset)
        rs.AddNew
        rs!PartLabel = Me.txtLabel
        rs!SortOrder = NextOrder("wq_part", "SortOrder")
    End If
    rs!PartText = t
    rs!Points = CDbl(Me.txtPartPoints)
    rs!ClassLevelID = Me.cboPartMark
    rs!ReferenceAnswer = ans
    rs.Update
    rs.Bookmark = rs.LastModified
    TempVars("PartID") = rs!id
    rs.Close
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgPart.btnOK_Click", m_id
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgPart.btnCancel_Click"
End Sub

' ---------------------------------------------------------------------------

' The line under the fields: this part's share of the question's points.
Private Sub ShowHint()
    Dim others As Double, s As String
    others = Nz(DSum("Points", "wq_part", "id<>" & m_id), 0)
    If Len(PointsProblem(m_total, "Points")) > 0 Then
        s = "The other parts give " & PointsText(others) & " points. Set the question's points in the editor so the parts can add up to them."
    Else
        s = "The question is worth " & PointsText(m_total) & " points and the other parts give " & PointsText(others) & _
            ", so this part's share is " & PointsText(CDbl(m_total) - others) & "."
    End If
    Me.lblPointsHint.Caption = s
End Sub

' What the question's points leave for a new part, "" when nothing (or
' the question's points aren't set).
Private Function OthersPoints() As String
    Dim remain As Double
    If Len(PointsProblem(m_total, "Points")) > 0 Then Exit Function
    remain = CDbl(m_total) - Nz(DSum("Points", "wq_part"), 0)
    If remain > 0 Then OthersPoints = PointsText(remain)
End Function

' The number after name= in OpenArgs ("id=3;total=4"), or 0.
Private Function ArgLong(ByVal name As String) As Long
    Dim v As String
    v = ArgText(name)
    If IsNumeric(v) Then ArgLong = CLng(v)
End Function

' The text after name= in OpenArgs, "" when it isn't there.
Private Function ArgText(ByVal name As String) As String
    Dim part As Variant
    For Each part In Split(Nz(Me.OpenArgs, ""), ";")
        If LCase$(Left$(part, Len(name) + 1)) = LCase$(name) & "=" Then ArgText = Mid$(part, Len(name) + 2)
    Next
End Function
