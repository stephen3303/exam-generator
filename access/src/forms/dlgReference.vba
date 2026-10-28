Option Compare Database
Option Explicit

' References (Build_dlgReference). OpenArgs mode:
'   "mode=cite"            the editor cites a reference: pick one already
'                          listed, or add a new one, and give the paragraph.
'                          Writes a citation to the working copy (wq_ref) and
'                          hands it back in TempVars("WorkRefID").
'   "mode=cite;id=<wq id>" the editor changes a citation's paragraph.
'   "mode=new"             the question bank adds a reference to the list,
'                          handed back in TempVars("RefID").
' A new reference goes into the shared list (tblSourceRefs) at once, so the
' other SMEs can cite it too, even if the question isn't saved. RunDialog
' sees DialogResult "ok".
'
' Owner: the question bank (build phase 5). Depends on: modErrors,
' modGlobals, modQuestions, modUI; the kit's modKitLocks, modKitUI,
' modKitUtil, modMenu, modUIKit.

Private Const MSG_TITLE As String = "Reference"

Private m_mode As String        ' cite or new
Private m_id As Long            ' the citation being changed (0 when citing anew)

Public Property Get HelpTopic() As String
    HelpTopic = "qedit.reference"
End Property

Private Sub Form_Load()
    Dim rs As DAO.Recordset
    On Error GoTo Err_Handler
    m_mode = LCase$(OpenArgText(Me, "mode"))
    m_id = OpenArgLong(Me, "id")
    If m_mode = "new" Then
        Me.Caption = "Add a reference"
        Me.lblHead.Caption = "Add a reference to the list"
        HideSegment
        ShowSource "new"
        ShowPara False
        Me.lblNote.Caption = "It goes into the shared list, where every SME can cite it. Give the number and revision " & _
            "as the document shows them."
        Me.txtTitle.SetFocus
    ElseIf m_id <> 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT RefID, Paragraph FROM wq_ref WHERE id=" & m_id, dbOpenSnapshot)
        If Not rs.EOF Then
            Me.lblHead.Caption = RefText(rs!RefID)
            Me.txtPara = rs!Paragraph
        End If
        rs.Close
        Me.Caption = "Change the paragraph"
        HideSegment
        ShowSource ""
        Me.capPara.Top = Me.capRef.Top: Me.txtPara.Top = Me.txtRef.Top
        Me.lblParaHint.Top = Me.txtRef.Top - px(4)
        Me.lblNote.Caption = "To cite another reference, remove this one and cite that."
        Me.txtPara.SetFocus
    Else
        Me.Caption = "Cite a reference"
        Me.lblHead.Caption = "Cite a reference"
        ShowSource "pick"
        Me.txtRef.SetFocus
    End If
    CenterOnApp Me
    Exit Sub
Err_Handler:
    HandleError "dlgReference.Form_Load", Nz(Me.OpenArgs, "")
End Sub

Private Sub txtRef_KeyDown(KeyCode As Integer, Shift As Integer)
    KitPickKey Me.txtRef, KeyCode
End Sub

' The switch moved (its onChange): pick a listed reference, or add one.
Public Sub SourceChanged()
    On Error GoTo Err_Handler
    ShowSource SegmentValue(Me, "segSource")
    Exit Sub
Err_Handler:
    HandleError "dlgReference.SourceChanged"
End Sub

Private Sub btnOK_Click()
    Dim rid As Long, problem As String, para As Variant
    On Error GoTo Err_Handler
    para = NullIfEmpty(Me.txtPara)
    If Len(Nz(para, "")) > 100 Then
        KitMsgBox "The paragraph can be at most 100 characters.", vbExclamation, MSG_TITLE
        Me.txtPara.SetFocus
        Exit Sub
    End If
    ' Changing a citation's paragraph.
    If m_id <> 0 Then
        KitExecute "UPDATE wq_ref SET Paragraph=" & IIf(IsNull(para), "Null", Q(para)) & " WHERE id=" & m_id
        TempVars("WorkRefID") = m_id
        Done
        Exit Sub
    End If
    ' The reference: picked, or new.
    If m_mode = "new" Or SegmentValue(Me, "segSource") = "new" Then
        problem = RefProblem(Me.txtTitle, Me.txtNumber, Me.txtRev)
        If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Sub
        rid = AddReference(Me.txtTitle, Me.txtNumber, Me.txtRev)
        TempVars("RefID") = rid
        If m_mode = "new" Then Done: Exit Sub
    Else
        rid = Nz(Me.txtRefKey, 0)
        If rid = 0 Then
            KitMsgBox "Type part of a reference's title or number, then pick it from the list under the box. " & _
                "If it isn't listed, choose Add a new reference.", vbInformation, MSG_TITLE
            Me.txtRef.SetFocus
            Exit Sub
        End If
        If DCount("*", "wq_ref", "RefID=" & rid) > 0 Then
            KitMsgBox "The question already cites " & RefText(rid) & ". Change its paragraph there instead.", vbInformation, MSG_TITLE
            Exit Sub
        End If
    End If
    TempVars("WorkRefID") = AddCitation(rid, para)
    Done
    Exit Sub
Err_Handler:
    If Err.Number = 3022 Then               ' the same title and number added at this moment (uxRefTitleNumber)
        KitMsgBox "Someone else has just added that reference, so it is in the list now. Pick it from the list instead.", _
            vbInformation, MSG_TITLE
        Exit Sub
    End If
    HandleError "dlgReference.btnOK_Click", rid
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgReference.btnCancel_Click"
End Sub

' ---------------------------------------------------------------------------

Private Sub Done()
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function AddCitation(ByVal rid As Long, ByVal para As Variant) As Long
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM wq_ref WHERE False", dbOpenDynaset)
    rs.AddNew
    rs!RefID = rid
    rs!Paragraph = para
    rs.Update
    rs.Bookmark = rs.LastModified
    AddCitation = rs!id
    rs.Close
End Function

' Show the picker (pick), the new reference's fields (new), or neither ("").
Private Sub ShowSource(ByVal k As String)
    Dim nm As Variant
    For Each nm In Array("capRef", "txtRef")
        SetVisible Me.Controls(nm), k = "pick"
    Next
    For Each nm In Array("capTitle", "txtTitle", "capNumber", "txtNumber", "capRev", "txtRev")
        SetVisible Me.Controls(nm), k = "new"
    Next
    If m_mode = "cite" And m_id = 0 Then
        Me.lblNote.Caption = IIf(k = "new", "OK adds it to the shared list at once, so other SMEs can cite it too, " & _
            "even if you don't save the question.", "Only references turned on are listed. The admin or an SME can turn " & _
            "one on again on the question bank's References tab.")
    End If
End Sub

Private Sub ShowPara(ByVal shown As Boolean)
    SetVisible Me.capPara, shown
    SetVisible Me.txtPara, shown
    SetVisible Me.lblParaHint, shown
End Sub

' The switch isn't used when adding to the list or changing a paragraph.
Private Sub HideSegment()
    Dim c As Control
    For Each c In Me.Controls
        If c.Name Like "segSource_*" Then c.Visible = False
    Next
End Sub
