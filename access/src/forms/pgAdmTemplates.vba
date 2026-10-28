Option Compare Database
Option Explicit

' Admin: Templates (see Build_pgAdmTemplates). The templates on the left
' with how many topics and questions each gives; on the right the open
' template in subTemplate (bound to tblTemplates) and its topics in exam
' order (tblTemplateTopics SortOrder) with their question counts.
'
' New template asks for the name and adds it, active and empty. Add topic
' and Change open dlgTemplateTopic; the arrows move the selected topic;
' Remove takes it off (the topic itself stays). More, or a right-click on a
' row, copies the template with its topics, or deletes it while no exam
' was made from it.

Private Const MSG_TITLE As String = "Templates"

Private m_id As Long            ' the open template (0 when none)
Private m_search As String
Private m_stale As Boolean      ' subTemplate saved a change the list doesn't show yet

Private Sub Form_Load()
    On Error GoTo Err_Handler
    If Not DataLinked() Then
        ShowEmpty Me, "emptyList", True, "No data file linked", "This copy was built without the data file, so there are no templates to show."
        SetVisible Me.emptyList_b, False
        SetVisible Me.sfList, False
        SetEnabled Me.btnNew, False
        ClearTemplate
        Exit Sub
    End If
    RefreshList 0
    Exit Sub
Err_Handler:
    HandleError "pgAdmTemplates.Form_Load"
End Sub

' frmMain asks before it swaps or closes this page: save the open template,
' or stay while it can't be saved (subTemplate says why).
Public Function CanLeave() As Boolean
    On Error GoTo Err_Handler
    CanLeave = SaveDetail(Me, "sfDetail")
    Exit Function
Err_Handler:
    HandleError "pgAdmTemplates.CanLeave"
End Function

' Ctrl+S saves the open template, Ctrl+N adds one (modKitKeys).
Public Function Shortcut(ByVal action As String) As Boolean
    On Error GoTo Err_Handler
    Select Case action
        Case SHORTCUT_SAVE
            If m_id <> 0 Then Me.sfDetail.Form.SaveTemplate
            Shortcut = True
        Case SHORTCUT_NEW
            NewTemplate
            Shortcut = True
    End Select
    Exit Function
Err_Handler:
    HandleError "pgAdmTemplates.Shortcut", action
    Shortcut = True
End Function

' Row click: in the list, open that template once the open one is saved;
' in its topics, select one (a double-click changes its count).
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    Dim id As Long
    On Error GoTo Err_Handler
    If grid = "sfTopics" Then
        ShowTopicState
        If dbl Then ChangeTopic
        Exit Sub
    End If
    id = Nz(Me.sfList.Form.txtSel, 0)
    If id = 0 Then Exit Sub
    If Not ShowTemplate(id) Then
        GridSelect Me.sfList, m_id
        Exit Sub
    End If
    If m_stale Then RefreshList id          ' the template left behind was saved with a change
    Exit Sub
Err_Handler:
    HandleError "pgAdmTemplates.GridClick", id
End Sub

' Right-click on a row: the template's actions, or a topic's.
Public Sub GridRightClick(ByVal grid As String)
    On Error GoTo Err_Handler
    If grid = "sfTopics" Then
        Select Case KitMenuAt("chg|Change count|E70F", "up|Move up|E70E", "down|Move down|E70D", "-", _
                "rem|!Remove from template|E74D")
            Case "chg": ChangeTopic
            Case "up": MoveTopic -1
            Case "down": MoveTopic 1
            Case "rem": RemoveTopic
        End Select
    Else
        TemplateActions True
    End If
    Exit Sub
Err_Handler:
    HandleError "pgAdmTemplates.GridRightClick", grid
End Sub

' The search box's handler (SearchBox): filter the list as the user types.
Public Sub SearchFor(ByVal text As String)
    On Error GoTo Err_Handler
    If Not SaveDetail(Me, "sfDetail") Then Exit Sub
    m_search = Trim$(text)
    RefreshList m_id
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "pgAdmTemplates.SearchFor", text
End Sub

' The empty state's button: show every template again.
Private Sub emptyList_b_Click()
    On Error GoTo Err_Handler
    SetSearch Me, "txtSearch", ""
    Exit Sub
Err_Handler:
    HandleError "pgAdmTemplates.emptyList_b_Click"
End Sub

Private Sub btnNew_Click()
    On Error GoTo Err_Handler
    NewTemplate
    Exit Sub
Err_Handler:
    HandleError "pgAdmTemplates.btnNew_Click"
End Sub

Private Sub btnMore_Click()
    On Error GoTo Err_Handler
    TemplateActions False
    Exit Sub
Err_Handler:
    HandleError "pgAdmTemplates.btnMore_Click", m_id
End Sub

Private Sub btnAddTopic_Click()
    On Error GoTo Err_Handler
    If Not Me.sfDetail.Form.SaveTemplate(True) Then Exit Sub
    If RunDialog("dlgTemplateTopic", "tpl=" & m_id) Then TopicsChanged Nz(TempVars("TemplateTopicID"), 0)
    Exit Sub
Err_Handler:
    HandleError "pgAdmTemplates.btnAddTopic_Click", m_id
End Sub

Private Sub btnChange_Click()
    On Error GoTo Err_Handler
    ChangeTopic
    Exit Sub
Err_Handler:
    HandleError "pgAdmTemplates.btnChange_Click", m_id
End Sub

Private Sub btnUp_Click()
    On Error GoTo Err_Handler
    MoveTopic -1
    Exit Sub
Err_Handler:
    HandleError "pgAdmTemplates.btnUp_Click", m_id
End Sub

Private Sub btnDown_Click()
    On Error GoTo Err_Handler
    MoveTopic 1
    Exit Sub
Err_Handler:
    HandleError "pgAdmTemplates.btnDown_Click", m_id
End Sub

Private Sub btnRemove_Click()
    On Error GoTo Err_Handler
    RemoveTopic
    Exit Sub
Err_Handler:
    HandleError "pgAdmTemplates.btnRemove_Click", m_id
End Sub

' subTemplate saved a change (its AfterUpdate). The list catches up once
' the save that caused it is over (GridClick, TemplateSaved).
Public Sub TemplateChanged()
    m_stale = True
End Sub

' subTemplate's Save or Active switch saved template id: show it in the list.
Public Sub TemplateSaved(ByVal id As Long)
    On Error GoTo Err_Handler
    m_stale = False
    GridSelect Me.sfList, id, True
    Exit Sub
Err_Handler:
    HandleError "pgAdmTemplates.TemplateSaved", id
End Sub

' ---------------------------------------------------------------------------

' Ask for a new template's name and add it, active and empty.
Private Sub NewTemplate()
    Dim v As Variant, problem As String, id As Long
    If Not DataLinked() Then Exit Sub
    RequireExamAdmin                        ' the right is checked now, not as of sign-in
    If Not SaveDetail(Me, "sfDetail") Then Exit Sub
    v = KitInputBox("Name of the new template:", "New template")
    If IsNull(v) Then Exit Sub
    problem = NameProblem(v, 100, "tblTemplates", "TemplateName", "template")
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Sub
    KitExecute "INSERT INTO tblTemplates (TemplateName, IsActive) VALUES (" & Q(Trim$(v)) & ", True)"
    id = DLookup("TemplateID", "tblTemplates", "TemplateName=" & Q(Trim$(v)))
    If Len(m_search) > 0 Then m_search = "": SetSearch Me, "txtSearch", ""
    RefreshList id
    KitToast "Template added. Add its topics with Add topic.", "success"
End Sub

' The open template's actions: copy it, or delete it.
Private Sub TemplateActions(ByVal atPointer As Boolean)
    Dim nm As String, k As String, n As Long, v As Variant, problem As String, id As Long, a(2) As String
    If m_id = 0 Then Exit Sub
    If Not SaveDetail(Me, "sfDetail") Then Exit Sub
    nm = Nz(DLookup("TemplateName", "tblTemplates", "TemplateID=" & m_id), "")
    a(0) = "copy|Copy template|E8C8"
    a(1) = "-"
    a(2) = "del|!Delete template|E74D"
    If atPointer Then k = KitMenuAt(a(0), a(1), a(2)) Else k = KitMenu(Me.btnMore, a(0), a(1), a(2))
    If Len(k) > 0 Then RequireExamAdmin    ' the right is checked now, not as of sign-in
    Select Case k
        Case "copy"
            v = KitInputBox("Name of the copy:", "Copy template", Left$("Copy of " & nm, 100))
            If IsNull(v) Then Exit Sub
            problem = NameProblem(v, 100, "tblTemplates", "TemplateName", "template")
            If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Sub
            KitExecute "INSERT INTO tblTemplates (TemplateName, Description, IsActive) SELECT " & Q(Trim$(v)) & _
                ", Description, True FROM tblTemplates WHERE TemplateID=" & m_id
            id = DLookup("TemplateID", "tblTemplates", "TemplateName=" & Q(Trim$(v)))
            KitExecute "INSERT INTO tblTemplateTopics (TemplateID, TopicID, QuestionCount, SortOrder) SELECT " & id & _
                ", TopicID, QuestionCount, SortOrder FROM tblTemplateTopics WHERE TemplateID=" & m_id
            If Len(m_search) > 0 Then m_search = "": SetSearch Me, "txtSearch", ""
            RefreshList id
            KitToast "Copied with its topics.", "success"
        Case "del"
            n = DCount("*", "tblExams", "TemplateID=" & m_id)
            If n > 0 Then
                KitMsgBox """" & nm & """ can't be deleted: " & n & IIf(n = 1, " exam was", " exams were") & _
                    " made from it. Turn it off instead, so no new exam can use it.", vbInformation, MSG_TITLE
                Exit Sub
            End If
            If Not KitConfirmTyped("Delete the template """ & nm & """?", "No exam was made from it. Its topic list goes " & _
                "with it; the topics themselves stay. This can't be undone.", PURGE_WORD, "Delete", MSG_TITLE) Then Exit Sub
            id = m_id
            ClearTemplate                   ' the panel lets go of the record first
            KitExecute "DELETE FROM tblTemplates WHERE TemplateID=" & id
            RefreshList 0
            KitToast "Template deleted.", "destructive"
    End Select
End Sub

' Change the selected topic's question count (dlgTemplateTopic).
Private Sub ChangeTopic()
    Dim tid As Long
    tid = Nz(Me.sfTopics.Form.txtSel, 0)
    If m_id = 0 Or tid = 0 Then Exit Sub
    If Not Me.sfDetail.Form.SaveTemplate(True) Then Exit Sub
    If RunDialog("dlgTemplateTopic", "tpl=" & m_id & ";topic=" & tid) Then TopicsChanged tid
End Sub

' Move the selected topic up (-1) or down (1) in exam order.
Private Sub MoveTopic(ByVal delta As Long)
    Dim tid As Long
    tid = Nz(Me.sfTopics.Form.txtSel, 0)
    If m_id = 0 Or tid = 0 Then Exit Sub
    RequireExamAdmin                        ' the right is checked now, not as of sign-in
    If MoveRow("tblTemplateTopics", "TopicID", "SortOrder", tid, delta, "TemplateID=" & m_id) Then TopicsChanged tid
End Sub

' Take the selected topic off the open template.
Private Sub RemoveTopic()
    Dim tid As Long, nm As String
    tid = Nz(Me.sfTopics.Form.txtSel, 0)
    If m_id = 0 Or tid = 0 Then Exit Sub
    RequireExamAdmin                        ' the right is checked now, not as of sign-in
    nm = Nz(DLookup("TopicName", "tblTopics", "TopicID=" & tid), "")
    If KitMsgBox("Take """ & nm & """ off this template? Exams already made keep it; new ones won't have it.", _
        vbYesNo + vbQuestion + vbDefaultButton2, MSG_TITLE) <> vbYes Then Exit Sub
    KitExecute "DELETE FROM tblTemplateTopics WHERE TemplateID=" & m_id & " AND TopicID=" & tid
    TopicsChanged 0
End Sub

' The topic list changed: its table (tid selected, 0 for none), the total
' and the list's row.
Private Sub TopicsChanged(ByVal tid As Long)
    GridSelect Me.sfTopics, tid, True
    ShowTotal
    ShowTopicState
    GridSelect Me.sfList, m_id, True
End Sub

' "6 topics, 24 questions an exam" above the topics.
Private Sub ShowTotal()
    Dim n As Long, q As Long
    n = DCount("*", "tblTemplateTopics", "TemplateID=" & m_id)
    q = Nz(DSum("QuestionCount", "tblTemplateTopics", "TemplateID=" & m_id), 0)
    If n = 0 Then
        Me.lblTotal.Caption = "None yet: Add topic"
    Else
        Me.lblTotal.Caption = n & " " & Plural(n, "topic") & ", " & q & " " & Plural(q, "question") & " an exam"
    End If
End Sub

Private Sub ShowTopicState()
    Dim sel As Boolean, g As Form, first As Boolean, last As Boolean
    sel = m_id <> 0 And Not IsNull(Me.sfTopics.Form.txtSel)
    If sel Then
        Set g = Me.sfTopics.Form
        With g.RecordsetClone
            .MoveFirst: first = (!id = g.txtSel)
            .MoveLast: last = (!id = g.txtSel)
        End With
    End If
    SetEnabled Me.btnChange, sel
    SetEnabled Me.btnRemove, sel
    SetEnabled Me.btnUp, sel And Not first
    SetEnabled Me.btnDown, sel And Not last
End Sub

' Load the list for the search and open openId (0, or not listed: the first
' template; none listed: nothing).
Private Sub RefreshList(ByVal openId As Long)
    Dim g As Form, n As Long
    m_stale = False
    Set g = Me.sfList.Form
    g.RecordSource = ListSql()
    GridHighlight Me.sfList, m_search
    n = g.RecordsetClone.RecordCount
    If Len(m_search) > 0 Then
        ShowEmpty Me, "emptyList", n = 0, "No templates match", "No template's name or description has those words."
    Else
        ShowEmpty Me, "emptyList", n = 0, "No templates yet", "Add the first one with New template."
        SetVisible Me.emptyList_b, False
    End If
    SetVisible Me.sfList, n > 0
    If GridSelect(Me.sfList, openId) Then
        ShowTemplate openId
    ElseIf IsNull(SelectFirstRow(Me.sfList)) Then
        ClearTemplate
    End If
End Sub

' Open template id on the right with its topics. False when the open
' template can't be saved first (it stays).
Private Function ShowTemplate(ByVal id As Long) As Boolean
    If Not ShowDetail(Me, "sfDetail", "subTemplate", "TemplateID=" & id) Then Exit Function
    m_id = id
    TopicParts True
    Me.sfTopics.Form.RecordSource = TopicsSql(id)
    Me.sfTopics.Form.txtSel = Null
    ShowTotal
    ShowTopicState
    ShowTemplate = True
End Function

' Nothing open: the hint in place of the template and its topics.
Private Sub ClearTemplate()
    ClearDetail Me, "sfDetail"
    m_id = 0
    TopicParts False
End Sub

' Show or hide the topics' part of the detail card.
Private Sub TopicParts(ByVal shown As Boolean)
    Dim nm As Variant
    For Each nm In Array("lblTopics", "lblTotal", "sfTopics", "btnAddTopic", "btnChange", "btnUp", "btnDown", "btnRemove")
        SetVisible Me.Controls(nm), shown
    Next
    For Each nm In Array("hdrTopics0", "hdrTopics1")
        SetVisible Me.Controls(nm), shown
    Next
    SetEnabled Me.btnMore, shown
End Sub

' The templates for the search, by name, with their topics and questions.
Private Function ListSql() As String
    Dim where As String, t As String
    If Len(m_search) > 0 Then
        t = Q("*" & LikeLiteral(m_search) & "*")
        where = " WHERE t.TemplateName LIKE " & t & " OR t.Description LIKE " & t
    End If
    ListSql = "SELECT t.TemplateID AS id, t.TemplateName, " & _
        "(SELECT Count(*) FROM tblTemplateTopics AS x WHERE x.TemplateID=t.TemplateID) AS topics, " & _
        "Nz((SELECT Sum(x.QuestionCount) FROM tblTemplateTopics AS x WHERE x.TemplateID=t.TemplateID),0) AS questions, " & _
        "IIf(Nz(t.IsActive,False),'Active','Off') AS state FROM tblTemplates AS t" & where & " ORDER BY t.TemplateName"
End Function

' The open template's topics in exam order; a topic turned off says so.
Private Function TopicsSql(ByVal id As Long) As String
    TopicsSql = "SELECT tt.TopicID AS id, t.TopicName & IIf(Nz(t.IsActive,False),'',' (off)') AS TopicName, tt.QuestionCount " & _
        "FROM tblTemplateTopics AS tt INNER JOIN tblTopics AS t ON tt.TopicID=t.TopicID WHERE tt.TemplateID=" & id & _
        " ORDER BY tt.SortOrder, t.TopicName"
End Function
