Option Compare Database
Option Explicit

' Admin: Topics (see Build_pgAdmTopics). The topics on the left with their
' SMEs, approved questions and status: Active, Short (fewer approved
' questions than the overlap rule needs for the most any active template
' gives it, modAdmin.QuestionsNeeded) or Off. On the right the open topic
' in subTopic (bound to tblTopics), its SMEs, and how its approved questions
' stand against the rule.
'
' New topic asks for the name and adds the topic, active, with the
' default question count from Rules and lists. Adding an SME gives the
' person the SME role (and Exam Generator access) if they lack it; removing
' one leaves their role alone. More, or a right-click on a row, turns the
' topic off or on, or deletes it while nothing uses it.
'
' Owner: the admin pages (build phase 4). Depends on: modAdmin, modErrors,
' modGlobals, modSignIn, modUI; the kit's modKitKeys, modKitLocks,
' modKitPurge, modKitUI, modKitUtil, modMenu, modPersonnel.

Private Const MSG_TITLE As String = "Topics"

Private m_id As Long            ' the open topic (0 when none)
Private m_search As String
Private m_isStale As Boolean    ' subTopic saved a change the list doesn't show yet

Private Sub Form_Load()
    On Error GoTo Err_Handler
    If Not DataLinked() Then
        ShowEmpty Me, "emptyList", True, "No data file linked", "This copy was built without the data file, so there are no topics to show."
        SetVisible Me.emptyList_b, False
        SetVisible Me.sfList, False
        SetEnabled Me.btnNew, False
        ClearTopic
        Exit Sub
    End If
    RefreshList 0
    Exit Sub
Err_Handler:
    HandleError "pgAdmTopics.Form_Load"
End Sub

' frmMain asks before it swaps or closes this page: save the open topic, or
' stay while it can't be saved (subTopic says why).
Public Function CanLeave() As Boolean
    On Error GoTo Err_Handler
    CanLeave = SaveDetail(Me, "sfDetail")
    Exit Function
Err_Handler:
    HandleError "pgAdmTopics.CanLeave"
End Function

' Ctrl+S saves the open topic, Ctrl+N adds one (modKitKeys).
Public Function Shortcut(ByVal action As String) As Boolean
    On Error GoTo Err_Handler
    Select Case action
        Case SHORTCUT_SAVE
            If m_id <> 0 Then Me.sfDetail.Form.SaveTopic
            Shortcut = True
        Case SHORTCUT_NEW
            NewTopic
            Shortcut = True
    End Select
    Exit Function
Err_Handler:
    HandleError "pgAdmTopics.Shortcut", action
    Shortcut = True
End Function

' Row click: in the list, open that topic once the open one is saved (when
' it can't be, it stays open and so does its row); in the SMEs, select one.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    Dim id As Long
    On Error GoTo Err_Handler
    If grid = "sfSMEs" Then ShowSMEState: Exit Sub
    id = Nz(Me.sfList.Form.txtSel, 0)
    If id = 0 Then Exit Sub
    If Not ShowTopic(id) Then
        GridSelect Me.sfList, m_id
        Exit Sub
    End If
    If m_isStale Then RefreshList id          ' the topic left behind was saved with a change
    Exit Sub
Err_Handler:
    HandleError "pgAdmTopics.GridClick", id
End Sub

' Right-click on a row: the topic's actions, or an SME's.
Public Sub GridRightClick(ByVal grid As String)
    On Error GoTo Err_Handler
    If grid = "sfSMEs" Then
        If KitMenuAt("rem|!Remove from this topic|E74D") = "rem" Then RemoveSME
    Else
        TopicActions True
    End If
    Exit Sub
Err_Handler:
    HandleError "pgAdmTopics.GridRightClick", grid
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
    HandleError "pgAdmTopics.SearchFor", text
End Sub

' The empty state's button: show every topic again.
Private Sub emptyList_b_Click()
    On Error GoTo Err_Handler
    SetSearch Me, "txtSearch", ""
    Exit Sub
Err_Handler:
    HandleError "pgAdmTopics.emptyList_b_Click"
End Sub

Private Sub btnNew_Click()
    On Error GoTo Err_Handler
    NewTopic
    Exit Sub
Err_Handler:
    HandleError "pgAdmTopics.btnNew_Click"
End Sub

Private Sub btnMore_Click()
    On Error GoTo Err_Handler
    TopicActions False
    Exit Sub
Err_Handler:
    HandleError "pgAdmTopics.btnMore_Click", m_id
End Sub

Private Sub txtSME_KeyDown(KeyCode As Integer, Shift As Integer)
    KitPickKey Me.txtSME, KeyCode
End Sub

Private Sub btnAddSME_Click()
    On Error GoTo Err_Handler
    AddSME
    Exit Sub
Err_Handler:
    HandleError "pgAdmTopics.btnAddSME_Click", m_id
End Sub

Private Sub btnRemoveSME_Click()
    On Error GoTo Err_Handler
    RemoveSME
    Exit Sub
Err_Handler:
    HandleError "pgAdmTopics.btnRemoveSME_Click", m_id
End Sub

' subTopic saved a change (its AfterUpdate). The list catches up once the
' save that caused it is over (GridClick, TopicSaved).
Public Sub TopicChanged()
    m_isStale = True
End Sub

' subTopic's Save or Active switch saved topic id: show it in the list.
Public Sub TopicSaved(ByVal id As Long)
    On Error GoTo Err_Handler
    m_isStale = False
    GridSelect Me.sfList, id, True
    ShowSupply
    Exit Sub
Err_Handler:
    HandleError "pgAdmTopics.TopicSaved", id
End Sub

' ---------------------------------------------------------------------------

' Ask for a new topic's name and add it, active, with the default count.
Private Sub NewTopic()
    Dim v As Variant, problem As String, id As Long
    If Not DataLinked() Then Exit Sub
    RequireExamAdmin                        ' the right is checked now, not as of sign-in
    If Not SaveDetail(Me, "sfDetail") Then Exit Sub
    v = KitInputBox("Name of the new topic:", "New topic")
    If IsNull(v) Then Exit Sub
    problem = NameProblem(v, MAX_TOPIC_NAME, "tblTopics", "TopicName", "topic")
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Sub
    id = InsertRow("INSERT INTO tblTopics (TopicName, DefaultCount, IsActive) VALUES (" & Q(Trim$(v)) & ", " & _
        DefaultQuestionCount() & ", True)")
    If Len(m_search) > 0 Then m_search = "": SetSearch Me, "txtSearch", ""
    RefreshList id
    KitToast "Topic added. Add its SMEs on the right.", "success"
End Sub

' The open topic's actions: turn it off or on, or delete it.
Private Sub TopicActions(ByVal atPointer As Boolean)
    Dim nm As String, isOn As Boolean, k As String, used As String, a(2) As String, id As Long
    If m_id = 0 Then Exit Sub
    If Not SaveDetail(Me, "sfDetail") Then Exit Sub
    nm = Nz(DLookup("TopicName", "tblTopics", "TopicID=" & m_id), "")
    isOn = Nz(DLookup("IsActive", "tblTopics", "TopicID=" & m_id), False)
    a(0) = "onoff|" & IIf(isOn, "Turn off", "Turn on") & "|E7E8"
    a(1) = "-"
    a(2) = "del|!Delete topic|E74D"
    If atPointer Then k = KitMenuAt(a(0), a(1), a(2)) Else k = KitMenu(Me.btnMore, a(0), a(1), a(2))
    If Len(k) > 0 Then RequireExamAdmin    ' the right is checked now, not as of sign-in
    Select Case k
        Case "onoff"
            KitExecute "UPDATE tblTopics SET IsActive=" & IIf(isOn, "False", "True") & " WHERE TopicID=" & m_id
            Me.sfDetail.Form.Requery
            RefreshList m_id
            KitToast IIf(isOn, "Turned off. Templates can't add it while it's off.", "Turned on."), "success"
        Case "del"
            used = TopicUse(m_id)
            If Len(used) > 0 Then
                KitMsgBox """" & nm & """ can't be deleted: " & used & ". Turn it off instead.", vbInformation, MSG_TITLE
                Exit Sub
            End If
            If Not KitConfirmTyped("Delete the topic """ & nm & """?", "No question, template or exam uses it. " & _
                "Its SMEs are taken off it. This can't be undone.", PURGE_WORD, "Delete", MSG_TITLE) Then Exit Sub
            id = m_id
            ClearTopic                      ' the panel lets go of the record first
            KitExecute "DELETE FROM tblTopics WHERE TopicID=" & id
            RefreshList 0
            KitToast "Topic deleted.", "destructive"
    End Select
End Sub

' What uses topic id ("3 questions and 1 template"), or "".
Private Function TopicUse(ByVal id As Long) As String
    Dim n(2) As Long, words As Variant, i As Long, s As String
    n(0) = DCount("*", "tblQuestions", "TopicID=" & id)
    n(1) = DCount("*", "tblTemplateTopics", "TopicID=" & id)
    n(2) = DCount("*", "tblExamTopics", "TopicID=" & id)
    words = Array("question", "template", "exam")
    For i = 0 To 2
        If n(i) > 0 Then s = s & IIf(Len(s) > 0, " and ", "") & n(i) & " " & Plural(n(i), CStr(words(i)))
    Next
    If Len(s) > 0 Then s = "it is on " & s
    TopicUse = s
End Function

' Make the picked person an SME of the open topic.
Private Sub AddSME()
    Dim pid As Long
    If m_id = 0 Then Exit Sub
    RequireExamAdmin                        ' the right is checked now, not as of sign-in
    pid = Nz(Me.txtSMEKey, 0)
    If pid = 0 Then
        KitMsgBox "Type part of a name in the box, then pick the person from the list under it.", vbInformation, MSG_TITLE
        Me.txtSME.SetFocus
        Exit Sub
    End If
    If DCount("*", "tblTopicSMEs", "TopicID=" & m_id & " AND PersonID=" & pid) > 0 Then
        KitToast "They're already an SME of this topic.", "info"
    Else
        SetPersonRole pid, "SME", True
        KitExecute "INSERT INTO tblTopicSMEs (TopicID, PersonID) VALUES (" & m_id & ", " & pid & ")"
        KitToast "SME added.", "success"
    End If
    Me.txtSMEKey = Null
    KitPickShow Me.txtSME
    SMEsChanged pid
End Sub

' Take the selected SME off the open topic. Their SME role stays.
Private Sub RemoveSME()
    Dim pid As Long
    pid = Nz(Me.sfSMEs.Form.txtSel, 0)
    If m_id = 0 Or pid = 0 Then Exit Sub
    RequireExamAdmin                        ' the right is checked now, not as of sign-in
    KitExecute "DELETE FROM tblTopicSMEs WHERE TopicID=" & m_id & " AND PersonID=" & pid
    SMEsChanged 0
    KitToast "Taken off this topic. They keep the SME role for any other topic.", "success"
End Sub

' The SMEs changed: their table (pid selected, 0 for none) and the list's row.
Private Sub SMEsChanged(ByVal pid As Long)
    GridSelect Me.sfSMEs, pid, True
    ShowSMEState
    GridSelect Me.sfList, m_id, True
End Sub

Private Sub ShowSMEState()
    SetEnabled Me.btnRemoveSME, m_id <> 0 And Not IsNull(Me.sfSMEs.Form.txtSel)
End Sub

' Load the list for the search and open openId (0, or not listed: the first
' topic; none listed: nothing).
Private Sub RefreshList(ByVal openId As Long)
    Dim g As Form, n As Long
    m_isStale = False
    Set g = Me.sfList.Form
    g.RecordSource = ListSql()
    GridHighlight Me.sfList, m_search
    n = g.RecordsetClone.RecordCount
    If Len(m_search) > 0 Then
        ShowEmpty Me, "emptyList", n = 0, "No topics match", "No topic's name or description has those words."
    Else
        ShowEmpty Me, "emptyList", n = 0, "No topics yet", "Add the first one with New topic."
        SetVisible Me.emptyList_b, False
    End If
    SetVisible Me.sfList, n > 0
    If GridSelect(Me.sfList, openId) Then
        ShowTopic openId
    ElseIf IsNull(SelectFirstRow(Me.sfList)) Then
        ClearTopic
    End If
End Sub

' Open topic id on the right with its SMEs and supply. False when the open
' topic can't be saved first (it stays).
Private Function ShowTopic(ByVal id As Long) As Boolean
    Dim nm As Variant
    If Not ShowDetail(Me, "sfDetail", "subTopic", "TopicID=" & id) Then Exit Function
    m_id = id
    For Each nm In Array("lblSMEs", "txtSME", "btnAddSME", "btnRemoveSME", "sfSMEs")
        SetVisible Me.Controls(nm), True
    Next
    SetEnabled Me.btnMore, True
    Me.sfSMEs.Form.RecordSource = SMESql(id)
    Me.sfSMEs.Form.txtSel = Null
    ShowSMEState
    ShowSupply
    ShowTopic = True
End Function

' Nothing open: the hint in place of the topic and its SMEs.
Private Sub ClearTopic()
    Dim nm As Variant
    ClearDetail Me, "sfDetail"
    m_id = 0
    For Each nm In Array("lblSMEs", "txtSME", "btnAddSME", "btnRemoveSME", "sfSMEs")
        SetVisible Me.Controls(nm), False
    Next
    SetEnabled Me.btnMore, False
    HideAlert Me, "alSupply"
End Sub

' The open topic's approved questions against what the overlap rule needs
' for the most questions any active template gives it.
Private Sub ShowSupply()
    Dim rs As DAO.Recordset, need As Long, have As Long, per As Long
    If m_id = 0 Then HideAlert Me, "alSupply": Exit Sub
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM (" & TopicSupplySql() & ") AS s WHERE s.TopicID=" & m_id, dbOpenSnapshot)
    If rs.EOF Then rs.Close: HideAlert Me, "alSupply": Exit Sub
    have = rs!on_hand: per = rs!per_exam
    rs.Close
    need = QuestionsNeeded(per, OverlapPct(), OverlapWindow())
    If per = 0 Then
        ShowAlert Me, "alSupply", have & " approved " & Plural(have, "question") & ". No active template gives this topic " & _
            "questions yet.", "info"
    ElseIf have >= need Then
        ShowAlert Me, "alSupply", have & " approved " & Plural(have, "question") & "; the overlap rule needs " & need & _
            " when an exam takes " & per & ".", "success"
    Else
        ShowAlert Me, "alSupply", "Short: " & have & " approved " & Plural(have, "question") & "; the overlap rule needs " & _
            need & " when an exam takes " & per & ".", "warning"
    End If
End Sub

' The topics for the search, by name, with their status under the rule.
Private Function ListSql() As String
    Dim where As String, t As String
    If Len(m_search) > 0 Then
        t = Q("*" & LikeLiteral(m_search) & "*")
        where = " WHERE s.TopicID IN (SELECT TopicID FROM tblTopics WHERE TopicName LIKE " & t & " OR Description LIKE " & t & ")"
    End If
    ListSql = "SELECT s.TopicID AS id, s.TopicName, TopicSMEsText(s.TopicID) AS smes, s.on_hand, " & _
        "IIf(Not Nz(s.IsActive,False),'Off',IIf(s.on_hand<QuestionsNeeded(s.per_exam," & OverlapPct() & "," & _
        OverlapWindow() & "),'Short','Active')) AS state FROM (" & TopicSupplySql() & ") AS s" & where & _
        " ORDER BY s.TopicName"
End Function

' The open topic's SMEs, by name.
Private Function SMESql(ByVal id As Long) As String
    SMESql = "SELECT ts.PersonID AS id, " & PersonNameSql() & " AS person, p.username FROM tblTopicSMEs AS ts " & _
        "INNER JOIN " & PERSON_TABLE & " AS p ON ts.PersonID=p.person_id WHERE ts.TopicID=" & id & _
        " ORDER BY p.last_name, p.first_name"
End Function
