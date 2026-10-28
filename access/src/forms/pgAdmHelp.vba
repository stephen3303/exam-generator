Option Compare Database
Option Explicit

' Admin: Help topics (see Build_pgAdmHelp). The topics on the left; the open
' topic's key, title and text on the right. Save writes the topic with
' KitHelpSave: to the data file's help_topic table, which everyone's copy
' reads at start-up (KitHelpSync), and to this copy's zz_help, so F1 shows
' it at once. A front end with no data file linked edits zz_help alone.
'
' The fields are unbound, so the page keeps track of an unsaved edit itself
' (m_isDirty) and asks before it is lost: another topic, New topic, a search
' that hides the topic, or leaving the page (CanLeave). A topic's key is set
' when it is made and can't change after, because the pages and dialogs find
' their help by it.

Private Const MSG_TITLE As String = "Help topics"

Private m_table As String       ' help_topic, or zz_help with no data file linked
Private m_key As String         ' the open topic's key ("" when none is open)
Private m_lastKey As String     ' the topic open before New topic, for Undo changes
Private m_isNew As Boolean      ' the open topic is new and not saved yet
Private m_isDirty As Boolean    ' the fields have changes not saved yet
Private m_search As String

Private Sub Form_Load()
    On Error GoTo Err_Handler
    m_table = TopicTable()
    RefreshList
    ShowState
    Exit Sub
Err_Handler:
    HandleError "pgAdmHelp.Form_Load"
End Sub

' frmMain asks before it swaps or closes this page.
Public Function CanLeave() As Boolean
    On Error GoTo Err_Handler
    CanLeave = LeaveTopic()
    Exit Function
Err_Handler:
    HandleError "pgAdmHelp.CanLeave"
End Function

' Ctrl+S saves the open topic, Ctrl+N starts a new one (modKitKeys); Ctrl+F
' is left to the kit, which goes to the search box.
Public Function Shortcut(ByVal action As String) As Boolean
    On Error GoTo Err_Handler
    Select Case action
        Case SHORTCUT_SAVE
            If m_isDirty Then SaveTopic
            Shortcut = True
        Case SHORTCUT_NEW
            NewTopic
            Shortcut = True
    End Select
    Exit Function
Err_Handler:
    HandleError "pgAdmHelp.Shortcut", action
    Shortcut = True
End Function

' Row click in the list (subGrid): open that topic, once the open one's
' changes are saved or dropped. Cancel keeps the open topic and its row.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    Dim key As String
    On Error GoTo Err_Handler
    key = Nz(Me.sfList.Form.txtSel, "")
    If Len(key) = 0 Or (key = m_key And Not m_isNew) Then Exit Sub
    If Not LeaveTopic() Then
        GridSelect Me.sfList, m_key
        Exit Sub
    End If
    ShowTopic key
    Exit Sub
Err_Handler:
    HandleError "pgAdmHelp.GridClick", key
End Sub

' The search box's handler (SearchBox): filter the list as the user types.
' An open topic with unsaved changes stays open even when the search hides
' its row.
Public Sub SearchFor(ByVal text As String)
    On Error GoTo Err_Handler
    m_search = Trim$(text)
    RefreshList
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "pgAdmHelp.SearchFor", text
End Sub

' The empty state's button: show every topic again.
Private Sub emptyList_b_Click()
    On Error GoTo Err_Handler
    SetSearch Me, "txtSearch", ""
    Exit Sub
Err_Handler:
    HandleError "pgAdmHelp.emptyList_b_Click"
End Sub

Private Sub btnNew_Click()
    On Error GoTo Err_Handler
    NewTopic
    Exit Sub
Err_Handler:
    HandleError "pgAdmHelp.btnNew_Click"
End Sub

Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    SaveTopic
    Exit Sub
Err_Handler:
    HandleError "pgAdmHelp.btnSave_Click", m_key
End Sub

' Preview: the help window on the topic as saved. Unsaved changes are saved
' first when the admin says so (a new topic has to be).
Private Sub btnPreview_Click()
    On Error GoTo Err_Handler
    If m_isNew Then
        If KitMsgBox("Preview shows saved topics. Save this new topic first?", vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then Exit Sub
        If Not SaveTopic() Then Exit Sub
    ElseIf m_isDirty Then
        Select Case KitMsgBox("Preview shows the topic as saved. Save your changes first?", vbYesNoCancel + vbQuestion, MSG_TITLE)
            Case vbYes: If Not SaveTopic() Then Exit Sub
            Case vbCancel: Exit Sub
        End Select
    End If
    If Len(m_key) > 0 Then KitHelp m_key
    Exit Sub
Err_Handler:
    HandleError "pgAdmHelp.btnPreview_Click", m_key
End Sub

' Undo changes: the open topic as saved, or, for a new one, the topic open
' before it.
Private Sub btnUndo_Click()
    Dim key As String
    On Error GoTo Err_Handler
    If m_isNew Then key = m_lastKey Else key = m_key
    ShowTopic key
    Exit Sub
Err_Handler:
    HandleError "pgAdmHelp.btnUndo_Click", m_key
End Sub

Private Sub txtKey_Change()
    On Error GoTo Err_Handler
    Edited
    Exit Sub
Err_Handler:
    HandleError "pgAdmHelp.txtKey_Change"
End Sub

Private Sub txtTitle_Change()
    On Error GoTo Err_Handler
    Edited
    Exit Sub
Err_Handler:
    HandleError "pgAdmHelp.txtTitle_Change"
End Sub

Private Sub txtBody_Change()
    On Error GoTo Err_Handler
    Edited
    Exit Sub
Err_Handler:
    HandleError "pgAdmHelp.txtBody_Change"
End Sub

' ---------------------------------------------------------------------------

Private Sub Edited()
    If m_isDirty Then Exit Sub
    m_isDirty = True
    ShowState
End Sub

' Start a new topic, once the open one's changes are saved or dropped.
Private Sub NewTopic()
    If Not LeaveTopic() Then Exit Sub
    m_lastKey = m_key
    ShowTopic ""
    m_isNew = True
    Me.lblStamp.Caption = "New topic, not saved yet"
    ShowState
    Me.txtKey.SetFocus
End Sub

' Save or drop the open topic's changes, asking first. False when the admin
' cancels, or the topic can't be saved (the admin has been told why).
Private Function LeaveTopic() As Boolean
    Dim nm As String
    If Not m_isDirty Then LeaveTopic = True: Exit Function
    nm = Trim$(FieldText(Me, Me.txtTitle))
    If Len(nm) = 0 Then nm = IIf(m_isNew, "the new topic", m_key) Else nm = """" & nm & """"
    Select Case KitMsgBox("Save your changes to " & nm & "?", vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbYes: LeaveTopic = SaveTopic()
        Case vbNo: m_isDirty = False: LeaveTopic = True
    End Select
End Function

' Save the open topic (KitHelpSave). False when it can't be saved yet; the
' admin is told what to fix.
Private Function SaveTopic() As Boolean
    Dim key As String, title As String, body As String, problem As String
    RequireExamAdmin                        ' the right is checked now, not as of sign-in
    key = LCase$(Trim$(FieldText(Me, Me.txtKey)))
    title = Trim$(FieldText(Me, Me.txtTitle))
    body = FieldText(Me, Me.txtBody)
    problem = TopicProblem(key, title)
    If Len(problem) > 0 Then
        KitMsgBox problem, vbExclamation, MSG_TITLE
        Exit Function
    End If
    KitHelpSave key, title, body
    m_key = key: m_isNew = False: m_isDirty = False
    RefreshList
    ShowTopic key
    If m_table = HELP_SHARED_TABLE Then
        KitToast "Saved. Everyone else sees it the next time they open " & APP_NAME & ".", "success"
    Else
        KitToast "Saved in this copy only: no data file is linked.", "warning"
    End If
    SaveTopic = True
End Function

' What stops key and title being saved, or "".
Private Function TopicProblem(ByVal key As String, ByVal title As String) As String
    Dim dot As Long
    If Len(key) = 0 Then
        TopicProblem = "Give the topic a key. A page's topic uses the page's key (exams); " & _
            "a topic under another adds a dot and a word (exams.new)."
    ElseIf Len(key) > 32 Then
        TopicProblem = "A topic's key can be at most 32 characters."
    ElseIf key Like "*[!a-z0-9.]*" Or Left$(key, 1) = "." Or Right$(key, 1) = "." Or InStr(key, "..") > 0 Then
        TopicProblem = "A topic's key can only have letters, digits and dots between words, such as exams.new."
    ElseIf Len(title) = 0 Then
        TopicProblem = "Give the topic a title. The help window lists topics by their titles."
    ElseIf Len(title) > 80 Then
        TopicProblem = "A topic's title can be at most 80 characters."
    ElseIf m_isNew Then
        If TopicExists(key) Then
            TopicProblem = "There is already a topic with the key " & key & ". Open it from the list to change it."
        Else
            dot = InStr(key, ".")
            If dot > 0 Then
                If Not TopicExists(Left$(key, dot - 1)) Then
                    TopicProblem = "There is no topic " & Left$(key, dot - 1) & " to put this one under. " & _
                        "Use the key of a topic in the list, a dot and a word."
                End If
            End If
        End If
    End If
End Function

Private Function TopicExists(ByVal key As String) As Boolean
    TopicExists = DCount("*", m_table, "topic_key=" & Q(key)) > 0
End Function

' Show topic key in the fields ("" clears them), with no unsaved changes,
' and mark its row.
Private Sub ShowTopic(ByVal key As String)
    Dim rs As DAO.Recordset
    m_key = "": m_isNew = False: m_isDirty = False
    Me.txtKey = Null: Me.txtTitle = Null: Me.txtBody = Null
    Me.lblStamp.Caption = ""
    If Len(key) > 0 Then
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & m_table & " WHERE topic_key=" & Q(key), dbOpenSnapshot)
        If Not rs.EOF Then
            m_key = rs!topic_key
            Me.txtKey = rs!topic_key
            Me.txtTitle = rs!title
            ' Stored with line feeds (the help file's lines); a text box breaks lines on CR LF.
            Me.txtBody = Replace(Nz(rs!body, ""), vbLf, vbCrLf)
            If m_table = HELP_SHARED_TABLE Then Me.lblStamp.Caption = StampText(rs!updated_by, rs!updated_at)
        End If
        rs.Close
    End If
    GridSelect Me.sfList, m_key
    ShowState
End Sub

' Who saved the topic last, and when.
Private Function StampText(ByVal who As Variant, ByVal savedAt As Variant) As String
    If Nz(who, "") = "help.txt" Then
        StampText = "As the build's help.txt has it; not changed here yet"
    ElseIf Not IsNull(savedAt) Then
        StampText = "Last saved " & Format$(savedAt, "d mmm yyyy h:nn") & IIf(Len(Nz(who, "")) > 0, " by " & who, "")
    End If
End Function

' Lock what can't be typed in and enable the buttons that apply. The key is
' typed only for a new topic.
Private Sub ShowState()
    Dim isOpen As Boolean
    isOpen = m_isNew Or Len(m_key) > 0
    Me.txtKey.Locked = Not m_isNew
    Me.txtTitle.Locked = Not isOpen
    Me.txtBody.Locked = Not isOpen
    SetEnabled Me.btnSave, m_isDirty
    SetEnabled Me.btnUndo, m_isDirty Or m_isNew
    SetEnabled Me.btnPreview, isOpen
End Sub

' Load the list for the search, mark the matches, and keep the open topic
' open; with none open (or it no longer listed and unchanged), open the
' first topic.
Private Sub RefreshList()
    Dim g As Form, n As Long
    Set g = Me.sfList.Form
    g.RecordSource = ListSql()
    GridHighlight Me.sfList, m_search
    n = g.RecordsetClone.RecordCount
    ShowEmpty Me, "emptyList", n = 0
    ' The empty state shows in place of the list: a subform always covers a label.
    SetVisible Me.sfList, n > 0
    If GridSelect(Me.sfList, m_key) Or m_isDirty Or m_isNew Then Exit Sub
    If IsNull(SelectFirstRow(Me.sfList)) Then ShowTopic ""
End Sub

' The topics for the search, in help order; a topic under another is
' indented, as in the help window.
Private Function ListSql() As String
    Dim where As String, t As String
    If Len(m_search) > 0 Then
        t = Q("*" & LikeLiteral(m_search) & "*")
        where = " WHERE title LIKE " & t & " OR topic_key LIKE " & t & " OR search_text LIKE " & t
    End If
    ListSql = "SELECT topic_key AS id, IIf(InStr(topic_key,'.')>0,'    ','') & title AS topic, topic_key, sort_order " & _
        "FROM " & m_table & where & " ORDER BY sort_order, topic_key"
End Function

' help_topic when this copy is linked to it, else zz_help.
Private Function TopicTable() As String
    If TableExists(HELP_SHARED_TABLE) Then TopicTable = HELP_SHARED_TABLE Else TopicTable = HELP_TABLE
End Function
