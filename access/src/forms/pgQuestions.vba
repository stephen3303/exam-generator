Option Compare Database
Option Explicit

' Question bank (see Build_pgQuestions). Two tabs, whose controls the hidden
' labels grpQ and grpR list.
'
' Questions: every question, newest first, filtered by the search box, the
' status chips (All shows the current version of each question; Retired the
' versions an approved one replaced), the column funnels and, from the
' References tab, the questions citing one reference. The picked question
' shows on the right; SMEs open it in the editor (Edit, or a double-click),
' start a new version of an approved one, or delete a draft never submitted.
'
' References: the shared list of references with how many current questions
' cite each. SMEs and the admin change a reference's title, number and
' revision, turn it off (it can't be cited again; citations stay) or delete
' one nothing cites. New reference adds one.

Private Const MSG_TITLE As String = "Question bank"

Private m_tab As String         ' the open tab: q or r
Private m_chip As Long          ' the status chip on (0 All ... 5 Retired)
Private m_search As String
Private m_refFilter As Long     ' only the questions citing this reference (0: all)
Private m_qid As Long           ' the question shown on the right (0: none)
Private m_rid As Long           ' the reference shown on the right (0: none)
Private m_refSnap As String     ' its fields as shown, to tell whether they changed
Private m_canRef As Boolean     ' the user may change references

Public Property Get HelpTopic() As String
    HelpTopic = "questions"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    m_tab = "q"
    SelectTab Me, "q"
    If Not DataLinked() Then
        ShowGroup
        ShowEmpty Me, "emptyList", True, "No data file linked", "This copy was built without the data file, so there are no questions to show."
        SetVisible Me.emptyList_b, False
        SetVisible Me.sfList, False
        SetEnabled Me.btnNew, False
        SetEnabled Me.btnExport, False
        ShowQuestion 0
        Exit Sub
    End If
    m_canRef = HasRole("SME") Or IsExamAdmin()
    SetEnabled Me.btnNew, HasRole("SME")
    ChipOn Me, "chip", 0
    RefreshRefs 0
    RefreshList Nz(TempVars("BankQuestionID"), 0)
    TempVars("BankQuestionID") = Null
    ShowGroup
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.Form_Load"
End Sub

' frmMain asks before it swaps or closes this page: a reference with
' unsaved changes is saved, dropped or kept.
Public Function CanLeave() As Boolean
    On Error GoTo Err_Handler
    CanLeave = RefLeaveOK()
    Exit Function
Err_Handler:
    HandleError "pgQuestions.CanLeave"
End Function

' Ctrl+S saves the reference shown, Ctrl+N is New question (or New
' reference on the References tab).
Public Function Shortcut(ByVal action As String) As Boolean
    On Error GoTo Err_Handler
    Select Case action
        Case SHORTCUT_SAVE
            If m_tab = "r" Then SaveRef
            Shortcut = True
        Case SHORTCUT_NEW
            NewClicked
            Shortcut = True
    End Select
    Exit Function
Err_Handler:
    HandleError "pgQuestions.Shortcut", action
    Shortcut = True
End Function

' ---------------------------------------------------------------------------
' Tabs, chips, search (the kit's builders call these)
' ---------------------------------------------------------------------------

Public Function TabPick(ByVal key As String)
    On Error GoTo Err_Handler
    If key = m_tab Then Exit Function
    If Not RefLeaveOK() Then Exit Function
    m_tab = key
    SelectTab Me, key
    ShowGroup
    Exit Function
Err_Handler:
    HandleError "pgQuestions.TabPick", key
End Function

Public Function ChipPick(ByVal i As Long)
    On Error GoTo Err_Handler
    m_chip = i
    ChipOn Me, "chip", i
    ShowCounts
    ApplyChip
    Exit Function
Err_Handler:
    HandleError "pgQuestions.ChipPick", i
End Function

Public Sub SearchFor(ByVal text As String)
    On Error GoTo Err_Handler
    If Not RefLeaveOK() Then Exit Sub
    m_search = Trim$(text)
    RefreshList m_qid
    RefreshRefs m_rid
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.SearchFor", text
End Sub

' A column funnel changed the questions shown (GridHeaders' onFilter).
Public Sub ListFiltered(ByVal sfName As String)
    On Error GoTo Err_Handler
    ShowCounts
    AfterFilter
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.ListFiltered", sfName
End Sub

' "Show all questions": stop showing only one reference's questions.
Public Function ClearRefFilter()
    On Error GoTo Err_Handler
    m_refFilter = 0
    RefreshList m_qid
    Exit Function
Err_Handler:
    HandleError "pgQuestions.ClearRefFilter"
End Function

' ---------------------------------------------------------------------------
' Buttons and rows
' ---------------------------------------------------------------------------

Private Sub btnNew_Click()
    On Error GoTo Err_Handler
    NewClicked
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.btnNew_Click"
End Sub

Private Sub btnExport_Click()
    On Error GoTo Err_Handler
    If m_tab = "r" Then ExportGridAs Me.sfRefs.Form, "References" Else ExportGridAs Me.sfList.Form, "Question bank"
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.btnExport_Click"
End Sub

Private Sub btnOpen_Click()
    On Error GoTo Err_Handler
    If m_qid <> 0 Then EditQuestion m_qid, False
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.btnOpen_Click", m_qid
End Sub

Private Sub btnHistory_Click()
    On Error GoTo Err_Handler
    If m_qid <> 0 Then RunDialog "dlgHistory", "key=" & Nz(DLookup("QuestionKey", "tblQuestions", "QuestionID=" & m_qid), 0)
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.btnHistory_Click", m_qid
End Sub

Private Sub btnMore_Click()
    On Error GoTo Err_Handler
    QuestionActions False
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.btnMore_Click", m_qid
End Sub

Private Sub btnRefMore_Click()
    On Error GoTo Err_Handler
    RefActions False
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.btnRefMore_Click", m_rid
End Sub

Private Sub btnRefSave_Click()
    On Error GoTo Err_Handler
    SaveRef
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.btnRefSave_Click", m_rid
End Sub

Private Sub btnShowCiting_Click()
    On Error GoTo Err_Handler
    ShowCiting m_rid, 0
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.btnShowCiting_Click", m_rid
End Sub

' The empty states' button: show every question (or reference) again.
Private Sub emptyList_b_Click()
    On Error GoTo Err_Handler
    m_chip = 0
    ChipOn Me, "chip", 0
    m_refFilter = 0
    ClearGridFilters Me.sfList
    SetSearch Me, "txtSearch", ""
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.emptyList_b_Click"
End Sub

Private Sub emptyRefs_b_Click()
    On Error GoTo Err_Handler
    SetSearch Me, "txtSearch", ""
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.emptyRefs_b_Click"
End Sub

' A row clicked: show it on the right; a double-click on a question opens
' it, on a citing question shows it in the question list.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    Dim id As Long
    On Error GoTo Err_Handler
    Select Case grid
        Case "sfList"
            id = Nz(Me.sfList.Form.txtSel, 0)
            ShowQuestion id
            If dbl And id <> 0 And HasRole("SME") Then EditQuestion id, False
        Case "sfRefs"
            id = Nz(Me.sfRefs.Form.txtSel, 0)
            If id = m_rid Then Exit Sub
            If Not RefLeaveOK() Then MarkRow Me.sfRefs, m_rid: Exit Sub
            ShowRef id
        Case "sfRefQs"
            id = Nz(Me.sfRefQs.Form.txtSel, 0)
            If dbl And id <> 0 Then ShowCiting m_rid, id
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.GridClick", grid & " " & id
End Sub

Public Sub GridRightClick(ByVal grid As String)
    On Error GoTo Err_Handler
    Select Case grid
        Case "sfList": QuestionActions True
        Case "sfRefs": RefActions True
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgQuestions.GridRightClick", grid
End Sub

' New question (Questions tab) or New reference (References tab).
Private Sub NewClicked()
    If Not DataLinked() Then Exit Sub
    If m_tab = "r" Then
        If Not m_canRef Then Exit Sub
        If Not RefLeaveOK() Then Exit Sub
        If RunDialog("dlgReference", "mode=new") Then
            RefreshRefs Nz(TempVars("RefID"), 0)
            KitToast "Reference added. SMEs can cite it now.", "success"
        End If
    ElseIf HasRole("SME") Then
        EditQuestion 0, False
    End If
End Sub

' Open question qid in the editor (0: a new question), or a new version of it.
Private Sub EditQuestion(ByVal qid As Long, ByVal asNewVersion As Boolean)
    TempVars("EditQuestionID") = qid
    TempVars("EditMode") = IIf(asNewVersion, "newversion", "")
    NavTo "qedit"
End Sub

' The shown question's actions (More, or a right-click on its row).
Private Sub QuestionActions(ByVal atPointer As Boolean)
    Dim sme As Boolean, st As String, canNew As Boolean, canDel As Boolean, a(5) As String, k As String
    Dim problem As String, shown As Long, nm As String
    If m_qid = 0 Then Exit Sub
    sme = HasRole("SME")
    st = QStatusName(Nz(DLookup("StatusID", "tblQuestions", "QuestionID=" & m_qid), 0))
    canNew = sme And st = QS_APPROVED And Nz(DLookup("IsCurrent", "tblQuestions", "QuestionID=" & m_qid), False)
    canDel = sme And Len(DeleteProblem(m_qid)) = 0
    a(0) = IIf(sme, "", "~") & "open|" & IIf(StatusEditable(st), "Edit", "Open") & "|E70F"
    a(1) = IIf(canNew, "", "~") & "newver|Edit as new version|E8C8"
    a(2) = "hist|History|E81C"
    a(3) = "copy|Copy question text|E8C8"
    a(4) = "-"
    a(5) = IIf(canDel, "", "~") & "del|!Delete draft|E74D"
    If atPointer Then
        k = KitMenuAt(a(0), a(1), a(2), a(3), a(4), a(5))
    Else
        k = KitMenu(Me.btnMore, a(0), a(1), a(2), a(3), a(4), a(5))
    End If
    Select Case k
        Case "open": EditQuestion m_qid, False
        Case "newver": EditQuestion m_qid, True
        Case "hist": btnHistory_Click
        Case "copy"
            If KitCopyText(Nz(DLookup("QuestionText", "tblQuestions", "QuestionID=" & m_qid), "")) Then KitToast "Question text copied.", "success"
        Case "del"
            problem = DeleteProblem(m_qid)
            If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: Exit Sub
            nm = Me.lblQRef.Caption
            If Not KitConfirmTyped("Delete " & nm & "?", "It was never submitted and is on no exam. Its parts, options and " & _
                    "citations go with it. This can't be undone.", PURGE_WORD, "Delete", MSG_TITLE) Then Exit Sub
            shown = DeleteDraft(m_qid)
            m_qid = 0
            RefreshList shown
            KitToast IIf(shown <> 0, "Draft deleted. The approved version is back in use.", "Draft deleted."), "destructive"
    End Select
End Sub

' ---------------------------------------------------------------------------
' The questions tab
' ---------------------------------------------------------------------------

' Load the questions for the search and reference filter, and show openId
' (0, or not listed: the first).
Private Sub RefreshList(ByVal openId As Long)
    Dim title As String
    Me.sfList.Form.RecordSource = QuestionsSql(True)
    GridHighlight Me.sfList, m_search
    title = "Questions"
    If m_refFilter <> 0 Then title = "Questions citing " & Left$(Nz(DLookup("RefTitle", "tblSourceRefs", "RefID=" & m_refFilter), "?"), 40)
    Me.lblListHdr.Caption = title
    SetVisible Me.lnkAllQs, m_refFilter <> 0 And m_tab = "q"
    ShowCounts
    m_qid = openId
    ApplyChip
End Sub

' The chip's filter on the list, then the list's state.
Private Sub ApplyChip()
    SetGridFilter Me.sfList, ChipWhere(m_chip)
    AfterFilter
End Sub

' After the rows shown changed: the empty state, and the shown question
' kept when still listed (else the first).
Private Sub AfterFilter()
    Dim g As Form, n As Long, any As Boolean
    Set g = Me.sfList.Form
    n = g.RecordsetClone.RecordCount
    any = DCount("*", "tblQuestions") > 0
    If m_tab = "q" Then
        If Not any Then
            ShowEmpty Me, "emptyList", n = 0, "No questions yet", "SMEs add the first one with New question."
            SetVisible Me.emptyList_b, False
        Else
            ShowEmpty Me, "emptyList", n = 0, "No questions match", "No question here fits the search, chip and column filters."
            Me.emptyList_b.Caption = "Show all"
        End If
        SetVisible Me.sfList, n > 0
    End If
    If MarkRow(Me.sfList, m_qid) Then
        ShowQuestion m_qid
    ElseIf IsNull(SelectFirstRow(Me.sfList)) Then
        ShowQuestion 0
    End If
End Sub

' The counts on the chips and the Questions tab, for the rows the search,
' reference filter and column funnels leave.
Private Sub ShowCounts()
    Dim rs As DAO.Recordset, where As String, n(5) As Long, i As Long
    If Not DataLinked() Then Exit Sub
    where = GridFilterWhere(Me.sfList)
    Set rs = CurrentDb.OpenRecordset("SELECT x.status, x.IsCurrent FROM (" & QuestionsSql(False) & ") AS x" & _
        IIf(Len(where) > 0, " WHERE " & where, ""), dbOpenSnapshot)
    Do While Not rs.EOF
        If Nz(rs!IsCurrent, False) Then n(0) = n(0) + 1
        Select Case rs!status
            Case QS_DRAFT: n(1) = n(1) + 1
            Case QS_REVIEW: n(2) = n(2) + 1
            Case QS_APPROVED: n(3) = n(3) + 1
            Case QS_REJECTED: n(4) = n(4) + 1
            Case QS_RETIRED: n(5) = n(5) + 1
        End Select
        rs.MoveNext
    Loop
    rs.Close
    For i = 0 To 5
        StyleChip Me.Controls("chip_" & i), i = m_chip, n(i)
    Next
    SetCount Me.tab_q, n(0)
End Sub

' The list's filter for chip i: All shows each question's current version.
Private Function ChipWhere(ByVal i As Long) As String
    Select Case i
        Case 0: ChipWhere = "IsCurrent<>0"
        Case 1: ChipWhere = "status=" & Q(QS_DRAFT)
        Case 2: ChipWhere = "status=" & Q(QS_REVIEW)
        Case 3: ChipWhere = "status=" & Q(QS_APPROVED)
        Case 4: ChipWhere = "status=" & Q(QS_REJECTED)
        Case 5: ChipWhere = "status=" & Q(QS_RETIRED)
    End Select
End Function

' The questions for the search and reference filter, newest first.
Private Function QuestionsSql(ByVal ordered As Boolean) As String
    Dim where As String, t As String
    If Len(m_search) > 0 Then
        t = Q("*" & LikeLiteral(m_search) & "*")
        where = "(q.QuestionText LIKE " & t & " OR tp.TopicName LIKE " & t & " OR ('Q' & q.QuestionKey) LIKE " & t & ")"
    End If
    If m_refFilter <> 0 Then
        where = where & IIf(Len(where) > 0, " AND ", "") & "q.QuestionID IN (SELECT QuestionID FROM tblQuestionRefs WHERE RefID=" & m_refFilter & ")"
    End If
    QuestionsSql = "SELECT q.QuestionID AS id, q.QuestionKey, 'Q' & q.QuestionKey & ' v' & q.VersionNo AS ref, " & _
        "Replace(Replace(Nz(q.QuestionText,''),Chr(13),' '),Chr(10),'') AS QuestionText, Left(Nz(q.QuestionText,''),255) AS qsort, " & _
        "tp.TopicName, IIf(ty.TypeName='" & QT_MC & "','MC',IIf(ty.TypeName='" & QT_TF & "','T/F','SA')) AS type_short, " & _
        "d.LevelName, s.StatusName AS status, q.IsCurrent " & _
        "FROM (((tblQuestions AS q LEFT JOIN tblTopics AS tp ON q.TopicID=tp.TopicID) " & _
        "LEFT JOIN tblQuestionTypes AS ty ON q.QuestionTypeID=ty.QuestionTypeID) " & _
        "LEFT JOIN tblDifficultyLevels AS d ON q.DifficultyID=d.DifficultyID) " & _
        "LEFT JOIN tblStatuses AS s ON q.StatusID=s.StatusID" & IIf(Len(where) > 0, " WHERE " & where, "") & _
        IIf(ordered, " ORDER BY q.QuestionKey DESC, q.VersionNo DESC", "")
End Function

' Show question qid on the right (0: the hint instead).
Private Sub ShowQuestion(ByVal qid As Long)
    Dim rs As DAO.Recordset, isOpen As Boolean, nm As Variant, st As String, s As String, last As Long, sme As Boolean
    m_qid = qid
    isOpen = (qid <> 0)
    If isOpen Then
        Set rs = CurrentDb.OpenRecordset("SELECT q.*, tp.TopicName, ty.TypeName, d.LevelName, c.Marking, r.RoleCode " & _
            "FROM ((((tblQuestions AS q LEFT JOIN tblTopics AS tp ON q.TopicID=tp.TopicID) " & _
            "LEFT JOIN tblQuestionTypes AS ty ON q.QuestionTypeID=ty.QuestionTypeID) " & _
            "LEFT JOIN tblDifficultyLevels AS d ON q.DifficultyID=d.DifficultyID) " & _
            "LEFT JOIN tblClassLevels AS c ON q.ClassLevelID=c.ClassLevelID) " & _
            "LEFT JOIN tblRoles AS r ON q.CurrentStepRoleID=r.RoleID WHERE q.QuestionID=" & qid, dbOpenSnapshot)
        If rs.EOF Then
            isOpen = False: m_qid = 0
        Else
            st = QStatusName(Nz(rs!StatusID, 0))
            Me.lblQRef.Caption = QuestionRef(rs!QuestionKey, rs!VersionNo)
            s = st
            If st = QS_REVIEW And Not IsNull(rs!RoleCode) Then s = s & ", waiting on the " & rs!RoleCode
            last = Nz(DMax("AuditID", "tblAuditTrail", "ItemType='" & ITEM_QUESTION & "' AND ItemID=" & qid), 0)
            If last <> 0 Then s = s & ". " & DLookup("ActionType", "tblAuditTrail", "AuditID=" & last) & " by " & AuditWho(last)
            Me.lblQState.Caption = s & "."
            Me.txtQText = rs!QuestionText
            Me.lblQTopic.Caption = Nz(rs!TopicName, "")
            Me.lblQType.Caption = Nz(rs!TypeName, "")
            Me.lblQDiff.Caption = Nz(rs!LevelName, "")
            Me.lblQPoints.Caption = PointsText(rs!TotalPoints)
            Me.lblQMark.Caption = Nz(rs!Marking, "")
            Me.lblQParts.Caption = PartsText(qid)
            Me.lblQRefs.Caption = RefsText(qid)
            sme = HasRole("SME")
            Me.btnOpen.Caption = IIf(StatusEditable(st), "Edit", "Open")
            SetEnabled Me.btnOpen, sme
            SetEnabled Me.btnMore, True
        End If
        rs.Close
    End If
    For Each nm In Array("lblQRef", "lblQState", "txtQText", "capQTopic", "lblQTopic", "capQType", "lblQType", "capQDiff", "lblQDiff", _
            "capQPoints", "lblQPoints", "capQMark", "lblQMark", "capQParts", "lblQParts", "capQRefs", "lblQRefs", "btnOpen", "btnHistory", "btnMore")
        SetVisible Me.Controls(nm), isOpen And m_tab = "q"
    Next
    SetVisible Me.lblDetailEmpty, Not isOpen And m_tab = "q"
End Sub

' "A, B (2 parts)" or "None".
Private Function PartsText(ByVal qid As Long) As String
    Dim rs As DAO.Recordset, s As String, n As Long
    Set rs = CurrentDb.OpenRecordset("SELECT PartLabel FROM tblQuestionParts WHERE QuestionID=" & qid & " ORDER BY SortOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & IIf(n > 0, ", ", "") & rs!PartLabel
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    PartsText = IIf(n = 0, "None", s)
End Function

' The references the question cites, one a line.
Private Function RefsText(ByVal qid As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT RefID, Paragraph FROM tblQuestionRefs WHERE QuestionID=" & qid, dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & IIf(Len(s) > 0, vbCrLf, "") & RefText(rs!RefID) & IIf(IsNull(rs!Paragraph), "", ", para " & rs!Paragraph)
        rs.MoveNext
    Loop
    rs.Close
    RefsText = IIf(Len(s) = 0, "None yet", s)
End Function

' ---------------------------------------------------------------------------
' The references tab
' ---------------------------------------------------------------------------

' Load the references for the search and show openId (0, or not listed:
' the first).
Private Sub RefreshRefs(ByVal openId As Long)
    Dim g As Form, n As Long
    Set g = Me.sfRefs.Form
    g.RecordSource = RefsSql()
    GridHighlight Me.sfRefs, m_search
    n = g.RecordsetClone.RecordCount
    If m_tab = "r" Then
        If Len(m_search) > 0 Then
            ShowEmpty Me, "emptyRefs", n = 0, "No references match", "No reference's title or number has those words."
        Else
            ShowEmpty Me, "emptyRefs", n = 0, "No references yet", "Add one with New reference, or cite one from the question editor."
            SetVisible Me.emptyRefs_b, False
        End If
        SetVisible Me.sfRefs, n > 0
    End If
    SetCount Me.tab_r, DCount("*", "tblSourceRefs")
    If MarkRow(Me.sfRefs, openId) Then
        ShowRef openId
    ElseIf IsNull(SelectFirstRow(Me.sfRefs)) Then
        ShowRef 0
    End If
End Sub

Private Function RefsSql() As String
    Dim where As String, t As String
    If Len(m_search) > 0 Then
        t = Q("*" & LikeLiteral(m_search) & "*")
        where = " WHERE r.RefTitle LIKE " & t & " OR r.RefNumber LIKE " & t
    End If
    RefsSql = "SELECT r.RefID AS id, r.RefTitle, r.RefNumber, r.Revision, " & _
        "(SELECT Count(*) FROM tblQuestionRefs AS x INNER JOIN tblQuestions AS q ON x.QuestionID=q.QuestionID " & _
        "WHERE x.RefID=r.RefID AND q.IsCurrent) AS cited, IIf(Nz(r.IsActive,False),'Active','Off') AS state " & _
        "FROM tblSourceRefs AS r" & where & " ORDER BY r.RefTitle, r.RefNumber"
End Function

' Show reference rid on the right (0: the hint instead).
Private Sub ShowRef(ByVal rid As Long)
    Dim rs As DAO.Recordset, isOpen As Boolean, nm As Variant, n As Long, c As Control
    m_rid = rid
    isOpen = (rid <> 0)
    If isOpen Then
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblSourceRefs WHERE RefID=" & rid, dbOpenSnapshot)
        If rs.EOF Then
            isOpen = False: m_rid = 0
        Else
            Me.txtRefTitle = rs!RefTitle
            Me.txtRefNumber = rs!RefNumber
            Me.txtRefRev = rs!Revision
            Me.lblRefState.Caption = IIf(Nz(rs!IsActive, False), "", "Turned off: it can't be cited again.")
        End If
        rs.Close
    End If
    If isOpen Then
        Me.sfRefQs.Form.RecordSource = "SELECT q.QuestionID AS id, 'Q' & q.QuestionKey & ' v' & q.VersionNo AS ref, " & _
            "Replace(Replace(Nz(q.QuestionText,''),Chr(13),' '),Chr(10),'') AS QuestionText, s.StatusName AS status " & _
            "FROM (tblQuestionRefs AS x INNER JOIN tblQuestions AS q ON x.QuestionID=q.QuestionID) " & _
            "LEFT JOIN tblStatuses AS s ON q.StatusID=s.StatusID WHERE x.RefID=" & rid & " AND q.IsCurrent ORDER BY q.QuestionKey DESC"
        Me.sfRefQs.Form.txtSel = Null
        n = Me.sfRefQs.Form.RecordsetClone.RecordCount
        Me.capRefQs.Caption = "QUESTIONS CITING IT (" & n & ")"
        SetEnabled Me.btnShowCiting, n > 0
    End If
    For Each nm In Array("txtRefTitle", "txtRefNumber", "txtRefRev")
        Me.Controls(nm).Locked = Not m_canRef
    Next
    SetEnabled Me.btnRefSave, m_canRef And isOpen
    SetEnabled Me.btnRefMore, m_canRef And isOpen
    For Each nm In Array("capRefTitle", "txtRefTitle", "capRefNumber", "txtRefNumber", "capRefRev", "txtRefRev", "btnRefSave", _
            "lblRefState", "capRefQs", "sfRefQs", "btnShowCiting")
        SetVisible Me.Controls(nm), isOpen And m_tab = "r"
    Next
    For Each c In Me.Controls
        If c.Name Like "hdrRefQs*" Then SetVisible c, isOpen And m_tab = "r"
    Next
    SetVisible Me.lblRefEmpty, Not isOpen And m_tab = "r"
    m_refSnap = RefSnapshot()
End Sub

Private Function RefSnapshot() As String
    RefSnapshot = Nz(Me.txtRefTitle, "") & "|" & Nz(Me.txtRefNumber, "") & "|" & Nz(Me.txtRefRev, "")
End Function

' Save the shown reference's fields. False when they can't be saved (the
' user was told why) or the user backed out.
Private Function SaveRef() As Boolean
    Dim problem As String, n As Long
    If m_rid = 0 Or Not m_canRef Then SaveRef = True: Exit Function
    If RefSnapshot() = m_refSnap Then SaveRef = True: Exit Function
    problem = RefProblem(Me.txtRefTitle, Me.txtRefNumber, Me.txtRefRev, m_rid)
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Function
    n = DCount("*", "tblQuestionRefs", "RefID=" & m_rid)
    If n > 0 Then
        If KitMsgBox(n & IIf(n = 1, " question cites", " questions cite") & " this reference, approved ones included, and will show " & _
                "the change. Correcting a typo or a revision is fine; a different document should be a new reference." & vbCrLf & vbCrLf & _
                "Save the change?", vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then Exit Function
    End If
    KitExecute "UPDATE tblSourceRefs SET RefTitle=" & Q(Trim$(Me.txtRefTitle)) & ", RefNumber=" & _
        IIf(IsNull(NullIfEmpty(Me.txtRefNumber)), "Null", Q(NullIfEmpty(Me.txtRefNumber))) & ", Revision=" & _
        IIf(IsNull(NullIfEmpty(Me.txtRefRev)), "Null", Q(NullIfEmpty(Me.txtRefRev))) & " WHERE RefID=" & m_rid
    RefreshRefs m_rid
    Me.lblRefState.Caption = "Saved."
    SaveRef = True
End Function

' Before the shown reference is left: save, drop or stay.
Private Function RefLeaveOK() As Boolean
    If m_rid = 0 Or Not m_canRef Or RefSnapshot() = m_refSnap Then RefLeaveOK = True: Exit Function
    Select Case KitMsgBox("Save your changes to the reference """ & Nz(Me.txtRefTitle, "") & """?", vbYesNoCancel + vbQuestion, MSG_TITLE)
        Case vbYes: RefLeaveOK = SaveRef()
        Case vbNo: ShowRef m_rid: RefLeaveOK = True
    End Select
End Function

' The shown reference's actions: turn off or on, its questions, delete.
Private Sub RefActions(ByVal atPointer As Boolean)
    Dim isOn As Boolean, nAll As Long, a(3) As String, k As String, nm As String
    If m_rid = 0 Or Not m_canRef Then Exit Sub
    If Not RefLeaveOK() Then Exit Sub
    isOn = Nz(DLookup("IsActive", "tblSourceRefs", "RefID=" & m_rid), False)
    nAll = DCount("*", "tblQuestionRefs", "RefID=" & m_rid)
    a(0) = "onoff|" & IIf(isOn, "Turn off|E8D8", "Turn on|E8FB")
    a(1) = IIf(nAll > 0, "", "~") & "citing|Show citing questions|E8A5"
    a(2) = "-"
    a(3) = IIf(nAll = 0, "", "~") & "del|!Delete reference|E74D"
    If atPointer Then k = KitMenuAt(a(0), a(1), a(2), a(3)) Else k = KitMenu(Me.btnRefMore, a(0), a(1), a(2), a(3))
    nm = Nz(DLookup("RefTitle", "tblSourceRefs", "RefID=" & m_rid), "")
    Select Case k
        Case "onoff"
            KitExecute "UPDATE tblSourceRefs SET IsActive=" & IIf(isOn, "False", "True") & " WHERE RefID=" & m_rid
            RefreshRefs m_rid
            KitToast IIf(isOn, "Turned off. Questions citing it keep it; no new citations.", "Turned on."), "success"
        Case "citing"
            ShowCiting m_rid, 0
        Case "del"
            If DCount("*", "tblQuestionRefs", "RefID=" & m_rid) > 0 Then
                KitMsgBox "A question cites """ & nm & """, so it can't be deleted. Turn it off instead.", vbInformation, MSG_TITLE
                Exit Sub
            End If
            If KitMsgBox("Delete the reference """ & nm & """? No question cites it.", vbYesNo + vbQuestion + vbDefaultButton2, _
                    MSG_TITLE) <> vbYes Then Exit Sub
            KitExecute "DELETE FROM tblSourceRefs WHERE RefID=" & m_rid
            m_rid = 0
            RefreshRefs 0
            KitToast "Reference deleted.", "destructive"
    End Select
End Sub

' The Questions tab showing only the questions citing reference rid, with
' qid shown (0: the first).
Private Sub ShowCiting(ByVal rid As Long, ByVal qid As Long)
    If rid = 0 Then Exit Sub
    If Not RefLeaveOK() Then Exit Sub
    m_refFilter = rid
    m_chip = 0
    ChipOn Me, "chip", 0
    ClearGridFilters Me.sfList
    m_tab = "q"
    SelectTab Me, "q"
    ShowGroup
    RefreshList qid
End Sub

' ---------------------------------------------------------------------------

' Show the open tab's controls and hide the other's; then each tab's own
' parts (empty states, detail, the filter link) as they stand.
Private Sub ShowGroup()
    Dim nm As Variant, onQ As Boolean
    onQ = (m_tab = "q")
    For Each nm In Split(Me.grpQ.Caption, ";")
        If Len(nm) > 0 Then SetVisible Me.Controls(nm), onQ
    Next
    For Each nm In Split(Me.grpR.Caption, ";")
        If Len(nm) > 0 Then SetVisible Me.Controls(nm), Not onQ
    Next
    Me.btnNew.Caption = IIf(onQ, "New question", "New reference")
    Me.btnNew.ControlTipText = IIf(onQ, "New question (Ctrl+N)", "Add a reference to the shared list (Ctrl+N)")
    If DataLinked() Then SetEnabled Me.btnNew, IIf(onQ, HasRole("SME"), m_canRef)
    SetVisible Me.lnkAllQs, onQ And m_refFilter <> 0
    If onQ Then
        ShowEmpty Me, "emptyRefs", False
        If DataLinked() Then AfterFilter Else ShowQuestion 0
    Else
        ShowEmpty Me, "emptyList", False
        If DataLinked() Then RefreshRefs m_rid Else ShowRef 0
    End If
End Sub

' Select row id of a table. False when it isn't listed.
Private Function MarkRow(sf As Control, ByVal id As Long) As Boolean
    Dim g As Form
    Set g = sf.Form
    g.txtSel = Null
    If id <> 0 Then
        With g.RecordsetClone
            .FindFirst "id=" & id
            If Not .NoMatch Then
                g.Bookmark = .Bookmark
                g.txtSel = id
                MarkRow = True
            End If
        End With
    End If
    g.Recalc
End Function
