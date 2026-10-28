Option Compare Database
Option Explicit

' Exams (see Build_pgExams). Every exam, newest first
' (modExamApprovals.ExamsListSql), filtered by a status chip, the search box
' and the column funnels. The picked one shows on the right: where it is in
' the chain, who it is for and who generated it, where it stands (its
' overall marking once approved, the approver's reason when it was sent
' back), its topics with who signed each off, Download a copy (approved
' exams only: the Word document modExamDoc made at the SM's approval, made
' again here when missing) and History. Anyone who can use the app sees
' the page; only a DH sees New exam, which opens dlgNewExam, and the right
' is checked again in the data file when the exam is generated
' (modExams.RequireExamMaker). Another page can ask for an exam to be shown
' first through TempVars("ExamsExamID").

Private Const MSG_TITLE As String = "Exams"

Private m_chip As Long          ' the chip on (EXAM_CHIPS; 0: All)
Private m_search As String
Private m_id As Long            ' the exam shown on the right (0: none)

Public Property Get HelpTopic() As String
    HelpTopic = "exams"
End Property

Private Sub Form_Load()
    Dim openId As Long
    On Error GoTo Err_Handler
    SetVisible Me.btnNew, HasRole("DH") Or Not DataLinked()
    SetEnabled Me.btnNew, DataLinked()
    ChipOn Me, "chip", 0
    If Not DataLinked() Then
        ShowEmpty Me, "emptyList", True, "No data file linked", "This copy was built without the data file, so there are no exams."
        SetVisible Me.emptyList_b, False
        SetVisible Me.sfList, False
        ShowExam 0
        Exit Sub
    End If
    SendSignedOffExams                      ' an exam a failed step left signed off but not sent on
    openId = Nz(TempVars("ExamsExamID"), 0)
    TempVars("ExamsExamID") = Null
    RefreshList openId
    Exit Sub
Err_Handler:
    HandleError "pgExams.Form_Load"
End Sub

' ---------------------------------------------------------------------------
' Chips and search (the kit's builders call these)
' ---------------------------------------------------------------------------

Public Function ChipPick(ByVal i As Long)
    On Error GoTo Err_Handler
    If i = m_chip Then Exit Function
    m_chip = i
    ChipOn Me, "chip", i
    RefreshList m_id
    Exit Function
Err_Handler:
    HandleError "pgExams.ChipPick", i
End Function

Public Sub SearchFor(ByVal text As String)
    On Error GoTo Err_Handler
    m_search = Trim$(text)
    RefreshList m_id
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "pgExams.SearchFor", text
End Sub

' A column funnel changed the rows shown (GridHeaders' onFilter).
Public Sub ListFiltered(ByVal sfName As String)
    On Error GoTo Err_Handler
    AfterFilter
    Exit Sub
Err_Handler:
    HandleError "pgExams.ListFiltered", sfName
End Sub

' ---------------------------------------------------------------------------
' Buttons and rows
' ---------------------------------------------------------------------------

Private Sub btnNew_Click()
    Dim problem As String
    On Error GoTo Err_Handler
    problem = ExamMakerProblem()
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: Exit Sub
    If RunDialog("dlgNewExam") Then
        KitToast Nz(TempVars("NewExamDone"), "Exam generated."), "success"
        RefreshList Nz(DMax("ExamID", "tblExams", "CreatedByPersonID=" & SignedInPersonID()), 0)
    End If
    Exit Sub
Err_Handler:
    HandleError "pgExams.btnNew_Click"
End Sub

Private Sub btnDownload_Click()
    On Error GoTo Err_Handler
    DownloadClicked
    Exit Sub
Err_Handler:
    HandleError "pgExams.btnDownload_Click", m_id
End Sub

Private Sub btnHistory_Click()
    On Error GoTo Err_Handler
    If m_id <> 0 Then RunDialog "dlgHistory", "exam=" & m_id
    Exit Sub
Err_Handler:
    HandleError "pgExams.btnHistory_Click", m_id
End Sub

' The empty state's button: clear the chip, the search and the column filters.
Private Sub emptyList_b_Click()
    On Error GoTo Err_Handler
    m_chip = 0
    ChipOn Me, "chip", 0
    ClearGridFilters Me.sfList
    SetSearch Me, "txtSearch", ""
    Exit Sub
Err_Handler:
    HandleError "pgExams.emptyList_b_Click"
End Sub

' A row clicked: show it on the right.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    Dim id As Long
    On Error GoTo Err_Handler
    If grid <> "sfList" Then Exit Sub
    id = Nz(Me.sfList.Form.txtSel, 0)
    ShowExam id
    Exit Sub
Err_Handler:
    HandleError "pgExams.GridClick", grid & " " & id
End Sub

' A right-click on a row: its actions.
Public Sub GridRightClick(ByVal grid As String)
    Dim k As String, canDl As Boolean, waits As Boolean
    On Error GoTo Err_Handler
    If grid <> "sfList" Or m_id = 0 Then Exit Sub
    canDl = (Len(DownloadProblem(m_id)) = 0)
    waits = (Len(ExamApproveProblem(m_id)) = 0)
    k = KitMenuAt(IIf(canDl, "", "~") & "dl|Download a copy...|E896", "hist|History|E81C", _
        IIf(waits, "", "~") & "appr|Open in Approvals|E8FB")
    Select Case k
        Case "dl": DownloadClicked
        Case "hist": btnHistory_Click
        Case "appr": NavTo "approvals"
    End Select
    Exit Sub
Err_Handler:
    HandleError "pgExams.GridRightClick", grid
End Sub

' Save a copy of the shown exam's Word document where the person picks.
Private Sub DownloadClicked()
    Dim problem As String, dest As String
    If m_id = 0 Then Exit Sub
    problem = DownloadProblem(m_id)
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: RefreshList m_id: Exit Sub
    dest = KitSaveAs("Download a copy of the exam", ExamDocFileName(m_id), "Word document|*.docx", "exams")
    If Len(dest) = 0 Then Exit Sub
    problem = DownloadExam(m_id, dest)
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, MSG_TITLE: Exit Sub
    KitToast "Saved " & Mid$(dest, InStrRev(dest, "\") + 1) & ".", "success"
End Sub

' ---------------------------------------------------------------------------
' The list and the detail
' ---------------------------------------------------------------------------

' Load the exams for the chip and the search, with each chip's count, and
' show openId (0, or not listed: the first).
Private Sub RefreshList(ByVal openId As Long)
    Dim i As Long
    If Not DataLinked() Then Exit Sub
    Me.sfList.Form.RecordSource = ExamsListSql(m_chip, m_search)
    SetGridFilter Me.sfList, ""             ' the column funnels, on the new rows
    GridHighlight Me.sfList, m_search
    For i = 0 To UBound(Split(EXAM_CHIPS, "|"))
        StyleChip Me.Controls("chip_" & i), i = m_chip, ExamCount(i)
    Next
    m_id = openId
    AfterFilter
End Sub

' After the rows shown changed: the empty state, and the shown exam kept
' when still listed (else the first).
Private Sub AfterFilter()
    Dim n As Long
    n = Me.sfList.Form.RecordsetClone.RecordCount
    If ExamCount(0) = 0 Then
        ShowEmpty Me, "emptyList", n = 0, "No exams yet", "When the DH generates an exam it shows here, from SME review to final approval."
        SetVisible Me.emptyList_b, False
    Else
        ShowEmpty Me, "emptyList", n = 0, "No exams match", "No exam fits the chip, the search and the column filters."
        Me.emptyList_b.Caption = "Show all"
    End If
    SetVisible Me.sfList, n > 0
    If MarkRow(Me.sfList, m_id) Then
        ShowExam m_id
    ElseIf IsNull(SelectFirstRow(Me.sfList)) Then
        ShowExam 0
    Else
        ShowExam Nz(Me.sfList.Form.txtSel, 0)
    End If
End Sub

' Show exam examId on the right (0: the hint instead).
Private Sub ShowExam(ByVal examId As Long)
    Dim rs As DAO.Recordset, isOpen As Boolean, c As Control, s As String, by As String, st As String, mk As String
    Dim nT As Long, nDone As Long, nQ As Long
    m_id = examId
    isOpen = (examId <> 0)
    If isOpen Then
        Set rs = CurrentDb.OpenRecordset("SELECT e.VersionNo, e.RecipientName, e.StatusID, e.CurrentStepRoleID, e.CreatedByPersonID, " & _
            "e.CreatedOn, e.FinalApprovedOn, e.OverallClassLevelID, tm.TemplateName FROM tblExams AS e " & _
            "LEFT JOIN tblTemplates AS tm ON e.TemplateID=tm.TemplateID WHERE e.ExamID=" & examId, dbOpenSnapshot)
        If rs.EOF Then
            isOpen = False: m_id = 0
        Else
            Me.lblExam.Caption = ExamRef(examId, rs!VersionNo) & " for " & rs!RecipientName
            nQ = DCount("*", "tblExamQuestions", "ExamID=" & examId)
            by = PersonName(rs!CreatedByPersonID)
            Me.lblExamBy.Caption = "From the template " & Nz(rs!TemplateName, "(deleted)") & ", generated by " & _
                IIf(Len(by) > 0, by, "someone no longer listed") & " on " & Format$(rs!CreatedOn, "d mmm yyyy") & ". " & _
                nQ & IIf(nQ = 1, " question.", " questions.")
            nT = DCount("*", "tblExamTopics", "ExamID=" & examId)
            nDone = DCount("*", "tblExamTopics", "ExamID=" & examId & " AND SignedOffByPersonID Is Not Null")
            st = ExamStatusName(rs!StatusID)
            Select Case st
                Case ES_SME_REVIEW
                    s = "In SME review: " & nT - nDone & " of " & nT & IIf(nT = 1, " topic waits", " topics wait") & _
                        " for a sign-off. Then it goes to the DH."
                Case ES_REJECTED
                    s = ExamRejectionText(examId) & " Its SMEs review it again; then it goes to the DH."
                Case ES_REVIEW
                    s = "Waiting on the " & RoleCodeOf(rs!CurrentStepRoleID) & "'s approval. Every topic is signed off."
                Case ES_APPROVED
                    mk = LevelMarking(rs!OverallClassLevelID)
                    s = "Approved on " & Format$(rs!FinalApprovedOn, "d mmm yyyy") & IIf(Len(mk) > 0, ", marked " & mk & " overall", "") & _
                        ". Download a copy saves its Word document."
                Case Else
                    s = "In no known state."
            End Select
            Me.lblState.Caption = s
            Me.lblTopics.Caption = nDone & " of " & nT & " signed off"
            Me.txtPct = IIf(nT = 0, 0, Int(100 * nDone / nT))
            Me.sfTopics.Form.RecordSource = ExamTopicsSql(examId)
            SetEnabled Me.btnDownload, st = ES_APPROVED
            SetStep Me, "stpExam", ExamStepNumber(examId)
        End If
        rs.Close
    End If
    For Each c In Me.Controls
        Select Case True
            Case c.Name Like "stpExam*", c.Name Like "hdrTopics*", c.Name = "sfTopics", c.Name = "lblExam", c.Name = "lblExamBy", _
                    c.Name = "lblState", c.Name = "capTopics", c.Name = "lblTopics", c.Name = "prgTopics", c.Name = "btnDownload", _
                    c.Name = "btnHistory"
                SetVisible c, isOpen
        End Select
    Next
    SetVisible Me.lblDetailEmpty, Not isOpen
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
