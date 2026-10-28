Option Compare Database
Option Explicit

' Swap a question on an exam topic (Build_dlgSwap), from My reviews.
' OpenArgs "eq=<ExamQuestionID>;stamp=<the topic's ChangeNo as My reviews
' showed it>", the question to swap out. Lists the approved questions of the
' topic in use with no version elsewhere on the exam, and the newer version
' of the one in the slot (modExams.SwapCandidatesSql), those given lately
' last, filtered by the search box. The picked one shows in full, and the
' alert says whether the overlap rule allows it (SwapProblem). Swap puts it on the exam
' (SwapQuestion), hands the toast's text back in TempVars("SwapDone"), and
' RunDialog sees DialogResult "ok". Cancel, or Esc, changes nothing. When
' another SME changed the topic since My reviews showed it, the swap is
' refused and the window closes, so My reviews shows the topic as it is.

Private Const MSG_TITLE As String = "Swap a question"

Private m_eq As Long            ' the exam question being swapped out
Private m_et As Long            ' its exam topic
Private m_keys As String        ' RecentKeysList for its exam
Private m_stamp As Long         ' the topic's change number My reviews showed
Private m_search As String
Private m_pick As Long          ' the question picked to swap in (0: none)

Public Property Get HelpTopic() As String
    HelpTopic = "review.swap"
End Property

Private Sub Form_Load()
    Dim qid As Long, ref As String
    On Error GoTo Err_Handler
    TempVars("SwapDone") = Null
    m_eq = OpenArgLong(Me, "eq")
    m_stamp = OpenArgLong(Me, "stamp")
    m_et = Nz(DLookup("ExamTopicID", "tblExamQuestions", "ExamQuestionID=" & m_eq), 0)
    qid = Nz(DLookup("QuestionID", "tblExamQuestions", "ExamQuestionID=" & m_eq), 0)
    ref = QuestionRef(DLookup("QuestionKey", "tblQuestions", "QuestionID=" & qid), DLookup("VersionNo", "tblQuestions", "QuestionID=" & qid))
    Me.Caption = "Swap " & ref
    Me.lblHead.Caption = "Swap " & ref & " for another question"
    Me.lblSub.Caption = ExamTopicName(m_et) & ". Pick an approved question of the topic to put in its place."
    m_keys = RecentKeysList(Nz(DLookup("ExamID", "tblExamTopics", "ExamTopicID=" & m_et), 0))
    RefreshList
    CenterOnApp Me
    Me.txtSearch.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgSwap.Form_Load", Nz(Me.OpenArgs, "")
End Sub

' The search box (SearchBox's handler).
Public Sub SearchFor(ByVal text As String)
    On Error GoTo Err_Handler
    m_search = Trim$(text)
    RefreshList
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "dlgSwap.SearchFor", text
End Sub

' A row clicked: show it; a double-click swaps it in when it may.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    ShowPick
    If dbl And Me.btnSwap.Enabled Then btnSwap_Click
    Exit Sub
Err_Handler:
    HandleError "dlgSwap.GridClick", grid
End Sub

' The empty state's button: clear the search.
Private Sub emptyCand_b_Click()
    On Error GoTo Err_Handler
    SetSearch Me, "txtSearch", ""
    Exit Sub
Err_Handler:
    HandleError "dlgSwap.emptyCand_b_Click"
End Sub

Private Sub btnSwap_Click()
    Dim problem As String, done As String
    On Error GoTo Err_Handler
    If m_pick = 0 Then
        KitMsgBox "Pick the question to put in its place.", vbInformation, MSG_TITLE
        Exit Sub
    End If
    problem = SwapQuestion(m_eq, m_pick, m_stamp, done)
    If Len(problem) > 0 Then
        KitMsgBox problem, vbInformation, MSG_TITLE
        If TopicChangeNo(m_et) <> m_stamp Then
            DoCmd.Close acForm, Me.Name, acSaveNo   ' My reviews reloads the topic as it is now
        Else
            RefreshList                     ' the list as it stands now
        End If
        Exit Sub
    End If
    TempVars("SwapDone") = done
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgSwap.btnSwap_Click", m_eq & " " & m_pick
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgSwap.btnCancel_Click"
End Sub

' Load the questions that could go in, for the search, and pick the first.
Private Sub RefreshList()
    Dim n As Long
    Me.sfCand.Form.RecordSource = SwapCandidatesSql(m_eq, m_keys, m_search)
    GridHighlight Me.sfCand, m_search
    n = Me.sfCand.Form.RecordsetClone.RecordCount
    If Len(m_search) = 0 Then
        ShowEmpty Me, "emptyCand", n = 0, "No question to swap in", "Every other approved question of this topic is on the exam already."
        SetVisible Me.emptyCand_b, False
    Else
        ShowEmpty Me, "emptyCand", n = 0, "No questions match", "No approved question of this topic has those words."
        Me.emptyCand_b.Caption = "Clear search"
    End If
    SetVisible Me.sfCand, n > 0
    If n > 0 Then SelectFirstRow Me.sfCand
    ShowPick
End Sub

' Show the picked question in full and whether it may go in.
Private Sub ShowPick()
    Dim problem As String, w As Long, key As Variant
    m_pick = 0
    If Me.sfCand.Visible Then m_pick = Nz(Me.sfCand.Form.txtSel, 0)
    If m_pick = 0 Then
        Me.txtPreview = Null
        HideAlert Me, "alSwap"
        SetEnabled Me.btnSwap, False
        Exit Sub
    End If
    Me.txtPreview = QuestionPreview(m_pick)
    problem = SwapProblem(m_eq, m_pick)
    If Len(problem) > 0 Then
        ShowAlert Me, "alSwap", problem, "warning"
    Else
        w = SettingLong(SET_OVERLAP_WINDOW, 10)
        key = DLookup("QuestionKey", "tblQuestions", "QuestionID=" & m_pick)
        If InKeyList(m_keys, key) Then
            ShowAlert Me, "alSwap", "It was on one of the " & w & " exams generated before this one, but the topic stays within " & _
                "the overlap rule with it.", "info"
        Else
            ShowAlert Me, "alSwap", "It wasn't on any of the " & w & " exams generated before this one.", "success"
        End If
    End If
    SetEnabled Me.btnSwap, Len(problem) = 0
End Sub
