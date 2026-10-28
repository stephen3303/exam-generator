Option Compare Database
Option Explicit

' New exam (Build_dlgNewExam), from the Exams page, for the department head.
' Picking a template lists the topics it draws from (modExams.DrawPlanSql):
' how many questions each gets, how many approved questions it has and how
' many of those weren't on the last exams, and its state. The alert says
' what stops the draw (a topic Short of questions or with No SME) or what
' it will do (Repeats: more questions from recent exams than the overlap
' rule allows, which Generate asks about first). Generate draws the exam
' (GenerateExam), hands the toast's text back in TempVars("NewExamDone"),
' and RunDialog sees DialogResult "ok". Cancel, or Esc, changes nothing.

Private Const MSG_TITLE As String = "New exam"

Private m_keys As String        ' RecentKeysList(), read when the window opened and again at Generate
Private m_tpl As Long           ' the template picked (0: none)
Private m_ready As Boolean      ' every topic can be drawn
Private m_repeats As String     ' the topics over the overlap rule, by name, "" for none
Private m_repeatIds As String   ' the same topics' TopicIDs, joined with commas

Public Property Get HelpTopic() As String
    HelpTopic = "exams.new"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    TempVars("NewExamDone") = Null
    m_keys = RecentKeysList()
    ShowPlan
    CenterOnApp Me
    Me.cboTemplate.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgNewExam.Form_Load"
End Sub

Private Sub cboTemplate_AfterUpdate()
    On Error GoTo Err_Handler
    ShowPlan
    Exit Sub
Err_Handler:
    HandleError "dlgNewExam.cboTemplate_AfterUpdate", Nz(Me.cboTemplate, "")
End Sub

' A row clicked (subGrid): nothing more to show than the row itself.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
End Sub

Private Sub btnGenerate_Click()
    Dim problem As String, examId As Long, allow As String, r As String, nQ As Long, nT As Long
    Dim num As Long, desc As String, src As String
    On Error GoTo Err_Handler
    problem = ExamMakerProblem()
    If Len(problem) > 0 Then KitMsgBox problem, vbInformation, MSG_TITLE: Exit Sub
    If m_tpl = 0 Then
        KitMsgBox "Pick a template first.", vbInformation, MSG_TITLE
        Me.cboTemplate.SetFocus
        Exit Sub
    End If
    r = Nz(CleanText(Me.txtRecipient), "")
    problem = RecipientProblem(r)
    If Len(problem) > 0 Then
        KitMsgBox problem, vbInformation, MSG_TITLE
        Me.txtRecipient.SetFocus
        Exit Sub
    End If
    m_keys = RecentKeysList()               ' an exam generated since the window opened counts
    ShowPlan                                ' the bank as it stands now
    If Not m_ready Then KitMsgBox "The exam can't be drawn as things stand. The note under the topics says why.", _
        vbInformation, MSG_TITLE: Exit Sub
    If Len(m_repeats) > 0 Then
        If KitMsgBox("Generate anyway? " & m_repeats & " can only be filled with more questions from the last " & _
                SettingLong(SET_OVERLAP_WINDOW, 10) & " exams than the overlap rule allows. Their SMEs can swap some out " & _
                "if other approved questions come in.", vbOKCancel + vbQuestion, MSG_TITLE) <> vbOK Then Exit Sub
        allow = m_repeatIds                 ' only these topics: another that goes over since is refused
    End If
    KitBusy "Drawing the questions..."
    On Error GoTo Busy_Fail
    problem = GenerateExam(m_tpl, r, allow, examId)
    KitBusyEnd
    On Error GoTo Err_Handler
    If Len(problem) > 0 Then
        KitMsgBox problem, vbExclamation, MSG_TITLE
        ShowPlan
        Exit Sub
    End If
    nQ = DCount("*", "tblExamQuestions", "ExamID=" & examId)
    nT = DCount("*", "tblExamTopics", "ExamID=" & examId)
    TempVars("NewExamDone") = "Exam for " & r & " generated: " & nQ & IIf(nQ = 1, " question", " questions") & " in " & nT & _
        IIf(nT = 1, " topic", " topics") & ". Each topic's SMEs find it under My reviews."
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Busy_Fail:
    num = Err.Number: desc = Err.Description: src = Err.Source
    KitBusyEnd                              ' resets Err, so it is read first
    Err.Number = num: Err.Description = desc: Err.Source = src
Err_Handler:
    HandleError "dlgNewExam.btnGenerate_Click", m_tpl & " " & r
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgNewExam.btnCancel_Click"
End Sub

' List the picked template's topics and say what the draw would do.
Private Sub ShowPlan()
    Dim rs As DAO.Recordset, nTopics As Long, nQ As Long, noSME As String, shortOnes As String, off As Long
    Dim s As String, c As Control, w As Long
    m_tpl = Nz(Me.cboTemplate, 0)
    m_ready = False: m_repeats = "": m_repeatIds = ""
    If m_tpl <> 0 Then
        Me.sfPlan.Form.RecordSource = DrawPlanSql(m_tpl, m_keys)
        Set rs = Me.sfPlan.Form.RecordsetClone
        If rs.RecordCount > 0 Then rs.MoveFirst
        Do While Not rs.EOF
            nTopics = nTopics + 1
            nQ = nQ + rs!per_exam
            Select Case rs!state
                Case "No SME": noSME = noSME & IIf(Len(noSME) > 0, ", ", "") & rs!TopicName
                Case "Short": shortOnes = shortOnes & IIf(Len(shortOnes) > 0, ", ", "") & rs!TopicName
                Case "Repeats"
                    m_repeats = m_repeats & IIf(Len(m_repeats) > 0, ", ", "") & rs!TopicName
                    m_repeatIds = m_repeatIds & IIf(Len(m_repeatIds) > 0, ",", "") & rs!id
            End Select
            rs.MoveNext
        Loop
        off = TopicsOffCount(m_tpl)
    End If
    For Each c In Me.Controls
        If c.Name Like "hdrPlan*" Then SetVisible c, m_tpl <> 0
    Next
    SetVisible Me.sfPlan, nTopics > 0
    SetVisible Me.lblPlanEmpty, nTopics = 0
    If m_tpl = 0 Then
        Me.lblPlanEmpty.Caption = "Pick a template to see the topics it draws from."
        HideAlert Me, "alPlan"
    ElseIf nTopics = 0 Then
        Me.lblPlanEmpty.Caption = "This template has no topics that are turned on."
        ShowAlert Me, "alPlan", "Add topics to this template, or turn its topics on, on the admin's Templates and Topics pages.", _
            "warning", "Nothing to draw"
    ElseIf Len(noSME) > 0 Or Len(shortOnes) > 0 Then
        If Len(noSME) > 0 Then s = "No active SME reviews " & noSME & "; the admin adds one on the Topics page."
        If Len(shortOnes) > 0 Then s = s & IIf(Len(s) > 0, " ", "") & shortOnes & IIf(InStr(shortOnes, ",") > 0, " have", " has") & _
            " fewer approved questions than the template gives it."
        ShowAlert Me, "alPlan", s, "warning", "Can't generate yet"
    Else
        m_ready = True
        w = SettingLong(SET_OVERLAP_WINDOW, 10)
        If Len(m_repeats) > 0 Then
            ShowAlert Me, "alPlan", m_repeats & " can only be filled with more questions from the last " & w & " exams than the " & _
                "overlap rule allows. Generate asks before it does that.", "warning", nQ & " questions from " & nTopics & " topics"
        Else
            ShowAlert Me, "alPlan", "Every topic can be drawn within the overlap rule: at most " & SettingLong(SET_OVERLAP_PCT, 30) & _
                "% of a topic's questions from the last " & w & " exams.", "success", nQ & " questions from " & nTopics & " topics"
        End If
    End If
    If off > 0 And m_tpl <> 0 Then
        Me.lblPlanEmpty.Caption = Me.lblPlanEmpty.Caption & " " & off & IIf(off = 1, " topic is", " topics are") & _
            " turned off and left out."
        If nTopics > 0 Then Me.alPlan_msg.Caption = Me.alPlan_msg.Caption & " " & off & IIf(off = 1, " topic is", " topics are") & _
            " turned off and left out."
    End If
    SetEnabled Me.btnGenerate, m_ready
End Sub
