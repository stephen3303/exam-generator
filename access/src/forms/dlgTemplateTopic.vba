Option Compare Database
Option Explicit

' Add a topic to a template, or change a topic's question count on it
' (Build_dlgTemplateTopic). OpenArgs: "tpl=<TemplateID>" to add, or
' "tpl=<TemplateID>;topic=<TopicID>" to change (the topic is then fixed).
' Picking a topic sets the slider to the topic's own questions per exam;
' the line under the slider says whether the topic has the approved
' questions the overlap rule needs for that count. OK writes
' tblTemplateTopics itself and hands back the topic in
' TempVars("TemplateTopicID"); RunDialog sees DialogResult "ok".
'
' A pick in the picker fires no event, so a timer watches for it.

Private Const MSG_TITLE As String = "Templates"

Private m_tpl As Long           ' the template
Private m_topic As Long         ' the topic being changed (0 when adding)
Private m_seen As Long          ' the pick the timer last saw

Private Sub Form_Load()
    Dim tplName As String, nm As String
    On Error GoTo Err_Handler
    m_tpl = ArgLong("tpl")
    m_topic = ArgLong("topic")
    tplName = Nz(DLookup("TemplateName", "tblTemplates", "TemplateID=" & m_tpl), "")
    If m_topic <> 0 Then
        nm = Nz(DLookup("TopicName", "tblTopics", "TopicID=" & m_topic), "")
        Me.Caption = "Change a topic's count"
        Me.lblHead.Caption = "Questions from " & nm & " on " & tplName
        Me.txtTopicKey = m_topic
        KitPickShow Me.txtTopic
        Me.txtTopic.Locked = True
        SetSlider Me, "sldCount", Nz(DLookup("QuestionCount", "tblTemplateTopics", "TemplateID=" & m_tpl & " AND TopicID=" & m_topic), _
            DefaultQuestionCount())
        Me.btnOK.Caption = "Save"
        m_seen = m_topic
    Else
        Me.Caption = "Add a topic"
        Me.lblHead.Caption = "Add a topic to " & tplName
        SetSlider Me, "sldCount", DefaultQuestionCount()
        Me.TimerInterval = 300
    End If
    ShowNeed
    CenterOnApp Me
    If m_topic = 0 Then Me.txtTopic.SetFocus
    Exit Sub
Err_Handler:
    HandleError "dlgTemplateTopic.Form_Load", Nz(Me.OpenArgs, "")
End Sub

' A new pick: start the count at the topic's own questions per exam.
Private Sub Form_Timer()
    Dim k As Long
    On Error GoTo Err_Handler
    k = Nz(Me.txtTopicKey, 0)
    If k = m_seen Then Exit Sub
    m_seen = k
    If k <> 0 Then SetSlider Me, "sldCount", Nz(DLookup("DefaultCount", "tblTopics", "TopicID=" & k), DefaultQuestionCount())
    ShowNeed
    Exit Sub
Err_Handler:
    Me.TimerInterval = 0                    ' one message, not one every 300 ms
    HandleError "dlgTemplateTopic.Form_Timer", k
End Sub

' Left and Right on the slider move it a step.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim nm As String
    On Error Resume Next                    ' no control has the focus
    nm = Me.ActiveControl.Name
    On Error GoTo Err_Handler
    If KeyCode <> vbKeyLeft And KeyCode <> vbKeyRight Then Exit Sub
    If Left$(nm, 10) = "sldCount_b" Then
        SliderStep Me, "sldCount", IIf(KeyCode = vbKeyLeft, -1, 1)
        KeyCode = 0
    End If
    Exit Sub
Err_Handler:
    HandleError "dlgTemplateTopic.Form_KeyDown", nm
End Sub

Private Sub txtTopic_KeyDown(KeyCode As Integer, Shift As Integer)
    KitPickKey Me.txtTopic, KeyCode
End Sub

' The slider moved (its onChange).
Public Sub CountChanged()
    On Error GoTo Err_Handler
    ShowNeed
    Exit Sub
Err_Handler:
    HandleError "dlgTemplateTopic.CountChanged"
End Sub

Private Sub btnOK_Click()
    Dim tid As Long, n As Long
    On Error GoTo Err_Handler
    tid = Nz(Me.txtTopicKey, 0)
    n = SliderValue(Me, "sldCount")
    If tid = 0 Then
        KitMsgBox "Type part of a topic's name, then pick it from the list under the box.", vbInformation, MSG_TITLE
        Me.txtTopic.SetFocus
        Exit Sub
    End If
    If m_topic = 0 Then
        If DCount("*", "tblTemplateTopics", "TemplateID=" & m_tpl & " AND TopicID=" & tid) > 0 Then
            KitMsgBox "That topic is already on this template. Close this and use Change to set its count.", vbInformation, MSG_TITLE
            Exit Sub
        End If
        KitExecute "INSERT INTO tblTemplateTopics (TemplateID, TopicID, QuestionCount, SortOrder) VALUES (" & m_tpl & ", " & _
            tid & ", " & n & ", " & NextOrder("tblTemplateTopics", "SortOrder", "TemplateID=" & m_tpl) & ")"
    Else
        KitExecute "UPDATE tblTemplateTopics SET QuestionCount=" & n & " WHERE TemplateID=" & m_tpl & " AND TopicID=" & tid
    End If
    Me.TimerInterval = 0
    TempVars("TemplateTopicID") = tid
    TempVars("DialogResult") = "ok"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgTemplateTopic.btnOK_Click", tid
End Sub

Private Sub btnCancel_Click()
    On Error GoTo Err_Handler
    Me.TimerInterval = 0
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
Err_Handler:
    HandleError "dlgTemplateTopic.btnCancel_Click"
End Sub

' ---------------------------------------------------------------------------

' The line under the slider: the picked topic's approved questions against
' what the overlap rule needs for this count.
Private Sub ShowNeed()
    Dim k As Long, have As Long, need As Long, n As Long
    k = Nz(Me.txtTopicKey, 0)
    If k = 0 Then
        Me.lblNeed.Caption = "Pick a topic to see whether it has the approved questions this count needs."
        Exit Sub
    End If
    n = SliderValue(Me, "sldCount")
    have = DCount("*", "tblQuestions", "TopicID=" & k & " AND IsCurrent AND StatusID=" & ApprovedStatusID())
    need = QuestionsNeeded(n, SettingLong(SET_OVERLAP_PCT, 30), SettingLong(SET_OVERLAP_WINDOW, 10))
    Me.lblNeed.Caption = "It has " & have & " approved " & IIf(have = 1, "question", "questions") & "; the overlap rule needs " & _
        need & " when an exam takes " & n & "." & IIf(have < need, " Until more are approved it will be short.", "")
End Sub

' The number after name= in OpenArgs ("tpl=3;topic=12"), or 0.
Private Function ArgLong(ByVal name As String) As Long
    Dim part As Variant
    For Each part In Split(Nz(Me.OpenArgs, ""), ";")
        If LCase$(Left$(part, Len(name) + 1)) = LCase$(name) & "=" Then
            If IsNumeric(Mid$(part, Len(name) + 2)) Then ArgLong = CLng(Mid$(part, Len(name) + 2))
        End If
    Next
End Function
