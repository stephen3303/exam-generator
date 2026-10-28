Option Compare Database
Option Explicit

' Exam Generator (Access edition) - what reaches each person (build phase 9):
' the Outlook email notices, and the Dashboard's "waiting on you" list and
' recent activity.
'
' Each person has two switches (tblPeople, set on My settings):
'   NotifyDashboard  the Dashboard shows what waits on them (the tiles and
'                    the Waiting on you list); off, it shows neither
'   NotifyEmail      an Outlook email when something new waits on them
' New people start with the Dashboard on and email off (modAdmin.EnsurePerson).
'
' A notice goes out only for something that now waits on someone:
'   a question reaching a step        the step's role holders, not its author
'   a question sent back              its author
'   review tasks made                 each SME with an open review on the exam
'                                     (generated, or sent back to its SMEs)
'   an exam reaching a step           the step's role holders who signed off
'                                     none of its topics
' Final approvals send none. Each person gets their own email, and only when
' their Email switch is on, they are active here and in the Personnel file,
' have the app there and have an email address there; never the person acting.
' The emails name the question or exam (Q12 v2, E3 v1), its topic or template
' and the step, and never the question text, the recipient or a reason: those
' stay in the app.
'
' The Notice* subs are called after the step's transaction has committed and
' never raise. They only queue their emails: the page sends the queue with
' SendQueuedNotices straight after the action that queued it, once its own
' busy pop-up has ended and its dialog has closed, and on the action's error
' path too (SendNoticesWithToast), so no email waits for a later action. A
' sweep run as a page opens sends once the page has loaded and nothing else
' is open over the app (NoticesShouldWait). Outlook's own prompts (a warning
' asking to allow the send, a first-run profile picker) are Outlook's
' windows: the app's dim doesn't cover them, but they can open behind
' Access, so the busy text says Outlook may ask. A failed email is logged,
' the step stands, and TakeNoticeNote gives the page a line for its toast;
' a batch stops at its first failed email. Outlook is driven by late
' binding (no reference), from the copy that took the step: the running
' Outlook, else one started here and kept for the session (OpenOutlook).
'
' Each action clears the note first (StartNotices), so a line left by an
' earlier sweep never lands on its toast.

Private Const EMAIL_FOOTER As String = "You get these emails because Emails is on under My settings in the Exam Generator. " & _
    "Turn it off there to stop them."
Private Const TEST_FOOTER As String = "You get this one because you asked for a test under My settings. " & _
    "It is sent whatever the Emails switch says."
Private Const NOTE_SEND As String = "An email notice couldn't be sent, so check Outlook on this computer; it still shows in the app."
Private Const NOTE_OUTBOX As String = "Outlook was closed, so the emails may wait in your Outbox until you open it."
Private Const NOTE_OUTBOX_ONE As String = "Outlook was closed, so the email may wait in your Outbox until you open it."

Private m_note As String        ' why the last notices weren't all sent ("" when they were)
Private m_queue As Collection   ' the emails waiting for SendQueuedNotices: Array(person id, address, subject, body)
Private m_skipped As String     ' ";12;;34;": people whose unusable address was logged this session
Private m_ol As Object          ' an Outlook this app started (OpenOutlook), kept for the session; never quit

' ---------------------------------------------------------------------------
' The notices
' ---------------------------------------------------------------------------

' Question qid now waits at its current step: tell the step's role holders
' (not its author). Called after a submit and after an approval that moves
' it on. Never raises.
Public Sub NoticeQuestionAtStep(ByVal qid As Long)
    Dim rs As DAO.Recordset, stepId As Long, author As Long, what As String, stp As String
    On Error GoTo Fail
    Set rs = CurrentDb.OpenRecordset("SELECT q.QuestionKey, q.VersionNo, q.CurrentStepRoleID, q.AuthorPersonID, tp.TopicName " & _
        "FROM tblQuestions AS q LEFT JOIN tblTopics AS tp ON q.TopicID=tp.TopicID WHERE q.QuestionID=" & qid, dbOpenSnapshot)
    If Not rs.EOF Then
        stepId = Nz(rs!CurrentStepRoleID, 0): author = Nz(rs!AuthorPersonID, 0)
        what = QuestionRef(rs!QuestionKey, rs!VersionNo) & IIf(IsNull(rs!TopicName), "", " (" & rs!TopicName & ")")
    End If
    rs.Close
    If stepId = 0 Then Exit Sub
    stp = RoleCodeOf(stepId)
    SendToEach "SELECT PersonID FROM tblPersonRoles WHERE RoleID=" & stepId & " AND PersonID<>" & author, _
        what & " waits for your approval", _
        what & " waits for your approval as " & stp & " in the Exam Generator." & vbCrLf & vbCrLf & _
        "Open the Exam Generator and go to Approvals to approve it or send it back."
    Exit Sub
Fail:
    Failed "modNotify.NoticeQuestionAtStep", "question " & qid
End Sub

' Question qid was sent back: tell its author. Never raises.
Public Sub NoticeQuestionReturned(ByVal qid As Long)
    Dim rs As DAO.Recordset, author As Long, what As String
    On Error GoTo Fail
    Set rs = CurrentDb.OpenRecordset("SELECT q.QuestionKey, q.VersionNo, q.AuthorPersonID, tp.TopicName " & _
        "FROM tblQuestions AS q LEFT JOIN tblTopics AS tp ON q.TopicID=tp.TopicID WHERE q.QuestionID=" & qid, dbOpenSnapshot)
    If Not rs.EOF Then
        author = Nz(rs!AuthorPersonID, 0)
        what = QuestionRef(rs!QuestionKey, rs!VersionNo) & IIf(IsNull(rs!TopicName), "", " (" & rs!TopicName & ")")
    End If
    rs.Close
    If author = 0 Then Exit Sub
    SendToEach "SELECT AuthorPersonID FROM tblQuestions WHERE QuestionID=" & qid, _
        what & " was sent back to you", _
        "Your question " & what & " was sent back to you in the Exam Generator. The reason is in the app." & vbCrLf & vbCrLf & _
        "Open the Exam Generator: the question is in the Question bank. Change it and submit it again."
    Exit Sub
Fail:
    Failed "modNotify.NoticeQuestionReturned", "question " & qid
End Sub

' Exam examId has open review tasks (it was generated, or sent back to its
' SMEs): tell each SME with one, naming their topics. Never raises.
Public Sub NoticeReviews(ByVal examId As Long)
    Dim rs As DAO.Recordset, who As DAO.Recordset, what As String, topics As String, nTopics As Long
    On Error GoTo Fail
    what = ExamWhat(examId)
    If Len(what) = 0 Then Exit Sub
    Set who = Recipients("SELECT AssignedPersonID FROM tblTasks WHERE ItemType=" & Q(ITEM_EXAM) & " AND ItemID=" & examId & _
        " AND TaskType=" & Q(TASK_SME_REVIEW) & " AND TaskStatus=" & Q(TS_OPEN))
    Do While Not who.EOF
        topics = "": nTopics = 0
        Set rs = CurrentDb.OpenRecordset("SELECT tp.TopicName FROM (tblTasks AS tk INNER JOIN tblExamTopics AS et " & _
            "ON tk.ExamTopicID=et.ExamTopicID) INNER JOIN tblTopics AS tp ON et.TopicID=tp.TopicID WHERE tk.ItemType=" & _
            Q(ITEM_EXAM) & " AND tk.ItemID=" & examId & " AND tk.TaskStatus=" & Q(TS_OPEN) & " AND tk.AssignedPersonID=" & _
            who!PersonID & " ORDER BY tp.TopicName", dbOpenSnapshot)
        Do While Not rs.EOF
            topics = topics & IIf(Len(topics) > 0, ", ", "") & rs!TopicName
            nTopics = nTopics + 1
            rs.MoveNext
        Loop
        rs.Close
        QueueOne who!PersonID, who!email, what & " has " & IIf(nTopics > 1, "topics", "a topic") & " for you to review", _
            what & " has " & IIf(nTopics > 1, "these topics", "a topic") & " for you to review in the Exam Generator: " & _
            topics & "." & vbCrLf & vbCrLf & _
            "Open the Exam Generator and go to My reviews to check the questions drawn, swap any out and sign off. " & _
            "One SME's sign-off completes a topic."
        who.MoveNext
    Loop
    who.Close
    Exit Sub
Fail:
    Failed "modNotify.NoticeReviews", "exam " & examId
End Sub

' Exam examId now waits at its current step: tell the step's role holders
' who signed off none of its topics (the ones who can approve it). Called
' after the last sign-off sends it to the DH and after an approval that
' moves it on. Never raises.
Public Sub NoticeExamAtStep(ByVal examId As Long)
    Dim stepId As Long, what As String, stp As String
    On Error GoTo Fail
    stepId = Nz(DLookup("CurrentStepRoleID", "tblExams", "ExamID=" & examId & " AND StatusID=" & ExamStatusID(ES_REVIEW)), 0)
    If stepId = 0 Then Exit Sub
    what = ExamWhat(examId)
    stp = RoleCodeOf(stepId)
    SendToEach "SELECT PersonID FROM tblPersonRoles WHERE RoleID=" & stepId & " AND PersonID NOT IN " & _
        "(SELECT SignedOffByPersonID FROM tblExamTopics WHERE ExamID=" & examId & " AND SignedOffByPersonID Is Not Null)", _
        what & " waits for your approval", _
        what & " waits for your approval as " & stp & " in the Exam Generator." & vbCrLf & vbCrLf & _
        "Open the Exam Generator and go to Approvals, Exams, to approve it or send it back to its SMEs."
    Exit Sub
Fail:
    Failed "modNotify.NoticeExamAtStep", "exam " & examId
End Sub

' The start of an action that may queue notices: forget any line an earlier
' sweep or action left, so the action's toast reports only its own emails.
' The queue is kept: emails still in it are still due, and go with the next
' SendQueuedNotices.
Public Sub StartNotices()
    m_note = ""
End Sub

' Send the queued notices (the Notice* subs queue them), one email each,
' through one Outlook for the lot (OpenOutlook). Call it straight after the
' action that queued them, once its own busy pop-up has ended and its dialog
' has closed, before its toast; WithNoticeNote then says what wasn't sent.
' Shows "Sending N email notices..." while it works. The batch stops when
' the first email fails or two fail in a row, so a denied prompt isn't asked
' again for each one; a later lone failure (one person's address) is logged
' and the rest still go. One note says how many weren't sent.
' Errors: none (a failure is logged and noted).
Public Sub SendQueuedNotices()
    Dim batch As Collection, ol As Object, started As Boolean, isNew As Boolean, it As Variant, n As Long, nSent As Long
    Dim k As Long, nFail As Long, busy As Boolean
    On Error GoTo Fail
    If m_queue Is Nothing Then Exit Sub
    Set batch = m_queue
    Set m_queue = Nothing                   ' taken now, so nothing is sent twice
    n = batch.Count
    If n = 0 Then Exit Sub
    KitBusy "Sending " & n & IIf(n = 1, " email notice...", " email notices...") & _
        " Outlook may ask you to allow the send.": busy = True
    Set ol = OpenOutlook(started, isNew)
    If ol Is Nothing Then
        ' One log row (OpenOutlook's) and one note for the whole batch.
        AddNote "Outlook couldn't be opened on this computer, so the email notices weren't sent; the error log has the details."
    Else
        For Each it In batch
            k = k + 1
            If SendMail(ol, CLng(it(0)), CStr(it(1)), CStr(it(2)), CStr(it(3)), "; email " & k & " of " & n) Then
                nSent = nSent + 1: nFail = 0
            Else                                ' logged there
                nFail = nFail + 1
                If k = 1 Or nFail >= 2 Then Exit For    ' Outlook refuses them all: the rest aren't tried
            End If
        Next
        If nSent < n And n = 1 Then
            AddNote NOTE_SEND
        ElseIf nSent < n Then
            AddNote CStr(n - nSent) & " of " & n & " email notices " & IIf(n - nSent = 1, "wasn't", "weren't") & _
                " sent, so check Outlook on this computer; they still show in the app."
        End If
        If started And nSent > 0 Then FlushOutbox ol, IIf(isNew, NOTE_OUTBOX, "")
    End If
    Set ol = Nothing
    KitBusyEnd
    Exit Sub
Fail:
    LogError "modNotify.SendQueuedNotices", n & " emails"
    AddNote NOTE_SEND
    If busy Then KitBusyEnd
End Sub

' Why the notices since the last call weren't all sent, as one sentence for
' a toast, or "" when they were (or none was due). Clears it.
Public Function TakeNoticeNote() As String
    TakeNoticeNote = m_note
    m_note = ""
End Function

' A toast's text s with TakeNoticeNote's sentence after it, if any.
Public Function WithNoticeNote(ByVal s As String) As String
    Dim n As String
    n = TakeNoticeNote()
    WithNoticeNote = s & IIf(Len(n) > 0, " " & n, "")
End Function

' Send the queue now and show its note, if there is one, in a toast of its
' own after lead: for a path with no toast to carry it (an action's error
' handler, a page's opening sweep). The note says something went wrong or
' may wait, so the toast is a warning and stays up longer. Never raises.
Public Sub SendNoticesWithToast(Optional ByVal lead As String = "")
    Dim n As String
    On Error Resume Next
    SendQueuedNotices
    n = TakeNoticeNote()
    If Len(n) > 0 Then KitToast IIf(Len(lead) > 0, lead & " ", "") & n, "warning", 8000
End Sub

' True while something is open over the app (a dialog, a menu, a busy
' pop-up, another form), so a page's opening sweep holds its emails and its
' timer tries again. Never raises.
Public Function NoticesShouldWait() As Boolean
    Dim nm As String
    If AppDimmed() Or KitBusyDepth() > 0 Then NoticesShouldWait = True: Exit Function
    On Error Resume Next                    ' no active form: nothing is open over the app
    nm = Screen.ActiveForm.Name
    NoticesShouldWait = (Len(nm) > 0 And nm <> SHELL_FORM)
End Function

' Send the signed-in person a test email at the address the Personnel file
' has for them, whatever their Emails switch (My settings). Returns "" or why
' it couldn't be sent. When it went through an Outlook this app started,
' TakeNoticeNote says the email may wait in the Outbox.
' Errors: none (a failure is logged and described).
Public Function SendTestEmail() As String
    Dim addr As String, pid As Long, ol As Object, started As Boolean, isNew As Boolean
    On Error GoTo Fail
    pid = SignedInPersonID()
    addr = PersonEmail(pid)
    If Len(addr) = 0 Then
        SendTestEmail = PersonEmailProblem(pid)
        Exit Function
    End If
    Set ol = OpenOutlook(started, isNew)
    If ol Is Nothing Then
        SendTestEmail = "Outlook couldn't be opened. Check that Outlook is set up on this computer; the error log has the details."
        Exit Function
    End If
    If SendMail(ol, pid, addr, "Test email", "This is a test email from the Exam Generator. If you can read it, its emails reach you." & _
            vbCrLf & vbCrLf & TEST_FOOTER) Then
        If started Then FlushOutbox ol, IIf(isNew, NOTE_OUTBOX_ONE, "")
    Else
        SendTestEmail = "Outlook couldn't send it. Check that Outlook is set up and open on this computer; the error log has the details."
    End If
    Exit Function
Fail:
    LogError "modNotify.SendTestEmail"
    SendTestEmail = "It couldn't be sent; the error log has the details."
End Function

' The email address the Personnel file has for person pid, "" when none
' (or it isn't one usable address: IsAddress; PersonEmailProblem says which).
' Errors: the engine's.
Public Function PersonEmail(ByVal pid As Long) As String
    Dim v As String
    v = RawEmail(pid)
    If IsAddress(v) Then PersonEmail = v
End Function

' Why emails can't reach person pid, as sentences for them (My settings):
' the Personnel file has no email for them, or one that isn't a single
' usable address. "" when PersonEmail has one.
' Errors: the engine's.
Public Function PersonEmailProblem(ByVal pid As Long) As String
    Dim v As String
    v = RawEmail(pid)
    If Len(v) = 0 Then
        PersonEmailProblem = "The Personnel file has no email address for you. Ask the admin to add it."
    ElseIf InStr(v, ";") > 0 Or InStr(v, ",") > 0 Then
        PersonEmailProblem = "The Personnel file's email for you isn't a single address (it has ; or ,). Ask the admin to fix it."
    ElseIf Not IsAddress(v) Then
        PersonEmailProblem = "The Personnel file's email for you isn't a usable address. Ask the admin to fix it."
    End If
End Function

' The Personnel file's email field for person pid, trimmed ("" when empty).
Private Function RawEmail(ByVal pid As Long) As String
    RawEmail = Trim$(Nz(DLookup("email", PERSON_TABLE, "person_id=" & pid), ""))
End Function

' ---------------------------------------------------------------------------
' Sending
' ---------------------------------------------------------------------------

' "E12 v1 (Boat exam)": an exam and its template, or "" when there is no such exam.
Private Function ExamWhat(ByVal examId As Long) As String
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT e.VersionNo, tm.TemplateName FROM tblExams AS e LEFT JOIN tblTemplates AS tm " & _
        "ON e.TemplateID=tm.TemplateID WHERE e.ExamID=" & examId, dbOpenSnapshot)
    If Not rs.EOF Then ExamWhat = ExamRef(examId, rs!VersionNo) & IIf(IsNull(rs!TemplateName), "", " (" & rs!TemplateName & ")")
    rs.Close
End Function

' The people among idSql's PersonIDs whom an email may reach (see the
' header), with their address: PersonID, email.
Private Function Recipients(ByVal idSql As String) As DAO.Recordset
    Set Recipients = CurrentDb.OpenRecordset("SELECT e.PersonID, p.email FROM (tblPeople AS e INNER JOIN " & PERSON_TABLE & _
        " AS p ON e.PersonID=p.person_id) INNER JOIN " & PERSON_APP_TABLE & " AS a ON e.PersonID=a.person_id " & _
        "WHERE e.NotifyEmail AND e.IsActive AND p.is_active AND a.is_active AND a.app_code=" & Q(APP_CODE) & _
        " AND Nz(p.email,'')<>'' AND e.PersonID<>" & SignedInPersonID() & " AND e.PersonID IN (" & idSql & ")", dbOpenSnapshot)
End Function

' The same email to each person in idSql whom one may reach, one email each.
Private Sub SendToEach(ByVal idSql As String, ByVal subject As String, ByVal body As String)
    Dim rs As DAO.Recordset
    Set rs = Recipients(idSql)
    Do While Not rs.EOF
        QueueOne rs!PersonID, rs!email, subject, body
        rs.MoveNext
    Loop
    rs.Close
End Sub

' Queue one notice to person pid at addr, for SendQueuedNotices. An address
' that isn't one usable address (IsAddress) is skipped and logged once a
' session for that person, by their id alone: Personnel's to fix, and My
' settings tells them.
Private Sub QueueOne(ByVal pid As Long, ByVal addr As Variant, ByVal subject As String, ByVal body As String)
    addr = Trim$(Nz(addr, ""))
    If Not IsAddress(CStr(addr)) Then
        If InStr(m_skipped, ";" & pid & ";") = 0 Then
            m_skipped = m_skipped & ";" & pid & ";"
            On Error Resume Next            ' a log row of its own, as UpgradeSkip does
            Err.Raise vbObjectError + 1, "modNotify.QueueOne", "No usable email address"
            LogError "modNotify.QueueOne", "person " & pid & ": the Personnel file's email isn't one usable address, " & _
                "so their email notices are skipped"
            Err.Clear
            On Error GoTo 0
        End If
        Exit Sub
    End If
    If m_queue Is Nothing Then Set m_queue = New Collection
    m_queue.Add Array(pid, CStr(addr), subject, body & vbCrLf & vbCrLf & EMAIL_FOOTER)
End Sub

' True when s looks like one email address: something@something, with no
' ; or , that would make it two.
Private Function IsAddress(ByVal s As String) As Boolean
    IsAddress = (s Like "?*@?*") And InStr(s, ";") = 0 And InStr(s, ",") = 0
End Function

' Outlook for one batch of emails: the one this app started earlier in the
' session (m_ol) while it still runs, else the one already running, else a
' new one. A new one is kept in m_ol for the session, so its sends can finish
' and later batches don't start Outlook again; it is never quit (the person
' may be about to use it). started returns True for an Outlook this app
' started (now or earlier), which is asked to send its Outbox; isNew only
' when it was started by this call, when the emails may wait there (the
' note). Nothing when Outlook can't be started; that is logged once, here.
Private Function OpenOutlook(ByRef started As Boolean, ByRef isNew As Boolean) As Object
    Dim ol As Object, nm As String
    started = False: isNew = False
    On Error Resume Next
    If Not m_ol Is Nothing Then
        nm = m_ol.Name                      ' fails once that Outlook has closed
        If Err.Number = 0 Then
            started = True
            Set OpenOutlook = m_ol
            Exit Function
        End If
        Err.Clear
        Set m_ol = Nothing
    End If
    Set ol = GetObject(, "Outlook.Application")
    On Error GoTo Fail                      ' also clears GetObject's error when Outlook isn't running
    If ol Is Nothing Then
        Set ol = CreateObject("Outlook.Application")
        Set m_ol = ol                       ' kept for the session
        started = True: isNew = True
    End If
    Set OpenOutlook = ol
    Exit Function
Fail:
    LogError "modNotify.OpenOutlook"
    started = False: isNew = False
End Function

' After sending through an Outlook this code started: ask it to send the
' Outbox now, as a hidden Outlook may not do so by itself, and add note (the
' email or emails may wait there), unless it is "". Never raises.
Private Sub FlushOutbox(ol As Object, ByVal note As String)
    On Error Resume Next
    ol.Session.SendAndReceive False
    On Error GoTo 0
    If Len(note) > 0 Then AddNote note
End Sub

' Send one plain-text email through Outlook ol to person pid at addr. True
' when Outlook took it. A failure (no profile, a refused send) is logged, by
' the person's id, with context after the subject.
Private Function SendMail(ol As Object, ByVal pid As Long, ByVal addr As String, ByVal subject As String, _
        ByVal body As String, Optional ByVal context As String = "") As Boolean
    Dim mi As Object
    On Error GoTo Fail
    Set mi = ol.CreateItem(0)                   ' olMailItem
    mi.To = addr
    mi.subject = APP_NAME & ": " & subject
    mi.body = body
    mi.Send
    SendMail = True
    Exit Function
Fail:
    LogError "modNotify.SendMail", "to person " & pid & ": " & subject & context
End Function

' Add sentence s to the note for the toast, once.
Private Sub AddNote(ByVal s As String)
    If InStr(m_note, s) = 0 Then m_note = m_note & IIf(Len(m_note) > 0, " ", "") & s
End Sub

' A notice that failed before it was queued: log it and note it for the toast.
Private Sub Failed(ByVal source As String, ByVal context As String)
    LogError source, context
    AddNote "The email notices couldn't be worked out; the error log has the details."
End Sub

' ---------------------------------------------------------------------------
' The Dashboard
' ---------------------------------------------------------------------------

' True when the signed-in person's Dashboard switch is on (as at sign-in or
' as last set on My settings).
' Errors: none (False when the TempVars aren't set).
Public Function DashboardOn() As Boolean
    On Error Resume Next                    ' TempVars may be unset
    DashboardOn = CBool(Nz(TempVars("EGNotifyDash"), False))
End Function

' How many of the signed-in person's questions were sent back and wait on
' them (Rejected, their author).
' Errors: the engine's.
Public Function ReturnedCount() As Long
    If Not DataLinked() Then Exit Function
    ReturnedCount = DCount("*", "tblQuestions", "StatusID=" & QStatusID(QS_REJECTED) & " AND AuthorPersonID=" & SignedInPersonID())
End Function

' Questions the signed-in person can approve now: in review at a step whose
' role they hold, and not their own (ApproveProblem refuses those), with the
' exams they can approve (ExamApprovalCount). The Approvals page lists their
' own questions too, saying why; the Dashboard counts what they can act on.
' Errors: the engine's.
Public Function DashApprovalCount() As Long
    Dim rs As DAO.Recordset
    If Not DataLinked() Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM tblQuestions AS q WHERE " & QuestionWaitWhere(), dbOpenSnapshot)
    DashApprovalCount = Nz(rs.Fields(0).Value, 0) + ExamApprovalCount()
    rs.Close
End Function

' The questions (as q) the signed-in person can approve now (DashApprovalCount).
Private Function QuestionWaitWhere() As String
    Dim pid As Long
    pid = SignedInPersonID()
    QuestionWaitWhere = "q.StatusID=" & QStatusID(QS_REVIEW) & " AND q.AuthorPersonID<>" & pid & _
        " AND q.CurrentStepRoleID IN (SELECT RoleID FROM tblPersonRoles WHERE PersonID=" & pid & ")"
End Function

' The time of the last audit row on an item (its id is idExpr), as a
' subquery; with action, the last row of that action only.
Private Function LastActionSql(ByVal itemType As String, ByVal idExpr As String, Optional ByVal action As String = "") As String
    LastActionSql = "(SELECT Max(au.ActionOn) FROM tblAuditTrail AS au WHERE au.ItemType=" & Q(itemType) & " AND au.ItemID=" & idExpr & _
        IIf(Len(action) > 0, " AND au.ActionType=" & Q(action), "") & ")"
End Function

' Everything that waits on the signed-in person, the longest waiting first,
' for the Dashboard: id (a letter and a number: t TaskID, q QuestionID to
' approve, r QuestionID sent back, x ExamID to approve), kind (Review,
' Approve, Sent back), what ("Damage Control on E3", "Q12 v2 Damage
' Control", "E3 v2 Boat exam"), step (SME, DH, XO, SM or Author), since and
' since_txt ("3 Oct", with the year when it isn't this year). A question sent
' back waits since it was rejected, whatever the author saved since. The
' questions and exams to approve are those the person can act on
' (QuestionWaitWhere, modExamApprovals.ExamQueueWhere).
' Builds the SQL only (it reads the signed-in person's id and the status ids).
' Errors: none of its own; a data error passes up.
Public Function WaitingSql() As String
    Dim pid As Long
    pid = SignedInPersonID()
    WaitingSql = "SELECT w.*, IIf(Year(w.since)=Year(Date()),Format(w.since,'d mmm'),Format(w.since,'d mmm yyyy')) AS since_txt FROM (" & _
        "SELECT 't' & tk.TaskID AS id, 'Review' AS kind, tp.TopicName & ' on E' & tk.ItemID AS what, 'SME' AS step, " & _
        "tk.CreatedOn AS since FROM (tblTasks AS tk INNER JOIN tblExamTopics AS et ON tk.ExamTopicID=et.ExamTopicID) " & _
        "INNER JOIN tblTopics AS tp ON et.TopicID=tp.TopicID WHERE tk.AssignedPersonID=" & pid & " AND tk.TaskType=" & _
        Q(TASK_SME_REVIEW) & " AND tk.ItemType=" & Q(ITEM_EXAM) & " AND tk.TaskStatus=" & Q(TS_OPEN) & _
        " UNION ALL SELECT 'q' & q.QuestionID, 'Approve', 'Q' & q.QuestionKey & ' v' & q.VersionNo & ' ' & Nz(tp.TopicName,''), " & _
        "r.RoleCode, " & LastActionSql(ITEM_QUESTION, "q.QuestionID") & " FROM (tblQuestions AS q LEFT JOIN tblTopics AS tp " & _
        "ON q.TopicID=tp.TopicID) LEFT JOIN tblRoles AS r ON q.CurrentStepRoleID=r.RoleID WHERE " & QuestionWaitWhere() & _
        " UNION ALL SELECT 'x' & e.ExamID, 'Approve', 'E' & e.ExamID & ' v' & e.VersionNo & ' ' & Nz(tm.TemplateName,''), " & _
        "r.RoleCode, " & LastActionSql(ITEM_EXAM, "e.ExamID") & " FROM (tblExams AS e LEFT JOIN tblTemplates AS tm " & _
        "ON e.TemplateID=tm.TemplateID) LEFT JOIN tblRoles AS r ON e.CurrentStepRoleID=r.RoleID WHERE " & ExamQueueWhere() & _
        " UNION ALL SELECT 'r' & q.QuestionID, 'Sent back', 'Q' & q.QuestionKey & ' v' & q.VersionNo & ' ' & Nz(tp.TopicName,''), " & _
        "'Author', " & LastActionSql(ITEM_QUESTION, "q.QuestionID", AU_REJECTED) & " FROM tblQuestions AS q LEFT JOIN tblTopics AS tp " & _
        "ON q.TopicID=tp.TopicID WHERE q.StatusID=" & QStatusID(QS_REJECTED) & " AND q.AuthorPersonID=" & pid & _
        ") AS w ORDER BY w.since, w.id"
End Function

' The last n audit rows, newest first, for the Dashboard's recent activity:
' id (AuditID), when_txt, who, what ("Approved Q12 v2", "Generated E3 v1",
' "Changed a reference", "Deleted a question" once the draft is gone),
' ActionOn. An admin sees the whole app's. Anyone else sees the rows about
' what they had a hand in: a row on an item (a question, an exam or a
' reference) they did something to themselves, on a question they wrote,
' or on an exam with a review task for them. An approver's step they
' haven't acted on yet isn't counted; it shows under Waiting on you.
' Newest first by AuditID (the key, which rises with time), not ActionOn,
' which has no index.
' Builds the SQL only (it reads the signed-in person's id and IsExamAdmin).
' Errors: none.
Public Function RecentActivitySql(Optional ByVal n As Long = 20) As String
    Dim pid As Long, mine As String
    If Not IsExamAdmin() Then
        pid = SignedInPersonID()
        mine = " WHERE EXISTS (SELECT * FROM tblAuditTrail AS m WHERE m.ItemType=a.ItemType AND m.ItemID=a.ItemID " & _
            "AND m.PersonID=" & pid & ") OR (a.ItemType=" & Q(ITEM_QUESTION) & " AND q.AuthorPersonID=" & pid & ") " & _
            "OR (a.ItemType=" & Q(ITEM_EXAM) & " AND a.ItemID IN (SELECT tk.ItemID FROM tblTasks AS tk WHERE tk.ItemType=" & _
            Q(ITEM_EXAM) & " AND tk.AssignedPersonID=" & pid & "))"
    End If
    ' The question join ignores ItemType (Jet can't join on a constant): the
    ' IIf and the WHERE read q only on a question's row.
    RecentActivitySql = "SELECT TOP " & n & " a.AuditID AS id, Format(a.ActionOn,'d mmm h:nn') AS when_txt, " & _
        PersonNameSql() & " AS who, a.ActionType & ' ' & IIf(a.ItemType=" & Q(ITEM_QUESTION) & ", IIf(IsNull(q.QuestionKey), " & _
        "'a question', 'Q' & q.QuestionKey & ' v' & a.VersionNo), IIf(a.ItemType=" & Q(ITEM_EXAM) & ", 'E' & a.ItemID & ' v' & " & _
        "a.VersionNo, 'a reference')) AS what, a.ActionOn FROM (tblAuditTrail AS a LEFT JOIN " & PERSON_TABLE & _
        " AS p ON a.PersonID=p.person_id) LEFT JOIN tblQuestions AS q ON a.ItemID=q.QuestionID" & mine & " ORDER BY a.AuditID DESC"
End Function
