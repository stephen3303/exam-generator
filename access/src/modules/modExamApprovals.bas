Option Compare Database
Option Explicit

' Exam Generator (Access edition) - exam approvals, from build phase 8: the
' chain an exam goes through once its topics are signed off, sending it
' back, and what the Exams page and the Approvals page's Exams tab show.
' The Word document is modExamDoc's.
'
' When the last topic of an exam is signed off (modExams.SignOffTopic) it
' waits at the DH's step (In Review). Approve moves it to the next step of
' the chain (tblRoles.ChainOrder: DH, XO, SM); the last step's approval
' makes it Approved, stamps FinalApprovedOn and the overall classification
' (the highest LevelRank of any question or part on it), and the page then
' makes the Word document. Reject sends it back to the SMEs with a reason
' (required): the approver picks one topic or all of them, those topics'
' sign-offs are cleared and their active SMEs get new reviews, and the exam
' is Rejected, its version one higher, until they sign off again and it
' goes back to the DH.
'
' Who may approve or reject is read from the data file at the moment of the
' write (RequireExamApprover): an active person who holds the step's role
' now and may still use the app, and who signed off none of the exam's
' topics (no one approves their own work). One person may hold the DH, XO
' and SM roles and approve every step. The DH who generated the exam may
' approve it: the draw is random and the SMEs check it. Every write is
' guarded by the state it expects, so two approvers acting at once can't
' both win, and writes an audit row (modQuestions.AddAudit).

Public Const AU_DOWNLOADED As String = "Downloaded"
Public Const AU_DOC_MADE As String = "Document made"
Public Const EXAM_REASON_WHO As String = "its SMEs know"   ' ReasonProblem's who, for an exam
Private Const PREVIEW_MAX As Long = 56000
Private Const MOVED_ON_X As String = "Someone else approved or sent back this exam a moment ago. The list now shows where it stands."

' The Exams page's chips, in order.
Public Const EXAM_CHIPS As String = "All|SME review|In review|Rejected|Approved"

' ---------------------------------------------------------------------------
' Naming and where an exam stands
' ---------------------------------------------------------------------------

' "E12 v2": how an exam version is named everywhere.
Public Function ExamRef(ByVal examId As Variant, ByVal ver As Variant) As String
    ExamRef = "E" & Nz(examId, "?") & " v" & Nz(ver, "?")
End Function

' An exam status's name, "" when unknown.
Public Function ExamStatusName(ByVal statusId As Variant) As String
    If IsNull(statusId) Then Exit Function
    ExamStatusName = Nz(DLookup("StatusName", "tblStatuses", "StatusID=" & CLng(statusId)), "")
End Function

' Where exam examId is in the chain, for the Stepper (SME, DH, XO, SM): 1
' while its topics are reviewed (or sent back), the step's place while it
' waits on an approver, and 5 (every step done) once approved.
Public Function ExamStepNumber(ByVal examId As Long) As Long
    Dim rs As DAO.Recordset, st As String
    ExamStepNumber = 1
    Set rs = CurrentDb.OpenRecordset("SELECT StatusID, CurrentStepRoleID FROM tblExams WHERE ExamID=" & examId, dbOpenSnapshot)
    If Not rs.EOF Then
        st = ExamStatusName(rs!StatusID)
        If st = ES_APPROVED Then
            ExamStepNumber = 5
        ElseIf st = ES_REVIEW Then
            ExamStepNumber = StepNumber(Nz(rs!CurrentStepRoleID, 0))
            If ExamStepNumber < 1 Then ExamStepNumber = 1
        End If
    End If
    rs.Close
End Function

' A count from sql ("SELECT Count(*) ..."), read through CurrentDb so that,
' inside a transaction, it sees the transaction's own writes.
Private Function CountOf(ByVal sql As String) As Long
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    CountOf = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function

' ---------------------------------------------------------------------------
' Who may approve (checked at the write)
' ---------------------------------------------------------------------------

' Why the signed-in person may not approve or reject exam examId now, or ""
' when they may. Reads the data file, so a role taken away since sign-in, or
' a step someone else moved on, counts at once. expectStep and expectVer,
' when given, are the step and version the person saw: one moved on, or sent
' back and resubmitted, since counts as acted on.
Public Function ExamApproveProblem(ByVal examId As Long, Optional ByVal expectStep As Long = 0, _
        Optional ByVal expectVer As Long = 0) As String
    Dim rs As DAO.Recordset, pid As Long, st As String, stepId As Long, code As String, ver As Long
    If Not DataLinked() Then ExamApproveProblem = "This copy has no data file.": Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT StatusID, CurrentStepRoleID, VersionNo FROM tblExams WHERE ExamID=" & examId, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        ExamApproveProblem = "This exam was deleted."
        Exit Function
    End If
    st = ExamStatusName(rs!StatusID)
    stepId = Nz(rs!CurrentStepRoleID, 0)
    ver = Nz(rs!VersionNo, 1)
    rs.Close
    If st <> ES_REVIEW Or stepId = 0 Then
        Select Case st
            Case ES_SME_REVIEW: st = "Its SMEs are still reviewing it"
            Case ES_REJECTED: st = "It was sent back to its SMEs"
            Case ES_APPROVED: st = "It already has final approval"
            Case Else: st = "It is in no known state"
        End Select
        ExamApproveProblem = st & ", so it isn't waiting for an approval."
        Exit Function
    End If
    If expectStep <> 0 And stepId <> expectStep Then ExamApproveProblem = MOVED_ON_X: Exit Function
    If expectVer <> 0 And ver <> expectVer Then ExamApproveProblem = MOVED_ON_X: Exit Function
    code = RoleCodeOf(stepId)
    pid = SignedInPersonID()
    If pid = 0 Or Not Nz(DLookup("IsActive", "tblPeople", "PersonID=" & pid), False) Or Not PersonHasRole(pid, code) Then
        ExamApproveProblem = "It is waiting on the " & code & ", and you aren't the " & code & " now. Your role may have been " & _
            "changed since you signed in."
        Exit Function
    End If
    If Not PersonCanUse(pid, APP_CODE) Then
        ExamApproveProblem = "Your access to the Exam Generator has been turned off in Personnel."
        Exit Function
    End If
    If DCount("*", "tblExamTopics", "ExamID=" & examId & " AND SignedOffByPersonID=" & pid) > 0 Then
        ExamApproveProblem = "You signed off a topic of this exam, so another " & code & " has to approve or reject it."
    End If
End Function

' Raise ERR_NOT_ALLOWED unless the signed-in person may approve or reject
' exam examId now (ExamApproveProblem). No side effects.
Public Sub RequireExamApprover(ByVal examId As Long, Optional ByVal expectStep As Long = 0)
    Dim problem As String
    problem = ExamApproveProblem(examId, expectStep)
    If Len(problem) > 0 Then Err.Raise ERR_NOT_ALLOWED, "modExamApprovals.RequireExamApprover", problem
End Sub

' ---------------------------------------------------------------------------
' Approve and send back
' ---------------------------------------------------------------------------

' Approve exam examId at its step, as the signed-in person. Before the last
' step it moves on to the next approver; at the last it becomes Approved,
' with FinalApprovedOn and its overall classification (OverallLevelID), in
' one transaction. isFinal returns True for that last approval, when the
' caller makes the Word document (modExamDoc.MakeExamDocument). Returns ""
' and what happened, for the toast, in doneText; or why it can't be
' approved now (someone else acted on it first).
' expectStep and expectVer are the step and version the person saw (0: as
' it stands); moved on or sent back and resubmitted since, it returns why
' instead, so nobody approves a version they haven't seen.
' Errors: ERR_NOT_ALLOWED when the person may not approve it.
Public Function ApproveExam(ByVal examId As Long, ByRef doneText As String, ByRef isFinal As Boolean, _
        Optional ByVal expectStep As Long = 0, Optional ByVal expectVer As Long = 0) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, stepId As Long, nextId As Long, ver As Long, n As Long, lvl As Long
    Dim num As Long, desc As String, src As String
    doneText = "": isFinal = False
    ' One step for the rights check and the guarded write: the one shown.
    stepId = Nz(DLookup("CurrentStepRoleID", "tblExams", "ExamID=" & examId), 0)
    If expectStep <> 0 And stepId <> expectStep Then ApproveExam = MOVED_ON_X: Exit Function
    ver = Nz(DLookup("VersionNo", "tblExams", "ExamID=" & examId), 1)
    If expectVer <> 0 And ver <> expectVer Then ApproveExam = MOVED_ON_X: Exit Function
    RequireExamApprover examId, stepId
    nextId = NextStepRoleID(stepId)
    If nextId = 0 Then
        ApproveExam = UnmarkedProblem(examId)
        If Len(ApproveExam) > 0 Then Exit Function
    End If
    Set ws = DBEngine.Workspaces(0)
    On Error GoTo Fail
    ws.BeginTrans: inTrans = True
    If nextId <> 0 Then
        n = KitExecute("UPDATE tblExams SET CurrentStepRoleID=" & nextId & StillAtStep(examId, stepId, ver))
    Else
        lvl = OverallLevelID(examId)
        n = KitExecute("UPDATE tblExams SET StatusID=" & ExamStatusID(ES_APPROVED) & ", CurrentStepRoleID=Null, " & _
            "FinalApprovedOn=Now(), OverallClassLevelID=" & IIf(lvl = 0, "Null", lvl) & StillAtStep(examId, stepId, ver))
    End If
    If n = 0 Then
        ws.Rollback: inTrans = False
        ApproveExam = "Someone else approved or sent back this exam a moment ago. The list now shows where it stands."
        Exit Function
    End If
    AddAudit ITEM_EXAM, examId, ver, AU_APPROVED, stepId
    ws.CommitTrans: inTrans = False
    If nextId <> 0 Then
        NoticeExamAtStep examId             ' queues the next step's emails; never raises
        doneText = "Approved. " & ExamRef(examId, ver) & " goes to the " & RoleCodeOf(nextId) & " next."
    Else
        isFinal = True
        doneText = "Approved. " & ExamRef(examId, ver) & " is final, and anyone can download a copy from the Exams page."
    End If
    Exit Function
Fail:
    num = Err.Number: desc = Err.Description: src = Err.Source
    If inTrans Then ws.Rollback
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modExamApprovals.ApproveExam", examId
End Function

' Send exam examId back from its step, as the signed-in person, with reason
' (required, in the audit row). etId names the exam topic to reopen; 0
' reopens every topic. Each reopened topic loses its sign-off and its
' active SMEs get a new open review; the exam is Rejected at the SME step,
' its version one higher. Returns "" or why it can't be sent back now (no
' reason, a topic with no SME to review it, or someone else acted first).
' expectStep and expectVer are the step and version the person saw (0: as
' it stands); moved on or sent back and resubmitted since, it returns why
' instead.
' Errors: ERR_NOT_ALLOWED when the person may not reject it.
Public Function RejectExam(ByVal examId As Long, ByVal reason As Variant, ByVal etId As Long, _
        Optional ByVal expectStep As Long = 0, Optional ByVal expectVer As Long = 0) As String
    Dim ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean, stepId As Long, ver As Long, n As Long
    Dim r As String, where As String, topics As String, noSme As String, ids As Variant, v As Variant, sql As String
    Dim plan As Collection, p As Variant, num As Long, desc As String, src As String
    ' One step for the rights check and the guarded write: the one shown.
    stepId = Nz(DLookup("CurrentStepRoleID", "tblExams", "ExamID=" & examId), 0)
    If expectStep <> 0 And stepId <> expectStep Then RejectExam = MOVED_ON_X: Exit Function
    If expectVer <> 0 And Nz(DLookup("VersionNo", "tblExams", "ExamID=" & examId), 1) <> expectVer Then
        RejectExam = MOVED_ON_X
        Exit Function
    End If
    RequireExamApprover examId, stepId
    RejectExam = ReasonProblem(reason, EXAM_REASON_WHO)
    If Len(RejectExam) > 0 Then Exit Function
    r = CleanText(reason)
    where = "et.ExamID=" & examId & IIf(etId <> 0, " AND et.ExamTopicID=" & etId, "")
    ' Who reviews each topic that goes back, read before anything is written.
    Set plan = New Collection
    Set rs = CurrentDb.OpenRecordset("SELECT et.ExamTopicID, et.TopicID, tp.TopicName FROM tblExamTopics AS et " & _
        "LEFT JOIN tblTopics AS tp ON et.TopicID=tp.TopicID WHERE " & where & " ORDER BY et.ExamTopicID", dbOpenSnapshot)
    Do While Not rs.EOF
        ids = TopicReviewers(rs!TopicID)
        If Len(ids) = 0 Then
            noSme = noSme & IIf(Len(noSme) > 0, ", ", "") & Nz(rs!TopicName, "A topic")
        Else
            plan.Add Array(CLng(rs!ExamTopicID), ids)
        End If
        topics = topics & IIf(Len(topics) > 0, ", ", "") & Nz(rs!TopicName, "a topic")
        rs.MoveNext
    Loop
    rs.Close
    If plan.Count = 0 And Len(noSme) = 0 Then
        RejectExam = "That topic isn't on this exam any more. Pick again."
        Exit Function
    End If
    If Len(noSme) > 0 Then
        RejectExam = "No active SME reviews " & noSme & ", so no one could change it. The admin adds one on the Topics page; " & _
            "then send it back."
        Exit Function
    End If
    ver = Nz(DLookup("VersionNo", "tblExams", "ExamID=" & examId), 1)
    If expectVer <> 0 And ver <> expectVer Then RejectExam = MOVED_ON_X: Exit Function
    Set ws = DBEngine.Workspaces(0)
    On Error GoTo Fail
    ws.BeginTrans: inTrans = True
    n = KitExecute("UPDATE tblExams SET StatusID=" & ExamStatusID(ES_REJECTED) & ", CurrentStepRoleID=" & RoleID("SME") & _
        ", VersionNo=" & ver + 1 & StillAtStep(examId, stepId, ver))
    If n = 0 Then
        ws.Rollback: inTrans = False
        RejectExam = "Someone else approved or sent back this exam a moment ago, so it wasn't sent back."
        Exit Function
    End If
    For Each p In plan
        KitExecute "UPDATE tblExamTopics SET SignedOffByPersonID=Null, SignedOffOn=Null, ChangeNo=Nz(ChangeNo,0)+1 WHERE ExamTopicID=" & p(0)
        For Each v In Split(p(1), ",")
            sql = "INSERT INTO tblTasks (TaskType, ItemType, ItemID, ExamTopicID, AssignedPersonID, TaskStatus, CreatedOn) VALUES (" & _
                Q(TASK_SME_REVIEW) & ", " & Q(ITEM_EXAM) & ", " & examId & ", " & p(0) & ", " & CLng(v) & ", " & Q(TS_OPEN) & ", Now())"
            KitExecute sql
        Next
    Next
    ' The reason first, so the SMEs see it however many topics went back.
    AddAudit ITEM_EXAM, examId, ver, AU_REJECTED, stepId, r & IIf(Right$(r, 1) Like "[.!?]", "", ".") & " Sent back: " & topics & "."
    ws.CommitTrans: inTrans = False
    NoticeReviews examId                    ' queues the SMEs' emails; never raises
    Exit Function
Fail:
    num = Err.Number: desc = Err.Description: src = Err.Source
    If inTrans Then ws.Rollback
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modExamApprovals.RejectExam", examId
End Function

' Why exam examId can't be marked, or "": a question or part on it whose
' portion marking the classification list doesn't know. The final approval
' and the Word document both refuse such an exam.
Public Function UnmarkedProblem(ByVal examId As Long) As String
    Dim rs As DAO.Recordset, unmarked As String
    Set rs = CurrentDb.OpenRecordset("SELECT q.QuestionKey, q.VersionNo FROM (tblExamQuestions AS eq INNER JOIN tblQuestions AS q " & _
        "ON eq.QuestionID=q.QuestionID) LEFT JOIN tblClassLevels AS c ON q.ClassLevelID=c.ClassLevelID WHERE eq.ExamID=" & examId & _
        " AND c.ClassLevelID Is Null UNION SELECT q.QuestionKey, q.VersionNo FROM ((tblExamQuestions AS eq INNER JOIN tblQuestions AS q " & _
        "ON eq.QuestionID=q.QuestionID) INNER JOIN tblQuestionParts AS pt ON pt.QuestionID=q.QuestionID) LEFT JOIN tblClassLevels AS c " & _
        "ON pt.ClassLevelID=c.ClassLevelID WHERE eq.ExamID=" & examId & " AND c.ClassLevelID Is Null", dbOpenSnapshot)
    Do While Not rs.EOF
        unmarked = unmarked & IIf(Len(unmarked) > 0, ", ", "") & QuestionRef(rs!QuestionKey, rs!VersionNo)
        rs.MoveNext
    Loop
    rs.Close
    If Len(unmarked) > 0 Then
        UnmarkedProblem = unmarked & IIf(InStr(unmarked, ",") > 0, " have", " has") & " a question or part with no portion " & _
            "marking the classification list knows, so the document can't be marked. The admin checks the list on the Rules and " & _
            "lists page."
    End If
End Function

' " WHERE" for an update that only touches exam examId while it is still in
' review at step stepId and, when ver is given, at that version (Null reads
' as 1, as everywhere else).
Private Function StillAtStep(ByVal examId As Long, ByVal stepId As Long, Optional ByVal ver As Long = 0) As String
    StillAtStep = " WHERE ExamID=" & examId & " AND StatusID=" & ExamStatusID(ES_REVIEW) & " AND CurrentStepRoleID=" & stepId & _
        IIf(ver <> 0, " AND (VersionNo=" & ver & IIf(ver = 1, " OR VersionNo Is Null", "") & ")", "")
End Function

' The classification level with the highest LevelRank of any question or
' part on exam examId (0 when none is marked): the exam's overall marking.
Public Function OverallLevelID(ByVal examId As Long) As Long
    Dim rk As Long, rk2 As Long
    rk = CountOf("SELECT Max(c.LevelRank) FROM (tblExamQuestions AS eq INNER JOIN tblQuestions AS q ON eq.QuestionID=q.QuestionID) " & _
        "INNER JOIN tblClassLevels AS c ON q.ClassLevelID=c.ClassLevelID WHERE eq.ExamID=" & examId)
    rk2 = CountOf("SELECT Max(c.LevelRank) FROM (tblExamQuestions AS eq INNER JOIN tblQuestionParts AS pt ON eq.QuestionID=pt.QuestionID) " & _
        "INNER JOIN tblClassLevels AS c ON pt.ClassLevelID=c.ClassLevelID WHERE eq.ExamID=" & examId)
    If rk2 > rk Then rk = rk2
    If rk > 0 Then OverallLevelID = Nz(DMin("ClassLevelID", "tblClassLevels", "LevelRank=" & rk), 0)
End Function

' A classification level's marking ("CONFIDENTIAL"), "" when none.
Public Function LevelMarking(ByVal levelId As Variant) As String
    If Nz(levelId, 0) = 0 Then Exit Function
    LevelMarking = Nz(DLookup("Marking", "tblClassLevels", "ClassLevelID=" & CLng(levelId)), "")
End Function

' The topics of exam examId (etId: that one only) that no active SME
' reviews now, "Navigation, Damage control", or "": they can't be sent back.
Public Function ExamTopicsWithoutSme(ByVal examId As Long, ByVal etId As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT et.TopicID, tp.TopicName FROM tblExamTopics AS et LEFT JOIN tblTopics AS tp " & _
        "ON et.TopicID=tp.TopicID WHERE et.ExamID=" & examId & IIf(etId <> 0, " AND et.ExamTopicID=" & etId, "") & _
        " ORDER BY tp.TopicName", dbOpenSnapshot)
    Do While Not rs.EOF
        If Len(TopicReviewers(rs!TopicID)) = 0 Then s = s & IIf(Len(s) > 0, ", ", "") & Nz(rs!TopicName, "A topic")
        rs.MoveNext
    Loop
    rs.Close
    ExamTopicsWithoutSme = s
End Function

' Who a sent-back exam goes back to, for the reject dialog: the SMEs of the
' topic etId (0: of every topic) who would get a new review, or which
' topics have none, so it can't go back yet.
Public Function ExamReturnsToText(ByVal examId As Long, ByVal etId As Long) As String
    Dim rs As DAO.Recordset, s As String, ids As String, n As Long, nm As String, v As Variant, noSme As String
    noSme = ExamTopicsWithoutSme(examId, etId)
    If Len(noSme) > 0 Then
        ExamReturnsToText = "No active SME reviews " & noSme & " now, so " & IIf(etId <> 0, "it", "the exam") & _
            " can't go back yet. The admin adds one on the Topics page first."
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT et.TopicID FROM tblExamTopics AS et WHERE et.ExamID=" & examId & _
        IIf(etId <> 0, " AND et.ExamTopicID=" & etId, ""), dbOpenSnapshot)
    Do While Not rs.EOF
        For Each v In Split(TopicReviewers(rs!TopicID), ",")
            If InStr("," & ids & ",", "," & v & ",") = 0 Then ids = ids & IIf(Len(ids) > 0, ",", "") & v
        Next
        rs.MoveNext
    Loop
    rs.Close
    For Each v In Split(ids, ",")
        nm = PersonName(CLng(v))
        If Len(nm) > 0 Then
            n = n + 1
            If n <= 4 Then s = s & IIf(Len(s) > 0, "; ", "") & nm
        End If
    Next
    If n > 4 Then s = s & " and " & n - 4 & " more"
    If n = 0 Then
        ExamReturnsToText = "No active SME reviews " & IIf(etId <> 0, "this topic", "its topics") & " now. The admin adds one on the " & _
            "Topics page first."
    Else
        ExamReturnsToText = s & ". Each gets a new review under My reviews; once signed off again, the exam starts again at the DH."
    End If
End Function

' Why exam examId was last sent back, for its SMEs: "The XO sent the exam
' back (Smith, J. at 14:05 on 27 Oct). <reason>. Sent back: Navigation.".
' "" unless it is rejected now.
Public Function ExamRejectionText(ByVal examId As Long) As String
    Dim id As Long, rs As DAO.Recordset, cm As String
    If ExamStatusName(DLookup("StatusID", "tblExams", "ExamID=" & examId)) <> ES_REJECTED Then Exit Function
    id = Nz(DMax("AuditID", "tblAuditTrail", "ItemType='" & ITEM_EXAM & "' AND ItemID=" & examId & " AND ActionType=" & Q(AU_REJECTED)), 0)
    If id = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT StepRoleID, Comments FROM tblAuditTrail WHERE AuditID=" & id, dbOpenSnapshot)
    If Not rs.EOF Then
        cm = Nz(rs!comments, "no reason given")
        ExamRejectionText = "The " & IIf(Len(RoleCodeOf(rs!StepRoleID)) > 0, RoleCodeOf(rs!StepRoleID), "approver") & _
            " sent the exam back (" & AuditWho(id) & "). " & cm & IIf(Right$(cm, 1) Like "[.!?]", "", ".")
    End If
    rs.Close
End Function

' ---------------------------------------------------------------------------
' What the pages show
' ---------------------------------------------------------------------------

' How many questions on exam examId are no longer the approved version in
' use (a newer version approved since the draw, or retired): the SMEs may
' have signed them off, so the approvers are told.
Public Function ExamReplacedCount(ByVal examId As Long) As Long
    ExamReplacedCount = CountOf("SELECT Count(*) FROM tblExamQuestions AS eq INNER JOIN tblQuestions AS q ON eq.QuestionID=q.QuestionID " & _
        "WHERE eq.ExamID=" & examId & " AND NOT (q.IsCurrent AND q.StatusID=" & ApprovedStatusID() & ")")
End Function

' The exams waiting for the signed-in person (ExamQueueWhere), the longest
' waiting first. Columns: id, ref, exam (who
' it is for), TemplateName, step, since, since_txt. search keeps those whose
' recipient, template or number (E12) has it.
Public Function ExamQueueSql(ByVal search As String) As String
    Dim where As String, t As String
    where = ExamQueueWhere()
    If Len(search) > 0 Then
        t = Q("*" & LikeLiteral(search) & "*")
        where = where & " AND (e.RecipientName LIKE " & t & " OR tm.TemplateName LIKE " & t & " OR ('E' & e.ExamID) LIKE " & t & ")"
    End If
    ExamQueueSql = "SELECT x.*, Format(x.since,'d mmm') AS since_txt FROM (SELECT e.ExamID AS id, 'E' & e.ExamID & ' v' & e.VersionNo AS ref, " & _
        "e.RecipientName AS exam, tm.TemplateName, r.RoleCode AS step, (SELECT Max(a.ActionOn) FROM tblAuditTrail AS a " & _
        "WHERE a.ItemType='" & ITEM_EXAM & "' AND a.ItemID=e.ExamID) AS since " & _
        "FROM (tblExams AS e LEFT JOIN tblTemplates AS tm ON e.TemplateID=tm.TemplateID) " & _
        "LEFT JOIN tblRoles AS r ON e.CurrentStepRoleID=r.RoleID WHERE " & where & ") AS x ORDER BY x.since, x.id"
End Function

' How many exams wait for the signed-in person (the Exams tab's count).
Public Function ExamApprovalCount() As Long
    If Not DataLinked() Then Exit Function
    ExamApprovalCount = CountOf("SELECT Count(*) FROM tblExams AS e WHERE " & ExamQueueWhere())
End Function

' The exams (as e) the signed-in person can act on: in review at a step
' whose role they hold, and none of whose topics they signed off (another
' holder of the role approves those; ExamApproveProblem). The Dashboard's
' list uses it too (modNotify.WaitingSql).
' Builds the SQL only (it reads the signed-in person's id and a status id).
Public Function ExamQueueWhere() As String
    Dim pid As Long
    pid = SignedInPersonID()
    ExamQueueWhere = "e.StatusID=" & ExamStatusID(ES_REVIEW) & " AND e.CurrentStepRoleID IN (SELECT RoleID FROM tblPersonRoles " & _
        "WHERE PersonID=" & pid & ") AND NOT EXISTS (SELECT * FROM tblExamTopics AS so WHERE so.ExamID=e.ExamID AND " & _
        "so.SignedOffByPersonID=" & pid & ")"
End Function

' The status a chip of the Exams page keeps (EXAM_CHIPS; 0, All: "").
Private Function ChipStatus(ByVal chip As Long) As String
    Select Case chip
        Case 1: ChipStatus = ES_SME_REVIEW
        Case 2: ChipStatus = ES_REVIEW
        Case 3: ChipStatus = ES_REJECTED
        Case 4: ChipStatus = ES_APPROVED
    End Select
End Function

' Every exam for the Exams page, the newest first: id, ref, exam (who it is
' for), TemplateName, status, step (who it waits on), CreatedOn and
' created_txt. chip keeps one status (EXAM_CHIPS); search keeps those whose
' recipient, template or number (E12) has it.
Public Function ExamsListSql(ByVal chip As Long, ByVal search As String) As String
    Dim where As String, t As String
    where = "True"
    If chip <> 0 Then where = where & " AND e.StatusID=" & ExamStatusID(ChipStatus(chip))
    If Len(search) > 0 Then
        t = Q("*" & LikeLiteral(search) & "*")
        where = where & " AND (e.RecipientName LIKE " & t & " OR tm.TemplateName LIKE " & t & " OR ('E' & e.ExamID) LIKE " & t & ")"
    End If
    ExamsListSql = "SELECT e.ExamID AS id, 'E' & e.ExamID & ' v' & e.VersionNo AS ref, e.RecipientName AS exam, tm.TemplateName, " & _
        "s.StatusName AS status, IIf(s.StatusName=" & Q(ES_APPROVED) & ",'',r.RoleCode) AS step, e.CreatedOn, " & _
        "Format(e.CreatedOn,'d mmm yy') AS created_txt " & _
        "FROM ((tblExams AS e LEFT JOIN tblTemplates AS tm ON e.TemplateID=tm.TemplateID) " & _
        "LEFT JOIN tblStatuses AS s ON e.StatusID=s.StatusID) LEFT JOIN tblRoles AS r ON e.CurrentStepRoleID=r.RoleID " & _
        "WHERE " & where & " ORDER BY e.CreatedOn DESC, e.ExamID DESC"
End Function

' How many exams a chip of the Exams page keeps (its count).
Public Function ExamCount(ByVal chip As Long) As Long
    If Not DataLinked() Then Exit Function
    If chip = 0 Then
        ExamCount = DCount("*", "tblExams")
    Else
        ExamCount = DCount("*", "tblExams", "StatusID=" & ExamStatusID(ChipStatus(chip)))
    End If
End Function

' The topics of exam examId, in the exam's order, for the Exams page's
' detail: id (ExamTopicID), TopicName, qn (questions), signer and status
' (Signed off or Waiting).
Public Function ExamTopicsSql(ByVal examId As Long) As String
    ExamTopicsSql = "SELECT x.* FROM (SELECT et.ExamTopicID AS id, tp.TopicName, et.QuestionCount AS qn, " & _
        "IIf(et.SignedOffByPersonID Is Null,'',IIf(p.person_id Is Null,'(no longer listed)'," & PersonNameSql() & ")) AS signer, " & _
        "IIf(et.SignedOffByPersonID Is Null,'Waiting','Signed off') AS status, " & _
        "(SELECT Min(eq.SortOrder) FROM tblExamQuestions AS eq WHERE eq.ExamTopicID=et.ExamTopicID) AS ord " & _
        "FROM (tblExamTopics AS et LEFT JOIN tblTopics AS tp ON et.TopicID=tp.TopicID) " & _
        "LEFT JOIN " & PERSON_TABLE & " AS p ON et.SignedOffByPersonID=p.person_id " & _
        "WHERE et.ExamID=" & examId & ") AS x ORDER BY x.ord, x.id"
End Function

' Exam examId as plain text for an approver to read: each topic with who
' signed it off, then each question in the exam's order, numbered, in full
' (modApprovals.QuestionPreview: parts, markings, the correct options and
' the reference answers, which the Word document never shows).
Public Function ExamPreview(ByVal examId As Long) As String
    Dim rs As DAO.Recordset, s As String, b As String, i As Long, lastEt As Long, signer As String
    Set rs = CurrentDb.OpenRecordset("SELECT eq.QuestionID, eq.ExamTopicID, q.QuestionKey, q.VersionNo, tp.TopicName, " & _
        "et.SignedOffByPersonID, et.SignedOffOn, (q.IsCurrent AND q.StatusID=" & ApprovedStatusID() & ") AS inUse FROM ((tblExamQuestions AS eq INNER JOIN tblQuestions AS q ON eq.QuestionID=q.QuestionID) " & _
        "INNER JOIN tblExamTopics AS et ON eq.ExamTopicID=et.ExamTopicID) LEFT JOIN tblTopics AS tp ON et.TopicID=tp.TopicID " & _
        "WHERE eq.ExamID=" & examId & " ORDER BY eq.SortOrder, eq.ExamQuestionID", dbOpenSnapshot)
    Do While Not rs.EOF
        ' Each question's text (with its topic's heading when the topic
        ' starts) is built first, so it is added only when it fits.
        b = ""
        If rs!ExamTopicID <> lastEt Then
            signer = PersonName(rs!SignedOffByPersonID)
            b = IIf(Len(s) > 0, vbCrLf & vbCrLf, "") & UCase$(Nz(rs!TopicName, "A topic")) & _
                IIf(IsNull(rs!SignedOffOn), " (not signed off)", " (signed off by " & IIf(Len(signer) > 0, signer, "someone no longer listed") & _
                " on " & Format$(rs!SignedOffOn, "d mmm") & ")")
        End If
        b = b & vbCrLf & vbCrLf & (i + 1) & ". " & QuestionRef(rs!QuestionKey, rs!VersionNo) & _
            IIf(Nz(rs!inUse, False), "", " (REPLACED OR RETIRED SINCE THE DRAW)") & vbCrLf & QuestionPreview(rs!QuestionID)
        ' A text box holds about 64,000 characters: stop before a question
        ' that would take the text past PREVIEW_MAX (the first one is cut
        ' to fit, so there is something to read).
        If Len(s) + Len(b) > PREVIEW_MAX Then
            If Len(s) = 0 Then s = Left$(b, PREVIEW_MAX)
            s = s & vbCrLf & vbCrLf & "(The exam goes on, but the rest is too long to show here.)"
            Exit Do
        End If
        s = s & b
        lastEt = rs!ExamTopicID
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    ExamPreview = s
End Function

' Every audit row of exam examId, newest first, for the history window:
' id, ActionOn, when_txt, ver, ActionType, person, step, Comments.
Public Function ExamHistorySql(ByVal examId As Long) As String
    ExamHistorySql = "SELECT a.AuditID AS id, a.ActionOn, Format(a.ActionOn,'dd mmm yyyy hh:nn') AS when_txt, " & _
        "'v' & a.VersionNo AS ver, a.ActionType, " & PersonNameSql() & " AS person, r.RoleCode AS step, a.Comments " & _
        "FROM (tblAuditTrail AS a LEFT JOIN " & PERSON_TABLE & " AS p ON a.PersonID=p.person_id) " & _
        "LEFT JOIN tblRoles AS r ON a.StepRoleID=r.RoleID " & _
        "WHERE a.ItemType='" & ITEM_EXAM & "' AND a.ItemID=" & examId & " ORDER BY a.ActionOn DESC, a.AuditID DESC"
End Function

' The exam's history as plain text, one line per row, for Copy.
Public Function ExamHistoryText(ByVal examId As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset(ExamHistorySql(examId), dbOpenSnapshot)
    s = "History of exam E" & examId & vbCrLf
    Do While Not rs.EOF
        s = s & rs!when_txt & vbTab & rs!ver & vbTab & rs!ActionType & vbTab & Nz(rs!person, "") & _
            IIf(IsNull(rs!step), "", " (" & rs!step & ")") & IIf(IsNull(rs!comments), "", vbTab & rs!comments) & vbCrLf
        rs.MoveNext
    Loop
    rs.Close
    ExamHistoryText = s
End Function

' ---------------------------------------------------------------------------
' Download a copy
' ---------------------------------------------------------------------------

' Why exam examId can't be downloaded now, or "": only an approved exam has
' a document, and anyone who can use the app may download it.
Public Function DownloadProblem(ByVal examId As Long) As String
    Dim st As String
    If Not DataLinked() Then DownloadProblem = "This copy has no data file.": Exit Function
    st = ExamStatusName(DLookup("StatusID", "tblExams", "ExamID=" & examId))
    If Len(st) = 0 Then
        DownloadProblem = "This exam was deleted."
    ElseIf st <> ES_APPROVED Then
        Select Case st
            Case ES_SME_REVIEW: st = "Its SMEs are still reviewing it."
            Case ES_REVIEW: st = "It is waiting on an approver."
            Case Else: st = "It was sent back to its SMEs."
        End Select
        DownloadProblem = "The document is made when the SM gives the exam final approval. " & st
    End If
End Function

' Copy exam examId's document to dest (a path the person chose) and write
' the Downloaded audit row. The document is made first when it is missing
' (modExamDoc.EnsureExamDocument; note then says if its pages carry the
' overall marking). Returns "" or why it couldn't be copied.
Public Function DownloadExam(ByVal examId As Long, ByVal dest As String, Optional ByRef note As String) As String
    Dim problem As String, srcPath As String
    problem = DownloadProblem(examId)
    If Len(problem) > 0 Then DownloadExam = problem: Exit Function
    problem = EnsureExamDocument(examId, srcPath, note)
    If Len(problem) > 0 Then DownloadExam = problem: Exit Function
    On Error GoTo Fail
    FileCopy srcPath, dest
    On Error GoTo 0
    AddAudit ITEM_EXAM, examId, DLookup("VersionNo", "tblExams", "ExamID=" & examId), AU_DOWNLOADED, 0, _
        Mid$(dest, InStrRev(dest, "\") + 1)
    Exit Function
Fail:
    DownloadExam = "The copy couldn't be saved there (" & Err.Description & "). Pick another folder, or close the file if it is " & _
        "open in Word."
End Function
