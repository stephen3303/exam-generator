Option Compare Database
Option Explicit

' Exam Generator (Access edition) - exams, from build phase 7: the draw with
' the overlap check, and the SMEs' review of each exam topic. The pages are
' dlgNewExam (from the Exams page), My reviews (pgReview) and dlgSwap.
'
' The department head (DH) picks a template and types who the exam is for.
' GenerateExam draws each of the template's active topics at random from
' the approved questions in use, and gives every SME of the topic a review
' task, in one transaction. The exam is then in SME Review.
'
' The overlap rule (see the top of modAdmin): of the n questions a topic
' gets, at most Int(n * OverlapMaxPct / 100) may have been on any of the
' OverlapWindowExams exams generated before it (for a new exam, the last
' ones), whatever their template or recipient; that is "given lately". The
' order is ExamID's (the autonumber goes up), not CreatedOn, which each
' copy takes from its own PC's clock. A question counts by its QuestionKey,
' so a new version of a question given lately counts as given. The draw
' takes questions not given lately first, at random, and fills the rest
' from the ones given lately; a topic that can only be filled by going over
' the limit is drawn that way only when the DH says so (repeatTopics).
' Generates and swaps take turns (LockOverlapWindow), since each changes
' what the other's overlap check reads.
'
' Review: an SME with an open task checks the questions drawn for the
' topic, swaps any for another approved question of the topic (a swap may
' not add questions given lately over the limit) and signs off. One sign-off
' completes the topic: the signer's task is Done and the other SMEs' tasks
' on it are Closed. When the exam's last topic is signed off, the exam goes
' to the DH for approval (In Review; modExamApprovals). An approver who
' sends it back reopens one topic or all of them (Rejected): their SMEs get
' new reviews, and the last sign-off sends it to the DH again.
'
' Every write checks the person's right in the data file at the moment of
' the write (RequireExamMaker, RequireReviewer) and raises ERR_NOT_ALLOWED
' when they have none. Two people acting at once can't both win: each
' update is guarded by the state it expects, and the loser is told. An exam
' topic's ChangeNo goes up with every swap and its sign-off; My reviews keeps
' the number it showed (TopicChangeNo), and a swap or sign-off goes through
' only while the topic still has it, so nobody signs off or swaps on a list
' another SME has changed since.

Public Const ITEM_EXAM As String = "Exam"

' Exam statuses (tblStatuses, AppliesTo Exam).
Public Const ES_SME_REVIEW As String = "SME Review"
Public Const ES_REVIEW As String = "In Review"
Public Const ES_APPROVED As String = "Approved"
Public Const ES_REJECTED As String = "Rejected"

' Review tasks (tblTasks).
Public Const TASK_SME_REVIEW As String = "SME Review"
Public Const TS_OPEN As String = "Open"
Public Const TS_DONE As String = "Done"
Public Const TS_CLOSED As String = "Closed"

' Audit actions on an exam.
Public Const AU_GENERATED As String = "Generated"
Public Const AU_SWAPPED As String = "Swapped"
Public Const AU_SIGNED_OFF As String = "Signed off"
Public Const AU_REOPENED As String = "Reopened"     ' a signed-off topic sent back to its SMEs (a question was replaced)

Public Const MAX_RECIPIENT As Long = 150        ' tblExams.RecipientName is TEXT(150)

' ---------------------------------------------------------------------------
' Lookups
' ---------------------------------------------------------------------------

' An exam status's id (tblStatuses, AppliesTo Exam), 0 when missing.
' Errors: none of its own; a data error passes up.
Public Function ExamStatusID(ByVal nm As String) As Long
    ExamStatusID = Nz(DLookup("StatusID", "tblStatuses", "StatusName=" & Q(nm) & " AND AppliesTo='Exam'"), 0)
End Function

' The exam statuses in which its topics are still with their SMEs (they can
' be swapped and signed off, and the last sign-off sends the exam to the DH):
' SME Review, and Rejected (sent back by an approver), as ids for IN (...).
' Every "still in SME review" check uses this list.
' Errors: none of its own; a data error passes up.
Public Function SmeReviewStatusIDs() As String
    SmeReviewStatusIDs = ExamStatusID(ES_SME_REVIEW) & "," & ExamStatusID(ES_REJECTED)
End Function

' How many of the n questions a topic gets may have been given lately.
' n <= 0 gives 0. Reads OverlapMaxPct (30 when it isn't set).
' Errors: none of its own; a data error passes up.
Public Function RepeatsAllowed(ByVal n As Long) As Long
    If n <= 0 Then Exit Function
    RepeatsAllowed = Int(n * OverlapPct() / 100)
End Function

' The QuestionKeys of the questions given lately for exam forExamId: those on
' the OverlapWindowExams exams generated before it (0, for a new exam: the
' last ones generated), as a list for IN (...): "12,15,40", or "0" when there
' are none. Exams generated after it don't count against it: their own
' windows hold it (see LaterExamProblem).
' Errors: none of its own; a data error passes up.
Public Function RecentKeysList(Optional ByVal forExamId As Long = 0) As String
    Dim w As Long, ids As String, keys As String
    w = OverlapWindow()
    If w > 0 Then
        ids = ColumnList("SELECT TOP " & w & " ExamID FROM tblExams" & IIf(forExamId <> 0, " WHERE ExamID<" & forExamId, "") & _
            " ORDER BY ExamID DESC")
        If Len(ids) > 0 Then
            keys = ColumnList("SELECT DISTINCT q.QuestionKey FROM tblExamQuestions AS eq INNER JOIN tblQuestions AS q " & _
                "ON eq.QuestionID=q.QuestionID WHERE eq.ExamID IN (" & ids & ") AND q.QuestionKey Is Not Null")
        End If
    End If
    RecentKeysList = IIf(Len(keys) > 0, keys, "0")
End Function

' True when key is in list (a RecentKeysList).
' A Null key is never in it.
' Errors: none.
Public Function InKeyList(ByVal list As String, ByVal key As Variant) As Boolean
    If IsNull(key) Then Exit Function
    InKeyList = InStr("," & list & ",", "," & CLng(key) & ",") > 0
End Function

' The first column of every row of sql, joined with commas ("" for none).
Private Function ColumnList(ByVal sql As String) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & IIf(Len(s) > 0, ",", "") & rs.Fields(0).Value
        rs.MoveNext
    Loop
    rs.Close
    ColumnList = s
End Function

' Lock the overlap window's settings row until the transaction commits, so
' Generates and swaps, which change what each other's overlap check reads,
' run one at a time. A missing row would lock nothing, so it is logged.
Private Sub LockOverlapWindow()
    If KitExecute("UPDATE tblSettings SET SettingValue=SettingValue WHERE SettingKey=" & Q(SET_OVERLAP_WINDOW)) = 0 Then
        LogError "modExams.LockOverlapWindow", "No " & SET_OVERLAP_WINDOW & " row in tblSettings: Generates and swaps aren't serialised."
    End If
End Sub

' A count from sql ("SELECT Count(*) ..."), read through CurrentDb so that,
' inside a transaction, it sees the transaction's own writes.
Private Function CountOf(ByVal sql As String) As Long
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    CountOf = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function

' "Damage Control on the exam for USS Example": an exam topic, named.
' "" when there is no such exam topic.
' Errors: none of its own; a data error passes up.
Public Function ExamTopicName(ByVal etId As Long) As String
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT tp.TopicName, e.RecipientName FROM (tblExamTopics AS et INNER JOIN tblExams AS e " & _
        "ON et.ExamID=e.ExamID) LEFT JOIN tblTopics AS tp ON et.TopicID=tp.TopicID WHERE et.ExamTopicID=" & etId, dbOpenSnapshot)
    If Not rs.EOF Then ExamTopicName = Nz(rs!TopicName, "A topic") & " on the exam for " & rs!RecipientName
    rs.Close
End Function

' ---------------------------------------------------------------------------
' Generating an exam (dlgNewExam)
' ---------------------------------------------------------------------------

' Why the signed-in person can't generate an exam now, or "": only an active
' DH whom the Personnel file lets use the app can.
' Errors: none of its own; a data error passes up.
Public Function ExamMakerProblem() As String
    Dim pid As Long
    If Not DataLinked() Then ExamMakerProblem = "This copy has no data file.": Exit Function
    pid = SignedInPersonID()
    If pid = 0 Or Not Nz(DLookup("IsActive", "tblPeople", "PersonID=" & pid), False) Or Not PersonHasRole(pid, "DH") Then
        ExamMakerProblem = "Only the department head (DH) generates exams, and you aren't the DH now. Your role may have " & _
            "been changed since you signed in."
    ElseIf Not PersonCanUse(pid, APP_CODE) Then
        ExamMakerProblem = "Your access to the Exam Generator has been turned off in Personnel."
    End If
End Function

' Raise ERR_NOT_ALLOWED unless the signed-in person may generate an exam
' now (ExamMakerProblem, read from the data file at this moment).
' Errors: ERR_NOT_ALLOWED with the reason for the user.
Public Sub RequireExamMaker()
    Dim problem As String
    problem = ExamMakerProblem()
    If Len(problem) > 0 Then Err.Raise ERR_NOT_ALLOWED, "modExams.RequireExamMaker", problem
End Sub

' Who the exam is for, cleaned (CleanText) for its one-line field: a line
' break pasted in becomes a space. "" for none.
' Errors: none.
Public Function CleanRecipient(ByVal v As Variant) As String
    CleanRecipient = Replace(Replace(Nz(CleanText(v), ""), vbCrLf & vbCrLf, " "), vbCrLf, " ")
End Function

' Why recipient (already cleaned, CleanRecipient) won't do as who the exam
' is for, or "".
' Errors: none.
Public Function RecipientProblem(ByVal recipient As String) As String
    If Len(recipient) = 0 Then
        RecipientProblem = "Type who the exam is for: the boat, schoolhouse or other command."
    ElseIf Len(recipient) > MAX_RECIPIENT Then
        RecipientProblem = "Who the exam is for can be at most " & MAX_RECIPIENT & " characters; it is " & Len(recipient) & "."
    End If
End Function

' The active SMEs of topic topicId whom the Personnel file lets use the
' app, as ids joined with commas ("" for none). Only they get review tasks.
Public Function TopicReviewers(ByVal topicId As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT s.PersonID FROM tblTopicSMEs AS s INNER JOIN tblPeople AS pp ON s.PersonID=pp.PersonID " & _
        "WHERE s.TopicID=" & topicId & " AND pp.IsActive AND s.PersonID IN (SELECT PersonID FROM tblPersonRoles WHERE RoleID=" & _
        RoleID("SME") & ")", dbOpenSnapshot)
    Do While Not rs.EOF
        If PersonCanUse(rs!PersonID, APP_CODE) Then s = s & IIf(Len(s) > 0, ",", "") & rs!PersonID
        rs.MoveNext
    Loop
    rs.Close
    TopicReviewers = s
End Function

' The topics template tplId draws from, in its order, for the new exam
' window's table: id (TopicID), TopicName, per_exam, on_hand (approved
' questions in use), fresh (of those, not given lately; keys is
' RecentKeysList()), smes (1 when an SME may review it now, as TopicReviewers
' counts them for the draw, else 0) and state: No SME, Short (fewer approved
' questions than it gets), Repeats (it can only be filled with more
' questions given lately than the overlap rule allows) or Ready. Topics
' turned off are left out, as the draw leaves them out.
' Errors: none of its own; a data error passes up.
Public Function DrawPlanSql(ByVal tplId As Long, ByVal keys As String) As String
    Dim ap As Long, rs As DAO.Recordset, withSME As String
    ap = ApprovedStatusID()
    ' The topics with a reviewer, by the same check as the draw (Personnel included).
    Set rs = CurrentDb.OpenRecordset("SELECT t.TopicID FROM tblTemplateTopics AS tt INNER JOIN tblTopics AS t " & _
        "ON tt.TopicID=t.TopicID WHERE tt.TemplateID=" & tplId & " AND t.IsActive", dbOpenSnapshot)
    Do While Not rs.EOF
        If Len(TopicReviewers(rs!TopicID)) > 0 Then withSME = withSME & IIf(Len(withSME) > 0, ",", "") & rs!TopicID
        rs.MoveNext
    Loop
    rs.Close
    If Len(withSME) = 0 Then withSME = "0"
    DrawPlanSql = "SELECT x.*, IIf(x.smes=0,'No SME',IIf(x.on_hand<x.per_exam,'Short'," & _
        "IIf(x.fresh+Int(x.per_exam*" & OverlapPct() & "/100)<x.per_exam,'Repeats','Ready'))) AS state FROM (" & _
        "SELECT t.TopicID AS id, t.TopicName, tt.QuestionCount AS per_exam, tt.SortOrder, " & _
        "(SELECT Count(*) FROM tblQuestions AS q WHERE q.TopicID=t.TopicID AND q.IsCurrent AND q.StatusID=" & ap & ") AS on_hand, " & _
        "(SELECT Count(*) FROM tblQuestions AS q WHERE q.TopicID=t.TopicID AND q.IsCurrent AND q.StatusID=" & ap & _
        " AND q.QuestionKey NOT IN (" & keys & ")) AS fresh, " & _
        "IIf(t.TopicID IN (" & withSME & "),1,0) AS smes " & _
        "FROM tblTemplateTopics AS tt INNER JOIN tblTopics AS t ON tt.TopicID=t.TopicID " & _
        "WHERE tt.TemplateID=" & tplId & " AND t.IsActive) AS x ORDER BY x.SortOrder, x.TopicName"
End Function

' How many of template tplId's topics are turned off (the draw leaves them out).
' Errors: none of its own; a data error passes up.
Public Function TopicsOffCount(ByVal tplId As Long) As Long
    TopicsOffCount = CountOf("SELECT Count(*) FROM tblTemplateTopics AS tt INNER JOIN tblTopics AS t ON tt.TopicID=t.TopicID " & _
        "WHERE tt.TemplateID=" & tplId & " AND NOT t.IsActive")
End Function

' Draw template tplId for recipient, as the signed-in DH, and give each
' topic's SMEs a review task, all in one transaction. repeatTopics lists the
' TopicIDs (joined with commas, "" for none) the DH agreed may take more
' questions given lately than the overlap rule allows, because they can't be
' filled otherwise; any other topic over the rule stops the draw. The draw
' is made inside the transaction, after its first write locks the overlap
' window's settings row, so two DHs generating at once run one after the
' other and each exam's draw sees the other; every question is inserted only
' while it is still approved and in use. Returns "" with the new exam's id
' in examId, or why it can't be generated (who it is for, the template, or
' a topic short of questions, of SMEs, or over the rule); then nothing is
' written.
' Writes: the exam (SME Review, at the SME step), its topics, questions and
' review tasks (one per reviewer of each topic), and an audit row on the
' exam (Generated, at the DH step).
' Errors: ERR_NOT_ALLOWED unless the signed-in person is an active DH now
'   (RequireExamMaker); a data error rolls back and passes up.
Public Function GenerateExam(ByVal tplId As Long, ByVal recipient As Variant, ByVal repeatTopics As String, _
        ByRef examId As Long) As String
    Dim ws As DAO.Workspace, db As DAO.Database, rs As DAO.Recordset, plan As Collection, p As Variant, v As Variant
    Dim r As String, tplName As String, problem As String, inTrans As Boolean, etId As Long, i As Long, ap As Long
    Dim num As Long, desc As String, src As String
    RequireExamMaker
    examId = 0
    r = CleanRecipient(recipient)
    GenerateExam = RecipientProblem(r)
    If Len(GenerateExam) > 0 Then Exit Function
    tplName = Nz(DLookup("TemplateName", "tblTemplates", "TemplateID=" & tplId & " AND IsActive"), "")
    If Len(tplName) = 0 Then
        GenerateExam = "Pick a template. Only templates that are turned on can be used."
        Exit Function
    End If
    ap = ApprovedStatusID()
    DBEngine.Idle dbRefreshCache            ' the draw reads other copies' latest commits, not Jet's cache
    Set ws = DBEngine.Workspaces(0)
    On Error GoTo Fail
    ws.BeginTrans: inTrans = True
    ' Another DH's Generate, or an SME's swap, waits here, then reads this
    ' exam in its window.
    LockOverlapWindow
    Set plan = DrawPlan(tplId, RecentKeysList(), repeatTopics, problem)
    If Len(problem) > 0 Then
        ws.Rollback: inTrans = False
        GenerateExam = problem
        Exit Function
    End If
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblExams", dbOpenDynaset)
    rs.AddNew
    rs!TemplateID = tplId
    rs!RecipientName = r
    rs!VersionNo = 1
    rs!StatusID = ExamStatusID(ES_SME_REVIEW)
    rs!CurrentStepRoleID = RoleID("SME")
    rs!CreatedByPersonID = SignedInPersonID()
    rs!CreatedOn = Now()
    rs.Update
    rs.Bookmark = rs.LastModified
    examId = rs!ExamID
    rs.Close
    For Each p In plan
        Set rs = db.OpenRecordset("tblExamTopics", dbOpenDynaset)
        rs.AddNew
        rs!ExamID = examId
        rs!TopicID = p(0)
        rs!QuestionCount = p(2)
        rs!ChangeNo = 0
        rs.Update
        rs.Bookmark = rs.LastModified
        etId = rs!ExamTopicID
        rs.Close
        For Each v In Split(p(3), ",")
            i = i + 1                       ' the exam's order, topic after topic
            ' Only while it is still approved and in use: a question retired since
            ' the draw stops the exam rather than land on it.
            If KitExecute("INSERT INTO tblExamQuestions (ExamID, ExamTopicID, QuestionID, SortOrder, WasSwapped) SELECT " & _
                    examId & ", " & etId & ", QuestionID, " & i & ", False FROM tblQuestions WHERE QuestionID=" & CLng(v) & _
                    " AND IsCurrent AND StatusID=" & ap) = 0 Then
                ws.Rollback: inTrans = False
                examId = 0
                GenerateExam = "A question was replaced or retired while the exam was being drawn, so nothing was saved. " & _
                    "Generate again."
                Exit Function
            End If
        Next
        For Each v In Split(p(4), ",")
            KitExecute "INSERT INTO tblTasks (TaskType, ItemType, ItemID, ExamTopicID, AssignedPersonID, TaskStatus, CreatedOn) " & _
                "VALUES (" & Q(TASK_SME_REVIEW) & ", " & Q(ITEM_EXAM) & ", " & examId & ", " & etId & ", " & CLng(v) & ", " & _
                Q(TS_OPEN) & ", Now())"
        Next
    Next
    AddAudit ITEM_EXAM, examId, 1, AU_GENERATED, RoleID("DH"), "For " & r & " from the template " & tplName & ": " & i & _
        IIf(i = 1, " question", " questions") & " in " & plan.Count & IIf(plan.Count = 1, " topic", " topics") & "."
    ws.CommitTrans: inTrans = False
    NoticeReviews examId                    ' queues the SMEs' emails; never raises
    Exit Function
Fail:
    num = Err.Number: desc = Err.Description: src = Err.Source
    If inTrans Then ws.Rollback
    examId = 0
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modExams.GenerateExam", tplId & " " & r
End Function

' The draw for template tplId (keys: RecentKeysList()): one Array(TopicID,
' TopicName, n, question ids, SME ids) per active topic in the template's
' order, each list joined with commas, the questions in a random order.
' repeatTopics: the TopicIDs that may go over the overlap rule (GenerateExam).
' problem returns why it can't be drawn, or "" when it can.
Private Function DrawPlan(ByVal tplId As Long, ByVal keys As String, ByVal repeatTopics As String, _
        ByRef problem As String) As Collection
    Dim rs As DAO.Recordset, plan As New Collection, fresh As Variant, given As Variant, smes As String, ap As Long
    Dim n As Long, nFresh As Long, nGiven As Long, need As Long, noSME As String, shortOnes As String, over As String
    ap = ApprovedStatusID()
    problem = ""
    Randomize
    Set rs = CurrentDb.OpenRecordset("SELECT t.TopicID, t.TopicName, tt.QuestionCount FROM tblTemplateTopics AS tt " & _
        "INNER JOIN tblTopics AS t ON tt.TopicID=t.TopicID WHERE tt.TemplateID=" & tplId & " AND t.IsActive " & _
        "ORDER BY tt.SortOrder, t.TopicName", dbOpenSnapshot)
    Do While Not rs.EOF
        n = rs!QuestionCount
        fresh = Shuffled("SELECT QuestionID FROM tblQuestions WHERE TopicID=" & rs!TopicID & " AND IsCurrent AND StatusID=" & ap & _
            " AND QuestionKey NOT IN (" & keys & ")")
        given = Shuffled("SELECT QuestionID FROM tblQuestions WHERE TopicID=" & rs!TopicID & " AND IsCurrent AND StatusID=" & ap & _
            " AND QuestionKey IN (" & keys & ")")
        nFresh = ItemCount(fresh): nGiven = ItemCount(given)
        smes = TopicReviewers(rs!TopicID)
        If Len(smes) = 0 Then
            noSME = noSME & IIf(Len(noSME) > 0, ", ", "") & rs!TopicName
        ElseIf nFresh + nGiven < n Then
            shortOnes = shortOnes & IIf(Len(shortOnes) > 0, ", ", "") & rs!TopicName & " (" & (nFresh + nGiven) & " of " & n & ")"
        Else
            need = n - IIf(nFresh < n, nFresh, n)       ' questions given lately it has to take
            If need > RepeatsAllowed(n) And Not InKeyList(repeatTopics, rs!TopicID) Then over = over & IIf(Len(over) > 0, ", ", "") & rs!TopicName
            plan.Add Array(CLng(rs!TopicID), CStr(rs!TopicName), n, PickIds(fresh, given, n), smes)
        End If
        rs.MoveNext
    Loop
    rs.Close
    If Len(noSME) > 0 Then problem = "No active SME reviews " & noSME & ". Ask the admin to add one on the Topics page."
    If Len(shortOnes) > 0 Then problem = problem & IIf(Len(problem) > 0, vbCrLf & vbCrLf, "") & _
        "Not enough approved questions in " & shortOnes & "."
    If Len(problem) = 0 And Len(over) > 0 Then problem = "These topics can only be filled with more questions from recent " & _
        "exams than the overlap rule allows: " & over & ". The plan now shows them; Generate asks before it goes over."
    If Len(problem) = 0 And plan.Count = 0 Then problem = "This template has no topics that are turned on."
    Set DrawPlan = plan
End Function

' The ids of n questions, those in fresh first and then those in given (each
' already in a random order), shuffled together and joined with commas.
Private Function PickIds(ByVal fresh As Variant, ByVal given As Variant, ByVal n As Long) As String
    Dim a() As Long, k As Long, i As Long, j As Long, t As Long, s As String
    ReDim a(0 To n - 1)
    For i = 0 To ItemCount(fresh) - 1
        If k = n Then Exit For
        a(k) = fresh(i): k = k + 1
    Next
    For i = 0 To ItemCount(given) - 1
        If k = n Then Exit For
        a(k) = given(i): k = k + 1
    Next
    For i = n - 1 To 1 Step -1
        j = Int(Rnd() * (i + 1))
        t = a(i): a(i) = a(j): a(j) = t
    Next
    For i = 0 To n - 1
        s = s & IIf(i > 0, ",", "") & a(i)
    Next
    PickIds = s
End Function

' The first column of sql's rows (Longs) in a random order, or Empty for none.
Private Function Shuffled(ByVal sql As String) As Variant
    Dim rs As DAO.Recordset, a() As Long, n As Long, i As Long, j As Long, t As Long
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
        ReDim Preserve a(0 To n)
        a(n) = rs.Fields(0).Value
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    For i = n - 1 To 1 Step -1
        j = Int(Rnd() * (i + 1))
        t = a(i): a(i) = a(j): a(j) = t
    Next
    If n > 0 Then Shuffled = a
End Function

' Items in an array Shuffled returned (0 for Empty).
Private Function ItemCount(ByVal v As Variant) As Long
    If IsArray(v) Then ItemCount = UBound(v) - LBound(v) + 1
End Function

' ---------------------------------------------------------------------------
' Reviewing an exam topic (pgReview and dlgSwap)
' ---------------------------------------------------------------------------

' The signed-in person's review tasks for the My reviews list: id (TaskID),
' ExamID, ExamTopicID, exam (who it is for), TemplateName, TopicName,
' CreatedOn, since_txt and status (Open, Done or Closed). done False lists
' the open ones, oldest first; True the done and closed ones, newest first.
' search matches who it is for, the template and the topic.
' Builds the SQL only (it reads the signed-in person's id).
' Errors: none.
Public Function ReviewTasksSql(ByVal done As Boolean, ByVal search As String) As String
    Dim where As String, t As String
    where = "tk.AssignedPersonID=" & SignedInPersonID() & " AND tk.TaskType=" & Q(TASK_SME_REVIEW) & " AND tk.ItemType=" & _
        Q(ITEM_EXAM) & " AND tk.TaskStatus" & IIf(done, "<>", "=") & Q(TS_OPEN)
    If Len(search) > 0 Then
        t = Q("*" & LikeLiteral(search) & "*")
        where = where & " AND (e.RecipientName LIKE " & t & " OR tm.TemplateName LIKE " & t & " OR tp.TopicName LIKE " & t & ")"
    End If
    ReviewTasksSql = "SELECT tk.TaskID AS id, tk.ItemID AS ExamID, tk.ExamTopicID, e.RecipientName AS exam, tm.TemplateName, " & _
        "tp.TopicName, tk.CreatedOn, Format(tk.CreatedOn,'d mmm') AS since_txt, tk.TaskStatus AS status " & _
        "FROM (((tblTasks AS tk INNER JOIN tblExams AS e ON tk.ItemID=e.ExamID) " & _
        "INNER JOIN tblExamTopics AS et ON tk.ExamTopicID=et.ExamTopicID) " & _
        "INNER JOIN tblTopics AS tp ON et.TopicID=tp.TopicID) LEFT JOIN tblTemplates AS tm ON e.TemplateID=tm.TemplateID " & _
        "WHERE " & where & " ORDER BY " & IIf(done, "tk.CompletedOn DESC, tk.TaskID DESC", "tk.CreatedOn, tk.TaskID")
End Function

' How many open (done False) or done and closed review tasks the signed-in
' person has (the chips' counts).
' 0 when no data file is linked.
' Errors: none of its own; a data error passes up.
Public Function ReviewCount(ByVal done As Boolean) As Long
    If Not DataLinked() Then Exit Function
    ReviewCount = DCount("*", "tblTasks", "AssignedPersonID=" & SignedInPersonID() & " AND TaskType=" & Q(TASK_SME_REVIEW) & _
        " AND TaskStatus" & IIf(done, "<>", "=") & Q(TS_OPEN))
End Function

' The questions on exam topic etId, in the exam's order, for the review
' table: id (ExamQuestionID), QuestionID, ref, QuestionText, LevelName, pts
' and note, one of: "Replaced" (no longer the approved version in use: a
' newer version was approved since the draw, or it was retired); "Swapped,
' given lately"; "Given lately" (its key is in keys, the RecentKeysList for
' the exam); "Swapped"; or "". A swapped question given lately says both,
' so the SME can see every question that counts against the overlap rule.
' Builds the SQL only (it reads the approved status's id).
' Errors: none of its own; a data error passes up.
Public Function ExamTopicQsSql(ByVal etId As Long, ByVal keys As String) As String
    ExamTopicQsSql = "SELECT eq.ExamQuestionID AS id, eq.QuestionID, 'Q' & q.QuestionKey & ' v' & q.VersionNo AS ref, " & _
        "Replace(Replace(Nz(q.QuestionText,''),Chr(13),' '),Chr(10),'') AS QuestionText, d.LevelName, q.TotalPoints AS pts, " & _
        "IIf(Not (q.IsCurrent AND q.StatusID=" & ApprovedStatusID() & "),'Replaced'," & _
        "IIf(q.QuestionKey IN (" & keys & "),IIf(eq.WasSwapped,'Swapped, given lately','Given lately')," & _
        "IIf(eq.WasSwapped,'Swapped',''))) AS note, eq.SortOrder " & _
        "FROM (tblExamQuestions AS eq INNER JOIN tblQuestions AS q ON eq.QuestionID=q.QuestionID) " & _
        "LEFT JOIN tblDifficultyLevels AS d ON q.DifficultyID=d.DifficultyID " & _
        "WHERE eq.ExamTopicID=" & etId & " ORDER BY eq.SortOrder, eq.ExamQuestionID"
End Function

' The approved questions in use that could replace exam question eqId, for
' the swap window's table: id (QuestionID), ref, QuestionText, qsort,
' LevelName, pts and note (Given lately, from keys, or ""). Questions with a
' version elsewhere on the exam are left out; the newer version of the
' question in the slot itself is listed, so a replaced one can be updated.
' search matches the text and the number (Q12).
' Errors: none of its own; a data error passes up.
Public Function SwapCandidatesSql(ByVal eqId As Long, ByVal keys As String, ByVal search As String) As String
    Dim examId As Long, etId As Long, topicId As Long, slotQid As Long, where As String, t As String
    etId = Nz(DLookup("ExamTopicID", "tblExamQuestions", "ExamQuestionID=" & eqId), 0)
    slotQid = Nz(DLookup("QuestionID", "tblExamQuestions", "ExamQuestionID=" & eqId), 0)
    examId = Nz(DLookup("ExamID", "tblExamTopics", "ExamTopicID=" & etId), 0)
    topicId = Nz(DLookup("TopicID", "tblExamTopics", "ExamTopicID=" & etId), 0)
    where = "q.TopicID=" & topicId & " AND q.IsCurrent AND q.StatusID=" & ApprovedStatusID() & " AND q.QuestionID<>" & slotQid & _
        " AND q.QuestionKey NOT IN " & _
        "(SELECT q2.QuestionKey FROM tblExamQuestions AS eq INNER JOIN tblQuestions AS q2 ON eq.QuestionID=q2.QuestionID " & _
        "WHERE eq.ExamID=" & examId & " AND eq.ExamQuestionID<>" & eqId & " AND q2.QuestionKey Is Not Null)"
    If Len(search) > 0 Then
        t = Q("*" & LikeLiteral(search) & "*")
        where = where & " AND (q.QuestionText LIKE " & t & " OR ('Q' & q.QuestionKey) LIKE " & t & ")"
    End If
    SwapCandidatesSql = "SELECT q.QuestionID AS id, 'Q' & q.QuestionKey & ' v' & q.VersionNo AS ref, q.QuestionKey, " & _
        "Replace(Replace(Nz(q.QuestionText,''),Chr(13),' '),Chr(10),'') AS QuestionText, Left(Nz(q.QuestionText,''),255) AS qsort, " & _
        "d.LevelName, q.TotalPoints AS pts, IIf(q.QuestionKey IN (" & keys & "),'Given lately','') AS note " & _
        "FROM tblQuestions AS q LEFT JOIN tblDifficultyLevels AS d ON q.DifficultyID=d.DifficultyID WHERE " & where & _
        " ORDER BY IIf(q.QuestionKey IN (" & keys & "),1,0), q.QuestionKey"
End Function

' Why the signed-in person can't change or sign off exam topic etId now, or
' "": the topic must still wait for its sign-off on an exam in SME review
' (or sent back to it),
' and the person must be an active SME of the topic with an open review
' task for it, whom the Personnel file lets use the app.
' Errors: none of its own; a data error passes up.
Public Function ReviewProblem(ByVal etId As Long) As String
    Dim rs As DAO.Recordset, pid As Long, topicId As Long, s As String
    If Not DataLinked() Then ReviewProblem = "This copy has no data file.": Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT et.TopicID, et.SignedOffByPersonID, et.SignedOffOn, e.StatusID " & _
        "FROM tblExamTopics AS et INNER JOIN tblExams AS e ON et.ExamID=e.ExamID WHERE et.ExamTopicID=" & etId, dbOpenSnapshot)
    If rs.EOF Then
        s = "This exam was deleted."
    ElseIf Not IsNull(rs!SignedOffByPersonID) Then
        s = "Signed off by " & SignerName(rs!SignedOffByPersonID) & " on " & Format$(rs!SignedOffOn, "d mmm yyyy") & _
            ", so its questions are settled."
    ElseIf Not InKeyList(SmeReviewStatusIDs(), rs!StatusID) Then
        s = "The exam is past SME review, so its questions can't be changed here."
    Else
        topicId = rs!TopicID
    End If
    rs.Close
    If Len(s) = 0 Then
        pid = SignedInPersonID()
        If pid = 0 Or Not Nz(DLookup("IsActive", "tblPeople", "PersonID=" & pid), False) Or Not PersonHasRole(pid, "SME") Then
            s = "Only an SME of this topic can review it, and you aren't an SME now. Your role may have been changed since you signed in."
        ElseIf DCount("*", "tblTopicSMEs", "TopicID=" & topicId & " AND PersonID=" & pid) = 0 Then
            s = "You aren't an SME of this topic any more. Ask the admin if that's wrong."
        ElseIf DCount("*", "tblTasks", "ExamTopicID=" & etId & " AND AssignedPersonID=" & pid & " AND TaskStatus=" & Q(TS_OPEN)) = 0 Then
            s = "You have no open review of this topic."
        ElseIf Not PersonCanUse(pid, APP_CODE) Then
            s = "Your access to the Exam Generator has been turned off in Personnel."
        End If
    End If
    ReviewProblem = s
End Function

' Exam topic etId's change number (ChangeNo), read past Jet's cache so another
' copy's swap or sign-off of a moment ago counts. Read it before the topic's
' questions, keep it, and pass it to SwapQuestion and SignOffTopic.
' 0 when there is no such topic (or its number is Null).
' Errors: none of its own; a data error passes up.
Public Function TopicChangeNo(ByVal etId As Long) As Long
    DBEngine.Idle dbRefreshCache
    TopicChangeNo = Nz(DLookup("ChangeNo", "tblExamTopics", "ExamTopicID=" & etId), 0)
End Function

' Person pid's name ("Lt A. Smith"), or "someone no longer listed" when the
' Personnel file has no such person: who signed off or generated an exam.
' Errors: none of its own; a data error passes up.
Public Function SignerName(ByVal pid As Variant) As String
    SignerName = PersonName(pid)
    If Len(SignerName) = 0 Then SignerName = "someone no longer listed"
End Function

' Raise ERR_NOT_ALLOWED unless the signed-in person may change or sign off
' exam topic etId now (ReviewProblem, read from the data file at this moment).
' Errors: ERR_NOT_ALLOWED with the reason for the user.
Public Sub RequireReviewer(ByVal etId As Long)
    Dim problem As String
    problem = ReviewProblem(etId)
    If Len(problem) > 0 Then Err.Raise ERR_NOT_ALLOWED, "modExams.RequireReviewer", problem
End Sub

' How many questions on exam topic etId are no longer the approved version in
' use (the Replaced note: a newer version was approved since the draw, or it
' was retired).
' Errors: none of its own; a data error passes up.
Public Function ReplacedCount(ByVal etId As Long) As Long
    ReplacedCount = CountOf("SELECT Count(*) FROM tblExamQuestions AS eq INNER JOIN tblQuestions AS q ON eq.QuestionID=q.QuestionID " & _
        "WHERE eq.ExamTopicID=" & etId & " AND NOT (q.IsCurrent AND q.StatusID=" & ApprovedStatusID() & ")")
End Function

' How many questions on exam topic etId were given lately (keys: the
' RecentKeysList for its exam).
' Errors: none of its own; a data error passes up.
Public Function GivenLatelyCount(ByVal etId As Long, ByVal keys As String) As Long
    GivenLatelyCount = CountOf("SELECT Count(*) FROM tblExamQuestions AS eq INNER JOIN tblQuestions AS q ON eq.QuestionID=q.QuestionID " & _
        "WHERE eq.ExamTopicID=" & etId & " AND q.QuestionKey IN (" & keys & ")")
End Function

' Why question newQid can't replace exam question eqId, or "": it must be an
' approved question in use of the same topic, other than the one in the slot,
' with no version elsewhere on the exam (the slot's own question may come
' back as its newer version), and the swap may not take the topic's
' questions given lately over what the overlap rule allows (nor over how
' many it has now, when it is over already), nor a topic of a later exam
' (LaterExamProblem).
' Reads only; also called inside SwapQuestion's transaction.
' Errors: none of its own; a data error passes up.
Public Function SwapProblem(ByVal eqId As Long, ByVal newQid As Long) As String
    Dim rs As DAO.Recordset, examId As Long, etId As Long, topicId As Long, n As Long, oldKey As Variant, newKey As Variant
    Dim keys As String, have As Long, after As Long, allowed As Long, oldQid As Long
    Set rs = CurrentDb.OpenRecordset("SELECT eq.ExamID, eq.ExamTopicID, eq.QuestionID, et.TopicID, et.QuestionCount, q.QuestionKey " & _
        "FROM (tblExamQuestions AS eq INNER JOIN tblExamTopics AS et ON eq.ExamTopicID=et.ExamTopicID) " & _
        "INNER JOIN tblQuestions AS q ON eq.QuestionID=q.QuestionID WHERE eq.ExamQuestionID=" & eqId, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        SwapProblem = "That question is no longer on the exam. The list now shows what is."
        Exit Function
    End If
    examId = rs!ExamID: etId = rs!ExamTopicID: oldQid = rs!QuestionID: topicId = rs!TopicID: n = rs!QuestionCount
    oldKey = rs!QuestionKey
    rs.Close
    If newQid = 0 Then SwapProblem = "Pick the question to put in its place.": Exit Function
    If newQid = oldQid Then SwapProblem = "That is the question in its place now. Pick another.": Exit Function
    If DCount("*", "tblQuestions", "QuestionID=" & newQid & " AND TopicID=" & topicId & " AND IsCurrent AND StatusID=" & _
            ApprovedStatusID()) = 0 Then
        SwapProblem = "That question isn't an approved question of this topic now. It may have been replaced or retired."
        Exit Function
    End If
    newKey = DLookup("QuestionKey", "tblQuestions", "QuestionID=" & newQid)
    If CountOf("SELECT Count(*) FROM tblExamQuestions AS eq INNER JOIN tblQuestions AS q ON eq.QuestionID=q.QuestionID " & _
            "WHERE eq.ExamID=" & examId & " AND eq.ExamQuestionID<>" & eqId & " AND q.QuestionKey=" & Nz(newKey, 0)) > 0 Then
        SwapProblem = "That question is on this exam already."
        Exit Function
    End If
    keys = RecentKeysList(examId)
    have = GivenLatelyCount(etId, keys)
    after = have - IIf(InKeyList(keys, oldKey), 1, 0) + IIf(InKeyList(keys, newKey), 1, 0)
    allowed = RepeatsAllowed(n)
    If after > allowed And after > have Then
        SwapProblem = "That question was on one of the " & OverlapWindow() & " exams generated before this " & _
            "one. With it, " & after & " of this topic's " & n & " questions would have been given lately; the overlap rule " & _
            "allows " & allowed & ". Pick one that wasn't given lately, or swap out one that was."
        Exit Function
    End If
    SwapProblem = LaterExamProblem(examId, newKey)
End Function

' Why putting question key newKey on exam examId would take a topic of an
' exam generated after it over the overlap rule, or "". The next
' OverlapWindowExams exams count this one in their window, so a question one
' of them holds becomes given lately for it (unless it was already). The
' question this swap takes out is still counted for them: that can only
' refuse a swap, never let one through over the rule.
' Errors: none of its own; a data error passes up.
Private Function LaterExamProblem(ByVal examId As Long, ByVal newKey As Variant) As String
    Dim rs As DAO.Recordset, w As Long, ids As String, keys As String, have As Long, n As Long, allowed As Long
    w = OverlapWindow()
    If w <= 0 Or IsNull(newKey) Then Exit Function
    ids = ColumnList("SELECT TOP " & w & " ExamID FROM tblExams WHERE ExamID>" & examId & " ORDER BY ExamID")
    If Len(ids) = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT eq.ExamID, eq.ExamTopicID, et.QuestionCount, e.RecipientName " & _
        "FROM ((tblExamQuestions AS eq INNER JOIN tblQuestions AS q ON eq.QuestionID=q.QuestionID) " & _
        "INNER JOIN tblExamTopics AS et ON eq.ExamTopicID=et.ExamTopicID) INNER JOIN tblExams AS e ON eq.ExamID=e.ExamID " & _
        "WHERE q.QuestionKey=" & CLng(newKey) & " AND eq.ExamID IN (" & ids & ") ORDER BY eq.ExamID", dbOpenSnapshot)
    Do While Not rs.EOF
        keys = RecentKeysList(rs!ExamID)
        If Not InKeyList(keys, newKey) Then
            have = GivenLatelyCount(rs!ExamTopicID, keys)
            n = rs!QuestionCount
            allowed = RepeatsAllowed(n)
            If have + 1 > allowed Then
                LaterExamProblem = "That question is on the exam for " & rs!RecipientName & ", generated after this one. With it " & _
                    "here too, " & (have + 1) & " of that exam's " & n & " questions in this topic would have been given lately; " & _
                    "the overlap rule allows " & allowed & ". Pick another question."
                Exit Do
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
End Function

' Put question newQid on the exam in place of exam question eqId, as the
' signed-in SME, while the topic still has change number stamp (the
' TopicChangeNo the page showed). Returns "" with doneText for the toast, or
' why it wasn't swapped (SwapProblem, or someone changed the topic since).
' Claiming the topic (ChangeNo + 1) is the transaction's first write, so two
' swaps of one topic run one after the other and the rule is checked again
' with nothing else changing it. Writes: the slot's QuestionID and
' WasSwapped, the topic's ChangeNo, and an audit row on the exam (Swapped,
' at the SME step, "<topic>: Q12 v1 out, Q15 v2 in.").
' Errors: ERR_NOT_ALLOWED when the person may not change the topic now
'   (RequireReviewer); a data error rolls back and passes up.
Public Function SwapQuestion(ByVal eqId As Long, ByVal newQid As Long, ByVal stamp As Long, ByRef doneText As String) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, etId As Long, examId As Long, oldQid As Long, n As Long
    Dim oldRef As String, newRef As String, topic As String, num As Long, desc As String, src As String
    doneText = ""
    etId = Nz(DLookup("ExamTopicID", "tblExamQuestions", "ExamQuestionID=" & eqId), 0)
    If etId = 0 Then SwapQuestion = "That question is no longer on the exam. The list now shows what is.": Exit Function
    RequireReviewer etId
    SwapQuestion = SwapProblem(eqId, newQid)
    If Len(SwapQuestion) > 0 Then Exit Function
    examId = Nz(DLookup("ExamID", "tblExamTopics", "ExamTopicID=" & etId), 0)
    oldQid = Nz(DLookup("QuestionID", "tblExamQuestions", "ExamQuestionID=" & eqId), 0)
    oldRef = RefOf(oldQid)
    newRef = RefOf(newQid)
    topic = Nz(DLookup("TopicName", "tblTopics", "TopicID=" & Nz(DLookup("TopicID", "tblExamTopics", "ExamTopicID=" & etId), 0)), "")
    DBEngine.Idle dbRefreshCache            ' the checks below read other copies' latest commits
    Set ws = DBEngine.Workspaces(0)
    On Error GoTo Fail
    ws.BeginTrans: inTrans = True
    ' Swaps and Generates change what was given lately across exams, so they
    ' take turns: a swap on another exam, or a Generate, waits here.
    LockOverlapWindow
    ' Claim the topic: only while nobody has swapped or signed it off since the
    ' SME's list was loaded. A second swap or a sign-off at the same moment
    ' changes no row here and is told to look again.
    n = KitExecute("UPDATE tblExamTopics SET ChangeNo=Nz(ChangeNo,0)+1 WHERE ExamTopicID=" & etId & _
        " AND SignedOffByPersonID Is Null AND Nz(ChangeNo,0)=" & stamp)
    If n = 0 Then
        ws.Rollback: inTrans = False
        SwapQuestion = "Another SME swapped a question on this topic or signed it off a moment ago. The list now shows where it stands."
        Exit Function
    End If
    ' With the topic claimed nothing else can change it: check the rule again.
    SwapQuestion = SwapProblem(eqId, newQid)
    If Len(SwapQuestion) > 0 Then ws.Rollback: inTrans = False: Exit Function
    ' Only while it still holds the question shown.
    n = KitExecute("UPDATE tblExamQuestions SET QuestionID=" & newQid & ", WasSwapped=True WHERE ExamQuestionID=" & eqId & _
        " AND QuestionID=" & oldQid)
    If n = 0 Then
        ws.Rollback: inTrans = False
        SwapQuestion = "Another SME swapped this question or signed off the topic a moment ago. The list now shows where it stands."
        Exit Function
    End If
    AddAudit ITEM_EXAM, examId, DLookup("VersionNo", "tblExams", "ExamID=" & examId), AU_SWAPPED, RoleID("SME"), _
        topic & ": " & oldRef & " out, " & newRef & " in."
    ws.CommitTrans: inTrans = False
    doneText = "Swapped. " & newRef & " is on the exam in place of " & oldRef & "."
    Exit Function
Fail:
    num = Err.Number: desc = Err.Description: src = Err.Source
    If inTrans Then ws.Rollback
    If num = 3022 Then                      ' uxExamQ: someone put it on the exam a moment ago
        SwapQuestion = "Another SME put that question on this exam a moment ago. Pick another."
        Exit Function
    End If
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modExams.SwapQuestion", eqId & " " & newQid
End Function

' "Q12 v2" for question qid.
Private Function RefOf(ByVal qid As Long) As String
    RefOf = QuestionRef(DLookup("QuestionKey", "tblQuestions", "QuestionID=" & qid), DLookup("VersionNo", "tblQuestions", "QuestionID=" & qid))
End Function

' Sign off exam topic etId as the signed-in SME, while it still has change
' number stamp (the TopicChangeNo the page showed, so the sign-off covers
' exactly the questions the SME saw): the topic is complete, the person's
' task is Done and the other SMEs' tasks on it are Closed. When it was the
' exam's last topic, the exam goes to the DH for approval. Returns "" with
' doneText for the toast, or why it wasn't signed off (another SME swapped
' a question or signed it off since, or a question was replaced since the
' draw: the SME swaps it first, ReplacedCount). Writes: the topic's sign-off and
' ChangeNo, the topic's open tasks (Done or Closed), an audit row on the
' exam (Signed off, at the SME step, the topic's name) and, for the last
' topic, the exam's status and step with its own audit row (Submitted).
' Errors: ERR_NOT_ALLOWED when the person may not sign it off now
'   (RequireReviewer); a data error rolls back and passes up.
Public Function SignOffTopic(ByVal etId As Long, ByVal stamp As Long, ByRef doneText As String) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, pid As Long, examId As Long, n As Long, nLeft As Long, ver As Variant
    Dim topic As String, sent As Boolean, reopened As Boolean, num As Long, desc As String, src As String
    RequireReviewer etId
    doneText = ""
    DBEngine.Idle dbRefreshCache            ' so a question replaced a moment ago in another copy counts
    pid = SignedInPersonID()
    examId = Nz(DLookup("ExamID", "tblExamTopics", "ExamTopicID=" & etId), 0)
    ver = DLookup("VersionNo", "tblExams", "ExamID=" & examId)
    topic = Nz(DLookup("TopicName", "tblTopics", "TopicID=" & Nz(DLookup("TopicID", "tblExamTopics", "ExamTopicID=" & etId), 0)), "")
    Set ws = DBEngine.Workspaces(0)
    On Error GoTo Fail
    ws.BeginTrans: inTrans = True
    ' Take the exam's row first, while it is still in SME review: two sign-offs
    ' of one exam's topics then run one after the other, so the later one sees
    ' the earlier one's topic signed off.
    If KitExecute("UPDATE tblExams SET VersionNo=VersionNo WHERE ExamID=" & examId & " AND StatusID IN (" & _
            SmeReviewStatusIDs() & ")") = 0 Then
        ws.Rollback: inTrans = False
        SignOffTopic = "The exam is past SME review, so this topic can't be signed off now."
        Exit Function
    End If
    ' Every question must still be the approved version in use: one replaced
    ' or retired since the draw is swapped first.
    n = ReplacedCount(etId)
    If n > 0 Then
        ws.Rollback: inTrans = False
        SignOffTopic = n & IIf(n = 1, " question on this topic was", " questions on this topic were") & " replaced by a newer " & _
            "version or retired since the exam was drawn. Swap " & IIf(n = 1, "it", "them") & " first (Swap lists the newer " & _
            "version), then sign off."
        Exit Function
    End If
    ' One sign-off wins, and only on the questions shown: only while no one
    ' has signed it off or swapped a question since the page loaded it.
    n = KitExecute("UPDATE tblExamTopics SET SignedOffByPersonID=" & pid & ", SignedOffOn=Now(), ChangeNo=Nz(ChangeNo,0)+1 " & _
        "WHERE ExamTopicID=" & etId & " AND SignedOffByPersonID Is Null AND Nz(ChangeNo,0)=" & stamp)
    If n = 0 Then
        ws.Rollback: inTrans = False
        SignOffTopic = "Another SME swapped a question on this topic or signed it off a moment ago. The list now shows where " & _
            "it stands; check the questions before you sign off."
        Exit Function
    End If
    KitExecute "UPDATE tblTasks SET TaskStatus=IIf(AssignedPersonID=" & pid & "," & Q(TS_DONE) & "," & Q(TS_CLOSED) & "), " & _
        "CompletedOn=Now(), CompletedByPersonID=" & pid & " WHERE ExamTopicID=" & etId & " AND TaskStatus=" & Q(TS_OPEN)
    AddAudit ITEM_EXAM, examId, ver, AU_SIGNED_OFF, RoleID("SME"), topic
    sent = SendWhenSignedOff(examId, reopened)
    ws.CommitTrans: inTrans = False
    If sent Then NoticeExamAtStep examId    ' queues the DHs' emails; never raises
    If reopened Then NoticeReviews examId   ' queues the SMEs' emails; never raises
    nLeft = SendAfterSignOff(examId)        ' and a last topic signed off at the same moment as this one; never raises
    If reopened Then
        doneText = "Signed off. Another topic had a question replaced or retired since the draw, so it went back to its " & _
            "SMEs to swap before the exam goes to the DH."
    ElseIf nLeft < 0 Then
        doneText = "Signed off."
    ElseIf nLeft = 0 Then
        doneText = "Signed off. That was the exam's last topic, so it goes to the DH for approval."
    Else
        doneText = "Signed off. " & nLeft & IIf(nLeft = 1, " more topic on this exam waits", " more topics on this exam wait") & _
            " for their SMEs."
    End If
    Exit Function
Fail:
    num = Err.Number: desc = Err.Description: src = Err.Source
    If inTrans Then ws.Rollback
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modExams.SignOffTopic", etId
End Function

' Send exam examId to the DH for approval (In Review, at the DH's step) when
' every topic is signed off and it is still in SME review or sent back
' (SmeReviewStatusIDs). A signed-off topic that now holds a question
' replaced or retired since the draw goes back to its SMEs instead
' (ReopenReplacedTopics), so the exam never reaches the DH with one; then
' reopened returns True (the caller sends those SMEs' emails once it has
' committed). Safe to call twice: only the call that moves it writes the
' audit row (Submitted), in the name of whoever signed off the last topic,
' at that time, whoever's copy made the move. Call it inside a transaction,
' so the move and its audit rows are written together. Returns True when
' this call moved it (the caller sends the DHs' emails once it has
' committed).
' Errors: a data error passes up (the caller rolls back).
Private Function SendWhenSignedOff(ByVal examId As Long, Optional ByRef reopened As Boolean) As Boolean
    Dim n As Long, sme As String, rs As DAO.Recordset, who As Long, whenOn As Variant
    sme = SmeReviewStatusIDs()
    ' Take the exam's row first. A send-back writes the same row first, so it
    ' runs wholly before this (the count then sees its reopened topics) or
    ' wholly after. In SignOffTopic the row is held already.
    reopened = False
    If KitExecute("UPDATE tblExams SET VersionNo=VersionNo WHERE ExamID=" & examId & " AND StatusID IN (" & sme & ")") = 0 Then Exit Function
    If CountOf("SELECT Count(*) FROM tblExamTopics WHERE ExamID=" & examId & " AND SignedOffByPersonID Is Null") > 0 Then Exit Function
    If ReopenReplacedTopics(examId) > 0 Then reopened = True: Exit Function
    ' The condition again in the move itself, so the count and the move can't disagree.
    n = KitExecute("UPDATE tblExams SET StatusID=" & ExamStatusID(ES_REVIEW) & ", CurrentStepRoleID=" & RoleID("DH") & _
        " WHERE ExamID=" & examId & " AND StatusID IN (" & sme & ") AND NOT EXISTS (SELECT * FROM tblExamTopics AS et " & _
        "WHERE et.ExamID=" & examId & " AND et.SignedOffByPersonID Is Null)")
    If n = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 SignedOffByPersonID, SignedOffOn FROM tblExamTopics WHERE ExamID=" & examId & _
        " ORDER BY SignedOffOn DESC, ExamTopicID DESC", dbOpenSnapshot)
    If Not rs.EOF Then who = Nz(rs!SignedOffByPersonID, 0): whenOn = rs!SignedOffOn
    rs.Close
    AddAudit ITEM_EXAM, examId, DLookup("VersionNo", "tblExams", "ExamID=" & examId), AU_SUBMITTED, RoleID("DH"), _
        "Every topic signed off", who, whenOn
    SendWhenSignedOff = True
End Function

' Exam examId's signed-off topics that hold a question replaced or retired
' since the draw, sent back to their SMEs while the exam is in SME review:
' the sign-off is cleared, the change number moves on (a page showing the
' old questions can't sign it off), its active SMEs get new open review
' tasks, and the exam gets an audit row (Reopened, the topic's name). A topic
' with no active SME gets no task now; the audit row says so, and the sweep
' (ReissueReviews) gives it tasks once the topic has an SME again.
' Returns how many topics went back. Call it inside a transaction.
' Errors: a data error passes up (the caller rolls back).
Private Function ReopenReplacedTopics(ByVal examId As Long) As Long
    Dim rs As DAO.Recordset, ver As Variant, n As Long, smes As String, note As String
    If KitExecute("UPDATE tblExams SET VersionNo=VersionNo WHERE ExamID=" & examId & " AND StatusID IN (" & _
            SmeReviewStatusIDs() & ")") = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT et.ExamTopicID, et.TopicID, tp.TopicName FROM tblExamTopics AS et " & _
        "LEFT JOIN tblTopics AS tp ON et.TopicID=tp.TopicID WHERE et.ExamID=" & examId & _
        " AND Not et.SignedOffByPersonID Is Null AND EXISTS (SELECT * FROM tblExamQuestions AS eq INNER JOIN tblQuestions AS q " & _
        "ON eq.QuestionID=q.QuestionID WHERE eq.ExamTopicID=et.ExamTopicID AND NOT (q.IsCurrent AND q.StatusID=" & _
        ApprovedStatusID() & "))", dbOpenSnapshot)
    ver = DLookup("VersionNo", "tblExams", "ExamID=" & examId)
    Do While Not rs.EOF
        ' Only while it is still signed off, so it is reopened (and its tasks made) once.
        If KitExecute("UPDATE tblExamTopics SET SignedOffByPersonID=Null, SignedOffOn=Null, ChangeNo=Nz(ChangeNo,0)+1 " & _
                "WHERE ExamTopicID=" & rs!ExamTopicID & " AND Not SignedOffByPersonID Is Null") > 0 Then
            smes = TopicReviewers(rs!TopicID)
            AddReviewTasks examId, rs!ExamTopicID, smes
            note = Nz(rs!TopicName, "(topic deleted)") & ": a question was replaced or retired since the draw"
            If Len(smes) = 0 Then note = note & ". It has no active SME, so add one on the Topics page to review it"
            AddAudit ITEM_EXAM, examId, ver, AU_REOPENED, RoleID("SME"), note
            n = n + 1
        End If
        rs.MoveNext
    Loop
    rs.Close
    ReopenReplacedTopics = n
End Function

' Open SME review tasks on exam topic etId of exam examId, one for each
' person in smes (ids joined with commas; "" makes none).
' Errors: a data error passes up.
Private Sub AddReviewTasks(ByVal examId As Long, ByVal etId As Long, ByVal smes As String)
    Dim v As Variant
    For Each v In Split(smes, ",")
        KitExecute "INSERT INTO tblTasks (TaskType, ItemType, ItemID, ExamTopicID, AssignedPersonID, TaskStatus, CreatedOn) " & _
            "VALUES (" & Q(TASK_SME_REVIEW) & ", " & Q(ITEM_EXAM) & ", " & examId & ", " & etId & ", " & CLng(v) & ", " & _
            Q(TS_OPEN) & ", Now())"
    Next
End Sub

' Exam examId's topics that wait on a sign-off, given an open review task
' for each of the topic's active SMEs (TopicReviewers) who holds none on it:
' a topic reopened when it had no active SME, or whose SMEs have changed
' since (the tasks of SMEs who left stay as they are). Checked first without
' a lock, then again under the exam's row lock, so two copies make each task
' once. Returns how many tasks were made; runs in its own transaction.
' Never raises: a failure (another copy holding the exam a moment) is logged
' and the next sweep tries again.
Private Function ReissueReviews(ByVal examId As Long) As Long
    Dim ws As DAO.Workspace, inTrans As Boolean, n As Long
    On Error GoTo Fail
    If MissingReviewTasks(examId, False) = 0 Then Exit Function    ' nothing to do: no lock taken
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    If KitExecute("UPDATE tblExams SET VersionNo=VersionNo WHERE ExamID=" & examId & " AND StatusID IN (" & _
            SmeReviewStatusIDs() & ")") > 0 Then n = MissingReviewTasks(examId, True)
    ws.CommitTrans: inTrans = False
    ReissueReviews = n
    Exit Function
Fail:
    If inTrans Then ws.Rollback
    LogError "modExams.ReissueReviews", "exam " & examId
End Function

' How many review tasks exam examId's unsigned topics lack: one for each of
' a topic's active SMEs who holds no open task on it. With doAdd, makes them.
' Errors: a data error passes up.
Private Function MissingReviewTasks(ByVal examId As Long, ByVal doAdd As Boolean) As Long
    Dim rs As DAO.Recordset, v As Variant, n As Long
    Set rs = CurrentDb.OpenRecordset("SELECT ExamTopicID, TopicID FROM tblExamTopics WHERE ExamID=" & examId & _
        " AND SignedOffByPersonID Is Null", dbOpenSnapshot)
    Do While Not rs.EOF
        For Each v In Split(TopicReviewers(rs!TopicID), ",")
            If CountOf("SELECT Count(*) FROM tblTasks WHERE ExamTopicID=" & rs!ExamTopicID & " AND AssignedPersonID=" & CLng(v) & _
                    " AND TaskStatus=" & Q(TS_OPEN)) = 0 Then
                If doAdd Then AddReviewTasks examId, rs!ExamTopicID, CStr(CLng(v))
                n = n + 1
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    MissingReviewTasks = n
End Function

' After a sign-off of exam examId has committed: read past Jet's cache, so
' another copy's sign-off of a moment ago counts, and send the exam to the DH
' when every topic is signed off now (SendWhenSignedOff, in its own
' transaction). Returns how many topics still wait, or -1 when this failed.
' Never raises: the sign-off is saved whatever happens here, so a failure is
' logged and the sweep (SendSignedOffExams) moves the exam later.
Private Function SendAfterSignOff(ByVal examId As Long) As Long
    Dim ws As DAO.Workspace, inTrans As Boolean, sent As Boolean, reopened As Boolean
    On Error GoTo Fail
    DBEngine.Idle dbRefreshCache
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    sent = SendWhenSignedOff(examId, reopened)
    ws.CommitTrans: inTrans = False
    If sent Then NoticeExamAtStep examId    ' queues the DHs' emails; never raises
    If reopened Then NoticeReviews examId   ' queues the SMEs' emails; never raises
    SendAfterSignOff = CountOf("SELECT Count(*) FROM tblExamTopics WHERE ExamID=" & examId & " AND SignedOffByPersonID Is Null")
    Exit Function
Fail:
    If inTrans Then ws.Rollback
    LogError "modExams.SendAfterSignOff", "exam " & examId
    SendAfterSignOff = -1
End Function

' Send every exam still in SME review whose topics are all signed off to the
' DH. Idempotent and guarded by status, so any page may run it at any time
' (My reviews does when it opens; the Exams list should too): it moves on an
' exam that two last sign-offs at the same moment, or a failed step after a
' sign-off, left behind. The audit row names whoever signed off the last
' topic, at the time of that sign-off (SendWhenSignedOff). Then it gives
' each topic that waits on a sign-off a review task for every active SME who
' holds none on it (ReissueReviews). It only queues the emails (the DHs',
' and the SMEs' for reopened or reissued topics): the page sends them once
' it has loaded (SendQueuedNotices).
' Never raises: a failure is logged.
Public Sub SendSignedOffExams()
    Dim ids As String, v As Variant, lastTask As Long
    On Error GoTo Fail
    If Not DataLinked() Then Exit Sub
    DBEngine.Idle dbRefreshCache
    ids = ColumnList("SELECT e.ExamID FROM tblExams AS e WHERE e.StatusID IN (" & SmeReviewStatusIDs() & ") " & _
        "AND EXISTS (SELECT * FROM tblExamTopics AS et WHERE et.ExamID=e.ExamID) " & _
        "AND NOT EXISTS (SELECT * FROM tblExamTopics AS et WHERE et.ExamID=e.ExamID AND et.SignedOffByPersonID Is Null)")
    For Each v In Split(ids, ",")
        SendAfterSignOff CLng(v)
    Next
    ' Topics that wait on a sign-off get a task for each active SME who holds
    ' none, so a topic whose SMEs left or changed isn't stuck (ReissueReviews).
    ids = ColumnList("SELECT DISTINCT et.ExamID FROM tblExamTopics AS et INNER JOIN tblExams AS e ON et.ExamID=e.ExamID " & _
        "WHERE e.StatusID IN (" & SmeReviewStatusIDs() & ") AND et.SignedOffByPersonID Is Null")
    For Each v In Split(ids, ",")
        lastTask = Nz(DMax("TaskID", "tblTasks"), 0)
        If ReissueReviews(CLng(v)) > 0 Then NoticeReviews CLng(v), lastTask ' queues the new tasks' emails; never raises
    Next
    Exit Sub
Fail:
    LogError "modExams.SendSignedOffExams"
End Sub
