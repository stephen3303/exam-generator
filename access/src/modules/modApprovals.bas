Option Compare Database
Option Explicit

' Exam Generator (Access edition) - question approvals (build phase 6): the
' approval chain, who may act at each step, approving, rejecting, and the
' text the Approvals page shows.
'
' A submitted question waits at one step of the chain at a time
' (tblQuestions.CurrentStepRoleID): the DH, then the XO, then the SM, in
' tblRoles.ChainOrder. Each approver's queue is the questions in review
' whose step is a role they hold. Approve moves the question to the next
' step; the last step's approval makes it Approved and current, and, for a
' new version, retires the version it replaces, in one transaction, so the
' bank never has two current versions of a question or none. Reject sends
' it back to its author with a reason (required): Rejected, no step. Its
' author, or an SME on its topic, changes it and submits it again, and it
' starts again at the DH.
'
' Who may approve or reject is checked in the data file at the moment of the
' write (RequireApprover), not from the roles read at sign-in: an active
' person who holds the step's role now and may still use the app, and who
' did not write the question. Every approval, rejection and retirement
' writes an audit row (modQuestions.AddAudit); the reason for a rejection is
' its Comments.
'
' Owner: question approvals (build phase 6). Depends on: modAdmin,
' modErrors, modGlobals, modQuestions, modSignIn; the kit's modKitLocks,
' modPersonnel.

Public Const AU_APPROVED As String = "Approved"
Public Const AU_REJECTED As String = "Rejected"

Public Const MAX_REASON As Long = 2000          ' longest rejection reason

' ---------------------------------------------------------------------------
' The chain
' ---------------------------------------------------------------------------

' The step after role stepRoleID in the chain (tblRoles.ChainOrder), or 0
' when stepRoleID is the last step.
Public Function NextStepRoleID(ByVal stepRoleID As Long) As Long
    Dim ord As Variant, nxt As Variant
    ord = DLookup("ChainOrder", "tblRoles", "RoleID=" & stepRoleID)
    If IsNull(ord) Then Exit Function
    nxt = DMin("ChainOrder", "tblRoles", "ChainOrder>" & ord)
    If IsNull(nxt) Then Exit Function
    NextStepRoleID = Nz(DLookup("RoleID", "tblRoles", "ChainOrder=" & nxt), 0)
End Function

' The step's place in the chain, 1 for the SME: the Approvals page's
' Stepper (SME, DH, XO, SM) shows it.
Public Function StepNumber(ByVal stepRoleID As Long) As Long
    StepNumber = Nz(DLookup("ChainOrder", "tblRoles", "RoleID=" & stepRoleID), 0)
End Function

Public Function RoleCodeOf(ByVal roleId As Variant) As String
    If IsNull(roleId) Then Exit Function
    RoleCodeOf = Nz(DLookup("RoleCode", "tblRoles", "RoleID=" & CLng(roleId)), "")
End Function

' ---------------------------------------------------------------------------
' Who may approve (checked at the write)
' ---------------------------------------------------------------------------

' Why the signed-in person may not approve or reject question qid now, or
' "" when they may. Reads the data file, so a role taken away since sign-in,
' or a step someone else moved on, counts at once.
Public Function ApproveProblem(ByVal qid As Long) As String
    Dim rs As DAO.Recordset, pid As Long, st As String, stepId As Long, code As String, author As Long
    If Not DataLinked() Then ApproveProblem = "This copy has no data file.": Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT StatusID, CurrentStepRoleID, AuthorPersonID FROM tblQuestions WHERE QuestionID=" & qid, _
        dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        ApproveProblem = "This question was deleted."
        Exit Function
    End If
    st = QStatusName(Nz(rs!StatusID, 0))
    stepId = Nz(rs!CurrentStepRoleID, 0)
    author = Nz(rs!AuthorPersonID, 0)
    rs.Close
    If st <> QS_REVIEW Or stepId = 0 Then
        ApproveProblem = "It is " & LCase$(st) & " now, so it isn't waiting for an approval."
        Exit Function
    End If
    code = RoleCodeOf(stepId)
    pid = SignedInPersonID()
    If pid = 0 Or Not Nz(DLookup("IsActive", "tblPeople", "PersonID=" & pid), False) Or Not PersonHasRole(pid, code) Then
        ApproveProblem = "It is waiting on the " & code & ", and you aren't the " & code & " now. Your role may have been " & _
            "changed since you signed in."
        Exit Function
    End If
    If Not PersonCanUse(pid, APP_CODE) Then
        ApproveProblem = "Your access to the Exam Generator has been turned off in Personnel."
        Exit Function
    End If
    If pid = author Then
        ApproveProblem = "You wrote this question, so another " & code & " has to approve or reject it."
    End If
End Function

' Raise ERR_NOT_ALLOWED unless the signed-in person may approve or reject
' question qid now (ApproveProblem). Called first by ApproveQuestion and
' RejectQuestion. No side effects.
Public Sub RequireApprover(ByVal qid As Long)
    Dim problem As String
    problem = ApproveProblem(qid)
    If Len(problem) > 0 Then Err.Raise ERR_NOT_ALLOWED, "modApprovals.RequireApprover", problem
End Sub

' True when the signed-in person may approve or reject question qid now:
' for the page, to grey out what they can't do.
Public Function CanApprove(ByVal qid As Long) As Boolean
    CanApprove = (Len(ApproveProblem(qid)) = 0)
End Function

' ---------------------------------------------------------------------------
' Approve and reject
' ---------------------------------------------------------------------------

' Approve question qid at its step, as the signed-in person. Before the last
' step it moves on to the next approver; at the last it becomes Approved and
' current, and any other current version of it is retired (Retired,
' RetiredOn, not current), all in one transaction. Returns "" and what
' happened, for the toast, in doneText; or why it can't be approved now
' (someone else acted on it first).
' Errors: ERR_NOT_ALLOWED when the person may not approve it (RequireApprover).
Public Function ApproveQuestion(ByVal qid As Long, ByRef doneText As String) As String
    Dim ws As DAO.Workspace, rs As DAO.Recordset, inTrans As Boolean, stepId As Long, nextId As Long
    Dim key As Long, ver As Long, n As Long, oldVers As String, num As Long, desc As String, src As String
    RequireApprover qid
    doneText = ""
    stepId = Nz(DLookup("CurrentStepRoleID", "tblQuestions", "QuestionID=" & qid), 0)
    nextId = NextStepRoleID(stepId)
    key = Nz(DLookup("QuestionKey", "tblQuestions", "QuestionID=" & qid), 0)
    ver = Nz(DLookup("VersionNo", "tblQuestions", "QuestionID=" & qid), 0)
    Set ws = DBEngine.Workspaces(0)
    On Error GoTo Fail
    ws.BeginTrans: inTrans = True
    ' Only while it still waits at this step: someone may have acted on it since.
    If nextId <> 0 Then
        n = KitExecute("UPDATE tblQuestions SET CurrentStepRoleID=" & nextId & StillAtStep(qid, stepId))
    Else
        n = KitExecute("UPDATE tblQuestions SET StatusID=" & QStatusID(QS_APPROVED) & ", CurrentStepRoleID=Null, IsCurrent=True" & _
            StillAtStep(qid, stepId))
    End If
    If n = 0 Then
        ws.Rollback: inTrans = False
        ApproveQuestion = "Someone else approved or rejected it a moment ago. The list now shows where it stands."
        Exit Function
    End If
    AddAudit ITEM_QUESTION, qid, ver, AU_APPROVED, stepId
    If nextId = 0 Then
        ' The version it replaces stops being drawn now, not before.
        Set rs = CurrentDb.OpenRecordset("SELECT QuestionID, VersionNo FROM tblQuestions WHERE QuestionKey=" & key & _
            " AND IsCurrent AND QuestionID<>" & qid & " ORDER BY VersionNo", dbOpenSnapshot)
        Do While Not rs.EOF
            KitExecute "UPDATE tblQuestions SET IsCurrent=False, StatusID=" & QStatusID(QS_RETIRED) & ", RetiredOn=Now() " & _
                "WHERE QuestionID=" & rs!QuestionID
            AddAudit ITEM_QUESTION, rs!QuestionID, rs!VersionNo, AU_RETIRED, , "Replaced by version " & ver
            oldVers = oldVers & IIf(Len(oldVers) > 0, ", ", "") & "v" & rs!VersionNo
            rs.MoveNext
        Loop
        rs.Close
    End If
    ws.CommitTrans: inTrans = False
    If nextId <> 0 Then
        doneText = "Approved. " & QuestionRef(key, ver) & " goes to the " & RoleCodeOf(nextId) & " next."
    ElseIf Len(oldVers) > 0 Then
        doneText = "Approved. " & QuestionRef(key, ver) & " is in use now; " & oldVers & " is retired."
    Else
        doneText = "Approved. " & QuestionRef(key, ver) & " is in the bank and can be drawn for exams."
    End If
    Exit Function
Fail:
    num = Err.Number: desc = Err.Description: src = Err.Source
    If inTrans Then ws.Rollback
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modApprovals.ApproveQuestion", qid
End Function

' Reject question qid at its step, as the signed-in person, with reason (in
' the audit row). It goes back to its author: Rejected, at no step. Returns
' "" or why it can't be rejected now (no reason, or someone else acted on it
' first).
' Errors: ERR_NOT_ALLOWED when the person may not reject it (RequireApprover).
Public Function RejectQuestion(ByVal qid As Long, ByVal reason As Variant) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, stepId As Long, ver As Long, n As Long, r As String
    Dim num As Long, desc As String, src As String
    RequireApprover qid
    RejectQuestion = ReasonProblem(reason)
    If Len(RejectQuestion) > 0 Then Exit Function
    r = CleanText(reason)
    stepId = Nz(DLookup("CurrentStepRoleID", "tblQuestions", "QuestionID=" & qid), 0)
    ver = Nz(DLookup("VersionNo", "tblQuestions", "QuestionID=" & qid), 0)
    Set ws = DBEngine.Workspaces(0)
    On Error GoTo Fail
    ws.BeginTrans: inTrans = True
    n = KitExecute("UPDATE tblQuestions SET StatusID=" & QStatusID(QS_REJECTED) & ", CurrentStepRoleID=Null" & StillAtStep(qid, stepId))
    If n = 0 Then
        ws.Rollback: inTrans = False
        RejectQuestion = "Someone else approved or rejected it a moment ago, so it wasn't rejected."
        Exit Function
    End If
    AddAudit ITEM_QUESTION, qid, ver, AU_REJECTED, stepId, r
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    num = Err.Number: desc = Err.Description: src = Err.Source
    If inTrans Then ws.Rollback
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modApprovals.RejectQuestion", qid
End Function

' What is wrong with a rejection's reason ("" when fine): it is required,
' and at most MAX_REASON characters once tidied.
Public Function ReasonProblem(ByVal reason As Variant) As String
    Dim r As String
    r = Nz(CleanText(reason), "")
    If Len(r) = 0 Then
        ReasonProblem = "Give the reason, so its author knows what to change."
    ElseIf Len(r) > MAX_REASON Then
        ReasonProblem = "The reason can be at most " & MAX_REASON & " characters; it is " & Len(r) & "."
    End If
End Function

' " WHERE" for an update that only touches question qid while it is still
' in review at step stepId.
Private Function StillAtStep(ByVal qid As Long, ByVal stepId As Long) As String
    StillAtStep = " WHERE QuestionID=" & qid & " AND StatusID=" & QStatusID(QS_REVIEW) & " AND CurrentStepRoleID=" & stepId
End Function

' ---------------------------------------------------------------------------
' What the pages show
' ---------------------------------------------------------------------------

' The questions waiting for the signed-in person: in review at a step whose
' role they hold now, the longest waiting first. Columns: id, QuestionKey,
' ref, QuestionText, qsort, TopicName, author, kind (New or New version),
' step, since (when it was last submitted or moved on), AuthorPersonID.
' search keeps those whose text, topic or Q number has it.
Public Function ApprovalQueueSql(ByVal search As String) As String
    Dim where As String, t As String
    where = "q.StatusID=" & QStatusID(QS_REVIEW) & " AND q.CurrentStepRoleID IN (SELECT RoleID FROM tblPersonRoles WHERE PersonID=" & _
        SignedInPersonID() & ")"
    If Len(search) > 0 Then
        t = Q("*" & LikeLiteral(search) & "*")
        where = where & " AND (q.QuestionText LIKE " & t & " OR tp.TopicName LIKE " & t & " OR ('Q' & q.QuestionKey) LIKE " & t & ")"
    End If
    ApprovalQueueSql = "SELECT x.* FROM (SELECT q.QuestionID AS id, q.QuestionKey, 'Q' & q.QuestionKey & ' v' & q.VersionNo AS ref, " & _
        "Replace(Replace(Nz(q.QuestionText,''),Chr(13),' '),Chr(10),'') AS QuestionText, Left(Nz(q.QuestionText,''),255) AS qsort, " & _
        "tp.TopicName, " & PersonNameSql() & " AS author, IIf(Nz(q.ReplacesQuestionID,0)=0,'New','New version') AS kind, " & _
        "r.RoleCode AS step, q.AuthorPersonID, (SELECT Max(a.ActionOn) FROM tblAuditTrail AS a WHERE a.ItemType='" & ITEM_QUESTION & _
        "' AND a.ItemID=q.QuestionID) AS since " & _
        "FROM ((tblQuestions AS q LEFT JOIN tblTopics AS tp ON q.TopicID=tp.TopicID) " & _
        "LEFT JOIN " & PERSON_TABLE & " AS p ON q.AuthorPersonID=p.person_id) " & _
        "LEFT JOIN tblRoles AS r ON q.CurrentStepRoleID=r.RoleID WHERE " & where & ") AS x ORDER BY x.since, x.id"
End Function

' How many questions wait for the signed-in person (the Questions tab's count).
Public Function ApprovalCount() As Long
    If Not DataLinked() Then Exit Function
    ApprovalCount = DCount("*", "tblQuestions", "StatusID=" & QStatusID(QS_REVIEW) & " AND CurrentStepRoleID IN " & _
        "(SELECT RoleID FROM tblPersonRoles WHERE PersonID=" & SignedInPersonID() & ")")
End Function

' A person's name as the pickers show it, "" when they aren't listed.
Public Function PersonName(ByVal pid As Variant) As String
    Dim rs As DAO.Recordset
    If Nz(pid, 0) = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT " & PersonNameSql() & " AS nm FROM " & PERSON_TABLE & " AS p WHERE p.person_id=" & _
        CLng(pid), dbOpenSnapshot)
    If Not rs.EOF Then PersonName = Nz(rs!nm, "")
    rs.Close
End Function

' Who a rejected question goes back to, for the reject dialog: its author,
' and the SMEs on its topic, who may change it too.
Public Function ReturnsToText(ByVal qid As Long) As String
    Dim author As String, topic As String
    author = PersonName(DLookup("AuthorPersonID", "tblQuestions", "QuestionID=" & qid))
    topic = Nz(DLookup("TopicName", "tblTopics", "TopicID=" & Nz(DLookup("TopicID", "tblQuestions", "QuestionID=" & qid), 0)), "")
    ReturnsToText = IIf(Len(author) > 0, author & ", who wrote it", "Its author") & ". The SMEs on " & _
        IIf(Len(topic) > 0, topic, "its topic") & " can also change it and submit it again; it then starts again at the DH."
End Function

' Why question qid was last sent back, for its author: "The XO sent it back
' (Smith, J. at 14:05 on 27 Oct): <reason>". "" unless it is rejected now.
Public Function RejectionText(ByVal qid As Long) As String
    Dim id As Long, rs As DAO.Recordset
    If QStatusName(Nz(DLookup("StatusID", "tblQuestions", "QuestionID=" & qid), 0)) <> QS_REJECTED Then Exit Function
    id = Nz(DMax("AuditID", "tblAuditTrail", "ItemType='" & ITEM_QUESTION & "' AND ItemID=" & qid & " AND ActionType=" & Q(AU_REJECTED)), 0)
    If id = 0 Then Exit Function
    Set rs = CurrentDb.OpenRecordset("SELECT StepRoleID, Comments FROM tblAuditTrail WHERE AuditID=" & id, dbOpenSnapshot)
    If Not rs.EOF Then
        RejectionText = "The " & IIf(Len(RoleCodeOf(rs!StepRoleID)) > 0, RoleCodeOf(rs!StepRoleID), "approver") & " sent it back (" & _
            AuditWho(id) & "): " & Nz(rs!comments, "no reason given") & IIf(Right$(Nz(rs!comments, ""), 1) Like "[.!?]", "", ".")
    End If
    rs.Close
End Function

' Question qid as plain text for the approver to read: its text, parts with
' points and markings, answer options with the correct one marked, the
' reference answers and the references it cites.
Public Function QuestionPreview(ByVal qid As Long) As String
    Dim rs As DAO.Recordset, s As String, ans As String
    Set rs = CurrentDb.OpenRecordset("SELECT q.QuestionText, q.TotalPoints, q.ReferenceAnswer, c.Marking FROM tblQuestions AS q " & _
        "LEFT JOIN tblClassLevels AS c ON q.ClassLevelID=c.ClassLevelID WHERE q.QuestionID=" & qid, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Exit Function
    s = Nz(rs!QuestionText, "") & vbCrLf & "(" & PointsText(rs!TotalPoints) & IIf(IsNull(rs!Marking), "", ", " & rs!Marking) & ")"
    ans = Nz(rs!ReferenceAnswer, "")
    rs.Close
    s = s & OptionsText(qid, 0)
    Set rs = CurrentDb.OpenRecordset("SELECT p.PartID, p.PartLabel, p.PartText, p.Points, p.ReferenceAnswer, c.Marking " & _
        "FROM tblQuestionParts AS p LEFT JOIN tblClassLevels AS c ON p.ClassLevelID=c.ClassLevelID WHERE p.QuestionID=" & qid & _
        " ORDER BY p.SortOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & vbCrLf & vbCrLf & rs!PartLabel & ". " & Nz(rs!PartText, "") & vbCrLf & "(" & PointsText(rs!Points) & _
            IIf(IsNull(rs!Marking), "", ", " & rs!Marking) & ")" & OptionsText(qid, rs!PartID)
        If Len(Nz(rs!ReferenceAnswer, "")) > 0 Then s = s & vbCrLf & "Reference answer: " & rs!ReferenceAnswer
        rs.MoveNext
    Loop
    rs.Close
    If Len(ans) > 0 Then s = s & vbCrLf & vbCrLf & "Reference answer: " & ans
    Set rs = CurrentDb.OpenRecordset("SELECT RefID, Paragraph FROM tblQuestionRefs WHERE QuestionID=" & qid, dbOpenSnapshot)
    If Not rs.EOF Then s = s & vbCrLf & vbCrLf & "References:"
    Do While Not rs.EOF
        s = s & vbCrLf & "- " & RefText(rs!RefID) & IIf(IsNull(rs!Paragraph), "", ", para " & rs!Paragraph)
        rs.MoveNext
    Loop
    rs.Close
    QuestionPreview = s
End Function

' The answer options of question qid's part partId (0: of the question
' itself), one a line, the correct one marked; "" when it has none.
Private Function OptionsText(ByVal qid As Long, ByVal partId As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT OptionLabel, OptionText, IsCorrect FROM tblMCOptions WHERE QuestionID=" & qid & _
        IIf(partId = 0, " AND PartID Is Null", " AND PartID=" & partId) & " ORDER BY SortOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & vbCrLf & "   " & rs!OptionLabel & ") " & Nz(rs!OptionText, "") & IIf(Nz(rs!IsCorrect, False), "   [correct]", "")
        rs.MoveNext
    Loop
    rs.Close
    OptionsText = s
End Function
