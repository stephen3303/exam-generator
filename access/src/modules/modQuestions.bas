Option Compare Database
Option Explicit

' Exam Generator (Access edition) - the question bank (build phase 5):
' statuses and types, the editor's working copy, the save and submit rules,
' versions and audit rows.
'
' The editor works on a copy of one question. Its own fields are on the page
' (pgQuestionEdit); its parts, answer options and references are in three
' local tables of this front end (wq_part, wq_option, wq_ref), which the
' editor's tables show and its dialogs change. Nothing reaches the shared
' data file until Save, which checks every rule (QuestionProblems) and then
' writes the question, its parts, options and references in one transaction
' (SaveQuestion). So no one else ever sees a half-entered question, and a
' multipart question is never saved with parts that don't add up or a
' portion marking missing. The working copy is safe only because each user
' runs their own copy of the front end (README, Release): two people sharing
' one copy on a share would edit each other's working copy.
'
' Versions: a question is changed in place while it is a draft or rejected.
' An approved question is changed by saving a new version: VersionNo + 1,
' the same QuestionKey, ReplacesQuestionID the approved one, IsCurrent off.
' The approved version stays current and can still be drawn while the new
' one is a draft and goes through the whole approval chain again. Only the
' new version's final approval retires it (IsCurrent off, status Retired,
' RetiredOn) and makes the new one current, in one transaction (the
' approvals' last step). A key has at most one open version (a version that is not
' current and not retired: a draft, in review or rejected) at a time
' (OpenVersionID). Every save, submission, approval, rejection and
' retirement writes an audit row; audit rows are never changed or deleted.
'
' The statuses are the data file's own names (tblStatuses): Draft,
' In Review, Approved, Rejected and Retired.
'
' Who may change a question is checked in the data file at the moment of
' every write (RequireQuestionWriter), not from the roles read at sign-in:
' an active SME who wrote the question, or who is an SME on its topic, and
' only into a topic they are an SME on. The pages grey out what the person
' can't do (CanWriteQuestion), but the check at the write is the one that
' counts. References are shared by every question: any active SME, or the
' admin, may add or correct them (RequireReferenceWriter).
'
' Owner: the question bank (build phase 5). Depends on: modAdmin, modErrors,
' modGlobals, modSignIn; the kit's modKitLocks, modKitUI, modPersonnel.

Public Const QS_DRAFT As String = "Draft"
Public Const QS_REVIEW As String = "In Review"
Public Const QS_APPROVED As String = "Approved"
Public Const QS_REJECTED As String = "Rejected"
Public Const QS_RETIRED As String = "Retired"

Public Const QT_MC As String = "Multiple Choice"
Public Const QT_TF As String = "True/False"
Public Const QT_SA As String = "Short Answer"

' Audit actions (tblAuditTrail.ActionType).
Public Const AU_AUTHORED As String = "Authored"
Public Const AU_EDITED As String = "Edited"
Public Const AU_NEWVERSION As String = "New version"
Public Const AU_SUBMITTED As String = "Submitted"
Public Const AU_RETIRED As String = "Retired"
Public Const AU_RESTORED As String = "Restored"
Public Const AU_DELETED As String = "Deleted"
Public Const AU_ADDED As String = "Added"               ' a reference added to the shared list
Public Const AU_CHANGED As String = "Changed"           ' a reference's title, number or revision
Public Const AU_TURNED_OFF As String = "Turned off"
Public Const AU_TURNED_ON As String = "Turned on"

Public Const ITEM_QUESTION As String = "Question"
Public Const ITEM_REFERENCE As String = "Reference"     ' tblAuditTrail.ItemID is then the RefID

Public Const MAX_POINTS As Long = 100           ' most points one question (or part) may carry
Public Const MAX_TEXT As Long = 4000            ' longest question, part or option text
Private Const POINT_EPS As Double = 0.001       ' points compared to a thousandth
Private Const NEW_VERSION_RACE As String = "Someone else started a new version of this question at the same moment. " & _
    "Go back to the question bank and open theirs."

' tblStatuses, read once on first use (it doesn't change while the app runs):
' question status ids by name, and every status's name by "#" & id.
Private m_statusIds As Collection
Private m_statusNames As Collection

' ---------------------------------------------------------------------------
' Lookups
' ---------------------------------------------------------------------------

' A question status's id (tblStatuses, AppliesTo Question).
'   nm   a QS_ constant: Draft, In Review, Approved, Rejected, Retired
' Returns 0 when the data file has no such status.
' Errors: the engine's, the first time the statuses are read (LoadStatuses
'   then caches them for the session).
Public Function QStatusID(ByVal nm As String) As Long
    LoadStatuses
    On Error Resume Next                    ' a name that isn't there: 0
    QStatusID = m_statusIds(nm)
End Function

' A status's name (any AppliesTo).
'   id   tblStatuses.StatusID
' Returns "" when there is no such id.
' Errors: the engine's, the first time the statuses are read.
Public Function QStatusName(ByVal id As Long) As String
    LoadStatuses
    On Error Resume Next                    ' an id that isn't there: ""
    QStatusName = m_statusNames("#" & id)
End Function

' Fill the status cache from tblStatuses the first time it is needed. A
' failure leaves it empty, so the next call tries again.
Private Sub LoadStatuses()
    Dim rs As DAO.Recordset, ids As Collection, names As Collection
    If Not m_statusIds Is Nothing Then Exit Sub
    Set ids = New Collection
    Set names = New Collection
    Set rs = CurrentDb.OpenRecordset("SELECT StatusID, StatusName, AppliesTo FROM tblStatuses", dbOpenSnapshot)
    Do While Not rs.EOF
        If rs!AppliesTo = "Question" Then ids.Add CLng(rs!StatusID), CStr(rs!StatusName)
        names.Add CStr(rs!StatusName), "#" & rs!StatusID
        rs.MoveNext
    Loop
    rs.Close
    Set m_statusNames = names
    Set m_statusIds = ids
End Sub

' A question type's id.
'   nm   a QT_ constant: Multiple Choice, True/False, Short Answer
' Returns 0 when the data file has no such type. Errors: the engine's.
Public Function QTypeID(ByVal nm As String) As Long
    QTypeID = Nz(DLookup("QuestionTypeID", "tblQuestionTypes", "TypeName=" & Q(nm)), 0)
End Function

' A question type's name.
'   id   tblQuestionTypes.QuestionTypeID
' Returns "" when there is no such id. Errors: the engine's.
Public Function QTypeName(ByVal id As Long) As String
    QTypeName = Nz(DLookup("TypeName", "tblQuestionTypes", "QuestionTypeID=" & id), "")
End Function

' The editor's type switch key (Segmented segType) for a type name: mc, tf
' or sa; "" for anything else. Errors: none.
Public Function TypeKey(ByVal typeName As String) As String
    Select Case typeName
        Case QT_MC: TypeKey = "mc"
        Case QT_TF: TypeKey = "tf"
        Case QT_SA: TypeKey = "sa"
    End Select
End Function

' The type name for an editor type switch key (mc, tf, sa); "" for anything
' else. Errors: none.
Public Function TypeOfKey(ByVal key As String) As String
    Select Case key
        Case "mc": TypeOfKey = QT_MC
        Case "tf": TypeOfKey = QT_TF
        Case "sa": TypeOfKey = QT_SA
    End Select
End Function

' How a question version is named everywhere: "Q12 v2".
'   key, ver   QuestionKey and VersionNo ("?" for a Null)
' Errors: none.
Public Function QuestionRef(ByVal key As Variant, ByVal ver As Variant) As String
    QuestionRef = "Q" & Nz(key, "?") & " v" & Nz(ver, "?")
End Function

' A classification level's rank (tblClassLevels.LevelRank).
'   id   ClassLevelID, or Null
' Returns 0 for Null or no such level. Errors: the engine's.
Public Function ClassRank(ByVal id As Variant) As Long
    If IsNull(id) Then Exit Function
    ClassRank = Nz(DLookup("LevelRank", "tblClassLevels", "ClassLevelID=" & CLng(id)), 0)
End Function

' A classification level's portion mark, "(U)".
'   id   ClassLevelID, or Null
' Returns "" for Null or no such level. Errors: the engine's.
Public Function ClassPortion(ByVal id As Variant) As String
    If IsNull(id) Then Exit Function
    ClassPortion = Nz(DLookup("PortionMark", "tblClassLevels", "ClassLevelID=" & CLng(id)), "")
End Function

' The open new version of a question: a version that is not current and not
' retired (a draft, in review or rejected). A key has at most one, so a
' second "Edit as new version" is refused.
'   key   QuestionKey (0 gives 0)
' Returns its QuestionID, or 0 when there is none. Errors: the engine's.
Public Function OpenVersionID(ByVal key As Long) As Long
    If key = 0 Then Exit Function
    OpenVersionID = Nz(DMin("QuestionID", "tblQuestions", "QuestionKey=" & key & " AND Not IsCurrent AND StatusID<>" & _
        QStatusID(QS_RETIRED)), 0)
End Function

' Why a question can't be started again as a new version, or "" when it can:
' it must be the approved, current version, with no other new version open.
'   qid   the version the new one would start from
' Errors: the engine's.
Public Function NewVersionProblem(ByVal qid As Long) As String
    Dim st As String, other As Long
    st = QStatusName(Nz(DLookup("StatusID", "tblQuestions", "QuestionID=" & qid), 0))
    If st <> QS_APPROVED Or Not Nz(DLookup("IsCurrent", "tblQuestions", "QuestionID=" & qid), False) Then
        NewVersionProblem = "Only the approved, current version of a question can be started again as a new version. This one is " & _
            LCase$(st) & "."
        Exit Function
    End If
    other = OpenVersionID(Nz(DLookup("QuestionKey", "tblQuestions", "QuestionID=" & qid), 0))
    If other <> 0 Then
        NewVersionProblem = "Version " & DLookup("VersionNo", "tblQuestions", "QuestionID=" & other) & " of this question is already " & _
            "open (" & LCase$(QStatusName(DLookup("StatusID", "tblQuestions", "QuestionID=" & other))) & "). Change that one instead."
    End If
End Function

' True when a question in this status may be changed in place (Draft,
' Rejected).
'   nm   a status name (QStatusName)
' Errors: none.
Public Function StatusEditable(ByVal nm As String) As Boolean
    StatusEditable = (nm = QS_DRAFT Or nm = QS_REJECTED)
End Function

' ---------------------------------------------------------------------------
' Who may change what (checked at the write)
' ---------------------------------------------------------------------------

' Why the signed-in person may not change question qid now, or "" when
' they may. Reads the data file, so a role or topic taken away since sign-in
' counts at once.
'   qid      the question being changed; for a new version, the approved
'            version it starts from; 0 for a new question.
'   topicId  the topic it is being saved under (0: leave the topic out of
'            the check).
' Allowed: an active SME (tblPeople, tblPersonRoles, and the Personnel file
' still letting them use the app) who wrote qid or is an SME on its topic;
' and, when topicId is another topic than qid's, an SME on topicId.
' A front end with no data file has nothing to protect: "".
Public Function QuestionWriteProblem(ByVal qid As Long, Optional ByVal topicId As Long = 0) As String
    Dim pid As Long, author As Long, ownTopic As Long
    If Not DataLinked() Then Exit Function
    pid = SignedInPersonID()
    If Not IsActiveSME(pid) Then
        QuestionWriteProblem = "Only an active SME can change questions. Your SME role may have been taken away since you signed in."
        Exit Function
    End If
    If qid <> 0 Then
        author = Nz(DLookup("AuthorPersonID", "tblQuestions", "QuestionID=" & qid), 0)
        ownTopic = Nz(DLookup("TopicID", "tblQuestions", "QuestionID=" & qid), 0)
        If author <> pid And Not IsTopicSME(pid, ownTopic) Then
            QuestionWriteProblem = "Only the question's author or an SME on its topic can change it."
            Exit Function
        End If
    End If
    If topicId <> 0 And topicId <> ownTopic Then
        If Not IsTopicSME(pid, topicId) Then
            QuestionWriteProblem = "You aren't an SME on " & Nz(DLookup("TopicName", "tblTopics", "TopicID=" & topicId), "that topic") & _
                ", so you can't put a question in it. The admin adds SMEs to topics on the Topics page."
        End If
    End If
End Function

' Raise ERR_NOT_ALLOWED unless the signed-in person may change question qid
' now (QuestionWriteProblem; same arguments). Called first by every write to
' a question: SaveQuestion, SubmitQuestion, DeleteDraft.
' Errors: ERR_NOT_ALLOWED with the reason for the user. No side effects.
Public Sub RequireQuestionWriter(ByVal qid As Long, Optional ByVal topicId As Long = 0)
    Dim problem As String
    problem = QuestionWriteProblem(qid, topicId)
    If Len(problem) > 0 Then Err.Raise ERR_NOT_ALLOWED, "modQuestions.RequireQuestionWriter", problem
End Sub

' True when the signed-in person may change a question now
' (QuestionWriteProblem is ""), for the pages to grey out what they can't
' do. The write itself checks again (RequireQuestionWriter).
'   qid   the question, or 0 for a new one
' Errors: the engine's.
Public Function CanWriteQuestion(ByVal qid As Long) As Boolean
    CanWriteQuestion = (Len(QuestionWriteProblem(qid)) = 0)
End Function

' True when the signed-in person is an SME on a topic now: the editor's
' topic picker lists only these. True in a front end with no data file.
'   topicId   TopicID, or Null (False)
' Errors: the engine's.
Public Function TopicWritable(ByVal topicId As Variant) As Boolean
    If Not DataLinked() Then TopicWritable = True: Exit Function
    If IsNull(topicId) Then Exit Function
    TopicWritable = IsTopicSME(SignedInPersonID(), CLng(topicId))
End Function

' Raise ERR_NOT_ALLOWED unless the signed-in person is an active SME or the
' admin now. Called first by every write to the shared reference list
' (AddReference, ChangeReference, SetReferenceActive, DeleteReference).
' Errors: ERR_NOT_ALLOWED with the reason for the user.
' Side effects: none of its own (RequireExamAdmin sets TempVars EGIsAdmin).
Public Sub RequireReferenceWriter()
    If Not DataLinked() Then Exit Sub
    If IsActiveSME(SignedInPersonID()) Then Exit Sub
    RequireExamAdmin                        ' raises when they aren't the admin either
End Sub

' pid is active here, holds the SME role, and may still use the app.
Private Function IsActiveSME(ByVal pid As Long) As Boolean
    If pid = 0 Then Exit Function
    If Not Nz(DLookup("IsActive", "tblPeople", "PersonID=" & pid), False) Then Exit Function
    If Not PersonHasRole(pid, "SME") Then Exit Function
    IsActiveSME = PersonCanUse(pid, APP_CODE)
End Function

Private Function IsTopicSME(ByVal pid As Long, ByVal topicId As Long) As Boolean
    If pid = 0 Or topicId = 0 Then Exit Function
    IsTopicSME = DCount("*", "tblTopicSMEs", "TopicID=" & topicId & " AND PersonID=" & pid) > 0
End Function

' ---------------------------------------------------------------------------
' Audit rows
' ---------------------------------------------------------------------------

' Write one audit row (tblAuditTrail), by the signed-in person, now, unless
' personId and actionOn say otherwise. Called
' inside the caller's transaction, so the row and the change it records go
' together.
'   itemType     ITEM_QUESTION, ITEM_REFERENCE or ITEM_EXAM
'   itemId       the question's QuestionID, the reference's RefID or the ExamID
'   ver          the question's VersionNo, Null for a reference
'   action       an AU_ constant
'   stepRoleID   the approval step's RoleID, 0 for none
'   comments     free text, "" for none
'   personId     who acted, when not the signed-in person (0: them)
'   actionOn     when, for an act recorded later (Null: now)
' Errors: KitExecute's lock and write errors pass up.
Public Sub AddAudit(ByVal itemType As String, ByVal itemId As Long, ByVal ver As Variant, ByVal action As String, _
        Optional ByVal stepRoleID As Long = 0, Optional ByVal comments As String = "", Optional ByVal personId As Long = 0, _
        Optional ByVal actionOn As Variant = Null)
    If personId = 0 Then personId = SignedInPersonID()
    KitExecute "INSERT INTO tblAuditTrail (ItemType, ItemID, VersionNo, ActionType, StepRoleID, PersonID, ActionOn, Comments) " & _
        "VALUES (" & Q(itemType) & ", " & itemId & ", " & IIf(IsNull(ver), "Null", CLng(Nz(ver, 0))) & ", " & Q(action) & ", " & _
        IIf(stepRoleID = 0, "Null", stepRoleID) & ", " & personId & ", " & _
        IIf(IsDate(actionOn), "#" & Format$(Nz(actionOn, Now()), "yyyy-mm-dd hh\:nn\:ss") & "#", "Now()") & ", " & _
        IIf(Len(comments) > 0, Q(comments), "Null") & ")"
End Sub

' The newest audit row of any version of a question, the editor's stamp to
' spot a save someone else made while it was open.
'   key   QuestionKey (0 gives 0)
' Returns its AuditID, 0 for none. Errors: the engine's.
Public Function LastQuestionAudit(ByVal key As Long) As Long
    If key = 0 Then Exit Function
    LastQuestionAudit = Nz(DMax("AuditID", "tblAuditTrail", "ItemType='" & ITEM_QUESTION & "' AND ItemID IN " & _
        "(SELECT QuestionID FROM tblQuestions WHERE QuestionKey=" & key & ")"), 0)
End Function

' Who wrote an audit row and when: "Smith, John (LT) at 14:05 on 27 Oct".
'   id   AuditID
' Returns "" when there is no such row. Errors: the engine's.
Public Function AuditWho(ByVal id As Long) As String
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT a.ActionOn, " & PersonNameSql() & " AS person FROM tblAuditTrail AS a " & _
        "LEFT JOIN " & PERSON_TABLE & " AS p ON a.PersonID=p.person_id WHERE a.AuditID=" & id, dbOpenSnapshot)
    If Not rs.EOF Then AuditWho = Nz(rs!person, "someone") & " at " & Format$(rs!ActionOn, "hh:nn") & " on " & Format$(rs!ActionOn, "d mmm")
    rs.Close
End Function

' SQL for every audit row of a question's versions, newest first, for a
' history table: id, ActionOn, when_txt, ver, ActionType, person, step,
' Comments.
'   key   QuestionKey
' Errors: none (the query's errors come when it is opened).
Public Function HistorySql(ByVal key As Long) As String
    HistorySql = "SELECT a.AuditID AS id, a.ActionOn, Format(a.ActionOn,'dd mmm yyyy hh:nn') AS when_txt, " & _
        "'v' & a.VersionNo AS ver, a.ActionType, " & PersonNameSql() & " AS person, r.RoleCode AS step, a.Comments " & _
        "FROM ((tblAuditTrail AS a INNER JOIN tblQuestions AS q ON a.ItemID=q.QuestionID) " & _
        "LEFT JOIN " & PERSON_TABLE & " AS p ON a.PersonID=p.person_id) LEFT JOIN tblRoles AS r ON a.StepRoleID=r.RoleID " & _
        "WHERE a.ItemType='" & ITEM_QUESTION & "' AND q.QuestionKey=" & key & " ORDER BY a.ActionOn DESC, a.AuditID DESC"
End Function

' A question's history as plain text, one line per row, for Copy.
'   key   QuestionKey
' Errors: the engine's.
Public Function HistoryText(ByVal key As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset(HistorySql(key), dbOpenSnapshot)
    s = "History of Q" & key & vbCrLf
    Do While Not rs.EOF
        s = s & rs!when_txt & vbTab & rs!ver & vbTab & rs!ActionType & vbTab & Nz(rs!person, "") & _
            IIf(IsNull(rs!step), "", " (" & rs!step & ")") & IIf(IsNull(rs!comments), "", vbTab & rs!comments) & vbCrLf
        rs.MoveNext
    Loop
    rs.Close
    HistoryText = s
End Function

' ---------------------------------------------------------------------------
' The editor's working copy (local tables)
' ---------------------------------------------------------------------------

' Make the editor's three local tables (wq_part, wq_option, wq_ref) when
' this front end lacks them: BuildAll's "local tables" step, and the editor
' before it loads a question. They live in each user's own copy of the front
' end, which is why each user needs their own copy.
' Errors: the engine's.
' Side effects: creates tables in the front end.
Public Sub EnsureWorkTables()
    If Not TableExists("wq_part", True) Then
        CurrentDb.Execute "CREATE TABLE wq_part (id COUNTER CONSTRAINT pk_wq_part PRIMARY KEY, PartLabel TEXT(5), " & _
            "PartText MEMO, Points DOUBLE, ClassLevelID LONG, ReferenceAnswer MEMO, SortOrder LONG)", dbFailOnError
    End If
    If Not TableExists("wq_option", True) Then
        CurrentDb.Execute "CREATE TABLE wq_option (id COUNTER CONSTRAINT pk_wq_option PRIMARY KEY, part_id LONG, " & _
            "OptionText MEMO, IsCorrect YESNO, SortOrder LONG)", dbFailOnError
    End If
    If Not TableExists("wq_ref", True) Then
        CurrentDb.Execute "CREATE TABLE wq_ref (id COUNTER CONSTRAINT pk_wq_ref PRIMARY KEY, RefID LONG, Paragraph TEXT(100))", dbFailOnError
    End If
End Sub

' Empty the editor's working copy (wq_option, wq_part, wq_ref).
' Errors: the engine's.
Public Sub ClearWork()
    CurrentDb.Execute "DELETE FROM wq_option", dbFailOnError
    CurrentDb.Execute "DELETE FROM wq_part", dbFailOnError
    CurrentDb.Execute "DELETE FROM wq_ref", dbFailOnError
End Sub

' Copy a question's parts, options and references into the working copy,
' replacing what was there, with the options numbered 1, 2, ... per group.
'   qid   the question's QuestionID (0 leaves the working copy empty)
' Errors: the engine's.
' Side effects: rewrites the three local tables.
Public Sub LoadWork(ByVal qid As Long)
    Dim db As DAO.Database, src As DAO.Recordset, dst As DAO.Recordset, map As Object, k As Variant
    Set db = CurrentDb
    ClearWork
    Set map = CreateObject("Scripting.Dictionary")
    Set src = db.OpenRecordset("SELECT * FROM tblQuestionParts WHERE QuestionID=" & qid & " ORDER BY SortOrder, PartLabel", dbOpenSnapshot)
    Set dst = db.OpenRecordset("wq_part", dbOpenDynaset)
    Do While Not src.EOF
        dst.AddNew
        dst!PartLabel = src!PartLabel
        dst!PartText = src!PartText
        dst!Points = src!Points
        dst!ClassLevelID = src!ClassLevelID
        dst!ReferenceAnswer = src!ReferenceAnswer
        dst!SortOrder = map.Count + 1
        k = dst!id
        dst.Update
        map(CLng(src!PartID)) = k
        src.MoveNext
    Loop
    src.Close: dst.Close
    Set src = db.OpenRecordset("SELECT * FROM tblMCOptions WHERE QuestionID=" & qid & " ORDER BY PartID, SortOrder, OptionLabel", dbOpenSnapshot)
    Set dst = db.OpenRecordset("wq_option", dbOpenDynaset)
    Do While Not src.EOF
        dst.AddNew
        If Not IsNull(src!PartID) Then
            If map.Exists(CLng(src!PartID)) Then dst!part_id = map(CLng(src!PartID))
        End If
        dst!OptionText = src!OptionText
        dst!IsCorrect = Nz(src!IsCorrect, False)
        dst!SortOrder = Nz(src!SortOrder, 0)
        dst.Update
        src.MoveNext
    Loop
    src.Close: dst.Close
    db.Execute "INSERT INTO wq_ref (RefID, Paragraph) SELECT RefID, Paragraph FROM tblQuestionRefs WHERE QuestionID=" & qid, dbFailOnError
    RenumberOptions
End Sub

' Number each option group's options 1, 2, ... in their present order (the
' working copy's wq_option). Errors: the engine's.
Public Sub RenumberOptions()
    Dim rs As DAO.Recordset, grp As Variant, n As Long
    Set rs = CurrentDb.OpenRecordset("SELECT part_id, SortOrder FROM wq_option ORDER BY Nz(part_id,0), SortOrder, id", dbOpenDynaset)
    grp = -1
    Do While Not rs.EOF
        If Nz(rs!part_id, 0) <> grp Then grp = Nz(rs!part_id, 0): n = 0
        n = n + 1
        If Nz(rs!SortOrder, 0) <> n Then rs.Edit: rs!SortOrder = n: rs.Update
        rs.MoveNext
    Loop
    rs.Close
End Sub

' How many parts the working copy has. Errors: the engine's.
Public Function WorkPartCount() As Long
    WorkPartCount = DCount("*", "wq_part")
End Function

' The working copy's parts' points added up (0 for no parts). Errors: the engine's.
Public Function WorkPointsSum() As Double
    WorkPointsSum = Nz(DSum("Points", "wq_part"), 0)
End Function

' The first letter A to Z no part in the working copy uses yet, "" when all
' 26 are used. Errors: the engine's.
Public Function NextPartLabel() As String
    Dim i As Long
    For i = 0 To 25
        If DCount("*", "wq_part", "PartLabel=" & Q(Chr$(65 + i))) = 0 Then NextPartLabel = Chr$(65 + i): Exit Function
    Next
End Function

' Letter the working copy's parts A, B, C... in their order and number them
' 1, 2, 3..., after a move or a removal. Errors: the engine's.
Public Sub RelabelParts()
    Dim rs As DAO.Recordset, n As Long
    Set rs = CurrentDb.OpenRecordset("SELECT PartLabel, SortOrder FROM wq_part ORDER BY SortOrder, PartLabel", dbOpenDynaset)
    Do While Not rs.EOF
        n = n + 1
        If Nz(rs!PartLabel, "") <> Chr$(64 + n) Or Nz(rs!SortOrder, 0) <> n Then
            rs.Edit: rs!PartLabel = Chr$(64 + n): rs!SortOrder = n: rs.Update
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

' Open a question in the editor, or say it is gone. An entry point (Ctrl+K's
' recent list calls it too), so it handles its own errors.
'   qid   the question's QuestionID
' Side effects: sets TempVars EditQuestionID and EditMode and goes to the
'   qedit page.
Public Function OpenQuestion(ByVal qid As Long)
    On Error GoTo Err_Handler
    If DCount("*", "tblQuestions", "QuestionID=" & qid) = 0 Then
        KitMsgBox "That question isn't in the bank any more.", vbInformation, "Question bank"
        Exit Function
    End If
    TempVars("EditQuestionID") = qid
    TempVars("EditMode") = ""
    NavTo "qedit"
    Exit Function
Err_Handler:
    HandleError "modQuestions.OpenQuestion", qid
End Function

' The highest marking among the working copy's parts (the question's, rolled
' up): a ClassLevelID, or Null when there are no parts or none is marked.
' Errors: the engine's.
Public Function RolledUpMarking() As Variant
    Dim rs As DAO.Recordset
    RolledUpMarking = Null
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 w.ClassLevelID FROM wq_part AS w INNER JOIN tblClassLevels AS c " & _
        "ON w.ClassLevelID=c.ClassLevelID ORDER BY c.LevelRank DESC", dbOpenSnapshot)
    If Not rs.EOF Then RolledUpMarking = rs!ClassLevelID
    rs.Close
End Function

' A working-copy part's letter for an option's group: "A", or "" for the
' question itself (Null) or no such part.
'   partId   wq_part.id, or Null
' Errors: the engine's.
Public Function PartLabelOf(ByVal partId As Variant) As String
    If IsNull(partId) Then Exit Function
    PartLabelOf = Nz(DLookup("PartLabel", "wq_part", "id=" & CLng(partId)), "")
End Function

' An option's letter in its group: 1 is A, 26 is Z; "" outside 1 to 26.
' Errors: none.
Public Function OptionLetter(ByVal n As Variant) As String
    If Nz(n, 0) < 1 Or Nz(n, 0) > 26 Then Exit Function
    OptionLetter = Chr$(64 + CLng(n))
End Function

' SQL for the working copy's parts, for the editor's grid: id, PartLabel,
' PartText, Points, pts, mark, ans, SortOrder. Errors: none.
Public Function WorkPartsSql() As String
    WorkPartsSql = "SELECT w.id, w.PartLabel, w.PartText, w.Points, PointsText(w.Points) AS pts, c.Marking AS mark, " & _
        "IIf(Len(Nz(w.ReferenceAnswer,''))>0,'Yes','') AS ans, w.SortOrder FROM wq_part AS w " & _
        "LEFT JOIN tblClassLevels AS c ON w.ClassLevelID=c.ClassLevelID ORDER BY w.SortOrder, w.PartLabel"
End Function

' SQL for the working copy's options, for the editor's grid: id, part,
' letter, OptionText, correct, psort, SortOrder. Errors: none.
Public Function WorkOptionsSql() As String
    WorkOptionsSql = "SELECT o.id, Nz(p.PartLabel,'') AS part, OptionLetter(o.SortOrder) AS letter, o.OptionText, " & _
        "IIf(o.IsCorrect,'Correct','') AS correct, Nz(p.SortOrder,0) AS psort, o.SortOrder FROM wq_option AS o " & _
        "LEFT JOIN wq_part AS p ON o.part_id=p.id ORDER BY Nz(p.SortOrder,0), o.SortOrder"
End Function

' SQL for the working copy's citations, for the editor's grid: id,
' RefTitle, RefNumber, Revision, Paragraph. Errors: none.
Public Function WorkRefsSql() As String
    WorkRefsSql = "SELECT w.id, r.RefTitle, r.RefNumber, r.Revision, w.Paragraph FROM wq_ref AS w " & _
        "INNER JOIN tblSourceRefs AS r ON w.RefID=r.RefID ORDER BY r.RefTitle"
End Function

' ---------------------------------------------------------------------------
' Rules
' ---------------------------------------------------------------------------

' Text as it is saved: tabs as spaces, each line's trailing spaces and the
' whole's leading and trailing blank lines dropped, at most one empty line
' in a row.
'   v   the text, or Null
' Returns Null when nothing is left. Errors: none.
Public Function CleanText(ByVal v As Variant) As Variant
    Dim s As String, lines() As String, i As Long, out As String, blank As Long
    CleanText = Null
    If IsNull(v) Then Exit Function
    s = Replace(Replace(Replace(CStr(v), vbCrLf, vbLf), vbCr, vbLf), vbTab, " ")
    lines = Split(s, vbLf)
    For i = 0 To UBound(lines)
        lines(i) = RTrim$(lines(i))
        If Len(lines(i)) = 0 Then
            blank = blank + 1
            If blank = 1 And Len(out) > 0 Then out = out & vbCrLf
        Else
            blank = 0
            out = out & IIf(Len(out) > 0 And Right$(out, 2) <> vbCrLf, vbCrLf, "") & lines(i)
        End If
    Next
    out = Trim$(out)
    Do While Right$(out, 2) = vbCrLf
        out = Left$(out, Len(out) - 2)
    Loop
    If Len(out) > 0 Then CleanText = out
End Function

' What is wrong with a points value, or "" when it is fine: a number above
' 0, at most MAX_POINTS, in whole or half points.
'   v      the value as typed
'   what   names it in the message: "Points", "Part B's points"
' Errors: none.
Public Function PointsProblem(ByVal v As Variant, ByVal what As String) As String
    Dim d As Double
    If IsNull(v) Then PointsProblem = what & " are missing.": Exit Function
    If Not IsNumeric(v) Then PointsProblem = what & " must be a number (" & v & " isn't).": Exit Function
    d = CDbl(v)
    If d <= 0 Then
        PointsProblem = what & " must be more than 0."
    ElseIf d > MAX_POINTS Then
        PointsProblem = what & " can be at most " & MAX_POINTS & "."
    ElseIf Abs(d * 2 - Int(d * 2 + POINT_EPS)) > POINT_EPS Then
        PointsProblem = what & " must be in whole or half points (" & d & " isn't)."
    End If
End Function

' Points as the app shows them: 4, 1.5; "" for Null or text. Errors: none.
Public Function PointsText(ByVal v As Variant) As String
    If IsNull(v) Then Exit Function
    PointsText = Format$(CDbl(v), "0.##")
    If Right$(PointsText, 1) = "." Then PointsText = Left$(PointsText, Len(PointsText) - 1)
End Function

' Every rule a question breaks, one per line, or "" when none. The
' question's own fields come from the editor; its parts, options and
' references from the working copy.
'   topicId .. text   the editor's fields as typed (any may be Null)
'   forSubmit         True adds the rules a question must meet before it
'                     goes for approval (references, and multiple choice
'                     options with one correct answer)
' Errors: the engine's.
Public Function QuestionProblems(ByVal topicId As Variant, ByVal typeName As String, ByVal difficultyId As Variant, _
        ByVal points As Variant, ByVal classLevelId As Variant, ByVal text As Variant, ByVal forSubmit As Boolean) As String
    Dim s As String, p As String, rs As DAO.Recordset, nParts As Long, sum As Double, isMC As Boolean
    Dim nOpts As Long, nCorrect As Long, grp As String, partsOK As Boolean
    ' The question's own fields.
    If IsNull(topicId) Then
        s = s & "Topic: pick the topic the question belongs to." & vbCrLf
    ElseIf Not Nz(DLookup("IsActive", "tblTopics", "TopicID=" & CLng(topicId)), False) Then
        s = s & "Topic: """ & Nz(DLookup("TopicName", "tblTopics", "TopicID=" & CLng(topicId)), "?") & _
            """ is turned off. Pick another, or ask the admin to turn it on." & vbCrLf
    End If
    If Len(typeName) = 0 Then s = s & "Type: pick multiple choice, true/false or short answer." & vbCrLf
    If IsNull(difficultyId) Then s = s & "Difficulty: pick how hard the question is." & vbCrLf
    p = PointsProblem(points, "Points")
    If Len(p) > 0 Then s = s & p & vbCrLf
    If IsNull(CleanText(text)) Then
        s = s & "Question text: write the question." & vbCrLf
    ElseIf Len(CleanText(text)) > MAX_TEXT Then
        s = s & "Question text can be at most " & MAX_TEXT & " characters; it is " & Len(CleanText(text)) & "." & vbCrLf
    End If
    ' Parts.
    nParts = WorkPartCount()
    If nParts = 0 Then
        If IsNull(classLevelId) Then s = s & "Portion marking: pick the question's marking." & vbCrLf
    Else
        If nParts = 1 Then s = s & "Parts: a question with parts needs at least two (A and B). Remove part " & _
            DLookup("PartLabel", "wq_part") & " or add another." & vbCrLf
        partsOK = True
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM wq_part ORDER BY SortOrder, PartLabel", dbOpenSnapshot)
        Do While Not rs.EOF
            p = PointsProblem(rs!Points, "Part " & rs!PartLabel & "'s points")
            If Len(p) > 0 Then
                s = s & p & vbCrLf
                partsOK = False
            Else
                sum = sum + rs!Points
            End If
            If IsNull(CleanText(rs!PartText)) Then s = s & "Part " & rs!PartLabel & ": write the part's text." & vbCrLf
            If IsNull(rs!ClassLevelID) Then s = s & "Part " & rs!PartLabel & ": pick its portion marking." & vbCrLf
            rs.MoveNext
        Loop
        rs.Close
        ' Only once the total and every part's points are numbers.
        If partsOK And Len(PointsProblem(points, "Points")) = 0 Then
            If Abs(sum - CDbl(points)) > POINT_EPS Then
                s = s & "Part points (" & PointsText(sum) & ") don't add up to the question total (" & PointsText(points) & ")." & vbCrLf
            End If
        End If
    End If
    ' Answer options.
    isMC = (typeName = QT_MC)
    nOpts = DCount("*", "wq_option")
    If nOpts > 0 And Not isMC And Len(typeName) > 0 Then
        s = s & "Answer options: only multiple choice questions have them. Remove the " & nOpts & _
            IIf(nOpts = 1, " option", " options") & ", or make the question multiple choice." & vbCrLf
    End If
    If isMC Then
        If nParts > 0 And DCount("*", "wq_option", "part_id Is Null") > 0 Then
            s = s & "Answer options: once a question has parts, each option belongs to a part. Remove the options " & _
                "without one and add them to their parts." & vbCrLf
        End If
        Set rs = CurrentDb.OpenRecordset(OptionGroupsSql(nParts > 0), dbOpenSnapshot)
        Do While Not rs.EOF
            grp = IIf(IsNull(rs!PartLabel), "The question", "Part " & rs!PartLabel)
            nOpts = Nz(rs!n, 0): nCorrect = Nz(rs!nCorrect, 0)
            If nCorrect > 1 Then
                s = s & "Answer options: " & grp & " has " & nCorrect & " options marked correct; mark only one." & vbCrLf
            ElseIf forSubmit Then
                If nOpts < 2 Then
                    s = s & "Answer options: " & grp & " needs at least two options (it has " & nOpts & ")." & vbCrLf
                ElseIf nCorrect = 0 Then
                    s = s & "Answer options: mark " & LCase$(Left$(grp, 1)) & Mid$(grp, 2) & "'s correct option." & vbCrLf
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    ' References.
    If forSubmit And DCount("*", "wq_ref") = 0 Then s = s & "References: cite at least one reference." & vbCrLf
    If Right$(s, 2) = vbCrLf Then s = Left$(s, Len(s) - 2)
    QuestionProblems = s
End Function

' The option groups of the working copy with their counts: one row for the
' question (no parts), or one per part. Columns PartLabel, n, nCorrect.
Private Function OptionGroupsSql(ByVal byPart As Boolean) As String
    If byPart Then
        OptionGroupsSql = "SELECT p.PartLabel, (SELECT Count(*) FROM wq_option AS o WHERE o.part_id=p.id) AS n, " & _
            "(SELECT Count(*) FROM wq_option AS o WHERE o.part_id=p.id AND o.IsCorrect) AS nCorrect " & _
            "FROM wq_part AS p ORDER BY p.SortOrder, p.PartLabel"
    Else
        OptionGroupsSql = "SELECT Null AS PartLabel, Count(*) AS n, Sum(IIf(IsCorrect,1,0)) AS nCorrect FROM wq_option"
    End If
End Function

' ---------------------------------------------------------------------------
' Save, submit, delete
' ---------------------------------------------------------------------------

' Write the editor's question to the data file, with the working copy's
' parts, options and references, in one transaction. Its rules must already
' pass (QuestionProblems).
'   qid            the question being changed, 0 for a new one
'   baseId         for a new version, the approved version it replaces (qid
'                  is then 0); else 0. The approved version stays current
'                  until the new one is approved (Phase 6).
'   topicId .. refAnswer   the editor's fields; with parts the marking is
'                  the parts' highest (RolledUpMarking)
'   savedId        returns the saved question's QuestionID (0 when not saved)
' Returns "" when saved, else why it can't be saved now (someone moved it on,
' deleted it, or started a new version at the same moment).
' Errors: ERR_NOT_ALLOWED unless the signed-in person may write it
'   (RequireQuestionWriter); others pass up through RaiseAgain after the
'   rollback.
' Side effects: replaces the question's parts, options and citations, and
'   writes an Authored, Edited or New version audit row.
Public Function SaveQuestion(ByVal qid As Long, ByVal baseId As Long, ByVal topicId As Long, ByVal typeName As String, _
        ByVal difficultyId As Long, ByVal points As Double, ByVal classLevelId As Variant, ByVal text As Variant, _
        ByVal refAnswer As Variant, ByRef savedId As Long) As String
    Dim ws As DAO.Workspace, db As DAO.Database, rs As DAO.Recordset, inTrans As Boolean
    Dim st As String, key As Long, ver As Long, action As String, oldVer As Long, mark As Variant
    Dim num As Long, desc As String, src As String
    RequireQuestionWriter IIf(qid <> 0, qid, baseId), topicId
    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb
    savedId = 0
    ' The marking: rolled up from the parts when there are any.
    mark = classLevelId
    If WorkPartCount() > 0 Then mark = RolledUpMarking()
    ' Still where the editor found it?
    If qid <> 0 Then
        st = QStatusName(Nz(DLookup("StatusID", "tblQuestions", "QuestionID=" & qid), 0))
        If Len(st) = 0 Then SaveQuestion = "This question was deleted while you had it open.": Exit Function
        If Not StatusEditable(st) Then SaveQuestion = "This question is " & LCase$(st) & " now, so it can't be changed. " & _
            "Undo your changes to see it as it stands.": Exit Function
    ElseIf baseId <> 0 Then
        st = NewVersionProblem(baseId)
        If Len(st) > 0 Then
            SaveQuestion = st & " So this new version can't be saved. Go back to the question bank and start from the current version."
            Exit Function
        End If
    End If
    On Error GoTo Fail
    ws.BeginTrans: inTrans = True
    ' Again inside the transaction, as part of a write: the checks above may
    ' be out of date by now. The UPDATE that only matches a row still in the
    ' state the save needs holds that row until the commit, so an approver's
    ' action or a second new version waits for this save instead of slipping
    ' in between the check and the write.
    If qid <> 0 Then
        If KitExecute("UPDATE tblQuestions SET StatusID=StatusID WHERE QuestionID=" & qid & " AND StatusID IN (" & _
                QStatusID(QS_DRAFT) & "," & QStatusID(QS_REJECTED) & ")") = 0 Then
            ws.Rollback: inTrans = False
            SaveQuestion = "Someone else moved this question on while you had it open, so your changes can't be saved over it. " & _
                "Undo your changes to see it as it stands."
            Exit Function
        End If
    ElseIf baseId <> 0 Then
        If KitExecute("UPDATE tblQuestions SET IsCurrent=IsCurrent WHERE QuestionID=" & baseId & " AND IsCurrent AND StatusID=" & _
                QStatusID(QS_APPROVED)) = 0 Then
            ws.Rollback: inTrans = False
            SaveQuestion = "The version this one starts from is no longer the approved, current version, so this new version " & _
                "can't be saved. Go back to the question bank and start from the current version."
            Exit Function
        End If
        If OpenVersionID(DLookup("QuestionKey", "tblQuestions", "QuestionID=" & baseId)) <> 0 Then
            ws.Rollback: inTrans = False
            SaveQuestion = NEW_VERSION_RACE
            Exit Function
        End If
    End If
    Set rs = db.OpenRecordset("SELECT * FROM tblQuestions WHERE QuestionID=" & qid, dbOpenDynaset)
    If qid = 0 Then
        rs.AddNew
        If baseId <> 0 Then
            key = DLookup("QuestionKey", "tblQuestions", "QuestionID=" & baseId)
            oldVer = DLookup("VersionNo", "tblQuestions", "QuestionID=" & baseId)
            ver = Nz(DMax("VersionNo", "tblQuestions", "QuestionKey=" & key), 0) + 1
            rs!QuestionKey = key
            rs!ReplacesQuestionID = baseId
            action = AU_NEWVERSION
        Else
            ver = 1
            action = AU_AUTHORED
        End If
        rs!VersionNo = ver
        rs!StatusID = QStatusID(QS_DRAFT)
        rs!AuthorPersonID = SignedInPersonID()
        rs!CreatedOn = Now()
        rs!IsCurrent = (baseId = 0)             ' a new version waits for its final approval
    Else
        rs.Edit
        key = rs!QuestionKey
        ver = rs!VersionNo
        action = AU_EDITED
    End If
    rs!TopicID = topicId
    rs!QuestionTypeID = QTypeID(typeName)
    rs!DifficultyID = difficultyId
    rs!ClassLevelID = mark
    rs!QuestionText = CleanText(text)
    rs!TotalPoints = points
    rs!ReferenceAnswer = CleanText(refAnswer)
    rs.Update
    rs.Bookmark = rs.LastModified
    savedId = rs!QuestionID
    If key = 0 Then key = savedId: rs.Edit: rs!QuestionKey = key: rs.Update
    rs.Close
    ' Its parts, options and references: replaced by the working copy's.
    db.Execute "DELETE FROM tblMCOptions WHERE QuestionID=" & savedId, dbFailOnError
    db.Execute "DELETE FROM tblQuestionParts WHERE QuestionID=" & savedId, dbFailOnError
    db.Execute "DELETE FROM tblQuestionRefs WHERE QuestionID=" & savedId, dbFailOnError
    WriteParts db, savedId
    db.Execute "INSERT INTO tblQuestionRefs (QuestionID, RefID, Paragraph) SELECT " & savedId & ", RefID, Paragraph FROM wq_ref", dbFailOnError
    ' The approved version stays in use until this one is approved.
    If baseId <> 0 Then
        AddAudit ITEM_QUESTION, savedId, ver, action, , "Replaces version " & oldVer & " once approved"
    Else
        AddAudit ITEM_QUESTION, savedId, ver, action
    End If
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    num = Err.Number: desc = Err.Description: src = Err.Source
    If inTrans Then ws.Rollback
    savedId = 0
    ' Two new versions of one question saved at the same moment: the
    ' version number is DMax + 1, and uxQKeyVer refuses the second. Only
    ' that clash is the race; any other duplicate is raised as it is.
    If num = 3022 And baseId <> 0 And key <> 0 And ver <> 0 Then
        DBEngine.Idle dbRefreshCache
        If DCount("*", "tblQuestions", "QuestionKey=" & key & " AND VersionNo=" & ver) > 0 Then SaveQuestion = NEW_VERSION_RACE: Exit Function
    End If
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modQuestions.SaveQuestion", IIf(qid <> 0, "question " & qid, IIf(baseId <> 0, "new version of " & baseId, "new question"))
End Function

' The working copy's parts and options as question qid's (inside
' SaveQuestion's transaction). Options are lettered A, B... in each group.
Private Sub WriteParts(db As DAO.Database, ByVal qid As Long)
    Dim src As DAO.Recordset, dst As DAO.Recordset, map As Object, n As Long
    Set map = CreateObject("Scripting.Dictionary")
    Set src = db.OpenRecordset("SELECT * FROM wq_part ORDER BY SortOrder, PartLabel", dbOpenSnapshot)
    Set dst = db.OpenRecordset("SELECT * FROM tblQuestionParts WHERE False", dbOpenDynaset)
    Do While Not src.EOF
        n = n + 1
        dst.AddNew
        dst!QuestionID = qid
        dst!PartLabel = src!PartLabel
        dst!PartText = CleanText(src!PartText)
        dst!Points = src!Points
        dst!ClassLevelID = src!ClassLevelID
        dst!ReferenceAnswer = CleanText(src!ReferenceAnswer)
        dst!SortOrder = n
        dst.Update
        dst.Bookmark = dst.LastModified
        map(CLng(src!id)) = CLng(dst!PartID)
        src.MoveNext
    Loop
    src.Close: dst.Close
    RenumberOptions
    Set src = db.OpenRecordset("SELECT * FROM wq_option ORDER BY Nz(part_id,0), SortOrder", dbOpenSnapshot)
    Set dst = db.OpenRecordset("SELECT * FROM tblMCOptions WHERE False", dbOpenDynaset)
    Do While Not src.EOF
        dst.AddNew
        dst!QuestionID = qid
        If Not IsNull(src!part_id) Then dst!PartID = map(CLng(src!part_id))
        dst!OptionLabel = OptionLetter(src!SortOrder)
        dst!OptionText = CleanText(src!OptionText)
        dst!IsCorrect = Nz(src!IsCorrect, False)
        dst!SortOrder = src!SortOrder
        dst.Update
        src.MoveNext
    Loop
    src.Close: dst.Close
End Sub

' Send a question to the approval chain: In Review, at the DH's step. Its
' rules must already pass for submission. The status change and its
' Submitted audit row are one transaction, and the change only matches a
' question still in draft or rejected.
'   qid   the question's QuestionID
' Returns "" when submitted, else why not (nothing is written).
' Errors: ERR_NOT_ALLOWED unless the signed-in person may write it; others
'   pass up through RaiseAgain after the rollback.
Public Function SubmitQuestion(ByVal qid As Long) As String
    Dim st As String, dh As Long, ver As Long, ws As DAO.Workspace, inTrans As Boolean
    RequireQuestionWriter qid
    st = QStatusName(Nz(DLookup("StatusID", "tblQuestions", "QuestionID=" & qid), 0))
    If Not StatusEditable(st) Then SubmitQuestion = "It is " & LCase$(st) & ", so it can't be submitted.": Exit Function
    dh = RoleID("DH")
    ver = Nz(DLookup("VersionNo", "tblQuestions", "QuestionID=" & qid), 0)
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    If KitExecute("UPDATE tblQuestions SET StatusID=" & QStatusID(QS_REVIEW) & ", CurrentStepRoleID=" & dh & _
            " WHERE QuestionID=" & qid & " AND StatusID IN (" & QStatusID(QS_DRAFT) & "," & QStatusID(QS_REJECTED) & ")") = 0 Then
        ws.Rollback: inTrans = False
        SubmitQuestion = "Someone else moved this question on while you had it open, so it wasn't submitted. " & _
            "Undo your changes to see it as it stands."
        Exit Function
    End If
    AddAudit ITEM_QUESTION, qid, ver, AU_SUBMITTED, dh
    ws.CommitTrans: inTrans = False
    NoticeQuestionAtStep qid                ' the DHs' emails; never raises
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modQuestions.SubmitQuestion", qid
End Function

' Why a question can't be deleted, or "" when it can: only a draft that was
' never submitted and is on no exam.
'   qid   the question's QuestionID
' Errors: the engine's.
Public Function DeleteProblem(ByVal qid As Long) As String
    Dim st As String
    st = QStatusName(Nz(DLookup("StatusID", "tblQuestions", "QuestionID=" & qid), 0))
    If Len(st) = 0 Then
        DeleteProblem = "It has already been deleted."
    ElseIf st <> QS_DRAFT Then
        DeleteProblem = "Only a draft can be deleted; this question is " & LCase$(st) & "."
    ElseIf DCount("*", "tblAuditTrail", "ItemType='" & ITEM_QUESTION & "' AND ItemID=" & qid & " AND ActionType=" & Q(AU_SUBMITTED)) > 0 Then
        DeleteProblem = "It has been submitted for approval before, so its history has to stay."
    ElseIf DCount("*", "tblExamQuestions", "QuestionID=" & qid) > 0 Then
        DeleteProblem = "It is on an exam."
    End If
End Function

' Delete a draft and its parts, options and citations (cascade), with a
' Deleted audit row, in one transaction. The delete rules (DeleteProblem)
' are checked again inside it, and only a draft is deleted.
'   qid   the draft's QuestionID
' Returns the question to show instead: for a new version's draft, the
' approved version it would have replaced, which stayed in use all along;
' else 0.
' Errors: ERR_NOT_ALLOWED when the signed-in person may not write it or the
'   rules now refuse (the reason is the message); others pass up through
'   RaiseAgain after the rollback.
Public Function DeleteDraft(ByVal qid As Long) As Long
    Dim ws As DAO.Workspace, inTrans As Boolean, baseId As Long, ver As Long, problem As String, qkey As String
    Dim num As Long, desc As String, src As String
    RequireQuestionWriter qid
    Set ws = DBEngine.Workspaces(0)
    baseId = Nz(DLookup("ReplacesQuestionID", "tblQuestions", "QuestionID=" & qid), 0)
    ver = Nz(DLookup("VersionNo", "tblQuestions", "QuestionID=" & qid), 0)
    qkey = Nz(DLookup("QuestionKey", "tblQuestions", "QuestionID=" & qid), "")
    On Error GoTo Fail
    ws.BeginTrans: inTrans = True
    ' The delete rules again inside the transaction, and a delete that only
    ' matches a draft: it may have been submitted since the caller checked.
    problem = DeleteProblem(qid)
    If Len(problem) = 0 Then
        If KitExecute("DELETE FROM tblQuestions WHERE QuestionID=" & qid & " AND StatusID=" & QStatusID(QS_DRAFT)) = 0 Then
            problem = "It was moved on or deleted while you had it open."
        End If
    End If
    If Len(problem) > 0 Then
        ws.Rollback: inTrans = False
        Err.Raise ERR_NOT_ALLOWED, "modQuestions.DeleteDraft", "It can't be deleted now. " & problem
    End If
    ' The key in the comment: the row it names is gone, so the audit trail
    ' can't look it up any more (Admin: Logs).
    AddAudit ITEM_QUESTION, qid, ver, AU_DELETED, , "Draft Q" & qkey & " v" & ver & " deleted"
    ws.CommitTrans: inTrans = False
    If baseId <> 0 Then
        If Nz(DLookup("IsCurrent", "tblQuestions", "QuestionID=" & baseId), False) Then DeleteDraft = baseId
    End If
    Exit Function
Fail:
    num = Err.Number: desc = Err.Description: src = Err.Source
    If inTrans Then ws.Rollback
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modQuestions.DeleteDraft", qid
End Function

' ---------------------------------------------------------------------------
' References
' ---------------------------------------------------------------------------

' What is wrong with a reference's fields, one per line, or "" when fine:
' a title, the columns' sizes, and no other reference with the same title
' and number.
'   title, number, rev   the fields as typed (any may be Null)
'   exceptId             the reference being changed, 0 for a new one
' Errors: the engine's.
Public Function RefProblem(ByVal title As Variant, ByVal number As Variant, ByVal rev As Variant, _
        Optional ByVal exceptId As Long = 0) As String
    Dim t As String, n As String, s As String, dup As Variant
    t = Trim$(Nz(title, "")): n = Trim$(Nz(number, ""))
    If Len(t) = 0 Then s = s & "Title: give the reference's title." & vbCrLf
    If Len(t) > 200 Then s = s & "Title can be at most 200 characters; it is " & Len(t) & "." & vbCrLf
    If Len(n) > 100 Then s = s & "Number can be at most 100 characters; it is " & Len(n) & "." & vbCrLf
    If Len(Trim$(Nz(rev, ""))) > 50 Then s = s & "Revision can be at most 50 characters." & vbCrLf
    If Len(s) = 0 Then
        dup = DLookup("RefID", "tblSourceRefs", "RefTitle=" & Q(t) & " AND Nz(RefNumber,'')=" & Q(n) & _
            IIf(exceptId <> 0, " AND RefID<>" & exceptId, ""))
        If Not IsNull(dup) Then s = "There is already a reference """ & t & """" & IIf(Len(n) > 0, " (" & n & ")", "") & _
            ". Pick it from the list instead of adding it again." & vbCrLf
    End If
    If Right$(s, 2) = vbCrLf Then s = Left$(s, Len(s) - 2)
    RefProblem = s
End Function

' Add a reference to the shared list, turned on, with an Added audit row, in
' one transaction. Its fields must pass RefProblem.
'   title, number, rev   the fields as typed; an empty number or revision is
'                        saved as Null
' Returns the new RefID.
' Errors: ERR_NOT_ALLOWED unless the signed-in person is an SME or the admin
'   now; 3022 when the same title and number was added at the same moment;
'   others pass up through RaiseAgain after the rollback.
Public Function AddReference(ByVal title As Variant, ByVal number As Variant, ByVal rev As Variant) As Long
    Dim rs As DAO.Recordset, ws As DAO.Workspace, inTrans As Boolean, id As Long
    RequireReferenceWriter
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblSourceRefs WHERE False", dbOpenDynaset)
    rs.AddNew
    rs!RefTitle = Trim$(Nz(title, ""))
    rs!RefNumber = NullIfEmpty(number)
    rs!Revision = NullIfEmpty(rev)
    rs!IsActive = True
    rs.Update
    rs.Bookmark = rs.LastModified
    id = rs!RefID
    rs.Close: Set rs = Nothing
    AddAudit ITEM_REFERENCE, id, Null, AU_ADDED, , RefText(id)
    ws.CommitTrans: inTrans = False
    AddReference = id
    Exit Function
Fail:
    If Not rs Is Nothing Then rs.Close
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modQuestions.AddReference"
End Function

' Correct a reference's title, number and revision, with a Changed audit
' row holding the old and new text ("Was: ...; now: ..."), in one
' transaction. Every question citing it shows the change, approved ones
' included. The fields must pass RefProblem with exceptId refId.
'   refId                the reference's RefID
'   title, number, rev   the fields as typed
' Errors: ERR_NOT_ALLOWED unless the signed-in person is an SME or the admin
'   now; 3022 when another reference already has this title and number.
Public Sub ChangeReference(ByVal refId As Long, ByVal title As Variant, ByVal number As Variant, ByVal rev As Variant)
    Dim ws As DAO.Workspace, inTrans As Boolean, was As String
    RequireReferenceWriter
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    was = RefText(refId)
    KitExecute "UPDATE tblSourceRefs SET RefTitle=" & Q(Trim$(Nz(title, ""))) & ", RefNumber=" & SqlOrNull(number) & _
        ", Revision=" & SqlOrNull(rev) & " WHERE RefID=" & refId
    AddAudit ITEM_REFERENCE, refId, Null, AU_CHANGED, , "Was: " & was & "; now: " & RefText(refId)
    ws.CommitTrans: inTrans = False
    Exit Sub
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modQuestions.ChangeReference", refId
End Sub

' Turn a reference on (it can be cited) or off (it can't be cited again;
' questions citing it keep it), with a Turned on or Turned off audit row, in
' one transaction.
'   refId   the reference's RefID
'   isOn    True to turn it on
' Errors: ERR_NOT_ALLOWED unless the signed-in person is an SME or the admin
'   now; others pass up through RaiseAgain after the rollback.
Public Sub SetReferenceActive(ByVal refId As Long, ByVal isOn As Boolean)
    Dim ws As DAO.Workspace, inTrans As Boolean
    RequireReferenceWriter
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    KitExecute "UPDATE tblSourceRefs SET IsActive=" & IIf(isOn, "True", "False") & " WHERE RefID=" & refId
    AddAudit ITEM_REFERENCE, refId, Null, IIf(isOn, AU_TURNED_ON, AU_TURNED_OFF), , RefText(refId)
    ws.CommitTrans: inTrans = False
    Exit Sub
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modQuestions.SetReferenceActive", refId
End Sub

' Delete a reference no question cites, with a Deleted audit row that keeps
' its text, in one transaction.
'   refId   the reference's RefID
' Errors: ERR_NOT_ALLOWED unless the signed-in person is an SME or the admin
'   now; 3200 when a question cites it (the rQR_R relationship).
Public Sub DeleteReference(ByVal refId As Long)
    Dim ws As DAO.Workspace, inTrans As Boolean, was As String
    RequireReferenceWriter
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    was = RefText(refId)
    KitExecute "DELETE FROM tblSourceRefs WHERE RefID=" & refId
    AddAudit ITEM_REFERENCE, refId, Null, AU_DELETED, , was
    ws.CommitTrans: inTrans = False
    Exit Sub
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modQuestions.DeleteReference", refId
End Sub

' v as SQL: Q of its trimmed text, or Null when it is empty.
Private Function SqlOrNull(ByVal v As Variant) As String
    If IsNull(NullIfEmpty(v)) Then SqlOrNull = "Null" Else SqlOrNull = Q(NullIfEmpty(v))
End Function

' A typed value as it is saved: its text trimmed, or Null when it is empty
' or only spaces.
'   v   the value, or Null
' Errors: none.
Public Function NullIfEmpty(ByVal v As Variant) As Variant
    If Len(Trim$(Nz(v, ""))) = 0 Then NullIfEmpty = Null Else NullIfEmpty = Trim$(v)
End Function

' A reference as one line: "NSTM 555 (S9086-S3-STM-010) rev 3".
'   refId   the reference's RefID
' Returns "" when there is no such reference. Errors: the engine's.
Public Function RefText(ByVal refId As Long) As String
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT RefTitle, RefNumber, Revision FROM tblSourceRefs WHERE RefID=" & refId, dbOpenSnapshot)
    If Not rs.EOF Then RefText = rs!RefTitle & IIf(IsNull(rs!RefNumber), "", " (" & rs!RefNumber & ")") & _
        IIf(IsNull(rs!Revision), "", " rev " & rs!Revision)
    rs.Close
End Function
