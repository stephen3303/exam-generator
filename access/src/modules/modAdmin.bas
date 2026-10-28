Option Compare Database
Option Explicit

' Exam Generator (Access edition) - the admin data the admin pages share
' (build phase 4): settings, people and their roles, and the overlap rule's
' "topics short of questions" check, which the dashboard and the draw use
' too.
'
' The overlap rule (tblSettings OverlapMaxPct and OverlapWindowExams): at
' most OverlapMaxPct percent of an exam's questions in a topic may have been
' given in the last OverlapWindowExams exams. So of the n questions a topic
' gets, at least n - Int(n * pct / 100) must be fresh, and the fresh ones of
' each exam stay in the window for the next OverlapWindowExams exams. A topic
' can meet the rule exam after exam only with at least
'     (OverlapWindowExams + 1) * (n - Int(n * pct / 100))
' approved questions, where n is the most questions any active template
' gives it (QuestionsNeeded). A topic with fewer is short.
'
' Owner: the admin pages (build phase 4). Depends on: modErrors, modGlobals,
' modQuestions, modSignIn, modUI; the kit's modKitLocks, modPersonnel; and
' tblSettings keys (the SET_ constants).

Public Const SET_OVERLAP_PCT As String = "OverlapMaxPct"
Public Const SET_OVERLAP_WINDOW As String = "OverlapWindowExams"
Public Const SET_DEFAULT_COUNT As String = "DefaultQuestionCount"
Public Const SET_ADMIN_BOOTSTRAPPED As String = "AdminBootstrapped"   ' "1" once the first admin was made (modSignIn)
Public Const SET_CLASS_ORDER_AT As String = "ClassOrderChangedAt"    ' the newest AuditID when the admin last reordered the classification levels (modExamDoc)

' The settings' values when tblSettings has no row for them. The back end's
' seed rows (backend\modBuildBackEnd.bas, SeedData) start with the same values.
Public Const DEFAULT_OVERLAP_PCT As Long = 30
Public Const DEFAULT_OVERLAP_WINDOW As Long = 10
Public Const DEFAULT_QUESTION_COUNT As Long = 4

' Questions per topic on one exam: a topic's DefaultCount, a template topic's
' QuestionCount and the DefaultQuestionCount setting all keep to this range.
Public Const MIN_QUESTIONS_PER_TOPIC As Long = 1
Public Const MAX_QUESTIONS_PER_TOPIC As Long = 50

' The longest names the admin pages accept. Each mirrors its column's size in
' the back end (backend\modBuildBackEnd.bas, CreateTables).
Public Const MAX_TOPIC_NAME As Long = 100       ' tblTopics.TopicName TEXT(100)
Public Const MAX_TEMPLATE_NAME As Long = 100    ' tblTemplates.TemplateName TEXT(100)
Public Const MAX_LEVEL_NAME As Long = 30        ' tblDifficultyLevels.LevelName and tblClassLevels.Marking TEXT(30)
Public Const MAX_PORTION_MARK As Long = 10      ' tblClassLevels.PortionMark TEXT(10)

' SQL for a person's name as the pickers show it: "Smith, John (LT)".
'   p   the alias of the Personnel table (kit_person) in the query
' Returns an SQL expression to put in a SELECT. Errors: none.
Public Function PersonNameSql(Optional ByVal p As String = "p") As String
    PersonNameSql = p & ".last_name & ', ' & " & p & ".first_name & IIf(Nz(" & p & ".rank,'')='','',' (' & " & p & ".rank & ')')"
End Function

' ---------------------------------------------------------------------------
' Settings
' ---------------------------------------------------------------------------

' A tblSettings value as text.
'   key    SettingKey (SET_ constants)
'   dflt   what to return when there is no row or its value is Null
' Errors: the engine's, when tblSettings can't be read.
Public Function SettingText(ByVal key As String, Optional ByVal dflt As String = "") As String
    SettingText = Nz(DLookup("SettingValue", "tblSettings", "SettingKey=" & Q(key)), dflt)
End Function

' A tblSettings value as a whole number.
'   key    SettingKey (SET_ constants)
'   dflt   what to return when there is no row or the value isn't a number
' Errors: the engine's, when tblSettings can't be read.
Public Function SettingLong(ByVal key As String, ByVal dflt As Long) As Long
    Dim s As String
    s = SettingText(key)
    If IsNumeric(s) Then SettingLong = CLng(s) Else SettingLong = dflt
End Function

' Save setting key as value, adding the row when it isn't there yet.
'   key     SettingKey (SET_ constants)
'   value   the text to store
' Errors: ERR_NOT_ALLOWED unless the signed-in person is an admin now
'   (RequireExamAdmin); KitExecute's lock and write errors pass up.
' Side effects: writes tblSettings in the data file.
Public Sub SetSetting(ByVal key As String, ByVal value As String)
    RequireExamAdmin
    If DCount("*", "tblSettings", "SettingKey=" & Q(key)) = 0 Then
        KitExecute "INSERT INTO tblSettings (SettingKey, SettingValue) VALUES (" & Q(key) & ", " & Q(value) & ")"
    Else
        KitExecute "UPDATE tblSettings SET SettingValue=" & Q(value) & " WHERE SettingKey=" & Q(key)
    End If
End Sub

' Questions per topic a new topic or template topic starts with (setting
' DefaultQuestionCount, DEFAULT_QUESTION_COUNT when it isn't set).
' Errors: the engine's, when tblSettings can't be read.
Public Function DefaultQuestionCount() As Long
    DefaultQuestionCount = SettingLong(SET_DEFAULT_COUNT, DEFAULT_QUESTION_COUNT)
End Function

' The overlap rule's most percent of an exam's questions that may repeat
' (setting OverlapMaxPct, DEFAULT_OVERLAP_PCT when it isn't set).
' Errors: the engine's, when tblSettings can't be read.
Public Function OverlapPct() As Long
    OverlapPct = SettingLong(SET_OVERLAP_PCT, DEFAULT_OVERLAP_PCT)
End Function

' The overlap rule's look-back, in exams (setting OverlapWindowExams,
' DEFAULT_OVERLAP_WINDOW when it isn't set).
' Errors: the engine's, when tblSettings can't be read.
Public Function OverlapWindow() As Long
    OverlapWindow = SettingLong(SET_OVERLAP_WINDOW, DEFAULT_OVERLAP_WINDOW)
End Function

' ---------------------------------------------------------------------------
' The overlap rule
' ---------------------------------------------------------------------------

' Approved questions a topic needs to meet the overlap rule exam after exam
' (see the top of the module).
'   n        questions the topic gets on one exam
'   pct      OverlapPct
'   window   OverlapWindow
' Returns the count; 0 when n is 0 or less (a topic in no template).
' Errors: none.
Public Function QuestionsNeeded(ByVal n As Long, ByVal pct As Long, ByVal window As Long) As Long
    If n <= 0 Then Exit Function
    QuestionsNeeded = (window + 1) * (n - Int(n * pct / 100))
End Function

' Approved current questions in a topic (TopicSupplySql's on_hand for one
' topic).
'   topicId   tblTopics.TopicID
' Errors: the engine's.
Public Function ApprovedOnHand(ByVal topicId As Long) As Long
    ApprovedOnHand = DCount("*", "tblQuestions", "TopicID=" & topicId & " AND IsCurrent AND StatusID=" & ApprovedStatusID())
End Function

' The id of the question status Approved (QStatusID(QS_APPROVED)), 0 when
' the data file has none.
' Errors: the engine's, the first time the statuses are read.
Public Function ApprovedStatusID() As Long
    ApprovedStatusID = QStatusID(QS_APPROVED)
End Function

' SQL for every topic's supply: TopicID, TopicName, IsActive, on_hand
' (approved current questions) and per_exam (the most questions any active
' template gives it, 0 for none). Use it as a subquery.
' Errors: the engine's, when the status ids can't be read.
Public Function TopicSupplySql() As String
    TopicSupplySql = "SELECT t.TopicID, t.TopicName, t.IsActive, " & _
        "(SELECT Count(*) FROM tblQuestions AS q WHERE q.TopicID=t.TopicID AND q.IsCurrent AND q.StatusID=" & ApprovedStatusID() & ") AS on_hand, " & _
        "Nz((SELECT Max(tt.QuestionCount) FROM tblTemplateTopics AS tt INNER JOIN tblTemplates AS tm ON tt.TemplateID=tm.TemplateID " & _
        "WHERE tt.TopicID=t.TopicID AND tm.IsActive), 0) AS per_exam FROM tblTopics AS t"
End Function

' The active topics short of approved questions under the overlap rule, one
' line each ("Damage Control: 12 approved, needs 28"), then "and 3 more".
'   pct, window   OverlapPct and OverlapWindow (or values being tried)
'   nShort        returns how many topics are short
'   maxLines      the most topic lines before "and N more"
'   nameLen       cut a long name short at this many characters (ClipText),
'                 e.g. 32 so a one-line alert holds two lines; 0 keeps it whole
' Returns "" when none is short. Errors: the engine's.
Public Function TopicsShortText(ByVal pct As Long, ByVal window As Long, Optional ByRef nShort As Long, _
        Optional ByVal maxLines As Long = 4, Optional ByVal nameLen As Long = 0) As String
    Dim rs As DAO.Recordset, need As Long, s As String
    nShort = 0
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM (" & TopicSupplySql() & ") AS x WHERE x.IsActive ORDER BY x.TopicName", dbOpenSnapshot)
    Do While Not rs.EOF
        need = QuestionsNeeded(rs!per_exam, pct, window)
        If rs!on_hand < need Then
            nShort = nShort + 1
            If nShort <= maxLines Then
                s = s & IIf(Len(s) > 0, vbCrLf, "") & ClipText(Nz(rs!TopicName, ""), nameLen) & ": " & rs!on_hand & " approved, needs " & need
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    If nShort > maxLines Then s = s & vbCrLf & "and " & (nShort - maxLines) & " more"
    TopicsShortText = s
End Function

' ---------------------------------------------------------------------------
' People and roles
' ---------------------------------------------------------------------------

' A role's id.
'   code   tblRoles.RoleCode: SME, DH, XO or SM
' Returns 0 when the data file has no such role. Errors: the engine's.
Public Function RoleID(ByVal code As String) As Long
    RoleID = Nz(DLookup("RoleID", "tblRoles", "RoleCode=" & Q(code)), 0)
End Function

' True when a person holds a role in this data file.
'   pid    the person's person_id in the Personnel file
'   code   role code: SME, DH, XO or SM
' Errors: the engine's.
Public Function PersonHasRole(ByVal pid As Long, ByVal code As String) As Boolean
    PersonHasRole = DCount("*", "tblPersonRoles", "PersonID=" & pid & " AND RoleID=" & RoleID(code)) > 0
End Function

' Make sure a person can sign in to the Exam Generator: a tblPeople row
' (active, dashboard notices on) when they have none, and access to the app
' in the Personnel file when they lack it. Both writes are one transaction.
'   pid   the person's person_id in the Personnel file
' Errors: ERR_NOT_ALLOWED unless the signed-in person is an admin now;
'   KitExecute's and SetPersonApp's errors pass up after the rollback.
' Side effects: may write tblPeople and the Personnel file's kit_person_app.
Public Sub EnsurePerson(ByVal pid As Long)
    Dim ws As DAO.Workspace, inTrans As Boolean
    RequireExamAdmin
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    If DCount("*", "tblPeople", "PersonID=" & pid) = 0 Then
        KitExecute "INSERT INTO tblPeople (PersonID, IsAdmin, IsActive, NotifyDashboard, NotifyEmail) " & _
            "VALUES (" & pid & ", False, True, True, False)"
    End If
    If Not PersonCanUse(pid, APP_CODE) Then SetPersonApp pid, APP_CODE, True
    ws.CommitTrans: inTrans = False
    Exit Sub
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modAdmin.EnsurePerson", pid
End Sub

' Give a person an Exam Generator role, or take it away.
'   pid    the person's person_id in the Personnel file
'   code   role code: SME, DH, XO or SM (tblRoles.RoleCode)
'   hasIt  True to give the role, False to take it away
' Errors: ERR_NOT_ALLOWED unless the signed-in person is an admin now; 5
'   when the data file has no role code; KitExecute's lock and write errors
'   pass up (after the rollback when giving).
' Side effects: giving a role also runs EnsurePerson, in one transaction
'   with the role row. Taking SME away leaves their topics (the People page
'   removes those first).
' Example: SetPersonRole 1042, "DH", True
Public Sub SetPersonRole(ByVal pid As Long, ByVal code As String, ByVal hasIt As Boolean)
    Dim rid As Long, ws As DAO.Workspace, inTrans As Boolean
    RequireExamAdmin
    rid = RoleID(code)
    If rid = 0 Then Err.Raise 5, "modAdmin.SetPersonRole", "There is no role " & code & " in the data file."
    If hasIt Then
        On Error GoTo Fail
        Set ws = DBEngine.Workspaces(0)
        ws.BeginTrans: inTrans = True          ' the person, their access and the role together
        EnsurePerson pid
        If Not PersonHasRole(pid, code) Then KitExecute "INSERT INTO tblPersonRoles (PersonID, RoleID) VALUES (" & pid & ", " & rid & ")"
        ws.CommitTrans: inTrans = False
    Else
        KitExecute "DELETE FROM tblPersonRoles WHERE PersonID=" & pid & " AND RoleID=" & rid
    End If
    Exit Sub
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modAdmin.SetPersonRole", code & " " & pid
End Sub

' The active people who hold a role, "Smith, John (LT); Jones, Ann", or "".
'   code   role code: SME, DH, XO or SM
' Errors: the engine's.
Public Function RoleHolders(ByVal code As String) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT " & PersonNameSql() & " AS nm FROM (tblPersonRoles AS pr INNER JOIN tblPeople AS e " & _
        "ON pr.PersonID=e.PersonID) INNER JOIN " & PERSON_TABLE & " AS p ON e.PersonID=p.person_id " & _
        "WHERE e.IsActive AND pr.RoleID=" & RoleID(code) & " ORDER BY p.last_name, p.first_name", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & IIf(Len(s) > 0, "; ", "") & rs!nm
        rs.MoveNext
    Loop
    rs.Close
    RoleHolders = s
End Function

' Exam Generator admins who can sign in now: admin and active in tblPeople,
' active in the Personnel file, with access to the app there. An admin turned
' off in the Personnel file doesn't count, so the last-admin checks keep one
' who can actually use the app.
' Errors: the engine's.
Public Function AdminCount() As Long
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM (tblPeople AS e INNER JOIN " & PERSON_TABLE & " AS p " & _
        "ON e.PersonID=p.person_id) INNER JOIN " & PERSON_APP_TABLE & " AS a ON e.PersonID=a.person_id " & _
        "WHERE e.IsAdmin AND e.IsActive AND p.is_active AND a.is_active AND a.app_code=" & Q(APP_CODE), dbOpenSnapshot)
    AdminCount = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function

' The role codes a person holds, in chain order, "SME,DH", or "" for none.
' The admin flag is not a role code (tblPeople.IsAdmin).
'   pid   the person's person_id in the Personnel file
' Errors: the engine's.
Public Function PersonRoleCodes(ByVal pid As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT r.RoleCode FROM tblPersonRoles AS pr INNER JOIN tblRoles AS r " & _
        "ON pr.RoleID=r.RoleID WHERE pr.PersonID=" & pid & " ORDER BY r.ChainOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & IIf(Len(s) > 0, ",", "") & rs!RoleCode
        rs.MoveNext
    Loop
    rs.Close
    PersonRoleCodes = s
End Function

' A person's roles for the people list: "Admin, SME, DH", or "No role yet".
'   pid   the person's person_id in the Personnel file
' Errors: the engine's. Called once per row by the list's SQL.
Public Function PersonRolesText(ByVal pid As Long) As String
    Dim s As String, codes As String
    If Nz(DLookup("IsAdmin", "tblPeople", "PersonID=" & pid), False) Then s = "Admin"
    codes = Replace(PersonRoleCodes(pid), ",", ", ")
    If Len(codes) > 0 Then s = s & IIf(Len(s) > 0, ", ", "") & codes
    If Len(s) = 0 Then s = "No role yet"
    PersonRolesText = s
End Function

' The topics a person is an SME for, "Damage Control, Firefighting", or "".
'   pid   the person's person_id in the Personnel file
' Errors: the engine's.
Public Function PersonTopicsText(ByVal pid As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT t.TopicName FROM tblTopicSMEs AS ts INNER JOIN tblTopics AS t " & _
        "ON ts.TopicID=t.TopicID WHERE ts.PersonID=" & pid & " ORDER BY t.TopicName", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & IIf(Len(s) > 0, ", ", "") & rs!TopicName
        rs.MoveNext
    Loop
    rs.Close
    PersonTopicsText = s
End Function

' A topic's SMEs for the topics list: "Smith, Jones", "Smith and 2 more",
' or "None yet".
'   topicId   tblTopics.TopicID
' Errors: the engine's. Called once per row by the list's SQL.
Public Function TopicSMEsText(ByVal topicId As Long) As String
    Dim rs As DAO.Recordset, s As String, n As Long
    Set rs = CurrentDb.OpenRecordset("SELECT p.last_name FROM tblTopicSMEs AS ts INNER JOIN " & PERSON_TABLE & _
        " AS p ON ts.PersonID=p.person_id WHERE ts.TopicID=" & topicId & " ORDER BY p.last_name, p.first_name", dbOpenSnapshot)
    Do While Not rs.EOF
        n = n + 1
        If n <= 2 Then s = s & IIf(Len(s) > 0, ", ", "") & rs!last_name
        rs.MoveNext
    Loop
    rs.Close
    If n = 0 Then
        s = "None yet"
    ElseIf n > 2 Then
        s = Left$(s, InStr(s, ",") - 1) & " and " & (n - 1) & " more"
    End If
    TopicSMEsText = s
End Function

' ---------------------------------------------------------------------------
' Ordered lists
' ---------------------------------------------------------------------------

' Move a row up or down one place in an ordered list. The rows are numbered
' 1, 2, 3... in their present order first, so ties and gaps sort themselves
' out, and the renumbering is one transaction.
'   tbl, idField, orderField   the table, its key field and its order field
'                              (plain names: they go into SQL as they are)
'   id      the row's key
'   delta   -1 for up, 1 for down
'   where   the list's rows when the table holds several lists: "" for all,
'           "field=<whole number>" or "field Is Null"
' Returns False when the row is already at that end (nothing changes).
' Errors: 5 when a name or where isn't of that shape; KitExecute's errors
'   pass up after the rollback.
' Side effects: rewrites orderField for every row of the list.
Public Function MoveRow(ByVal tbl As String, ByVal idField As String, ByVal orderField As String, ByVal id As Long, _
        ByVal delta As Long, Optional ByVal where As String = "") As Boolean
    Dim rs As DAO.Recordset, ids() As Long, n As Long, i As Long, at As Long
    Dim ws As DAO.Workspace, inTrans As Boolean
    CheckSqlParts "modAdmin.MoveRow", where, tbl, idField, orderField
    Set rs = CurrentDb.OpenRecordset("SELECT " & idField & " FROM " & tbl & IIf(Len(where) > 0, " WHERE " & where, "") & _
        " ORDER BY " & orderField & ", " & idField, dbOpenSnapshot)
    at = -1
    Do While Not rs.EOF
        ReDim Preserve ids(n)
        ids(n) = rs.Fields(0).Value
        If ids(n) = id Then at = n
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    If at < 0 Or at + delta < 0 Or at + delta >= n Then Exit Function
    ids(at) = ids(at + delta): ids(at + delta) = id
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    For i = 0 To n - 1
        KitExecute "UPDATE " & tbl & " SET " & orderField & "=" & (i + 1) & " WHERE " & idField & "=" & ids(i) & _
            IIf(Len(where) > 0, " AND " & where, "")
    Next
    ws.CommitTrans: inTrans = False
    MoveRow = True
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modAdmin.MoveRow", tbl & " " & id
End Function

' The order number for a row added at the end of a list.
'   tbl, orderField   the table and its order field (plain names)
'   where             as for MoveRow: "", "field=<n>" or "field Is Null"
' Returns the highest order number in the list plus 1 (1 for an empty list).
' Errors: 5 when a name or where isn't of that shape.
Public Function NextOrder(ByVal tbl As String, ByVal orderField As String, Optional ByVal where As String = "") As Long
    CheckSqlParts "modAdmin.NextOrder", where, tbl, orderField
    If Len(where) > 0 Then
        NextOrder = Nz(DMax(orderField, tbl, where), 0) + 1
    Else
        NextOrder = Nz(DMax(orderField, tbl), 0) + 1
    End If
End Function

' What is wrong with a name for a list entry, or "" when it is fine: it must
' be filled in, at most maxLen characters, and unique in the table.
'   nm            the name as typed (any text; it goes through Q)
'   maxLen        the column's size (MAX_ constants)
'   tbl, fld      the table and its name field (plain names)
'   what          the thing, for the message: "topic"
'   exceptWhere   the row being renamed, "field=<n>", or "" for a new one
' Errors: 5 when a name or exceptWhere isn't of the shape MoveRow takes.
Public Function NameProblem(ByVal nm As String, ByVal maxLen As Long, ByVal tbl As String, ByVal fld As String, _
        ByVal what As String, Optional ByVal exceptWhere As String = "") As String
    CheckSqlParts "modAdmin.NameProblem", exceptWhere, tbl, fld
    nm = Trim$(nm)
    If Len(nm) = 0 Then
        NameProblem = "Give the " & what & " a name."
    ElseIf Len(nm) > maxLen Then
        NameProblem = "A " & what & "'s name can be at most " & maxLen & " characters; this one is " & Len(nm) & "."
    ElseIf DCount("*", tbl, fld & "=" & Q(nm) & IIf(Len(exceptWhere) > 0, " AND NOT (" & exceptWhere & ")", "")) > 0 Then
        NameProblem = "There is already a " & what & " called """ & nm & """."
    End If
End Function

' The ordered-list helpers put table and field names and a where test into
' SQL as they are. Every caller passes constants and Long ids today; this
' keeps it so. Raise error 5 (from proc) unless each of names is a plain
' name (a letter or _, then letters, digits or _) and where is "", one
' "name=<whole number>" or one "name Is Null".
Private Sub CheckSqlParts(ByVal proc As String, ByVal where As String, ParamArray names() As Variant)
    Dim nm As Variant, eq As Long, ok As Boolean
    For Each nm In names
        If Not IsPlainName(CStr(nm)) Then Err.Raise 5, proc, "Not a plain table or field name: " & nm
    Next
    If Len(where) = 0 Then Exit Sub
    eq = InStr(where, "=")
    If eq > 0 Then
        ok = IsPlainName(Left$(where, eq - 1)) And IsWholeNumber(Mid$(where, eq + 1))
    ElseIf Right$(where, 8) = " Is Null" Then
        ok = IsPlainName(Left$(where, Len(where) - 8))
    End If
    If Not ok Then Err.Raise 5, proc, "Not one simple test: " & where
End Sub

Private Function IsPlainName(ByVal s As String) As Boolean
    IsPlainName = (s Like "[A-Za-z_]*") And Not (s Like "*[!A-Za-z0-9_]*")
End Function

Private Function IsWholeNumber(ByVal s As String) As Boolean
    IsWholeNumber = Len(s) > 0 And Len(s) <= 10 And Not (s Like "*[!0-9]*")
End Function
