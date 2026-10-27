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

Public Const SET_OVERLAP_PCT As String = "OverlapMaxPct"
Public Const SET_OVERLAP_WINDOW As String = "OverlapWindowExams"
Public Const SET_DEFAULT_COUNT As String = "DefaultQuestionCount"

' The person's name as the pickers show it: "Smith, John (LT)". p is the
' alias of the Personnel table (kit_person) in the query.
Public Function PersonNameSql(Optional ByVal p As String = "p") As String
    PersonNameSql = p & ".last_name & ', ' & " & p & ".first_name & IIf(Nz(" & p & ".rank,'')='','',' (' & " & p & ".rank & ')')"
End Function

' ---------------------------------------------------------------------------
' Settings
' ---------------------------------------------------------------------------

Public Function SettingText(ByVal key As String, Optional ByVal dflt As String = "") As String
    SettingText = Nz(DLookup("SettingValue", "tblSettings", "SettingKey=" & Q(key)), dflt)
End Function

Public Function SettingLong(ByVal key As String, ByVal dflt As Long) As Long
    Dim s As String
    s = SettingText(key)
    If IsNumeric(s) Then SettingLong = CLng(s) Else SettingLong = dflt
End Function

Public Sub SetSetting(ByVal key As String, ByVal value As String)
    If DCount("*", "tblSettings", "SettingKey=" & Q(key)) = 0 Then
        KitExecute "INSERT INTO tblSettings (SettingKey, SettingValue) VALUES (" & Q(key) & ", " & Q(value) & ")"
    Else
        KitExecute "UPDATE tblSettings SET SettingValue=" & Q(value) & " WHERE SettingKey=" & Q(key)
    End If
End Sub

' Questions per topic for a new topic or template topic.
Public Function DefaultQuestionCount() As Long
    DefaultQuestionCount = SettingLong(SET_DEFAULT_COUNT, 4)
End Function

' ---------------------------------------------------------------------------
' The overlap rule
' ---------------------------------------------------------------------------

' Approved questions a topic that gets n questions an exam needs to meet the
' rule (see the top of the module). 0 for a topic in no template.
Public Function QuestionsNeeded(ByVal n As Long, ByVal pct As Long, ByVal window As Long) As Long
    If n <= 0 Then Exit Function
    QuestionsNeeded = (window + 1) * (n - Int(n * pct / 100))
End Function

' The id of the question status Approved.
Public Function ApprovedStatusID() As Long
    ApprovedStatusID = Nz(DLookup("StatusID", "tblStatuses", "StatusName='Approved' AND AppliesTo='Question'"), 0)
End Function

' Approved questions on hand in each topic, and the most questions any active
' template gives it: TopicID, TopicName, IsActive, on_hand, per_exam.
Public Function TopicSupplySql() As String
    TopicSupplySql = "SELECT t.TopicID, t.TopicName, t.IsActive, " & _
        "(SELECT Count(*) FROM tblQuestions AS q WHERE q.TopicID=t.TopicID AND q.IsCurrent AND q.StatusID=" & ApprovedStatusID() & ") AS on_hand, " & _
        "Nz((SELECT Max(tt.QuestionCount) FROM tblTemplateTopics AS tt INNER JOIN tblTemplates AS tm ON tt.TemplateID=tm.TemplateID " & _
        "WHERE tt.TopicID=t.TopicID AND tm.IsActive), 0) AS per_exam FROM tblTopics AS t"
End Function

' The active topics short of approved questions under pct and window, one
' line each, "Damage Control: 12 approved, needs 28", at most maxLines then
' "and 3 more". "" when none is short. nShort returns how many are.
Public Function TopicsShortText(ByVal pct As Long, ByVal window As Long, Optional ByRef nShort As Long, _
        Optional ByVal maxLines As Long = 4) As String
    Dim rs As DAO.Recordset, need As Long, s As String
    nShort = 0
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM (" & TopicSupplySql() & ") AS x WHERE x.IsActive ORDER BY x.TopicName", dbOpenSnapshot)
    Do While Not rs.EOF
        need = QuestionsNeeded(rs!per_exam, pct, window)
        If rs!on_hand < need Then
            nShort = nShort + 1
            If nShort <= maxLines Then
                s = s & IIf(Len(s) > 0, vbCrLf, "") & rs!TopicName & ": " & rs!on_hand & " approved, needs " & need
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

Public Function RoleID(ByVal code As String) As Long
    RoleID = Nz(DLookup("RoleID", "tblRoles", "RoleCode=" & Q(code)), 0)
End Function

Public Function PersonHasRole(ByVal pid As Long, ByVal code As String) As Boolean
    PersonHasRole = DCount("*", "tblPersonRoles", "PersonID=" & pid & " AND RoleID=" & RoleID(code)) > 0
End Function

' Give pid a row in tblPeople if they have none (active, dashboard notices
' on), and access to the Exam Generator in the Personnel file, so a role
' given to them lets them sign in.
Public Sub EnsurePerson(ByVal pid As Long)
    If DCount("*", "tblPeople", "PersonID=" & pid) = 0 Then
        KitExecute "INSERT INTO tblPeople (PersonID, IsAdmin, IsActive, NotifyDashboard, NotifyEmail) " & _
            "VALUES (" & pid & ", False, True, True, False)"
    End If
    If Not PersonCanUse(pid, APP_CODE) Then SetPersonApp pid, APP_CODE, True
End Sub

' Give pid role code, or take it away.
Public Sub SetPersonRole(ByVal pid As Long, ByVal code As String, ByVal hasIt As Boolean)
    Dim rid As Long
    rid = RoleID(code)
    If rid = 0 Then Err.Raise 5, "modAdmin.SetPersonRole", "There is no role " & code & " in the data file."
    If hasIt Then
        EnsurePerson pid
        If Not PersonHasRole(pid, code) Then KitExecute "INSERT INTO tblPersonRoles (PersonID, RoleID) VALUES (" & pid & ", " & rid & ")"
    Else
        KitExecute "DELETE FROM tblPersonRoles WHERE PersonID=" & pid & " AND RoleID=" & rid
    End If
End Sub

' The people who hold role code and are active, "Smith, John (LT); Jones,
' Ann", or "".
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

' Active Exam Generator admins.
Public Function AdminCount() As Long
    AdminCount = DCount("*", "tblPeople", "IsAdmin AND IsActive")
End Function

' The roles pid holds, for the people list: "Admin, SME, DH", or "No role
' yet".
Public Function PersonRolesText(ByVal pid As Long) As String
    Dim rs As DAO.Recordset, s As String
    If Nz(DLookup("IsAdmin", "tblPeople", "PersonID=" & pid), False) Then s = "Admin"
    Set rs = CurrentDb.OpenRecordset("SELECT r.RoleCode FROM tblPersonRoles AS pr INNER JOIN tblRoles AS r " & _
        "ON pr.RoleID=r.RoleID WHERE pr.PersonID=" & pid & " ORDER BY r.ChainOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & IIf(Len(s) > 0, ", ", "") & rs!RoleCode
        rs.MoveNext
    Loop
    rs.Close
    If Len(s) = 0 Then s = "No role yet"
    PersonRolesText = s
End Function

' The topics pid is an SME for, "Damage Control, Firefighting", or "".
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

' A topic's SMEs for the topics list: "Smith, Jones", "Smith and 2 more", or
' "None yet".
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

' Move row id of tbl (key field idField, order field orderField) up (delta
' -1) or down (1) a place among the rows matching where ("" for all). The
' rows are numbered 1, 2, 3... in their present order first, so ties and
' gaps sort themselves out. False when the row is already at that end.
Public Function MoveRow(ByVal tbl As String, ByVal idField As String, ByVal orderField As String, ByVal id As Long, _
        ByVal delta As Long, Optional ByVal where As String = "") As Boolean
    Dim rs As DAO.Recordset, ids() As Long, n As Long, i As Long, at As Long
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
    For i = 0 To n - 1
        KitExecute "UPDATE " & tbl & " SET " & orderField & "=" & (i + 1) & " WHERE " & idField & "=" & ids(i) & _
            IIf(Len(where) > 0, " AND " & where, "")
    Next
    MoveRow = True
End Function

' The next order number for a row added at the end of tbl's rows matching
' where.
Public Function NextOrder(ByVal tbl As String, ByVal orderField As String, Optional ByVal where As String = "") As Long
    If Len(where) > 0 Then
        NextOrder = Nz(DMax(orderField, tbl, where), 0) + 1
    Else
        NextOrder = Nz(DMax(orderField, tbl), 0) + 1
    End If
End Function

' What is wrong with name for a list entry of at most maxLen characters,
' which must be unique in field of tbl (other than the row exceptWhere
' picks, "" for none), or "". what names the thing: "topic".
Public Function NameProblem(ByVal nm As String, ByVal maxLen As Long, ByVal tbl As String, ByVal fld As String, _
        ByVal what As String, Optional ByVal exceptWhere As String = "") As String
    nm = Trim$(nm)
    If Len(nm) = 0 Then
        NameProblem = "Give the " & what & " a name."
    ElseIf Len(nm) > maxLen Then
        NameProblem = "A " & what & "'s name can be at most " & maxLen & " characters; this one is " & Len(nm) & "."
    ElseIf DCount("*", tbl, fld & "=" & Q(nm) & IIf(Len(exceptWhere) > 0, " AND NOT (" & exceptWhere & ")", "")) > 0 Then
        NameProblem = "There is already a " & what & " called """ & nm & """."
    End If
End Function
