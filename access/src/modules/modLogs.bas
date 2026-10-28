Option Compare Database
Option Explicit

' Exam Generator (Access edition) - the Admin: Logs page's lists (build
' phase 10): the audit trail (tblAuditTrail: who authored, edited,
' submitted, reviewed, approved, sent back, generated and downloaded what)
' and the error log (error_log, written by modErrors). Both are read only
' here; the error log can be cleared (modErrors.ClearErrorLog), the audit
' trail never.
'
' Each list is the newest LOG_MAX rows that fit the search box and the
' dates, newest first. A day range is whole days, both included; a missing
' date leaves that end open.
'
' Owner: build phase 10. Depends on: modAdmin (PersonNameSql), modExams
' (ITEM_EXAM), modGlobals (Q, LikeLiteral), modQuestions (ITEM_QUESTION,
' ITEM_REFERENCE); the kit's PERSON_TABLE.

Public Const LOG_MAX As Long = 1000             ' rows a list shows at most

' The audit trail: id (AuditID), when_txt, ActionOn, who, ActionType, item
' ("Q12 v2", "E3 v1", "Reference: NSTM 555"), step (the role it was done
' as) and note (the comment's first line).
'   search   words to find in the who, action, item and comment ("" for all)
'   fromDay, toDay   whole days, both included; Null leaves that end open
' Builds the SQL only. Errors: none.
Public Function AuditLogSql(ByVal search As String, ByVal fromDay As Variant, ByVal toDay As Variant) As String
    Dim w As String
    w = DayWhere("a.ActionOn", fromDay, toDay)
    If Len(search) > 0 Then
        w = w & " AND (" & WhoSql() & " & ' ' & a.ActionType & ' ' & " & AuditItemSql() & " & ' ' & Nz(a.Comments,'')) LIKE " & _
            Q("*" & LikeLiteral(search) & "*")
    End If
    AuditLogSql = "SELECT TOP " & LOG_MAX & " a.AuditID AS id, Format(a.ActionOn,'d mmm yy h:nn') AS when_txt, a.ActionOn, " & _
        WhoSql() & " AS who, a.ActionType, " & AuditItemSql() & " AS item, " & _
        "r.RoleCode AS step, Left(Replace(Nz(a.Comments,''), Chr(13) & Chr(10), ' '), 120) AS note " & AuditFromSql() & _
        " WHERE " & Mid$(w, 6) & " ORDER BY a.ActionOn DESC, a.AuditID DESC"
End Function

' Everything one audit row holds, as text for the detail box and Copy; ""
' when there is no such row.
' Errors: the engine's.
Public Function AuditDetail(ByVal id As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT a.*, " & WhoSql() & " AS who, " & AuditItemSql() & " AS item, r.RoleCode AS step, " & _
        "tp.TopicName, tm.TemplateName " & AuditFromSql(True) & " WHERE a.AuditID=" & id, dbOpenSnapshot)
    If Not rs.EOF Then
        s = "Audit entry " & rs!AuditID & "  -  " & Format$(rs!ActionOn, "d mmm yyyy h:nn:ss") & vbCrLf & _
            rs!ActionType & ": " & rs!item & vbCrLf
        If Not IsNull(rs!TopicName) And rs!ItemType = ITEM_QUESTION Then s = s & "Topic: " & rs!TopicName & vbCrLf
        If Not IsNull(rs!TemplateName) And rs!ItemType = ITEM_EXAM Then s = s & "Template: " & rs!TemplateName & vbCrLf
        s = s & "By: " & rs!who & IIf(IsNull(rs!step), "", " as " & rs!step) & vbCrLf
        If Len(Nz(rs!Comments, "")) > 0 Then s = s & vbCrLf & rs!Comments
    End If
    rs.Close
    AuditDetail = s
End Function

' The error log: id, when_txt, logged_at, who, page, source (the procedure)
' and err_txt (the number and the description's first 120 characters).
'   search, fromDay, toDay   as AuditLogSql: the words are looked for in
'                            the description, procedure, page and user
' Builds the SQL only. Errors: none.
Public Function ErrorListSql(ByVal search As String, ByVal fromDay As Variant, ByVal toDay As Variant) As String
    Dim w As String
    w = DayWhere("l.logged_at", fromDay, toDay)
    If Len(search) > 0 Then
        w = w & " AND (l.error_number & ' ' & Nz(l.description,'') & ' ' & Nz(l.source,'') & ' ' & Nz(l.page,'') & ' ' & " & _
            "Nz(l.username,'') & ' ' & Nz(l.windows_user,'')) LIKE " & Q("*" & LikeLiteral(search) & "*")
    End If
    ErrorListSql = "SELECT TOP " & LOG_MAX & " l.id, Format(l.logged_at,'d mmm yy h:nn') AS when_txt, l.logged_at, " & _
        "IIf(Len(l.username & '')>0, l.username, '(' & l.windows_user & ')') AS who, l.page, l.source, " & _
        "l.error_number & ': ' & Left(Replace(Nz(l.description,''), Chr(13) & Chr(10), ' '), 120) AS err_txt " & _
        "FROM error_log AS l WHERE " & Mid$(w, 6) & " ORDER BY l.id DESC"
End Function

' How many rows a list's filters match in all, to say when LOG_MAX cut it.
' Errors: the engine's.
Public Function AuditLogCount(ByVal search As String, ByVal fromDay As Variant, ByVal toDay As Variant) As Long
    AuditLogCount = CountRows(Replace(AuditLogSql(search, fromDay, toDay), "SELECT TOP " & LOG_MAX & " ", "SELECT ", , 1))
End Function

Public Function ErrorListCount(ByVal search As String, ByVal fromDay As Variant, ByVal toDay As Variant) As Long
    ErrorListCount = CountRows(Replace(ErrorListSql(search, fromDay, toDay), "SELECT TOP " & LOG_MAX & " ", "SELECT ", , 1))
End Function

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

' Who did it, as SQL: their name, or "(person 12)" when the Personnel file
' no longer has them.
Private Function WhoSql() As String
    WhoSql = "IIf(IsNull(p.person_id), '(person ' & a.PersonID & ')', " & PersonNameSql() & ")"
End Function

' The audit row's item as SQL: a question's key and version, an exam's id
' and version, a reference's title, else its type and id. The joins ignore
' ItemType (Jet can't join on a constant); the IIf reads each table only on
' its own type's rows.
Private Function AuditItemSql() As String
    AuditItemSql = "IIf(a.ItemType=" & Q(ITEM_QUESTION) & ", 'Q' & q.QuestionKey & ' v' & a.VersionNo, " & _
        "IIf(a.ItemType=" & Q(ITEM_EXAM) & ", 'E' & a.ItemID & ' v' & a.VersionNo, " & _
        "IIf(a.ItemType=" & Q(ITEM_REFERENCE) & ", 'Reference: ' & Nz(s.RefTitle,'(deleted)'), a.ItemType & ' ' & a.ItemID)))"
End Function

' The audit list's FROM clause; withNames adds the question's topic and the
' exam's template.
Private Function AuditFromSql(Optional ByVal withNames As Boolean = False) As String
    Dim f As String
    f = "(((tblAuditTrail AS a LEFT JOIN " & PERSON_TABLE & " AS p ON a.PersonID=p.person_id) " & _
        "LEFT JOIN tblQuestions AS q ON a.ItemID=q.QuestionID) LEFT JOIN tblSourceRefs AS s ON a.ItemID=s.RefID) " & _
        "LEFT JOIN tblRoles AS r ON a.StepRoleID=r.RoleID"
    If withNames Then
        f = "(((" & f & ") LEFT JOIN tblTopics AS tp ON q.TopicID=tp.TopicID) LEFT JOIN tblExams AS e ON a.ItemID=e.ExamID) " & _
            "LEFT JOIN tblTemplates AS tm ON e.TemplateID=tm.TemplateID"
    End If
    AuditFromSql = "FROM " & f
End Function

' " AND <field> from fromDay to toDay", always starting " AND " (" AND True"
' for no limit), so callers can strip it with Mid$(w, 6).
Private Function DayWhere(ByVal fld As String, ByVal fromDay As Variant, ByVal toDay As Variant) As String
    Dim w As String
    If IsDate(fromDay) Then w = w & " AND " & fld & ">=" & Format$(DateValue(fromDay), "\#yyyy\-mm\-dd\#")
    If IsDate(toDay) Then w = w & " AND " & fld & "<" & Format$(DateAdd("d", 1, DateValue(toDay)), "\#yyyy\-mm\-dd\#")
    If Len(w) = 0 Then w = " AND True"
    DayWhere = w
End Function

Private Function CountRows(ByVal sql As String) As Long
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM (" & Left$(sql, InStr(sql, " ORDER BY ") - 1) & ")", dbOpenSnapshot)
    CountRows = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function
