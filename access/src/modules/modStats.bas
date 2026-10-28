Option Compare Database
Option Explicit

' Exam Generator (Access edition) - the Statistics page's figures (build
' phase 9): how many questions each topic gave over a period, and how much of
' each topic came from each reference, to show weak reference coverage.
'
' What counts: the questions on exams the SM gave final approval to (status
' Approved) between two dates, by FinalApprovedOn, from all templates or one
' (there is no series field, so a template stands in for a series). A
' question given on three exams counts three times. A question counts once
' for each reference it cites, so a topic's reference shares (questions
' citing it / questions given in the topic) can add up past 100%.
'
' RunStats works the figures out into three local tables in this front end,
' which the page's tables read:
'   ws_given   one row per question given (ExamID, TopicID, QuestionID)
'   ws_topic   one row per topic: id (TopicID), TopicName, nq (questions
'              given), nx (exams), nref (references cited), top_ref and top_n
'              (the reference cited most and how many questions cite it)
'   ws_ref     one row per topic and reference: TopicID, RefID, n
'
' Depends on: modExams (ExamStatusID), modGlobals (TableExists), modUI
' (ClipText).

Public Const WEAK_SHARE As Double = 0.8          ' one reference this share of a topic's questions or more: weak coverage
Public Const WEAK_MIN_QS As Long = 5            ' fewer questions given than this says too little to flag

Private Const ALERT_TOPIC_LEN As Long = 24      ' a weak coverage line's topic name and reference are cut to
Private Const ALERT_REF_LEN As Long = 28        ' these lengths (ClipText), so two lines fit the alert's one line

' Make the local tables when this front end lacks them: BuildAll's "local
' tables" step, and the page before it runs.
' Errors: the engine's.
' Side effects: creates tables in the front end.
Public Sub EnsureStatTables()
    If Not TableExists("ws_given", True) Then
        CurrentDb.Execute "CREATE TABLE ws_given (ExamID LONG, TopicID LONG, QuestionID LONG)", dbFailOnError
    End If
    If Not TableExists("ws_topic", True) Then
        CurrentDb.Execute "CREATE TABLE ws_topic (id LONG CONSTRAINT pk_ws_topic PRIMARY KEY, TopicName TEXT(100), nq LONG, " & _
            "nx LONG, nref LONG, top_ref LONG, top_n LONG)", dbFailOnError
    End If
    If Not TableExists("ws_ref", True) Then
        CurrentDb.Execute "CREATE TABLE ws_ref (TopicID LONG, RefID LONG, n LONG)", dbFailOnError
    End If
End Sub

' Work out the figures for exams finally approved from fromDay to toDay
' (whole days, both included), from template tplId (0: every template),
' into ws_given, ws_topic and ws_ref (see the header).
' Errors: the engine's.
' Side effects: rewrites the three local tables.
Public Sub RunStats(ByVal fromDay As Date, ByVal toDay As Date, ByVal tplId As Long)
    Dim db As DAO.Database, rs As DAO.Recordset, top As DAO.Recordset
    Set db = CurrentDb
    EnsureStatTables
    db.Execute "DELETE FROM ws_given", dbFailOnError
    db.Execute "DELETE FROM ws_topic", dbFailOnError
    db.Execute "DELETE FROM ws_ref", dbFailOnError
    db.Execute "INSERT INTO ws_given (ExamID, TopicID, QuestionID) SELECT eq.ExamID, et.TopicID, eq.QuestionID " & _
        "FROM (tblExamQuestions AS eq INNER JOIN tblExamTopics AS et ON eq.ExamTopicID=et.ExamTopicID) " & _
        "INNER JOIN tblExams AS e ON eq.ExamID=e.ExamID WHERE e.StatusID=" & ExamStatusID(ES_APPROVED) & _
        " AND e.FinalApprovedOn>=" & DayLiteral(fromDay) & " AND e.FinalApprovedOn<" & DayLiteral(DateAdd("d", 1, toDay)) & _
        IIf(tplId <> 0, " AND e.TemplateID=" & tplId, ""), dbFailOnError
    db.Execute "INSERT INTO ws_topic (id, TopicName, nq, nx, nref, top_n) SELECT g.TopicID, First(tp.TopicName), Count(*), 0, 0, 0 " & _
        "FROM ws_given AS g LEFT JOIN tblTopics AS tp ON g.TopicID=tp.TopicID GROUP BY g.TopicID", dbFailOnError
    db.Execute "INSERT INTO ws_ref (TopicID, RefID, n) SELECT g.TopicID, qr.RefID, Count(*) FROM ws_given AS g " & _
        "INNER JOIN tblQuestionRefs AS qr ON g.QuestionID=qr.QuestionID GROUP BY g.TopicID, qr.RefID", dbFailOnError
    ' The exams and references per topic, and its reference cited most (Jet
    ' can't update from a grouped query, so row by row).
    Set rs = db.OpenRecordset("SELECT id, nx, nref, top_ref, top_n FROM ws_topic", dbOpenDynaset)
    Do While Not rs.EOF
        Set top = db.OpenRecordset("SELECT TOP 1 RefID, n FROM ws_ref WHERE TopicID=" & rs!id & " ORDER BY n DESC, RefID", dbOpenSnapshot)
        rs.Edit
        rs!nx = CountOf("SELECT Count(*) FROM (SELECT DISTINCT ExamID FROM ws_given WHERE TopicID=" & rs!id & ")")
        rs!nref = CountOf("SELECT Count(*) FROM ws_ref WHERE TopicID=" & rs!id)
        If Not top.EOF Then rs!top_ref = top!RefID: rs!top_n = top!n
        rs.Update
        top.Close
        rs.MoveNext
    Loop
    rs.Close
End Sub

' The figures' totals: exams, questions given and references cited.
' Errors: the engine's.
Public Function StatExams() As Long
    StatExams = CountOf("SELECT Count(*) FROM (SELECT DISTINCT ExamID FROM ws_given)")
End Function

Public Function StatQuestions() As Long
    StatQuestions = CountOf("SELECT Count(*) FROM ws_given")
End Function

Public Function StatReferences() As Long
    StatReferences = CountOf("SELECT Count(*) FROM (SELECT DISTINCT RefID FROM ws_ref)")
End Function

' The topics table: id (TopicID), TopicName, nq, nx, nref, top_ref_txt (the
' reference cited most), top_pct (its share, 0 to 1) and top_txt ("86%"),
' by topic name. A topic whose questions cite no reference has no top
' reference: top_pct is Null and top_txt a dash.
' Builds the SQL only. Errors: none.
Public Function StatTopicsSql() As String
    StatTopicsSql = "SELECT t.id, t.TopicName, t.nq, t.nx, t.nref, " & RefTextSql("s") & " AS top_ref_txt, " & _
        "IIf(t.nq=0 Or t.top_n=0,Null,t.top_n/t.nq) AS top_pct, IIf(t.nq=0 Or t.top_n=0,'-',Format(t.top_n/t.nq,'0%')) AS top_txt " & _
        "FROM ws_topic AS t LEFT JOIN tblSourceRefs AS s ON t.top_ref=s.RefID ORDER BY t.TopicName"
End Function

' Topic topicId's references: id (RefID), ref (title and number), n
' (questions citing it), pct (its share of the topic's questions given) and
' pct_txt, the most cited first.
' Builds the SQL only. Errors: none.
Public Function StatRefsSql(ByVal topicId As Long) As String
    StatRefsSql = "SELECT r.RefID AS id, " & RefTextSql("s") & " AS ref, r.n, IIf(t.nq=0,0,r.n/t.nq) AS pct, " & _
        "Format(IIf(t.nq=0,0,r.n/t.nq),'0%') AS pct_txt FROM (ws_ref AS r INNER JOIN ws_topic AS t ON r.TopicID=t.id) " & _
        "LEFT JOIN tblSourceRefs AS s ON r.RefID=s.RefID WHERE r.TopicID=" & topicId & " ORDER BY r.n DESC, s.RefTitle"
End Function

' The topics whose questions came WEAK_SHARE or more from one reference
' (with WEAK_MIN_QS questions given or more), one line each ("Damage
' Control: 86% cite NSTM 555", long names cut short), then "and 3 more";
' "" when none.
'   nWeak   returns how many topics are weak
' Errors: the engine's.
Public Function WeakTopicsText(Optional ByRef nWeak As Long, Optional ByVal maxLines As Long = 3) As String
    Dim rs As DAO.Recordset, s As String
    nWeak = 0
    Set rs = CurrentDb.OpenRecordset("SELECT t.TopicName, t.top_n/t.nq AS share, " & RefTextSql("s") & " AS rt " & _
        "FROM ws_topic AS t LEFT JOIN tblSourceRefs AS s ON t.top_ref=s.RefID WHERE t.nq>=" & WEAK_MIN_QS & _
        " AND t.top_n>=" & Str$(WEAK_SHARE) & "*t.nq ORDER BY t.top_n/t.nq DESC, t.TopicName", dbOpenSnapshot)
    Do While Not rs.EOF
        nWeak = nWeak + 1
        If nWeak <= maxLines Then
            s = s & IIf(Len(s) > 0, vbCrLf, "") & ClipText(Nz(rs!TopicName, "A topic"), ALERT_TOPIC_LEN) & ": " & _
                Format$(rs!share, "0%") & " cite " & ClipText(Nz(rs!rt, "one reference"), ALERT_REF_LEN)
        End If
        rs.MoveNext
    Loop
    rs.Close
    If nWeak > maxLines Then s = s & vbCrLf & "and " & (nWeak - maxLines) & " more"
    WeakTopicsText = s
End Function

' The template picker's rows: id (0 for every template) and nm, the turned
' off ones too (their old exams still count), every template first.
' Builds the SQL only. Errors: none.
Public Function StatTemplatesSql() As String
    StatTemplatesSql = "SELECT 0 AS id, '(Every template)' AS nm, 0 AS k FROM tblSettings WHERE SettingKey='SchemaVersion' " & _
        "UNION ALL SELECT TemplateID, TemplateName & IIf(IsActive,'',' (turned off)'), 1 FROM tblTemplates ORDER BY k, nm"
End Function

' A reference's title and number as SQL, from tblSourceRefs as alias a.
Private Function RefTextSql(ByVal a As String) As String
    RefTextSql = a & ".RefTitle & IIf(IsNull(" & a & ".RefNumber),'',' (' & " & a & ".RefNumber & ')')"
End Function

' A day as a Jet date literal (#2026-10-28#), whatever the computer's date format.
Private Function DayLiteral(ByVal d As Date) As String
    DayLiteral = Format$(d, "\#yyyy\-mm\-dd\#")
End Function

Private Function CountOf(ByVal sql As String) As Long
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    CountOf = Nz(rs.Fields(0).Value, 0)
    rs.Close
End Function
