Attribute VB_Name = "modBuildBackEnd"
Option Compare Database
Option Explicit
' Exam Generator - back-end builder.
' Creates the shared data file (ExamGen_BE.accdb) with all tables, relationships,
' indexes and seed lookup data. Two ways in:
'   BuildBackEnd          run by hand from any Access database; asks for the path
'                         (default: same folder) and whether to overwrite.
'   BuildBackEndAt(path)  run by backend\new-backend.ps1; asks nothing, never
'                         overwrites, and returns "OK" or what failed.
'
' People: names, accounts and email live once in the kit's shared Personnel file
' (kit_person). This file keeps only each person's Exam Generator settings and
' roles, keyed on the Personnel person_id (tblPeople, tblPersonRoles), and every
' "who" column holds a person_id. No people are seeded: the first Personnel
' admin to open the app with no Exam Generator admin becomes its admin (sign-in,
' build phase 2).
'
' help_topic holds the help text the admin edits in the app (seeded from
' help.txt the first time, build phase 3); its columns match the kit's zz_help.
' error_log is the kit's error log (modErrors), so admins see every user's errors.
'
' The schema version (tblSettings SchemaVersion) is the app's DATA_VERSION
' (modSignIn). A change here bumps both and adds a step to modSignIn.UpgradeData,
' which brings an existing data file up to date. 2: tblExamTopics.ChangeNo.

Private db As DAO.Database

Public Sub BuildBackEnd()
    Dim p As String
    p = InputBox("Back-end path:", "Exam Generator", CurrentProject.Path & "\ExamGen_BE.accdb")
    If Len(p) = 0 Then Exit Sub
    If Len(Dir(p)) > 0 Then
        If MsgBox(p & " exists. Overwrite?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
        Kill p
    End If
    Set db = DBEngine.CreateDatabase(p, dbLangGeneral, dbVersion120)
    On Error GoTo Fail
    CreateAll
    db.Close: Set db = Nothing
    MsgBox "Back end created:" & vbCrLf & p, vbInformation
    Exit Sub
Fail:
    MsgBox "Build failed: " & Err.Description, vbCritical
    If Not db Is Nothing Then db.Close: Set db = Nothing
End Sub

' Script entry point: create the data file at path (a full path to a new
' .accdb). Returns "OK", or "ERROR: <description>" with nothing left behind.
Public Function BuildBackEndAt(ByVal p As String) As String
    On Error GoTo Fail
    If Len(Dir(p)) > 0 Then Err.Raise 58, , "There is already a file at " & p & "; nothing was changed."
    Set db = DBEngine.CreateDatabase(p, dbLangGeneral, dbVersion120)
    CreateAll
    db.Close: Set db = Nothing
    BuildBackEndAt = "OK"
    Exit Function
Fail:
    BuildBackEndAt = "ERROR: " & Err.Description
    If Not db Is Nothing Then
        db.Close: Set db = Nothing
        On Error Resume Next
        Kill p                                  ' don't leave a half-made file
    End If
End Function

Private Sub CreateAll()
    CreateTables
    CreateIndexes
    CreateRelationships
    SeedData
End Sub

Private Sub X(sql As String)
    db.Execute sql, dbFailOnError
End Sub

Private Sub CreateTables()
    ' ---- Lookups ----
    X "CREATE TABLE tblRoles (RoleID AUTOINCREMENT PRIMARY KEY, RoleCode TEXT(10) NOT NULL, RoleName TEXT(50) NOT NULL, ChainOrder INTEGER)"
    X "CREATE TABLE tblDifficultyLevels (DifficultyID AUTOINCREMENT PRIMARY KEY, LevelName TEXT(30) NOT NULL, SortOrder INTEGER NOT NULL, IsActive YESNO)"
    X "CREATE TABLE tblClassLevels (ClassLevelID AUTOINCREMENT PRIMARY KEY, Marking TEXT(30) NOT NULL, PortionMark TEXT(10) NOT NULL, LevelRank INTEGER NOT NULL)"
    X "CREATE TABLE tblQuestionTypes (QuestionTypeID AUTOINCREMENT PRIMARY KEY, TypeName TEXT(30) NOT NULL)"
    X "CREATE TABLE tblStatuses (StatusID AUTOINCREMENT PRIMARY KEY, StatusName TEXT(40) NOT NULL, AppliesTo TEXT(10) NOT NULL)"
    X "CREATE TABLE tblSettings (SettingKey TEXT(50) PRIMARY KEY, SettingValue TEXT(255), Description TEXT(255))"

    ' ---- People (keyed on the Personnel file's person_id) ----
    X "CREATE TABLE tblPeople (PersonID LONG PRIMARY KEY, IsAdmin YESNO, IsActive YESNO, NotifyDashboard YESNO, NotifyEmail YESNO)"
    X "CREATE TABLE tblPersonRoles (PersonID LONG NOT NULL, RoleID LONG NOT NULL, CONSTRAINT pkPersonRoles PRIMARY KEY (PersonID, RoleID))"

    ' ---- Topics & templates ----
    X "CREATE TABLE tblTopics (TopicID AUTOINCREMENT PRIMARY KEY, TopicName TEXT(100) NOT NULL, DefaultCount INTEGER, Description MEMO, IsActive YESNO)"
    X "CREATE TABLE tblTopicSMEs (TopicID LONG NOT NULL, PersonID LONG NOT NULL, CONSTRAINT pkTopicSMEs PRIMARY KEY (TopicID, PersonID))"
    X "CREATE TABLE tblTemplates (TemplateID AUTOINCREMENT PRIMARY KEY, TemplateName TEXT(100) NOT NULL, Description MEMO, IsActive YESNO)"
    X "CREATE TABLE tblTemplateTopics (TemplateID LONG NOT NULL, TopicID LONG NOT NULL, QuestionCount INTEGER NOT NULL, SortOrder INTEGER, " & _
      "CONSTRAINT pkTemplateTopics PRIMARY KEY (TemplateID, TopicID))"

    ' ---- Question bank (versioned) ----
    X "CREATE TABLE tblQuestions (QuestionID AUTOINCREMENT PRIMARY KEY, QuestionKey LONG, VersionNo INTEGER NOT NULL, ReplacesQuestionID LONG, " & _
      "TopicID LONG NOT NULL, QuestionTypeID LONG NOT NULL, DifficultyID LONG, ClassLevelID LONG, QuestionText MEMO NOT NULL, " & _
      "TotalPoints DOUBLE NOT NULL, ReferenceAnswer MEMO, StatusID LONG NOT NULL, CurrentStepRoleID LONG, " & _
      "AuthorPersonID LONG NOT NULL, CreatedOn DATETIME, IsCurrent YESNO, RetiredOn DATETIME)"
    X "CREATE TABLE tblQuestionParts (PartID AUTOINCREMENT PRIMARY KEY, QuestionID LONG NOT NULL, PartLabel TEXT(5) NOT NULL, PartText MEMO NOT NULL, " & _
      "Points DOUBLE NOT NULL, ClassLevelID LONG NOT NULL, ReferenceAnswer MEMO, SortOrder INTEGER)"
    X "CREATE TABLE tblMCOptions (OptionID AUTOINCREMENT PRIMARY KEY, QuestionID LONG, PartID LONG, OptionLabel TEXT(2), OptionText MEMO NOT NULL, IsCorrect YESNO, SortOrder INTEGER)"
    X "CREATE TABLE tblSourceRefs (RefID AUTOINCREMENT PRIMARY KEY, RefTitle TEXT(200) NOT NULL, RefNumber TEXT(100), Revision TEXT(50), IsActive YESNO)"
    X "CREATE TABLE tblQuestionRefs (QuestionID LONG NOT NULL, RefID LONG NOT NULL, Paragraph TEXT(100), CONSTRAINT pkQuestionRefs PRIMARY KEY (QuestionID, RefID))"

    ' ---- Exams ----
    ' tblExamTopics.ChangeNo counts the swaps and the sign-off on the topic, so a
    ' swap or sign-off only goes through on the questions the SME was shown.
    X "CREATE TABLE tblExams (ExamID AUTOINCREMENT PRIMARY KEY, TemplateID LONG NOT NULL, RecipientName TEXT(150) NOT NULL, VersionNo INTEGER, " & _
      "StatusID LONG NOT NULL, CurrentStepRoleID LONG, OverallClassLevelID LONG, CreatedByPersonID LONG NOT NULL, CreatedOn DATETIME, " & _
      "FinalApprovedOn DATETIME, OutputFile MEMO)"
    X "CREATE TABLE tblExamTopics (ExamTopicID AUTOINCREMENT PRIMARY KEY, ExamID LONG NOT NULL, TopicID LONG NOT NULL, QuestionCount INTEGER NOT NULL, " & _
      "SignedOffByPersonID LONG, SignedOffOn DATETIME, ChangeNo LONG)"
    X "CREATE TABLE tblExamQuestions (ExamQuestionID AUTOINCREMENT PRIMARY KEY, ExamID LONG NOT NULL, ExamTopicID LONG NOT NULL, QuestionID LONG NOT NULL, " & _
      "SortOrder INTEGER, WasSwapped YESNO)"
    X "CREATE TABLE tblTasks (TaskID AUTOINCREMENT PRIMARY KEY, TaskType TEXT(20) NOT NULL, ItemType TEXT(10) NOT NULL, ItemID LONG NOT NULL, " & _
      "ExamTopicID LONG, AssignedPersonID LONG NOT NULL, TaskStatus TEXT(20) NOT NULL, CreatedOn DATETIME, CompletedOn DATETIME, CompletedByPersonID LONG)"

    ' ---- Help text (admin-editable; same columns as the kit's zz_help) ----
    X "CREATE TABLE help_topic (topic_key TEXT(32) CONSTRAINT pk_help_topic PRIMARY KEY, title TEXT(80), sort_order LONG, " & _
      "body MEMO, search_text MEMO, updated_at DATETIME, updated_by TEXT(150))"

    ' ---- Audit trail (append-only by app) ----
    X "CREATE TABLE tblAuditTrail (AuditID AUTOINCREMENT PRIMARY KEY, ItemType TEXT(10) NOT NULL, ItemID LONG NOT NULL, VersionNo INTEGER, " & _
      "ActionType TEXT(20) NOT NULL, StepRoleID LONG, PersonID LONG NOT NULL, ActionOn DATETIME NOT NULL, Comments MEMO)"

    ' ---- Error log (the kit's modErrors writes it) ----
    X "CREATE TABLE error_log (id COUNTER CONSTRAINT pk_error_log PRIMARY KEY, logged_at DATETIME, " & _
      "error_number LONG, description LONGTEXT, err_source TEXT(255), source TEXT(255), call_trail LONGTEXT, " & _
      "context LONGTEXT, page TEXT(32), active_form TEXT(64), active_control TEXT(64), username TEXT(150), " & _
      "windows_user TEXT(64), computer TEXT(64), fe_build TEXT(32), access_version TEXT(32))"
End Sub

Private Sub CreateIndexes()
    X "CREATE UNIQUE INDEX uxTopicName ON tblTopics (TopicName)"
    X "CREATE UNIQUE INDEX uxTemplateName ON tblTemplates (TemplateName)"
    X "CREATE UNIQUE INDEX uxQKeyVer ON tblQuestions (QuestionKey, VersionNo)"
    X "CREATE INDEX ixQTopicStatus ON tblQuestions (TopicID, StatusID, IsCurrent)"
    X "CREATE UNIQUE INDEX uxPartLabel ON tblQuestionParts (QuestionID, PartLabel)"
    X "CREATE UNIQUE INDEX uxExamTopic ON tblExamTopics (ExamID, TopicID)"
    X "CREATE UNIQUE INDEX uxExamQ ON tblExamQuestions (ExamID, QuestionID)"
    X "CREATE INDEX ixTaskPerson ON tblTasks (AssignedPersonID, TaskStatus)"
    X "CREATE INDEX ixAuditItem ON tblAuditTrail (ItemType, ItemID)"
    X "CREATE INDEX ixExamApproved ON tblExams (FinalApprovedOn)"
    X "CREATE INDEX ixErrorLogged ON error_log (logged_at)"
End Sub

Private Sub Rel(nm As String, pt As String, pk As String, ct As String, fk As String, Optional cascadeDel As Boolean = False)
    Dim r As DAO.Relation
    Set r = db.CreateRelation(nm, pt, ct, IIf(cascadeDel, dbRelationDeleteCascade, 0))
    r.Fields.Append r.CreateField(pk)
    r.Fields(pk).ForeignName = fk
    db.Relations.Append r
End Sub

Private Sub CreateRelationships()
    Rel "rPR_P", "tblPeople", "PersonID", "tblPersonRoles", "PersonID", True
    Rel "rPR_R", "tblRoles", "RoleID", "tblPersonRoles", "RoleID"
    Rel "rTS_T", "tblTopics", "TopicID", "tblTopicSMEs", "TopicID", True
    Rel "rTS_P", "tblPeople", "PersonID", "tblTopicSMEs", "PersonID", True
    Rel "rTT_Tm", "tblTemplates", "TemplateID", "tblTemplateTopics", "TemplateID", True
    Rel "rTT_T", "tblTopics", "TopicID", "tblTemplateTopics", "TopicID"
    Rel "rQ_T", "tblTopics", "TopicID", "tblQuestions", "TopicID"
    Rel "rQ_Ty", "tblQuestionTypes", "QuestionTypeID", "tblQuestions", "QuestionTypeID"
    Rel "rQ_D", "tblDifficultyLevels", "DifficultyID", "tblQuestions", "DifficultyID"
    Rel "rQ_C", "tblClassLevels", "ClassLevelID", "tblQuestions", "ClassLevelID"
    Rel "rQ_S", "tblStatuses", "StatusID", "tblQuestions", "StatusID"
    Rel "rQ_A", "tblPeople", "PersonID", "tblQuestions", "AuthorPersonID"
    Rel "rQP_Q", "tblQuestions", "QuestionID", "tblQuestionParts", "QuestionID", True
    Rel "rQP_C", "tblClassLevels", "ClassLevelID", "tblQuestionParts", "ClassLevelID"
    Rel "rMC_Q", "tblQuestions", "QuestionID", "tblMCOptions", "QuestionID", True
    Rel "rMC_P", "tblQuestionParts", "PartID", "tblMCOptions", "PartID", True
    Rel "rQR_Q", "tblQuestions", "QuestionID", "tblQuestionRefs", "QuestionID", True
    Rel "rQR_R", "tblSourceRefs", "RefID", "tblQuestionRefs", "RefID"
    Rel "rE_Tm", "tblTemplates", "TemplateID", "tblExams", "TemplateID"
    Rel "rE_S", "tblStatuses", "StatusID", "tblExams", "StatusID"
    Rel "rE_P", "tblPeople", "PersonID", "tblExams", "CreatedByPersonID"
    Rel "rET_E", "tblExams", "ExamID", "tblExamTopics", "ExamID", True
    Rel "rET_T", "tblTopics", "TopicID", "tblExamTopics", "TopicID"
    Rel "rEQ_E", "tblExams", "ExamID", "tblExamQuestions", "ExamID", True
    Rel "rEQ_ET", "tblExamTopics", "ExamTopicID", "tblExamQuestions", "ExamTopicID"
    Rel "rEQ_Q", "tblQuestions", "QuestionID", "tblExamQuestions", "QuestionID"
    Rel "rTk_P", "tblPeople", "PersonID", "tblTasks", "AssignedPersonID"
    Rel "rAT_P", "tblPeople", "PersonID", "tblAuditTrail", "PersonID"
End Sub

Private Sub SeedData()
    X "INSERT INTO tblRoles (RoleCode, RoleName, ChainOrder) VALUES ('SME','Subject Matter Expert',1)"
    X "INSERT INTO tblRoles (RoleCode, RoleName, ChainOrder) VALUES ('DH','Department Head',2)"
    X "INSERT INTO tblRoles (RoleCode, RoleName, ChainOrder) VALUES ('XO','Executive Officer',3)"
    X "INSERT INTO tblRoles (RoleCode, RoleName, ChainOrder) VALUES ('SM','Senior Member',4)"

    X "INSERT INTO tblDifficultyLevels (LevelName, SortOrder, IsActive) VALUES ('Easy',1,True)"
    X "INSERT INTO tblDifficultyLevels (LevelName, SortOrder, IsActive) VALUES ('Medium',2,True)"
    X "INSERT INTO tblDifficultyLevels (LevelName, SortOrder, IsActive) VALUES ('Hard',3,True)"

    ' Placeholder levels - confirm list and order (open item).
    X "INSERT INTO tblClassLevels (Marking, PortionMark, LevelRank) VALUES ('UNCLASSIFIED','(U)',1)"
    X "INSERT INTO tblClassLevels (Marking, PortionMark, LevelRank) VALUES ('CUI','(CUI)',2)"
    X "INSERT INTO tblClassLevels (Marking, PortionMark, LevelRank) VALUES ('CONFIDENTIAL','(C)',3)"
    X "INSERT INTO tblClassLevels (Marking, PortionMark, LevelRank) VALUES ('SECRET','(S)',4)"

    X "INSERT INTO tblQuestionTypes (TypeName) VALUES ('Multiple Choice')"
    X "INSERT INTO tblQuestionTypes (TypeName) VALUES ('True/False')"
    X "INSERT INTO tblQuestionTypes (TypeName) VALUES ('Short Answer')"

    Dim s
    For Each s In Array("Draft", "In Review", "Approved", "Rejected", "Retired")
        X "INSERT INTO tblStatuses (StatusName, AppliesTo) VALUES ('" & s & "','Question')"
    Next
    For Each s In Array("SME Review", "In Review", "Approved", "Rejected")
        X "INSERT INTO tblStatuses (StatusName, AppliesTo) VALUES ('" & s & "','Exam')"
    Next

    X "INSERT INTO tblSettings VALUES ('OverlapMaxPct','30','Max % of an exam''s questions that may appear in the lookback window')"
    X "INSERT INTO tblSettings VALUES ('OverlapWindowExams','10','Number of most recent exams checked for overlap')"
    X "INSERT INTO tblSettings VALUES ('DefaultQuestionCount','4','Default questions per topic in a template')"
    X "INSERT INTO tblSettings VALUES ('SchemaVersion','2','Back-end schema version')"
End Sub
