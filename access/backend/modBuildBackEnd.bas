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
' admin to open the app while tblPeople is empty becomes its admin, once
' (modSignIn, tblSettings AdminBootstrapped).
'
' help_topic holds the help text the admin edits in the app (seeded from
' help.txt the first time, build phase 3); its columns match the kit's zz_help.
' error_log is the kit's error log (modErrors), so admins see every user's errors.
'
' The schema version (tblSettings SchemaVersion) is the app's DATA_VERSION
' (modSignIn). A change here bumps both and adds a step to modSignIn.UpgradeData,
' which brings an existing data file up to date. 2: tblExamTopics.ChangeNo.
' 3: the unique indexes and relationships from uxLevelName and rQ_Step on
' (CreateIndexes, CreateRelationships). 4: ixAuditOn and ixAuditPerson.

' Make the data file by hand from any Access database: asks for the path
' (default: the database's own folder) and whether to replace a file there.
' The new file is built under a temporary name next to it and only takes the
' old one's place once it is complete, so a failure leaves any old file as it
' was and no half-made file behind.
Public Sub BuildBackEnd()
    Dim p As String, tmp As String, db As DAO.Database, built As Boolean, msg As String
    On Error GoTo Fail
    p = InputBox("Back-end path:", "Exam Generator", CurrentProject.Path & "\ExamGen_BE.accdb")
    If Len(p) = 0 Then Exit Sub
    If Len(Dir(p)) > 0 Then
        If MsgBox(p & " exists. Overwrite?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    End If
    If InStrRev(p, ".") > InStrRev(p, "\") Then
        tmp = Left$(p, InStrRev(p, ".") - 1) & "_new" & Mid$(p, InStrRev(p, "."))
    Else
        tmp = p & "_new"
    End If
    If Len(Dir(tmp)) > 0 Then Kill tmp      ' left by an earlier failed run
    Set db = DBEngine.CreateDatabase(tmp, dbLangGeneral, dbVersion120)
    CreateAll db
    db.Close: Set db = Nothing
    built = True
    If Len(Dir(p)) > 0 Then Kill p
    Name tmp As p
    MsgBox "Back end created:" & vbCrLf & p, vbInformation
    Exit Sub
Fail:
    msg = "Build failed: " & Err.Description
    On Error Resume Next
    If Not db Is Nothing Then db.Close: Set db = Nothing
    If built Then                           ' complete, but the old file couldn't be replaced
        msg = msg & vbCrLf & "The new file is complete at " & tmp & ". Rename it to " & p & " once nothing has the old one open."
    Else
        If Len(tmp) > 0 Then Kill tmp
        msg = msg & vbCrLf & "Nothing was changed."
    End If
    MsgBox msg, vbCritical
End Sub

' Script entry point: create the data file at path (a full path to a new
' .accdb). Returns "OK", or "ERROR: <description>" with nothing left behind.
Public Function BuildBackEndAt(ByVal p As String) As String
    Dim db As DAO.Database
    On Error GoTo Fail
    If Len(Dir(p)) > 0 Then Err.Raise 58, , "There is already a file at " & p & "; nothing was changed."
    Set db = DBEngine.CreateDatabase(p, dbLangGeneral, dbVersion120)
    CreateAll db
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

Private Sub CreateAll(db As DAO.Database)
    CreateTables db
    CreateIndexes db
    CreateRelationships db
    SeedData db
End Sub

' Run one DDL or seed statement on the new file.
Private Sub ExecDdl(db As DAO.Database, ByVal sql As String)
    db.Execute sql, dbFailOnError
End Sub

Private Sub CreateTables(db As DAO.Database)
    ' ---- Lookups ----
    ExecDdl db, "CREATE TABLE tblRoles (RoleID AUTOINCREMENT PRIMARY KEY, RoleCode TEXT(10) NOT NULL, RoleName TEXT(50) NOT NULL, ChainOrder INTEGER)"
    ExecDdl db, "CREATE TABLE tblDifficultyLevels (DifficultyID AUTOINCREMENT PRIMARY KEY, LevelName TEXT(30) NOT NULL, SortOrder INTEGER NOT NULL, IsActive YESNO)"
    ExecDdl db, "CREATE TABLE tblClassLevels (ClassLevelID AUTOINCREMENT PRIMARY KEY, Marking TEXT(30) NOT NULL, PortionMark TEXT(10) NOT NULL, LevelRank INTEGER NOT NULL)"
    ExecDdl db, "CREATE TABLE tblQuestionTypes (QuestionTypeID AUTOINCREMENT PRIMARY KEY, TypeName TEXT(30) NOT NULL)"
    ExecDdl db, "CREATE TABLE tblStatuses (StatusID AUTOINCREMENT PRIMARY KEY, StatusName TEXT(40) NOT NULL, AppliesTo TEXT(10) NOT NULL)"
    ExecDdl db, "CREATE TABLE tblSettings (SettingKey TEXT(50) PRIMARY KEY, SettingValue TEXT(255), Description TEXT(255))"

    ' ---- People (keyed on the Personnel file's person_id) ----
    ExecDdl db, "CREATE TABLE tblPeople (PersonID LONG PRIMARY KEY, IsAdmin YESNO, IsActive YESNO, NotifyDashboard YESNO, NotifyEmail YESNO)"
    ExecDdl db, "CREATE TABLE tblPersonRoles (PersonID LONG NOT NULL, RoleID LONG NOT NULL, CONSTRAINT pkPersonRoles PRIMARY KEY (PersonID, RoleID))"

    ' ---- Topics & templates ----
    ExecDdl db, "CREATE TABLE tblTopics (TopicID AUTOINCREMENT PRIMARY KEY, TopicName TEXT(100) NOT NULL, DefaultCount INTEGER, Description MEMO, IsActive YESNO)"
    ExecDdl db, "CREATE TABLE tblTopicSMEs (TopicID LONG NOT NULL, PersonID LONG NOT NULL, CONSTRAINT pkTopicSMEs PRIMARY KEY (TopicID, PersonID))"
    ExecDdl db, "CREATE TABLE tblTemplates (TemplateID AUTOINCREMENT PRIMARY KEY, TemplateName TEXT(100) NOT NULL, Description MEMO, IsActive YESNO)"
    ExecDdl db, "CREATE TABLE tblTemplateTopics (TemplateID LONG NOT NULL, TopicID LONG NOT NULL, QuestionCount INTEGER NOT NULL, SortOrder INTEGER, " & _
      "CONSTRAINT pkTemplateTopics PRIMARY KEY (TemplateID, TopicID))"

    ' ---- Question bank (versioned) ----
    ExecDdl db, "CREATE TABLE tblQuestions (QuestionID AUTOINCREMENT PRIMARY KEY, QuestionKey LONG, VersionNo INTEGER NOT NULL, ReplacesQuestionID LONG, " & _
      "TopicID LONG NOT NULL, QuestionTypeID LONG NOT NULL, DifficultyID LONG, ClassLevelID LONG, QuestionText MEMO NOT NULL, " & _
      "TotalPoints DOUBLE NOT NULL, ReferenceAnswer MEMO, StatusID LONG NOT NULL, CurrentStepRoleID LONG, " & _
      "AuthorPersonID LONG NOT NULL, CreatedOn DATETIME, IsCurrent YESNO, RetiredOn DATETIME)"
    ExecDdl db, "CREATE TABLE tblQuestionParts (PartID AUTOINCREMENT PRIMARY KEY, QuestionID LONG NOT NULL, PartLabel TEXT(5) NOT NULL, PartText MEMO NOT NULL, " & _
      "Points DOUBLE NOT NULL, ClassLevelID LONG NOT NULL, ReferenceAnswer MEMO, SortOrder INTEGER)"
    ExecDdl db, "CREATE TABLE tblMCOptions (OptionID AUTOINCREMENT PRIMARY KEY, QuestionID LONG, PartID LONG, OptionLabel TEXT(2), OptionText MEMO NOT NULL, IsCorrect YESNO, SortOrder INTEGER)"
    ExecDdl db, "CREATE TABLE tblSourceRefs (RefID AUTOINCREMENT PRIMARY KEY, RefTitle TEXT(200) NOT NULL, RefNumber TEXT(100), Revision TEXT(50), IsActive YESNO)"
    ExecDdl db, "CREATE TABLE tblQuestionRefs (QuestionID LONG NOT NULL, RefID LONG NOT NULL, Paragraph TEXT(100), CONSTRAINT pkQuestionRefs PRIMARY KEY (QuestionID, RefID))"

    ' ---- Exams ----
    ' tblExamTopics.ChangeNo counts the swaps and the sign-off on the topic, so a
    ' swap or sign-off only goes through on the questions the SME was shown.
    ExecDdl db, "CREATE TABLE tblExams (ExamID AUTOINCREMENT PRIMARY KEY, TemplateID LONG NOT NULL, RecipientName TEXT(150) NOT NULL, VersionNo INTEGER, " & _
      "StatusID LONG NOT NULL, CurrentStepRoleID LONG, OverallClassLevelID LONG, CreatedByPersonID LONG NOT NULL, CreatedOn DATETIME, " & _
      "FinalApprovedOn DATETIME, OutputFile MEMO)"
    ExecDdl db, "CREATE TABLE tblExamTopics (ExamTopicID AUTOINCREMENT PRIMARY KEY, ExamID LONG NOT NULL, TopicID LONG NOT NULL, QuestionCount INTEGER NOT NULL, " & _
      "SignedOffByPersonID LONG, SignedOffOn DATETIME, ChangeNo LONG)"
    ExecDdl db, "CREATE TABLE tblExamQuestions (ExamQuestionID AUTOINCREMENT PRIMARY KEY, ExamID LONG NOT NULL, ExamTopicID LONG NOT NULL, QuestionID LONG NOT NULL, " & _
      "SortOrder INTEGER, WasSwapped YESNO)"
    ExecDdl db, "CREATE TABLE tblTasks (TaskID AUTOINCREMENT PRIMARY KEY, TaskType TEXT(20) NOT NULL, ItemType TEXT(10) NOT NULL, ItemID LONG NOT NULL, " & _
      "ExamTopicID LONG, AssignedPersonID LONG NOT NULL, TaskStatus TEXT(20) NOT NULL, CreatedOn DATETIME, CompletedOn DATETIME, CompletedByPersonID LONG)"

    ' ---- Help text (admin-editable; same columns as the kit's zz_help) ----
    ExecDdl db, "CREATE TABLE help_topic (topic_key TEXT(32) CONSTRAINT pk_help_topic PRIMARY KEY, title TEXT(80), sort_order LONG, " & _
      "body MEMO, search_text MEMO, updated_at DATETIME, updated_by TEXT(150))"

    ' ---- Audit trail (append-only by app) ----
    ExecDdl db, "CREATE TABLE tblAuditTrail (AuditID AUTOINCREMENT PRIMARY KEY, ItemType TEXT(10) NOT NULL, ItemID LONG NOT NULL, VersionNo INTEGER, " & _
      "ActionType TEXT(20) NOT NULL, StepRoleID LONG, PersonID LONG NOT NULL, ActionOn DATETIME NOT NULL, Comments MEMO)"

    ' ---- Error log (the kit's modErrors writes it) ----
    ExecDdl db, "CREATE TABLE error_log (id COUNTER CONSTRAINT pk_error_log PRIMARY KEY, logged_at DATETIME, " & _
      "error_number LONG, description LONGTEXT, err_source TEXT(255), source TEXT(255), call_trail LONGTEXT, " & _
      "context LONGTEXT, page TEXT(32), active_form TEXT(64), active_control TEXT(64), username TEXT(150), " & _
      "windows_user TEXT(64), computer TEXT(64), fe_build TEXT(32), access_version TEXT(32))"
End Sub

Private Sub CreateIndexes(db As DAO.Database)
    ExecDdl db, "CREATE UNIQUE INDEX uxTopicName ON tblTopics (TopicName)"
    ExecDdl db, "CREATE UNIQUE INDEX uxTemplateName ON tblTemplates (TemplateName)"
    ExecDdl db, "CREATE UNIQUE INDEX uxQKeyVer ON tblQuestions (QuestionKey, VersionNo)"
    ExecDdl db, "CREATE INDEX ixQTopicStatus ON tblQuestions (TopicID, StatusID, IsCurrent)"
    ExecDdl db, "CREATE UNIQUE INDEX uxPartLabel ON tblQuestionParts (QuestionID, PartLabel)"
    ExecDdl db, "CREATE UNIQUE INDEX uxExamTopic ON tblExamTopics (ExamID, TopicID)"
    ExecDdl db, "CREATE UNIQUE INDEX uxExamQ ON tblExamQuestions (ExamID, QuestionID)"
    ExecDdl db, "CREATE INDEX ixTaskPerson ON tblTasks (AssignedPersonID, TaskStatus)"
    ExecDdl db, "CREATE INDEX ixAuditItem ON tblAuditTrail (ItemType, ItemID)"
    ExecDdl db, "CREATE INDEX ixAuditOn ON tblAuditTrail (ActionOn)"
    ExecDdl db, "CREATE INDEX ixAuditPerson ON tblAuditTrail (PersonID, ItemType, ItemID)"
    ExecDdl db, "CREATE INDEX ixExamApproved ON tblExams (FinalApprovedOn)"
    ExecDdl db, "CREATE INDEX ixErrorLogged ON error_log (logged_at)"
    ' Names and codes the app keeps unique: the data file refuses a second one
    ' even when two people add it at the same moment. RefNumber may be empty;
    ' the app treats an empty number as one value, the index lets Nulls repeat.
    ExecDdl db, "CREATE UNIQUE INDEX uxLevelName ON tblDifficultyLevels (LevelName)"
    ExecDdl db, "CREATE UNIQUE INDEX uxMarking ON tblClassLevels (Marking)"
    ExecDdl db, "CREATE UNIQUE INDEX uxPortionMark ON tblClassLevels (PortionMark)"
    ExecDdl db, "CREATE UNIQUE INDEX uxRoleCode ON tblRoles (RoleCode)"
    ExecDdl db, "CREATE UNIQUE INDEX uxRefTitleNumber ON tblSourceRefs (RefTitle, RefNumber)"
End Sub

' Relationship nm: ct.fk refers to pt.pk, enforced, deleting the children
' with the parent when cascadeDel.
Private Sub AddRelation(db As DAO.Database, ByVal nm As String, ByVal pt As String, ByVal pk As String, ByVal ct As String, _
        ByVal fk As String, Optional ByVal cascadeDel As Boolean = False)
    Dim r As DAO.Relation
    Set r = db.CreateRelation(nm, pt, ct, IIf(cascadeDel, dbRelationDeleteCascade, 0))
    r.Fields.Append r.CreateField(pk)
    r.Fields(pk).ForeignName = fk
    db.Relations.Append r
End Sub

Private Sub CreateRelationships(db As DAO.Database)
    AddRelation db, "rPR_P", "tblPeople", "PersonID", "tblPersonRoles", "PersonID", True
    AddRelation db, "rPR_R", "tblRoles", "RoleID", "tblPersonRoles", "RoleID"
    AddRelation db, "rTS_T", "tblTopics", "TopicID", "tblTopicSMEs", "TopicID", True
    AddRelation db, "rTS_P", "tblPeople", "PersonID", "tblTopicSMEs", "PersonID", True
    AddRelation db, "rTT_Tm", "tblTemplates", "TemplateID", "tblTemplateTopics", "TemplateID", True
    AddRelation db, "rTT_T", "tblTopics", "TopicID", "tblTemplateTopics", "TopicID"
    AddRelation db, "rQ_T", "tblTopics", "TopicID", "tblQuestions", "TopicID"
    AddRelation db, "rQ_Ty", "tblQuestionTypes", "QuestionTypeID", "tblQuestions", "QuestionTypeID"
    AddRelation db, "rQ_D", "tblDifficultyLevels", "DifficultyID", "tblQuestions", "DifficultyID"
    AddRelation db, "rQ_C", "tblClassLevels", "ClassLevelID", "tblQuestions", "ClassLevelID"
    AddRelation db, "rQ_S", "tblStatuses", "StatusID", "tblQuestions", "StatusID"
    AddRelation db, "rQ_A", "tblPeople", "PersonID", "tblQuestions", "AuthorPersonID"
    AddRelation db, "rQP_Q", "tblQuestions", "QuestionID", "tblQuestionParts", "QuestionID", True
    AddRelation db, "rQP_C", "tblClassLevels", "ClassLevelID", "tblQuestionParts", "ClassLevelID"
    AddRelation db, "rMC_Q", "tblQuestions", "QuestionID", "tblMCOptions", "QuestionID", True
    AddRelation db, "rMC_P", "tblQuestionParts", "PartID", "tblMCOptions", "PartID", True
    AddRelation db, "rQR_Q", "tblQuestions", "QuestionID", "tblQuestionRefs", "QuestionID", True
    AddRelation db, "rQR_R", "tblSourceRefs", "RefID", "tblQuestionRefs", "RefID"
    AddRelation db, "rE_Tm", "tblTemplates", "TemplateID", "tblExams", "TemplateID"
    AddRelation db, "rE_S", "tblStatuses", "StatusID", "tblExams", "StatusID"
    AddRelation db, "rE_P", "tblPeople", "PersonID", "tblExams", "CreatedByPersonID"
    AddRelation db, "rET_E", "tblExams", "ExamID", "tblExamTopics", "ExamID", True
    AddRelation db, "rET_T", "tblTopics", "TopicID", "tblExamTopics", "TopicID"
    AddRelation db, "rEQ_E", "tblExams", "ExamID", "tblExamQuestions", "ExamID", True
    AddRelation db, "rEQ_ET", "tblExamTopics", "ExamTopicID", "tblExamQuestions", "ExamTopicID"
    AddRelation db, "rEQ_Q", "tblQuestions", "QuestionID", "tblExamQuestions", "QuestionID"
    AddRelation db, "rTk_P", "tblPeople", "PersonID", "tblTasks", "AssignedPersonID"
    AddRelation db, "rAT_P", "tblPeople", "PersonID", "tblAuditTrail", "PersonID"
    ' Links that keep their columns pointing at real rows. Each column may be
    ' Null (no step, no marking yet, not signed off), never 0.
    AddRelation db, "rQ_Step", "tblRoles", "RoleID", "tblQuestions", "CurrentStepRoleID"
    AddRelation db, "rE_Step", "tblRoles", "RoleID", "tblExams", "CurrentStepRoleID"
    AddRelation db, "rE_C", "tblClassLevels", "ClassLevelID", "tblExams", "OverallClassLevelID"
    AddRelation db, "rET_P", "tblPeople", "PersonID", "tblExamTopics", "SignedOffByPersonID"
    AddRelation db, "rTk_CP", "tblPeople", "PersonID", "tblTasks", "CompletedByPersonID"
    AddRelation db, "rQ_Rep", "tblQuestions", "QuestionID", "tblQuestions", "ReplacesQuestionID"
End Sub

Private Sub SeedData(db As DAO.Database)
    ExecDdl db, "INSERT INTO tblRoles (RoleCode, RoleName, ChainOrder) VALUES ('SME','Subject Matter Expert',1)"
    ExecDdl db, "INSERT INTO tblRoles (RoleCode, RoleName, ChainOrder) VALUES ('DH','Department Head',2)"
    ExecDdl db, "INSERT INTO tblRoles (RoleCode, RoleName, ChainOrder) VALUES ('XO','Executive Officer',3)"
    ExecDdl db, "INSERT INTO tblRoles (RoleCode, RoleName, ChainOrder) VALUES ('SM','Senior Member',4)"

    ExecDdl db, "INSERT INTO tblDifficultyLevels (LevelName, SortOrder, IsActive) VALUES ('Easy',1,True)"
    ExecDdl db, "INSERT INTO tblDifficultyLevels (LevelName, SortOrder, IsActive) VALUES ('Medium',2,True)"
    ExecDdl db, "INSERT INTO tblDifficultyLevels (LevelName, SortOrder, IsActive) VALUES ('Hard',3,True)"

    ' TODO(Stephen): confirm the classification levels and their order with
    ' security before the first real data file; these drive every exam's marking.
    ExecDdl db, "INSERT INTO tblClassLevels (Marking, PortionMark, LevelRank) VALUES ('UNCLASSIFIED','(U)',1)"
    ExecDdl db, "INSERT INTO tblClassLevels (Marking, PortionMark, LevelRank) VALUES ('CUI','(CUI)',2)"
    ExecDdl db, "INSERT INTO tblClassLevels (Marking, PortionMark, LevelRank) VALUES ('CONFIDENTIAL','(C)',3)"
    ExecDdl db, "INSERT INTO tblClassLevels (Marking, PortionMark, LevelRank) VALUES ('SECRET','(S)',4)"

    ExecDdl db, "INSERT INTO tblQuestionTypes (TypeName) VALUES ('Multiple Choice')"
    ExecDdl db, "INSERT INTO tblQuestionTypes (TypeName) VALUES ('True/False')"
    ExecDdl db, "INSERT INTO tblQuestionTypes (TypeName) VALUES ('Short Answer')"

    Dim s As Variant                        ' For Each over an Array needs a Variant
    For Each s In Array("Draft", "In Review", "Approved", "Rejected", "Retired")
        ExecDdl db, "INSERT INTO tblStatuses (StatusName, AppliesTo) VALUES ('" & s & "','Question')"
    Next
    For Each s In Array("SME Review", "In Review", "Approved", "Rejected")
        ExecDdl db, "INSERT INTO tblStatuses (StatusName, AppliesTo) VALUES ('" & s & "','Exam')"
    Next

    ' The front end falls back to the same three values when a row is missing
    ' (modAdmin DEFAULT_OVERLAP_PCT, DEFAULT_OVERLAP_WINDOW, DEFAULT_QUESTION_COUNT).
    ExecDdl db, "INSERT INTO tblSettings VALUES ('OverlapMaxPct','30','Max % of an exam''s questions that may appear in the lookback window')"
    ExecDdl db, "INSERT INTO tblSettings VALUES ('OverlapWindowExams','10','Number of most recent exams checked for overlap')"
    ExecDdl db, "INSERT INTO tblSettings VALUES ('DefaultQuestionCount','4','Default questions per topic in a template')"
    ExecDdl db, "INSERT INTO tblSettings VALUES ('SchemaVersion','4','Back-end schema version')"
End Sub
