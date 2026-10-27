Attribute VB_Name = "modBuildBackEnd"
Option Compare Database
Option Explicit
' Exam Generator - back-end builder.
' Run BuildBackEnd from any Access database. Creates ExamGen_BE.accdb with all tables,
' relationships, indexes and seed lookup data. Path is prompted (default: same folder).

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
    CreateTables
    CreateIndexes
    CreateRelationships
    SeedData
    db.Close: Set db = Nothing
    MsgBox "Back end created:" & vbCrLf & p, vbInformation
    Exit Sub
Fail:
    MsgBox "Build failed: " & Err.Description, vbCritical
    If Not db Is Nothing Then db.Close
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

    ' ---- People ----
    X "CREATE TABLE tblUsers (UserID AUTOINCREMENT PRIMARY KEY, WinUserName TEXT(64) NOT NULL, DisplayName TEXT(100), Email TEXT(150), " & _
      "IsAdmin YESNO, IsActive YESNO, NotifyDashboard YESNO, NotifyEmail YESNO)"
    X "CREATE TABLE tblUserRoles (UserID LONG NOT NULL, RoleID LONG NOT NULL, CONSTRAINT pkUserRoles PRIMARY KEY (UserID, RoleID))"

    ' ---- Topics & templates ----
    X "CREATE TABLE tblTopics (TopicID AUTOINCREMENT PRIMARY KEY, TopicName TEXT(100) NOT NULL, Description MEMO, IsActive YESNO)"
    X "CREATE TABLE tblTopicSMEs (TopicID LONG NOT NULL, UserID LONG NOT NULL, CONSTRAINT pkTopicSMEs PRIMARY KEY (TopicID, UserID))"
    X "CREATE TABLE tblTemplates (TemplateID AUTOINCREMENT PRIMARY KEY, TemplateName TEXT(100) NOT NULL, Description MEMO, IsActive YESNO)"
    X "CREATE TABLE tblTemplateTopics (TemplateID LONG NOT NULL, TopicID LONG NOT NULL, QuestionCount INTEGER NOT NULL, SortOrder INTEGER, " & _
      "CONSTRAINT pkTemplateTopics PRIMARY KEY (TemplateID, TopicID))"

    ' ---- Question bank (versioned) ----
    X "CREATE TABLE tblQuestions (QuestionID AUTOINCREMENT PRIMARY KEY, QuestionKey LONG, VersionNo INTEGER NOT NULL, ReplacesQuestionID LONG, " & _
      "TopicID LONG NOT NULL, QuestionTypeID LONG NOT NULL, DifficultyID LONG, ClassLevelID LONG, QuestionText MEMO NOT NULL, " & _
      "TotalPoints DOUBLE NOT NULL, ReferenceAnswer MEMO, StatusID LONG NOT NULL, CurrentStepRoleID LONG, " & _
      "AuthorUserID LONG NOT NULL, CreatedOn DATETIME, IsCurrent YESNO, RetiredOn DATETIME)"
    X "CREATE TABLE tblQuestionParts (PartID AUTOINCREMENT PRIMARY KEY, QuestionID LONG NOT NULL, PartLabel TEXT(5) NOT NULL, PartText MEMO NOT NULL, " & _
      "Points DOUBLE NOT NULL, ClassLevelID LONG NOT NULL, ReferenceAnswer MEMO, SortOrder INTEGER)"
    X "CREATE TABLE tblMCOptions (OptionID AUTOINCREMENT PRIMARY KEY, QuestionID LONG, PartID LONG, OptionLabel TEXT(2), OptionText MEMO NOT NULL, IsCorrect YESNO, SortOrder INTEGER)"
    X "CREATE TABLE tblSourceRefs (RefID AUTOINCREMENT PRIMARY KEY, RefTitle TEXT(200) NOT NULL, RefNumber TEXT(100), Revision TEXT(50), IsActive YESNO)"
    X "CREATE TABLE tblQuestionRefs (QuestionID LONG NOT NULL, RefID LONG NOT NULL, Paragraph TEXT(100), CONSTRAINT pkQuestionRefs PRIMARY KEY (QuestionID, RefID))"

    ' ---- Exams ----
    X "CREATE TABLE tblExams (ExamID AUTOINCREMENT PRIMARY KEY, TemplateID LONG NOT NULL, RecipientName TEXT(150) NOT NULL, VersionNo INTEGER, " & _
      "StatusID LONG NOT NULL, CurrentStepRoleID LONG, OverallClassLevelID LONG, CreatedByUserID LONG NOT NULL, CreatedOn DATETIME, " & _
      "FinalApprovedOn DATETIME, OutputFile MEMO)"
    X "CREATE TABLE tblExamTopics (ExamTopicID AUTOINCREMENT PRIMARY KEY, ExamID LONG NOT NULL, TopicID LONG NOT NULL, QuestionCount INTEGER NOT NULL, " & _
      "SignedOffByUserID LONG, SignedOffOn DATETIME)"
    X "CREATE TABLE tblExamQuestions (ExamQuestionID AUTOINCREMENT PRIMARY KEY, ExamID LONG NOT NULL, ExamTopicID LONG NOT NULL, QuestionID LONG NOT NULL, " & _
      "SortOrder INTEGER, WasSwapped YESNO)"
    X "CREATE TABLE tblTasks (TaskID AUTOINCREMENT PRIMARY KEY, TaskType TEXT(20) NOT NULL, ItemType TEXT(10) NOT NULL, ItemID LONG NOT NULL, " & _
      "ExamTopicID LONG, AssignedUserID LONG NOT NULL, TaskStatus TEXT(20) NOT NULL, CreatedOn DATETIME, CompletedOn DATETIME, CompletedByUserID LONG)"

    ' ---- Help text (admin-editable, keyed by form) ----
    X "CREATE TABLE tblHelp (HelpID AUTOINCREMENT PRIMARY KEY, FormName TEXT(64) NOT NULL, SectionKey TEXT(64), HelpTitle TEXT(150), HelpText MEMO, SortOrder INTEGER)"

    ' ---- Audit trail (append-only by app) ----
    X "CREATE TABLE tblAuditTrail (AuditID AUTOINCREMENT PRIMARY KEY, ItemType TEXT(10) NOT NULL, ItemID LONG NOT NULL, VersionNo INTEGER, " & _
      "ActionType TEXT(20) NOT NULL, StepRoleID LONG, UserID LONG NOT NULL, ActionOn DATETIME NOT NULL, Comments MEMO)"
End Sub

Private Sub CreateIndexes()
    X "CREATE UNIQUE INDEX uxWinUser ON tblUsers (WinUserName)"
    X "CREATE UNIQUE INDEX uxTopicName ON tblTopics (TopicName)"
    X "CREATE UNIQUE INDEX uxTemplateName ON tblTemplates (TemplateName)"
    X "CREATE UNIQUE INDEX uxQKeyVer ON tblQuestions (QuestionKey, VersionNo)"
    X "CREATE INDEX ixQTopicStatus ON tblQuestions (TopicID, StatusID, IsCurrent)"
    X "CREATE UNIQUE INDEX uxPartLabel ON tblQuestionParts (QuestionID, PartLabel)"
    X "CREATE UNIQUE INDEX uxExamTopic ON tblExamTopics (ExamID, TopicID)"
    X "CREATE UNIQUE INDEX uxExamQ ON tblExamQuestions (ExamID, QuestionID)"
    X "CREATE INDEX ixTaskUser ON tblTasks (AssignedUserID, TaskStatus)"
    X "CREATE INDEX ixAuditItem ON tblAuditTrail (ItemType, ItemID)"
    X "CREATE INDEX ixExamApproved ON tblExams (FinalApprovedOn)"
    X "CREATE INDEX ixHelpForm ON tblHelp (FormName)"
End Sub

Private Sub Rel(nm As String, pt As String, pk As String, ct As String, fk As String, Optional cascadeDel As Boolean = False)
    Dim r As DAO.Relation
    Set r = db.CreateRelation(nm, pt, ct, IIf(cascadeDel, dbRelationDeleteCascade, 0))
    r.Fields.Append r.CreateField(pk)
    r.Fields(pk).ForeignName = fk
    db.Relations.Append r
End Sub

Private Sub CreateRelationships()
    Rel "rUR_U", "tblUsers", "UserID", "tblUserRoles", "UserID", True
    Rel "rUR_R", "tblRoles", "RoleID", "tblUserRoles", "RoleID"
    Rel "rTS_T", "tblTopics", "TopicID", "tblTopicSMEs", "TopicID", True
    Rel "rTS_U", "tblUsers", "UserID", "tblTopicSMEs", "UserID", True
    Rel "rTT_Tm", "tblTemplates", "TemplateID", "tblTemplateTopics", "TemplateID", True
    Rel "rTT_T", "tblTopics", "TopicID", "tblTemplateTopics", "TopicID"
    Rel "rQ_T", "tblTopics", "TopicID", "tblQuestions", "TopicID"
    Rel "rQ_Ty", "tblQuestionTypes", "QuestionTypeID", "tblQuestions", "QuestionTypeID"
    Rel "rQ_D", "tblDifficultyLevels", "DifficultyID", "tblQuestions", "DifficultyID"
    Rel "rQ_C", "tblClassLevels", "ClassLevelID", "tblQuestions", "ClassLevelID"
    Rel "rQ_S", "tblStatuses", "StatusID", "tblQuestions", "StatusID"
    Rel "rQ_A", "tblUsers", "UserID", "tblQuestions", "AuthorUserID"
    Rel "rQP_Q", "tblQuestions", "QuestionID", "tblQuestionParts", "QuestionID", True
    Rel "rQP_C", "tblClassLevels", "ClassLevelID", "tblQuestionParts", "ClassLevelID"
    Rel "rMC_Q", "tblQuestions", "QuestionID", "tblMCOptions", "QuestionID", True
    Rel "rMC_P", "tblQuestionParts", "PartID", "tblMCOptions", "PartID", True
    Rel "rQR_Q", "tblQuestions", "QuestionID", "tblQuestionRefs", "QuestionID", True
    Rel "rQR_R", "tblSourceRefs", "RefID", "tblQuestionRefs", "RefID"
    Rel "rE_Tm", "tblTemplates", "TemplateID", "tblExams", "TemplateID"
    Rel "rE_S", "tblStatuses", "StatusID", "tblExams", "StatusID"
    Rel "rE_U", "tblUsers", "UserID", "tblExams", "CreatedByUserID"
    Rel "rET_E", "tblExams", "ExamID", "tblExamTopics", "ExamID", True
    Rel "rET_T", "tblTopics", "TopicID", "tblExamTopics", "TopicID"
    Rel "rEQ_E", "tblExams", "ExamID", "tblExamQuestions", "ExamID", True
    Rel "rEQ_ET", "tblExamTopics", "ExamTopicID", "tblExamQuestions", "ExamTopicID"
    Rel "rEQ_Q", "tblQuestions", "QuestionID", "tblExamQuestions", "QuestionID"
    Rel "rTk_U", "tblUsers", "UserID", "tblTasks", "AssignedUserID"
    Rel "rAT_U", "tblUsers", "UserID", "tblAuditTrail", "UserID"
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
    X "INSERT INTO tblSettings VALUES ('SchemaVersion','1','Back-end schema version')"

    ' First admin = whoever runs the build.
    X "INSERT INTO tblUsers (WinUserName, DisplayName, IsAdmin, IsActive, NotifyDashboard, NotifyEmail) VALUES ('" & _
      Replace(Environ$("USERNAME"), "'", "''") & "','" & Replace(Environ$("USERNAME"), "'", "''") & "',True,True,True,False)"
End Sub
