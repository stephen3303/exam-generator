Attribute VB_Name = "modSeedDemo"
Option Compare Database
Option Explicit
' Exam Generator - demo data for testing (tools\seed-demo.ps1 runs it).
'
' Fills a NEW, empty data file with made-up people, topics, templates,
' references, questions and exams, so every page has rows to show and the
' Word documents get made. Nothing here is real: every name, reference and
' question is invented for testing.
'
' It writes through the app's own functions, signed in as each made-up
' person in turn (ActAs), so every rule the pages enforce is enforced here
' too: an SME writes and submits, the DH, XO and SM approve, the DH
' generates exams, the SMEs swap and sign off, and the SM's approval makes
' the Word document (modExamDoc.MakeExamDocument). A rule the seed breaks
' stops it with the app's own message.
'
' This module is never built into the app: seed-demo.ps1 loads it, with the
' app's modules, into a throwaway database in TEMP that links the data file.
'
' What it leaves (see access\SEED-DEMO.md for the click-through):
'   people      the owner (admin, DH, XO, SM) and 8 made-up people
'   topics      6 in use and 1 turned off; templates 4 in use
'   questions   80 approved in the bank; 3 retired by an approved version 2;
'               a version 2 in draft and one waiting on the DH; drafts in
'               five topics; several waiting at each of the DH, XO and SM;
'               rejected by each of the DH, XO and SM; one rejected, fixed
'               and sent again
'   exams       13: four approved with their Word documents (one had a
'               question swapped); waiting on each of the DH, XO and SM (one
'               of them sent back once and signed off again); sent back to
'               their SMEs by each of the DH, XO and SM; two still with
'               their SMEs, one with a topic signed off and one untouched
' Then the dates are spread over the last 80 days so the Statistics page
' and the histories read like real use.

Private Const COMMAND_NAME As String = "Training Command (demo)"
Private Const EMAIL_DOMAIN As String = "@example.com"

Private m_step As String            ' what the seed was doing, for the error
Private m_log As String             ' the report seed-demo.ps1 prints
Private m_ids As Object             ' username -> person_id
Private m_names As Object           ' person_id -> username
Private m_topics As Object          ' topic code -> TopicID
Private m_refs As Object            ' reference code -> RefID
Private m_bank As Collection        ' QuestionKeys in the approved bank, oldest first
Private m_docsOK As Long
Private m_docsFailed As Long

' ---------------------------------------------------------------------------
' Entry point
' ---------------------------------------------------------------------------

' Seed the data file at bePath. ownerUser names the owner's sign-in name in
' the Personnel file when their Windows account isn't saved there ("" to
' find them by Windows account); ownerFirst, ownerLast and ownerRank are used
' only when there is no Personnel file yet and one is made. demoPassword is
' every made-up person's app password (and the owner's, when they have
' none). makeDocs False skips Word. personnelFile names the Personnel file
' to use ("" for the usual place: a Personnel folder beside the data file's
' folder, made there when there is none).
' Returns "OK" and the report, one line each, or "ERROR ..." and the report
' so far.
Public Function SeedDemoRun(ByVal bePath As String, ByVal ownerUser As String, ByVal ownerFirst As String, _
        ByVal ownerLast As String, ByVal ownerRank As String, ByVal demoPassword As String, ByVal makeDocs As Boolean, _
        ByVal personnelFile As String) As String
    Dim owner As Long
    On Error GoTo Fail
    g_Silent = True
    m_log = ""
    m_docsOK = 0: m_docsFailed = 0
    Set m_ids = CreateObject("Scripting.Dictionary")
    Set m_names = CreateObject("Scripting.Dictionary")
    Set m_topics = CreateObject("Scripting.Dictionary")
    Set m_refs = CreateObject("Scripting.Dictionary")
    Set m_bank = New Collection

    m_step = "link the data file": LinkDataFile bePath
    m_step = "check the data file is new": CheckEmpty
    m_step = "the work tables": EnsureWorkTables
    m_step = "the Personnel file"
    If Len(personnelFile) > 0 Then
        LinkPersonnelFile personnelFile
        Note "Personnel file named: " & personnelFile
    End If
    owner = SetUpOwner(ownerUser, ownerFirst, ownerLast, ownerRank, demoPassword)
    m_step = "people": AddPeople owner, demoPassword
    m_step = "topics and templates": AddTopicsAndTemplates
    m_step = "references": AddReferences
    m_step = "the question bank": AddBank
    m_step = "questions in the approval chain": AddWorkflowQuestions
    m_step = "exams": AddExams makeDocs
    m_step = "spreading the dates": SpreadDates
    m_step = "the summary": Summarize
    SignOutSeed
    g_Silent = False
    SeedDemoRun = "OK" & vbCrLf & m_log
    Exit Function
Fail:
    SeedDemoRun = "ERROR during " & m_step & ": " & Err.Number & " " & Err.Description & _
        IIf(Len(Err.Source) > 0, " [" & Err.Source & "]", "") & vbCrLf & m_log
    On Error Resume Next
    SignOutSeed
    g_Silent = False
End Function

Private Sub Note(ByVal s As String)
    m_log = m_log & s & vbCrLf
    Debug.Print s
End Sub

' ---------------------------------------------------------------------------
' The data file and the Personnel file
' ---------------------------------------------------------------------------

' Link every table in the data file into this throwaway database, as the
' build's LinkBackEnd does.
Private Sub LinkDataFile(ByVal bePath As String)
    Dim be As DAO.Database, td As DAO.TableDef, names As String
    KitCheckBackendPath bePath
    Set be = DBEngine.OpenDatabase(bePath, False, True)
    For Each td In be.TableDefs
        If Left$(td.Name, 4) <> "MSys" And Left$(td.Name, 4) <> "USys" And Left$(td.Name, 1) <> "~" Then
            names = names & IIf(Len(names) > 0, vbTab, "") & td.Name
        End If
    Next
    be.Close
    If Len(names) = 0 Then Err.Raise 5, "modSeedDemo", "The data file has no tables: " & bePath
    KitLinkTables bePath, Split(names, vbTab)
    Note "Data file: " & bePath
End Sub

' Demo data goes only into a data file with no questions and no exams, so it
' never mixes with real ones.
Private Sub CheckEmpty()
    Dim nQ As Long, nE As Long, problem As String
    problem = DataVersionProblem()
    If Len(problem) > 0 Then Err.Raise 5, "modSeedDemo", problem
    nQ = DCount("*", "tblQuestions"): nE = DCount("*", "tblExams")
    If nQ + nE > 0 Then
        Err.Raise 5, "modSeedDemo", "This data file already has " & Plural(nQ, "question") & " and " & Plural(nE, "exam") & _
            ". Demo data only goes into a new, empty data file: make one with backend\new-backend.ps1 and seed that."
    End If
End Sub

' Link the Personnel file (in its usual place beside the data folder), or
' make it with this Windows account as its first admin; then sign in as the
' owner and make them the Exam Generator admin. Returns the owner's id.
Private Function SetUpOwner(ByVal ownerUser As String, ByVal ownerFirst As String, ByVal ownerLast As String, _
        ByVal ownerRank As String, ByVal demoPassword As String) As Long
    Dim st As Long, pid As Long, def As String, made As Boolean
    st = PersonnelStatus("tblTopics")
    Select Case st
        Case PERSONNEL_READY
            Note "Personnel file: " & PersonnelPath()
        Case PERSONNEL_NOT_SET_UP
            def = PersonnelDefaultPath("tblTopics")
            pid = CreatePersonnelFile(def, APP_CODE, ownerFirst, ownerLast, ownerRank)
            made = True
            Note "Personnel file made: " & def & " (its key is in Personnel.key beside it)"
        Case PERSONNEL_LOST
            Err.Raise 5, "modSeedDemo", "The Personnel file this data file uses can't be reached."
        Case Else
            Err.Raise 5, "modSeedDemo", "The Personnel file has no active admin, so nobody can set up people in it."
    End Select
    If pid = 0 Then
        If Len(ownerUser) > 0 Then
            pid = PersonIdFor(ownerUser)
            If pid = 0 Then Err.Raise 5, "modSeedDemo", "The Personnel file has no active person with the sign-in name " & ownerUser & "."
        Else
            pid = PersonnelWindowsPersonId()
            If pid = 0 Then Err.Raise 5, "modSeedDemo", "Your Windows account isn't saved for anyone in the Personnel file. " & _
                "Run again with -OwnerUser and your sign-in name."
        End If
    End If
    If Not Nz(DLookup("is_admin", PERSON_TABLE, "person_id=" & pid), False) Then
        Err.Raise 5, "modSeedDemo", "The owner must be a Personnel admin, and " & PersonUser(pid) & " isn't."
    End If
    If Not PersonCanUse(pid, APP_CODE) Then SetPersonApp pid, APP_CODE, True
    RememberPerson pid
    If IsNull(DLookup("password_hash", PERSON_TABLE, "person_id=" & pid)) Then
        CheckPassword PersonnelSetPassword(pid, demoPassword)
        Note "Owner " & PersonUser(pid) & ": app password set to the demo password (it had none)."
    End If
    ActAsId pid                          ' a new data file makes them its first admin here (LoadRoles)
    RequireExamAdmin
    SetPersonRole pid, "DH", True
    SetPersonRole pid, "XO", True
    SetPersonRole pid, "SM", True
    ActAsId pid                          ' again, now with the roles
    Note "Owner: " & PersonUser(pid) & " (" & Nz(DLookup("display_name", PERSON_TABLE, "person_id=" & pid), "") & _
        "), Exam Generator admin with DH, XO and SM" & IIf(made, ", signed in by this Windows account.", ".")
    SetUpOwner = pid
End Function

' ---------------------------------------------------------------------------
' Signing in as someone
' ---------------------------------------------------------------------------

' Sign in as person pid, as the kit's StartSession and the app's LoadRoles
' do, so every check that reads the signed-in person reads them.
Private Sub ActAsId(ByVal pid As Long)
    Dim rs As DAO.Recordset, why As String
    Set rs = CurrentDb.OpenRecordset("SELECT username, display_name, is_admin FROM " & PERSON_TABLE & _
        " WHERE person_id=" & pid, dbOpenSnapshot)
    If rs.EOF Then rs.Close: Err.Raise 5, "modSeedDemo", "No person " & pid & " in the Personnel file."
    TempVars("PersonID") = pid
    TempVars("Username") = rs!username.Value
    TempVars("PersonName") = Nz(rs!display_name, rs!username)
    TempVars("PersonIsAdmin") = CBool(Nz(rs!is_admin, False))
    rs.Close
    why = LoadRoles()
    If Len(why) > 0 Then Err.Raise 5, "modSeedDemo", PersonUser(pid) & ": " & why
End Sub

Private Sub ActAs(ByVal username As String)
    If Not m_ids.Exists(username) Then Err.Raise 5, "modSeedDemo", "No made-up person " & username & "."
    ActAsId m_ids(username)
End Sub

Private Sub SignOutSeed()
    Dim v As Variant
    On Error Resume Next
    For Each v In Array("PersonID", "Username", "PersonName", "PersonIsAdmin", "EGRoles", "EGIsAdmin", "EGNotifyDash", "EGNotifyEmail")
        TempVars.Remove v
    Next
End Sub

Private Function PersonUser(ByVal pid As Long) As String
    PersonUser = Nz(DLookup("username", PERSON_TABLE, "person_id=" & pid), "person " & pid)
End Function

Private Sub RememberPerson(ByVal pid As Long)
    Dim u As String
    u = PersonUser(pid)
    m_ids(u) = pid
    m_names(pid) = u
End Sub

Private Sub CheckPassword(ByVal problem As String)
    If Len(problem) > 0 Then Err.Raise 5, "modSeedDemo", "The demo password won't do: " & problem
End Sub

' ---------------------------------------------------------------------------
' People
' ---------------------------------------------------------------------------

' The made-up people, as the owner: each in the Personnel file with the demo
' password and access to the app, then their Exam Generator roles. A sign-in
' name already in the file is used as it is (its password is left alone).
Private Sub AddPeople(ByVal owner As Long, ByVal demoPassword As String)
    Dim spec As Variant, f() As String, pid As Long, r As Variant, n As Long
    ActAsId owner
    For Each spec In Array( _
            "dreyes|CDR|Dana|Reyes|DH", _
            "mhale|CAPT|Morgan|Hale|XO", _
            "tbrooks|CMDCM|Taylor|Brooks|SM", _
            "achen|LT|Alex|Chen|SME", _
            "srivera|CWO2|Sam|Rivera|SME", _
            "jpatel|LTJG|Jordan|Patel|SME", _
            "ckim|ENS|Casey|Kim|SME", _
            "rford|ENS|Riley|Ford|SME")
        f = Split(spec, "|")
        pid = Nz(DLookup("person_id", PERSON_TABLE, "username=" & Q(f(0))), 0)
        If pid = 0 Then
            pid = AddPerson(f(2), f(3), f(1), "", f(0), False, "", COMMAND_NAME, LCase$(f(2) & "." & f(3)) & EMAIL_DOMAIN)
            CheckPassword PersonnelSetPassword(pid, demoPassword)
            n = n + 1
        End If
        If Not PersonCanUse(pid, APP_CODE) Then SetPersonApp pid, APP_CODE, True
        RememberPerson pid
        For Each r In Split(f(4), ",")
            SetPersonRole pid, CStr(r), True
        Next
    Next
    ' One SME whose Exam Generator access is turned off, as the People page does.
    KitExecute "UPDATE tblPeople SET IsActive=False WHERE PersonID=" & m_ids("rford")
    Note "People: " & n & " made-up people added to the Personnel file (8 in the demo), all with the demo password, " & _
        "emails off. rford's Exam Generator access is turned off."
End Sub

' ---------------------------------------------------------------------------
' Topics, templates, references
' ---------------------------------------------------------------------------

Private Sub AddTopicsAndTemplates()
    Dim t1 As Long, t2 As Long
    RequireExamAdmin
    AddTopic "DC", "Damage Control", "Fire classes, firefighting, flooding and material conditions.", "achen,srivera"
    AddTopic "NAV", "Navigation", "Charts, aids to navigation, fixes and compass error.", "jpatel,rford"
    AddTopic "ENG", "Engineering Fundamentals", "Pumps, lube oil, electrical protection and tagout.", "srivera,achen"
    AddTopic "COM", "Communications and Security", "Voice procedure, precedence and everyday information security.", "ckim,jpatel"
    AddTopic "SEA", "Seamanship", "Line handling, anchoring, small boats and deck safety.", "achen,ckim"
    AddTopic "HAZ", "Hazardous Materials", "Labels, storage, spills and the safety data sheet.", "srivera,jpatel"
    AddTopic "OLD", "Legacy Seamanship (turned off)", "A topic that is turned off, with no questions.", ""
    KitExecute "UPDATE tblTopics SET IsActive=False WHERE TopicID=" & m_topics("OLD")

    t1 = InsertRow("INSERT INTO tblTemplates (TemplateName, Description, IsActive) VALUES (" & Q("Basic Qualification Board") & _
        ", " & Q("Demo: all four topics.") & ", True)")
    AddTemplateTopic t1, "DC", 3
    AddTemplateTopic t1, "NAV", 2
    AddTemplateTopic t1, "ENG", 2
    AddTemplateTopic t1, "COM", 2
    t2 = InsertRow("INSERT INTO tblTemplates (TemplateName, Description, IsActive) VALUES (" & Q("Damage Control Refresher") & _
        ", " & Q("Demo: damage control with one engineering question.") & ", True)")
    AddTemplateTopic t2, "DC", 3
    AddTemplateTopic t2, "ENG", 1
    t1 = InsertRow("INSERT INTO tblTemplates (TemplateName, Description, IsActive) VALUES (" & Q("Deck and Hazmat Qualification") & _
        ", " & Q("Demo: seamanship and hazardous materials, with damage control.") & ", True)")
    AddTemplateTopic t1, "SEA", 3
    AddTemplateTopic t1, "HAZ", 3
    AddTemplateTopic t1, "DC", 2
    t2 = InsertRow("INSERT INTO tblTemplates (TemplateName, Description, IsActive) VALUES (" & Q("Comprehensive Board") & _
        ", " & Q("Demo: two questions from every topic.") & ", True)")
    AddTemplateTopic t2, "DC", 2
    AddTemplateTopic t2, "NAV", 2
    AddTemplateTopic t2, "ENG", 2
    AddTemplateTopic t2, "COM", 2
    AddTemplateTopic t2, "SEA", 2
    AddTemplateTopic t2, "HAZ", 2
    Note "Topics: 6 in use and 1 turned off. Templates: Basic Qualification Board (9 questions), Damage Control Refresher (4), " & _
        "Deck and Hazmat Qualification (8) and Comprehensive Board (12)."
End Sub

Private Sub AddTopic(ByVal code As String, ByVal nm As String, ByVal desc As String, ByVal smes As String)
    Dim id As Long, u As Variant
    id = InsertRow("INSERT INTO tblTopics (TopicName, DefaultCount, Description, IsActive) VALUES (" & Q(nm) & ", " & _
        DefaultQuestionCount() & ", " & Q(desc) & ", True)")
    m_topics(code) = id
    If Len(smes) > 0 Then
        For Each u In Split(smes, ",")
            KitExecute "INSERT INTO tblTopicSMEs (TopicID, PersonID) VALUES (" & id & ", " & m_ids(CStr(u)) & ")"
        Next
    End If
End Sub

Private Sub AddTemplateTopic(ByVal tplId As Long, ByVal code As String, ByVal n As Long)
    KitExecute "INSERT INTO tblTemplateTopics (TemplateID, TopicID, QuestionCount, SortOrder) VALUES (" & tplId & ", " & _
        m_topics(code) & ", " & n & ", " & NextOrder("tblTemplateTopics", "SortOrder", "TemplateID=" & tplId) & ")"
End Sub

' Made-up references, added by an SME as the question editor does.
Private Sub AddReferences()
    ActAs "achen"
    m_refs("DCTM") = AddReference("Damage Control Training Manual (demo)", "DCTM-100", "B")
    m_refs("SSP") = AddReference("Ship Safety Program Guide (demo)", "SSP-1", "C")
    m_refs("NFH") = AddReference("Navigation Fundamentals Handbook (demo)", "NFH-20", "3")
    m_refs("EPB") = AddReference("Engineering Plant Basics (demo)", "EPB-7", "A")
    m_refs("CPG") = AddReference("Communications Procedures Guide (demo)", "CPG-12", "2")
    m_refs("SH") = AddReference("Seamanship Handbook (demo)", "SH-3", "4")
    m_refs("HMC") = AddReference("Hazardous Material Control Guide (demo)", "HMC-5", "B")
    Note "References: 7 made-up references."
End Sub

' ---------------------------------------------------------------------------
' Questions
' ---------------------------------------------------------------------------
' A question is one line, its fields separated by |:
'   topic code | MC, TF or SA | E, M or H | points | U or CUI | question text |
'   reference answer | MC options, ; between them, * before the correct one |
'   reference code ~ paragraph | parts, ^ between them, each text ~ points ~
'   U or CUI ~ answer (labelled A, B...)

' The approved bank: each written by its topic's SMEs in turn, submitted,
' and approved by the DH, the XO and the SM.
Private Sub AddBank()
    Dim specs As Collection, spec As Variant, qid As Long, n As Long, authors As Object, a As Variant, who As String
    Set authors = CreateObject("Scripting.Dictionary")
    authors("DC") = Array("achen", "srivera")
    authors("NAV") = Array("jpatel")
    authors("ENG") = Array("srivera", "achen")
    authors("COM") = Array("ckim", "jpatel")
    authors("SEA") = Array("achen", "ckim")
    authors("HAZ") = Array("srivera", "jpatel")
    Set specs = BankSpecs()
    For Each spec In specs
        a = authors(Split(spec, "|")(0))
        who = a(n Mod (UBound(a) + 1))
        qid = WriteQuestion(CStr(spec), who)
        SubmitAs qid, who
        ApproveChain qid
        m_bank.Add DLookup("QuestionKey", "tblQuestions", "QuestionID=" & qid)
        n = n + 1
    Next
    Note "Question bank: " & n & " questions written, submitted and approved by the DH, XO and SM."
End Sub

Private Function BankSpecs() As Collection
    Dim c As New Collection
    ' Damage Control (12)
    c.Add "DC|MC|E|2|U|Which class of fire involves flammable liquids such as fuel or lubricating oil?||Class A;*Class B;Class C;Class D|DCTM~3.1|"
    c.Add "DC|MC|E|2|U|Which class of fire involves energized electrical equipment?||Class A;Class B;*Class C;Class D|DCTM~3.1|"
    c.Add "DC|TF|E|1|U|Water is the preferred agent for a fire in energized electrical equipment.|False||DCTM~3.4|"
    c.Add "DC|SA|M|5|U|List the three sides of the fire triangle.|Heat, fuel and oxygen.||DCTM~3.2|"
    c.Add "DC|MC|M|3|U|What is the first action on discovering a fire aboard ship?||Fight it alone before telling anyone;*Sound the alarm and report its location;Open the nearest hatch for air;Secure the ventilation and leave|DCTM~2.1|"
    c.Add "DC|TF|E|1|U|A watertight door should be left open when it is not in use.|False||DCTM~5.2|"
    c.Add "DC|SA|M|5|U|Explain the purpose of setting material condition Zebra.|It gives the ship its greatest watertight and fire integrity by closing every fitting marked X, Y and Z. It is set for general quarters and when damage is expected.||DCTM~5.4|"
    c.Add "DC|MC|H|4|U|A space reports a ruptured fuel line with no fire. What hazard is controlled first?||Flooding;*Sources of ignition;Smoke;Loss of lighting|DCTM~4.3|"
    c.Add "DC|SA|H|10|U|A main space fire is reported during the underway watch.|||DCTM~6.2|Describe the immediate actions of the watchstanders in the space.~5~U~Report the fire, fight it if it is small, leave the space if it is not, and secure the space.^Describe how the space is made ready before it is entered again.~5~CUI~Remove the smoke, test the atmosphere, post a reflash watch and confirm with the scene leader."
    c.Add "DC|TF|M|1|U|Portable dewatering pumps may be used to remove firefighting water from a compartment.|True||DCTM~7.1|"
    c.Add "DC|MC|M|2|U|Which tool measures the oxygen in a closed space before anyone enters it?||Thermal imager;Sound-powered phone;*Gas-free test meter;Pressure gauge|SSP~2.5|"
    c.Add "DC|SA|E|3|U|Name two items of personal protective equipment worn by a hose team.|Any two of: firefighter's ensemble, breathing apparatus, helmet, gloves, boots.||DCTM~3.6|"
    ' Navigation (8)
    c.Add "NAV|MC|E|2|U|How many minutes of latitude equal one nautical mile?||*One;Two;Ten;Sixty|NFH~1.2|"
    c.Add "NAV|TF|E|1|U|True north and magnetic north are the same everywhere on Earth.|False||NFH~2.1|"
    c.Add "NAV|MC|M|3|U|Returning from sea in IALA Region B, which side of the channel does a red buoy mark?||*The starboard side;The port side;Safe water on either side;An isolated danger|NFH~4.3|"
    c.Add "NAV|SA|M|5|U|Explain the difference between variation and deviation.|Variation is the angle between true and magnetic north at a place. Deviation is the error of the ship's own compass caused by the ship's magnetism.||NFH~2.2|"
    c.Add "NAV|MC|M|2|U|In navigation, what does the term set mean?||The speed of the current;*The direction toward which the current flows;The ship's heading;The depth of the water|NFH~3.4|"
    c.Add "NAV|TF|M|1|U|A fix is a position found from two or more lines of position taken at the same time.|True||NFH~3.1|"
    c.Add "NAV|SA|H|6|U|The ship is approaching a harbor in reduced visibility.|||NFH~5.1|List the aids the navigation team should use to fix the ship's position.~3~U~Radar ranges and bearings, GPS, soundings, and any aids that can be seen.^State how often fixes are taken in restricted waters.~3~CUI~At the interval the commanding officer's standing orders set for piloting waters."
    c.Add "NAV|MC|E|2|U|Which publication lists the characteristics of the lights along a coast?||Tide tables;*Light list;Pilot chart;Almanac|NFH~6.2|"
    ' Engineering Fundamentals (8)
    c.Add "ENG|MC|E|2|U|What does a gauge reading of zero psig mean?||A perfect vacuum;*Pressure equal to the atmosphere;The system's highest pressure;A failed gauge|EPB~1.3|"
    c.Add "ENG|TF|E|1|U|Lube oil both lubricates and cools the moving parts it reaches.|True||EPB~2.1|"
    c.Add "ENG|SA|M|5|U|Explain why a centrifugal pump must be primed before it is started.|It cannot move air well, so its casing must be full of liquid to build suction and to avoid damage from running dry.||EPB~3.2|"
    c.Add "ENG|MC|M|3|U|Which device protects a circuit by opening when the current is too high?||Relay;Transformer;*Circuit breaker;Rheostat|EPB~4.1|"
    c.Add "ENG|TF|M|1|U|Tags on a tagout may be removed by anyone once the work is complete.|False||SSP~3.2|"
    c.Add "ENG|MC|M|2|U|What is the main purpose of the cooler in a lube oil system?||To raise the oil pressure;*To take heat out of the oil;To filter the oil;To store spare oil|EPB~2.4|"
    c.Add "ENG|SA|H|8|U|A diesel generator trips offline on low lube oil pressure.|||EPB~5.3|List three likely causes.~4~U~Any three of: low sump level, a clogged strainer, a failed pump, a leaking line, a faulty sensor.^Describe the checks made before it is restarted.~4~U~Check the sump level, the strainers and lines, and the sensor, then get permission from the engineering officer of the watch."
    c.Add "ENG|MC|H|4|U|What does a sudden drop in a motor's insulation resistance most likely show?||Normal aging only;*Moisture or damage in the windings;Overspeed;A loose coupling|EPB~4.5|"
    ' Communications and Security (8)
    c.Add "COM|MC|E|2|U|What is the purpose of the phonetic alphabet?||To encrypt messages;*To spell words clearly over voice circuits;To shorten messages;To mark a message's priority|CPG~1.2|"
    c.Add "COM|TF|E|1|U|The proword OUT means the transmission is complete and no reply is expected.|True||CPG~1.4|"
    c.Add "COM|MC|M|3|U|Which of these message precedences is the highest?||Routine;Priority;Immediate;*Flash|CPG~2.1|"
    c.Add "COM|SA|M|4|U|Explain why classified matters are never discussed on an unsecured circuit.|An unsecured circuit can be monitored, so anything said on it can be collected and used by an adversary.||CPG~3.1|"
    c.Add "COM|TF|M|1|U|A user may share a smart card PIN with a trusted shipmate.|False||CPG~4.2|"
    c.Add "COM|MC|M|2|U|What should a user do with a suspicious email that has an attachment?||Open the attachment to check it;Forward it to the division;*Report it without opening the attachment;Reply and ask the sender|CPG~4.4|"
    c.Add "COM|SA|H|6|U|A watchstander finds an unknown removable drive in a secure space.|||CPG~4.6|State the immediate action.~3~U~Do not connect it. Secure it and report it to the security officer.^Explain the risk it poses.~3~CUI~It may carry malicious code or data that could compromise the network."
    c.Add "COM|MC|E|2|U|Which proword means I have received your last transmission satisfactorily?||Over;*Roger;Wait;Break|CPG~1.4|"
    MoreBankSpecs c
    Set BankSpecs = c
End Function

' The rest of the bank (44): more of the first four topics, then seamanship
' and hazardous materials. Kept apart so neither list grows too long to read.
Private Sub MoreBankSpecs(ByVal c As Collection)
    ' Damage Control (6 more)
    c.Add "DC|MC|M|2|U|Which material condition gives the least watertight integrity and is set in port during working hours?||*X-RAY;YOKE;ZEBRA;WILLIAM|DCTM~5.3|"
    c.Add "DC|TF|M|1|U|A reflash watch is posted after a fire is reported out.|True||DCTM~6.4|"
    c.Add "DC|SA|M|4|U|Explain why a space is tested for oxygen and toxic gases after a fire.|Burning uses up oxygen and gives off toxic gases, so the air must be proved safe before anyone enters without breathing apparatus.||SSP~2.6|"
    c.Add "DC|MC|E|2|U|What does a fitting marked with a black W stay as?||*Open at all times unless damage control needs it closed;Closed at general quarters;Closed in port;Open only in port|DCTM~5.5|"
    c.Add "DC|MC|H|4|U|Flooding is reported in a space below the waterline. What is done first?||Start the dewatering pumps;*Report it and set the boundaries;Open the space to inspect it;Shift ballast|DCTM~7.2|"
    c.Add "DC|SA|H|6|U|A compartment is flooding through a cracked pipe.|||DCTM~7.4|Name two ways to stop or slow the leak.~3~U~Any two of: secure the valves that feed the pipe, a pipe patch, a jubilee clamp, wedges and rags.^Explain why the boundaries are set before the patch.~3~U~So the water can't spread to the next spaces if the patch fails."
    ' Navigation (6 more)
    c.Add "NAV|MC|M|2|U|What is the purpose of a range made of two lights or markers?||To show an isolated danger;*To keep the ship on the centerline of a channel;To mark a wreck;To give the depth|NFH~4.6|"
    c.Add "NAV|TF|E|1|U|A bearing taken from the ship to an object is a line of position.|True||NFH~3.2|"
    c.Add "NAV|SA|M|4|U|Explain what a dead reckoning position is.|A position worked out from the last fix using only the ordered course and speed, with no allowance for current or wind.||NFH~3.5|"
    c.Add "NAV|MC|H|3|U|The compass shows 095 and the deviation is 3 degrees east. What is the magnetic heading?||092;*098;095;101|NFH~2.4|"
    c.Add "NAV|MC|E|2|U|What does a green can buoy mark when entering from sea in IALA Region B?||The starboard side;*The port side;A mid-channel point;A junction|NFH~4.3|"
    c.Add "NAV|TF|M|1|U|Charts should be corrected with the latest notices before they are used.|True||NFH~6.1|"
    ' Engineering Fundamentals (6 more)
    c.Add "ENG|MC|E|2|U|What does a strainer in a pump suction line do?||Raise the pressure;*Keep debris out of the pump;Measure the flow;Cool the liquid|EPB~3.1|"
    c.Add "ENG|TF|E|1|U|A ground in an electrical circuit can cause a shock to anyone who touches the equipment.|True||EPB~4.2|"
    c.Add "ENG|SA|M|4|U|Explain why an electric motor is checked for heat while it runs.|Too much heat shows overload, poor cooling or bearing trouble, and it damages the insulation if it goes on.||EPB~4.4|"
    c.Add "ENG|MC|M|3|U|Which reading shows a pump is losing suction?||Steady discharge pressure;*Falling discharge pressure and a noisy pump;A cool casing;Rising tank level|EPB~3.4|"
    c.Add "ENG|TF|M|1|U|A relief valve may be gagged shut during normal operation to stop it leaking.|False||EPB~3.5|"
    c.Add "ENG|MC|H|4|U|Why is a diesel engine run unloaded for a time before it is shut down?||To fill the fuel tanks;*To let its temperatures fall evenly;To charge the batteries;To test the governor|EPB~5.2|"
    ' Communications and Security (6 more)
    c.Add "COM|MC|M|2|U|Which proword means my transmission is ended and I expect a reply?||Out;*Over;Roger;Wilco|CPG~1.4|"
    c.Add "COM|TF|E|1|U|A user must lock the workstation before leaving it, even for a moment.|True||CPG~4.1|"
    c.Add "COM|SA|M|4|U|Explain what the need-to-know rule means.|A person sees protected information only when their duties require it, even if they hold the clearance for it.||CPG~3.2|"
    c.Add "COM|MC|M|2|U|What is the best way to dispose of paper with controlled unclassified information on it?||Put it in the trash;*Shred it in an approved shredder;Tear it in half;Recycle it|CPG~3.4|"
    c.Add "COM|MC|E|2|U|What does the proword WILCO mean?||*I have received your message, understand it and will comply;Wait for a reply;Say again;I am ending the call|CPG~1.4|"
    c.Add "COM|TF|M|1|U|Personal devices may be charged from a government computer's USB port.|False||CPG~4.5|"
    ' Seamanship (10)
    c.Add "SEA|MC|E|2|U|What is the line called that is thrown to the pier first to pass a heavier mooring line?||Spring line;*Heaving line;Breast line;Painter|SH~2.1|"
    c.Add "SEA|TF|E|1|U|Line handlers should stand in the bight of a line so they can control it.|False||SH~2.3|"
    c.Add "SEA|MC|M|2|U|Which knot makes a fixed loop at the end of a line that will not slip?||Clove hitch;*Bowline;Square knot;Half hitch|SH~1.4|"
    c.Add "SEA|SA|M|4|U|Explain what scope means when anchoring.|The length of anchor chain let out compared with the depth of the water; more scope holds better.||SH~5.2|"
    c.Add "SEA|MC|M|3|U|What is the main danger when a mooring line under strain parts?||It sinks;*It snaps back and can strike people;It catches fire;It fouls the screw|SH~2.5|"
    c.Add "SEA|TF|M|1|U|An inherently buoyant life jacket must be worn by everyone working over the side.|True||SH~6.1|"
    c.Add "SEA|MC|E|2|U|What is the forward end of a ship called?||Stern;*Bow;Beam;Quarter|SH~1.1|"
    c.Add "SEA|SA|H|6|U|The ship is lowering a small boat at sea.|||SH~7.3|List two checks made before the boat is lowered.~3~U~Any two of: the crew wear life jackets and helmets, the plug is in, the falls and gear are sound, the engine starts.^State who controls the lowering.~3~U~The safety officer or boat officer at the davit, on orders from the bridge."
    c.Add "SEA|MC|H|4|U|While anchored, the ship's position shows it is moving steadily downwind. What is happening?||Swinging on the anchor;*Dragging the anchor;Riding to a short stay;Veering chain|SH~5.4|"
    c.Add "SEA|TF|E|1|U|Port is the left side of the ship when facing forward.|True||SH~1.1|"
    ' Hazardous Materials (10)
    c.Add "HAZ|MC|E|2|U|Which document gives the hazards, handling and first aid for a hazardous material?||The purchase order;*The safety data sheet;The tag-out record;The ship's plan|HMC~2.1|"
    c.Add "HAZ|TF|E|1|U|Flammable liquids may be stored in a berthing space if they are in sealed cans.|False||HMC~3.2|"
    c.Add "HAZ|MC|M|2|U|Why are oxidizers stored apart from flammable materials?||They spoil faster;*Together they can start or feed a fire;They need refrigeration;They are heavier|HMC~3.4|"
    c.Add "HAZ|SA|M|4|U|List the first three actions on finding a hazardous material spill.|Warn others and leave the area if needed, report the spill, and keep people out until the spill team arrives.||HMC~5.1|"
    c.Add "HAZ|MC|M|3|U|What should be checked before using personal protective equipment for a chemical?||Its color;*That it is rated for that chemical and in good condition;Its weight;Who used it last|HMC~4.2|"
    c.Add "HAZ|TF|M|1|U|An unlabeled container of liquid should be treated as hazardous until it is identified.|True||HMC~2.4|"
    c.Add "HAZ|MC|E|2|U|What does a hazardous material label tell the user?||The price;*The material, its hazards and how to handle it;The shelf it came from;The supplier's address only|HMC~2.2|"
    c.Add "HAZ|SA|H|6|U|A sailor gets a cleaning chemical splashed in the eyes.|||HMC~6.1|Give the immediate first aid.~3~U~Flush the eyes with clean water at an eyewash for at least 15 minutes, holding the eyelids open.^State what is done next.~3~U~Report it, get medical help, and bring the safety data sheet for the chemical."
    c.Add "HAZ|MC|H|4|U|Two chemicals in the same locker begin to give off fumes. What is done first?||Mix in water to dilute them;*Leave, warn others and report it;Carry them to the weather deck;Open them to check the labels|HMC~5.3|"
    c.Add "HAZ|TF|M|1|U|Used rags with oil or solvent on them are kept in a closed metal container.|True||HMC~3.6|"
End Sub

' Questions in every other state, and one question's history of versions.
Private Sub AddWorkflowQuestions()
    Dim qid As Long, baseId As Long
    ' Drafts, not yet submitted.
    WriteQuestion "NAV|MC|M|2|U|Which chart projection shows a rhumb line as a straight line?||*Mercator;Gnomonic;Polar;Lambert|NFH~1.5|", "jpatel"
    WriteQuestion "COM|SA|E|3|U|Describe how a garbled message is handled on receipt.|Log the time and circuit, ask for a repeat of the garbled part, and note it in the log.||CPG~2.3|", "ckim"
    WriteQuestion "DC|TF|E|1|U|Carbon dioxide extinguishers are safe to use on energized electrical equipment.|True||DCTM~3.4|", "srivera"
    WriteQuestion "ENG|MC|M|2|U|What does a tachometer measure?||Pressure;Temperature;*Shaft speed;Current|EPB~1.5|", "achen"
    WriteQuestion "SEA|MC|E|2|U|Which line keeps a ship from moving forward or aft along the pier?||Breast line;*Spring line;Heaving line;Messenger|SH~2.2|", "ckim"
    WriteQuestion "HAZ|SA|M|4|U|Explain why hazardous materials are issued in the smallest amount needed.|Less material in work spaces means less to spill, burn or store, and less waste to get rid of.||HMC~1.3|", "jpatel"

    ' Waiting at the DH.
    qid = WriteQuestion("NAV|TF|E|1|U|Soundings on most U.S. charts show depths at mean lower low water.|True||NFH~6.4|", "jpatel")
    SubmitAs qid, "jpatel"
    qid = WriteQuestion("DC|MC|M|2|U|Which agent is used on a Class A fire in bedding and rags?||Dry chemical only;*Water;Carbon dioxide only;Sand|DCTM~3.3|", "achen")
    SubmitAs qid, "achen"
    qid = WriteQuestion("SEA|SA|M|4|U|Explain why line handlers stay clear of the snap-back zone.|A line that parts under strain whips back with great force along a path that can be predicted, and can kill or maim.||SH~2.5|", "achen")
    SubmitAs qid, "achen"
    qid = WriteQuestion("HAZ|MC|E|2|U|Who may issue hazardous material from a storeroom?||Anyone in the division;*Trained and authorized issuers;Only the commanding officer;The supplier|HMC~1.2|", "srivera")
    SubmitAs qid, "srivera"
    ' Waiting at the XO.
    qid = WriteQuestion("COM|MC|M|2|U|Which of these makes the strongest password?||password1;Summer2026;*A long passphrase of unrelated words;The ship's name|CPG~4.3|", "ckim")
    SubmitAs qid, "ckim"
    ApproveAs qid, "dreyes"
    qid = WriteQuestion("ENG|TF|M|1|U|An electrical safety check is needed before portable equipment is used aboard ship.|True||EPB~4.3|", "srivera")
    SubmitAs qid, "srivera"
    ApproveAs qid, "dreyes"
    qid = WriteQuestion("SEA|MC|M|3|U|What does the order avast mean to line handlers?||Pull harder;*Stop what you are doing;Let go the line;Take in slack|SH~2.4|", "ckim")
    SubmitAs qid, "ckim"
    ApproveAs qid, "dreyes"
    ' Waiting at the SM.
    qid = WriteQuestion("ENG|SA|M|4|U|Explain the purpose of a relief valve.|It opens at a set pressure to protect the system from too much pressure.||EPB~3.5|", "srivera")
    SubmitAs qid, "srivera"
    ApproveAs qid, "dreyes"
    ApproveAs qid, "mhale"
    qid = WriteQuestion("NAV|MC|M|2|U|What is the main use of a gyrocompass repeater on the bridge wing?||Steering the ship;*Taking visual bearings;Measuring speed;Plotting depth|NFH~3.3|", "jpatel")
    SubmitAs qid, "jpatel"
    ApproveAs qid, "dreyes"
    ApproveAs qid, "mhale"
    qid = WriteQuestion("HAZ|TF|E|1|U|Aerosol cans may be thrown in the general trash once they are empty.|False||HMC~7.1|", "jpatel")
    SubmitAs qid, "jpatel"
    ApproveAs qid, "dreyes"
    ApproveAs qid, "mhale"

    ' Rejected by the DH, the XO and the SM, back with their authors.
    qid = WriteQuestion("SEA|TF|E|1|U|The anchor is let go by the boatswain without an order from the bridge.|False||SH~5.1|", "achen")
    SubmitAs qid, "achen"
    RejectAs qid, "dreyes", "A true/false question this easy tells us little. Ask who gives the order instead."
    qid = WriteQuestion("DC|MC|M|2|U|Which extinguishing agent is best for a Class D fire?||Water;Foam;*Dry powder made for metal fires;Carbon dioxide|DCTM~3.5|", "achen")
    SubmitAs qid, "achen"
    ApproveAs qid, "dreyes"
    RejectAs qid, "mhale", "The correct option names metal fires, which gives the answer away. Reword the options so they read alike."
    qid = WriteQuestion("COM|SA|M|4|U|Explain why messages are numbered on a circuit.|So the receiver can tell when one is missed and ask for it again.||CPG~2.2|", "jpatel")
    SubmitAs qid, "jpatel"
    ApproveAs qid, "dreyes"
    ApproveAs qid, "mhale"
    RejectAs qid, "tbrooks", "Cite the paragraph that covers message serial numbers, not the general one."
    ' Rejected by the DH, fixed by its author and sent again: back at the DH.
    qid = WriteQuestion("HAZ|MC|M|2|U|Where is a safety data sheet kept?||In the supply office only;*Where the material is used and stored;With the purchase order;Nowhere aboard|HMC~2.1|", "srivera")
    SubmitAs qid, "srivera"
    RejectAs qid, "dreyes", "The wrong options are too easy to rule out. Make them things a sailor might really believe."
    ReviseQuestion qid, "srivera", "Where must a copy of the safety data sheet be available to the people using the material?"
    SubmitAs qid, "srivera"

    ' Versions 2, approved, which retire versions 1.
    baseId = BankQuestionId("DC", "Water is the preferred agent")
    qid = WriteNewVersion(baseId, "srivera", "Water from a fire hose is the preferred agent for a fire in energized electrical equipment.")
    SubmitAs qid, "srivera"
    ApproveChain qid
    baseId = BankQuestionId("NAV", "Which publication lists")
    qid = WriteNewVersion(baseId, "jpatel", "Which publication lists the color, period and range of the lights along a coast?")
    SubmitAs qid, "jpatel"
    ApproveChain qid
    baseId = BankQuestionId("HAZ", "Used rags with oil")
    qid = WriteNewVersion(baseId, "jpatel", "Used rags with oil or solvent on them are kept in a closed, labeled metal container until disposal.")
    SubmitAs qid, "jpatel"
    ApproveChain qid
    ' A version 2 in draft, and one waiting on the DH: versions 1 stay in use
    ' until they are approved.
    baseId = BankQuestionId("NAV", "How many minutes of latitude")
    WriteNewVersion baseId, "jpatel", "How many minutes of latitude make one nautical mile at any latitude?"
    baseId = BankQuestionId("ENG", "What does a gauge reading of zero psig")
    qid = WriteNewVersion(baseId, "achen", "What does a gauge reading of zero psig mean on a system open to the air?")
    SubmitAs qid, "achen"
    Note "Other questions: 6 drafts; 5 waiting on the DH (one of them sent back once and fixed), 3 on the XO and 3 on the SM; " & _
        "1 rejected by each of the DH, XO and SM; 3 approved versions 2 (versions 1 retired); a version 2 in draft and one " & _
        "waiting on the DH."
End Sub

Private Sub RejectAs(ByVal qid As Long, ByVal who As String, ByVal reason As String)
    Dim r As String
    ActAs who
    r = RejectQuestion(qid, reason)
    If Len(r) > 0 Then Err.Raise 5, "modSeedDemo", "Reject question " & qid & " as " & who & ": " & r
End Sub

' Write a question as the editor does: its parts, options and references in
' the working copy, the editor's checks for submitting, then SaveQuestion.
' Returns its QuestionID.
Private Function WriteQuestion(ByVal spec As String, ByVal author As String) As Long
    Dim f() As String, topicId As Long, typeName As String, diffId As Long, pts As Double, cls As Long
    Dim problems As String, r As String, qid As Long
    f = Split(spec, "|")
    If UBound(f) <> 9 Then Err.Raise 5, "modSeedDemo", "A question line needs 10 fields: " & Left$(spec, 60)
    ActAs author
    topicId = m_topics(f(0))
    typeName = TypeNameOf(f(1))
    diffId = DifficultyId(f(2))
    pts = Val(f(3))
    cls = ClassId(f(4))
    FillWork f(7), f(8), f(9)
    problems = QuestionProblems(topicId, typeName, diffId, pts, cls, f(5), True)
    If Len(problems) > 0 Then Err.Raise 5, "modSeedDemo", "Question """ & Left$(f(5), 50) & """: " & problems
    r = SaveQuestion(0, 0, topicId, typeName, diffId, pts, cls, f(5), NullIfEmpty(f(6)), qid)
    If Len(r) > 0 Then Err.Raise 5, "modSeedDemo", "Save: " & r
    WriteQuestion = qid
End Function

' Start a new version of approved question baseId as author, with new text
' and the rest as it stands. Returns the new version's QuestionID.
Private Function WriteNewVersion(ByVal baseId As Long, ByVal author As String, ByVal newText As String) As Long
    WriteNewVersion = SaveFrom(baseId, 0, baseId, author, newText)
End Function

' Change the text of draft or rejected question qid as author, as the
' editor does, keeping the rest as it stands.
Private Sub ReviseQuestion(ByVal qid As Long, ByVal author As String, ByVal newText As String)
    SaveFrom qid, qid, 0, author, newText
End Sub

' Load question srcId into the working copy and save it with newText as
' author: as qid (0 for a new row) and, for a new version, baseId. Returns
' the saved QuestionID.
Private Function SaveFrom(ByVal srcId As Long, ByVal qid As Long, ByVal baseId As Long, ByVal author As String, _
        ByVal newText As String) As Long
    Dim rs As DAO.Recordset, topicId As Long, typeName As String, diffId As Long, pts As Double, cls As Variant
    Dim ans As Variant, problems As String, r As String, savedId As Long
    ActAs author
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblQuestions WHERE QuestionID=" & srcId, dbOpenSnapshot)
    topicId = rs!TopicID: typeName = QTypeName(rs!QuestionTypeID): diffId = rs!DifficultyID
    pts = rs!TotalPoints: cls = rs!ClassLevelID: ans = rs!ReferenceAnswer
    rs.Close
    LoadWork srcId
    problems = QuestionProblems(topicId, typeName, diffId, pts, cls, newText, True)
    If Len(problems) > 0 Then Err.Raise 5, "modSeedDemo", "Question " & srcId & ": " & problems
    r = SaveQuestion(qid, baseId, topicId, typeName, diffId, pts, cls, newText, ans, savedId)
    If Len(r) > 0 Then Err.Raise 5, "modSeedDemo", "Save question " & srcId & ": " & r
    SaveFrom = savedId
End Function

' The editor's working copy for one question: options (none for TF or SA),
' one reference, and the parts.
Private Sub FillWork(ByVal opts As String, ByVal ref As String, ByVal parts As String)
    Dim db As DAO.Database, v As Variant, n As Long, isRight As Boolean, t As String, p() As String, rf() As String
    Set db = CurrentDb
    ClearWork
    If Len(opts) > 0 Then
        For Each v In Split(opts, ";")
            n = n + 1
            t = CStr(v)
            isRight = (Left$(t, 1) = "*")
            If isRight Then t = Mid$(t, 2)
            db.Execute "INSERT INTO wq_option (part_id, OptionText, IsCorrect, SortOrder) VALUES (Null, " & Q(t) & ", " & _
                IIf(isRight, "True", "False") & ", " & n & ")", dbFailOnError
        Next
    End If
    If Len(parts) > 0 Then
        n = 0
        For Each v In Split(parts, "^")
            p = Split(v, "~")
            n = n + 1
            db.Execute "INSERT INTO wq_part (PartLabel, PartText, Points, ClassLevelID, ReferenceAnswer, SortOrder) VALUES (" & _
                Q(Chr$(64 + n)) & ", " & Q(p(0)) & ", " & Trim$(Str$(Val(p(1)))) & ", " & ClassId(p(2)) & ", " & Q(p(3)) & ", " & n & ")", _
                dbFailOnError
        Next
    End If
    rf = Split(ref, "~")
    db.Execute "INSERT INTO wq_ref (RefID, Paragraph) VALUES (" & m_refs(rf(0)) & ", " & Q(rf(1)) & ")", dbFailOnError
End Sub

Private Function TypeNameOf(ByVal code As String) As String
    Select Case code
        Case "MC": TypeNameOf = QT_MC
        Case "TF": TypeNameOf = QT_TF
        Case "SA": TypeNameOf = QT_SA
        Case Else: Err.Raise 5, "modSeedDemo", "No question type " & code & "."
    End Select
End Function

Private Function DifficultyId(ByVal code As String) As Long
    Dim nm As String
    Select Case code
        Case "E": nm = "Easy"
        Case "M": nm = "Medium"
        Case "H": nm = "Hard"
    End Select
    DifficultyId = Nz(DLookup("DifficultyID", "tblDifficultyLevels", "LevelName=" & Q(nm)), 0)
    If DifficultyId = 0 Then Err.Raise 5, "modSeedDemo", "The data file has no difficulty " & nm & "."
End Function

' The classification level with portion mark (U) or (CUI).
Private Function ClassId(ByVal code As String) As Long
    ClassId = Nz(DLookup("ClassLevelID", "tblClassLevels", "PortionMark=" & Q("(" & code & ")")), 0)
    If ClassId = 0 Then Err.Raise 5, "modSeedDemo", "The data file has no portion marking (" & code & ")."
End Function

' The current approved question in topic code whose text starts with
' textStart.
Private Function BankQuestionId(ByVal code As String, ByVal textStart As String) As Long
    BankQuestionId = Nz(DLookup("QuestionID", "tblQuestions", "TopicID=" & m_topics(code) & " AND IsCurrent AND StatusID=" & _
        QStatusID(QS_APPROVED) & " AND QuestionText Like " & Q(textStart & "*")), 0)
    If BankQuestionId = 0 Then Err.Raise 5, "modSeedDemo", "No approved question starting """ & textStart & """."
End Function

Private Sub SubmitAs(ByVal qid As Long, ByVal who As String)
    Dim r As String
    ActAs who
    r = SubmitQuestion(qid)
    If Len(r) > 0 Then Err.Raise 5, "modSeedDemo", "Submit question " & qid & ": " & r
End Sub

Private Sub ApproveAs(ByVal qid As Long, ByVal who As String)
    Dim r As String, done As String
    ActAs who
    r = ApproveQuestion(qid, done)
    If Len(r) > 0 Then Err.Raise 5, "modSeedDemo", "Approve question " & qid & " as " & who & ": " & r
End Sub

Private Sub ApproveChain(ByVal qid As Long)
    ApproveAs qid, "dreyes"
    ApproveAs qid, "mhale"
    ApproveAs qid, "tbrooks"
End Sub

' ---------------------------------------------------------------------------
' Exams
' ---------------------------------------------------------------------------

Private Sub AddExams(ByVal makeDocs As Boolean)
    Dim basic As Long, refresher As Long, deck As Long, comp As Long, e As Long, s As String
    basic = TemplateIdOf("Basic Qualification Board")
    refresher = TemplateIdOf("Damage Control Refresher")
    deck = TemplateIdOf("Deck and Hazmat Qualification")
    comp = TemplateIdOf("Comprehensive Board")

    ' Approved, the whole way to their Word documents (oldest first, so the
    ' Statistics page has them). The second has a question swapped first.
    e = Generate(basic, "USS Harborview (demo)")
    SignOffAll e, ""
    ApproveExamChain e, makeDocs
    s = "approved E" & e
    e = Generate(basic, "USS Ridgeline (demo)")
    SwapOne e, "DC"
    SignOffAll e, ""
    ApproveExamChain e, makeDocs
    s = s & ", E" & e & " (a question swapped)"
    e = Generate(deck, "USS Bayfield (demo)")
    SignOffAll e, ""
    ApproveExamChain e, makeDocs
    s = s & ", E" & e
    e = Generate(comp, "Surface Warfare School (demo)")
    SignOffAll e, ""
    ApproveExamChain e, makeDocs
    s = s & ", E" & e & "; "

    ' Waiting on the DH, the XO and the SM.
    e = Generate(comp, "USS Kestrel (demo)")
    SignOffAll e, ""
    s = s & "waiting on the DH E" & e
    e = Generate(refresher, "Damage Control School (demo)")
    SignOffAll e, ""
    ApproveExamAs e, "dreyes", makeDocs
    s = s & ", on the XO E" & e
    e = Generate(deck, "USS Glenmore (demo)")
    SignOffAll e, ""
    ApproveExamAs e, "dreyes", makeDocs
    ApproveExamAs e, "mhale", makeDocs
    s = s & ", on the SM E" & e
    ' Sent back by the DH, signed off again by its SMEs and approved by the
    ' DH: version 2, waiting on the XO.
    e = Generate(refresher, "Fleet Training Center (demo)")
    SignOffAll e, ""
    RejectExamAs e, "dreyes", "The engineering question repeats one from last month's exam. Swap it.", "ENG"
    SwapOne e, "ENG"
    SignOffAll e, ""
    ApproveExamAs e, "dreyes", makeDocs
    s = s & ", on the XO again E" & e & " (version 2); "

    ' Sent back to their SMEs by the DH, the XO and the SM.
    e = Generate(basic, "USS Cedar Point (demo)")
    SignOffAll e, ""
    RejectExamAs e, "dreyes", "Two damage control questions cover the same fire class. Swap one for a question on flooding or " & _
        "material conditions.", "DC"
    s = s & "sent back by the DH E" & e
    e = Generate(deck, "Deck Department Training (demo)")
    SignOffAll e, ""
    ApproveExamAs e, "dreyes", makeDocs
    RejectExamAs e, "mhale", "This board is for new deck seamen. Every topic needs easier questions; send it back to me when " & _
        "they are in.", ""
    s = s & ", by the XO E" & e & " (every topic)"
    e = Generate(basic, "USS Northpoint (demo)")
    SignOffAll e, ""
    ApproveExamAs e, "dreyes", makeDocs
    ApproveExamAs e, "mhale", makeDocs
    RejectExamAs e, "tbrooks", "The navigation questions are both on buoys. Give the board one on fixes or compass error.", "NAV"
    s = s & ", by the SM E" & e & "; "

    ' Still with their SMEs: one topic signed off, and none yet.
    e = Generate(refresher, "Engineering Training Unit (demo)")
    SignOffAll e, "DC"
    s = s & "with their SMEs E" & e & " (engineering signed off)"
    e = Generate(comp, "Engineering Duty School (demo)")
    s = s & " and E" & e & " (nothing signed off)."

    Note "Exams: " & s
    If makeDocs Then
        Note "Word documents: " & m_docsOK & " made" & IIf(m_docsFailed > 0, ", " & m_docsFailed & " FAILED (see above)", "") & _
            ", in " & ExamDocFolder()
    Else
        Note "Word documents: skipped (-NoDocs). Download a copy on the Exams page makes them."
    End If
End Sub

Private Function TemplateIdOf(ByVal nm As String) As Long
    TemplateIdOf = Nz(DLookup("TemplateID", "tblTemplates", "TemplateName=" & Q(nm)), 0)
    If TemplateIdOf = 0 Then Err.Raise 5, "modSeedDemo", "No template " & nm & "."
End Function

' Send exam examId back to its SMEs as who: topic code's topic, or every
' topic when code is "".
Private Sub RejectExamAs(ByVal examId As Long, ByVal who As String, ByVal reason As String, ByVal code As String)
    Dim r As String, etId As Long
    If Len(code) > 0 Then etId = ExamTopicId(examId, code)
    ActAs who
    r = RejectExam(examId, reason, etId)
    If Len(r) > 0 Then Err.Raise 5, "modSeedDemo", "Send back exam " & examId & " as " & who & ": " & r
End Sub

' Generate an exam as the DH, letting every topic take questions given
' lately (the bank is small, so the overlap rule would stop the later exams).
Private Function Generate(ByVal tplId As Long, ByVal recipient As String) As Long
    Dim r As String, examId As Long, allIds As String, k As Variant
    For Each k In m_topics.Keys
        allIds = allIds & IIf(Len(allIds) > 0, ",", "") & m_topics(k)
    Next
    ActAs "dreyes"
    r = GenerateExam(tplId, recipient, allIds, examId)
    If Len(r) > 0 Then Err.Raise 5, "modSeedDemo", "Generate for " & recipient & ": " & r
    Generate = examId
End Function

Private Function ExamTopicId(ByVal examId As Long, ByVal code As String) As Long
    ExamTopicId = Nz(DLookup("ExamTopicID", "tblExamTopics", "ExamID=" & examId & " AND TopicID=" & m_topics(code)), 0)
    If ExamTopicId = 0 Then Err.Raise 5, "modSeedDemo", "Exam " & examId & " has no " & code & " topic."
End Function

' The SME with an open review of exam topic etId (the first, when several).
Private Function ReviewerOf(ByVal etId As Long) As Long
    ReviewerOf = Nz(DMin("AssignedPersonID", "tblTasks", "ExamTopicID=" & etId & " AND TaskStatus=" & Q(TS_OPEN)), 0)
    If ReviewerOf = 0 Then Err.Raise 5, "modSeedDemo", "Exam topic " & etId & " has no open review."
End Function

' Sign off every topic of examId that waits for it, but skipCode's, each by
' an SME with an open review of it, as My reviews does.
Private Sub SignOffAll(ByVal examId As Long, ByVal skipCode As String)
    Dim rs As DAO.Recordset, ids As New Collection, v As Variant, r As String, done As String, skipId As Long
    If Len(skipCode) > 0 Then skipId = m_topics(skipCode)
    Set rs = CurrentDb.OpenRecordset("SELECT ExamTopicID, TopicID FROM tblExamTopics WHERE ExamID=" & examId & _
        " AND SignedOffByPersonID Is Null ORDER BY ExamTopicID", dbOpenSnapshot)
    Do While Not rs.EOF
        If rs!TopicID <> skipId Then ids.Add CLng(rs!ExamTopicID)
        rs.MoveNext
    Loop
    rs.Close
    For Each v In ids
        ActAsId ReviewerOf(v)
        r = SignOffTopic(v, TopicChangeNo(v), done)
        If Len(r) > 0 Then Err.Raise 5, "modSeedDemo", "Sign off exam topic " & v & ": " & r
    Next
End Sub

' Swap the first question of examId's topic code for the first approved
' question the swap rules allow, as the Swap dialog does.
Private Sub SwapOne(ByVal examId As Long, ByVal code As String)
    Dim etId As Long, eqId As Long, rs As DAO.Recordset, stamp As Long, r As String, done As String
    etId = ExamTopicId(examId, code)
    eqId = DMin("ExamQuestionID", "tblExamQuestions", "ExamTopicID=" & etId)
    ActAsId ReviewerOf(etId)
    stamp = TopicChangeNo(etId)
    Set rs = CurrentDb.OpenRecordset("SELECT QuestionID FROM tblQuestions WHERE TopicID=" & m_topics(code) & " AND IsCurrent AND " & _
        "StatusID=" & QStatusID(QS_APPROVED) & " AND QuestionID NOT IN (SELECT QuestionID FROM tblExamQuestions WHERE ExamID=" & _
        examId & ") ORDER BY QuestionID DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        If Len(SwapProblem(eqId, rs!QuestionID)) = 0 Then
            r = SwapQuestion(eqId, rs!QuestionID, stamp, done)
            rs.Close
            If Len(r) > 0 Then Err.Raise 5, "modSeedDemo", "Swap on exam " & examId & ": " & r
            Exit Sub
        End If
        rs.MoveNext
    Loop
    rs.Close
    Err.Raise 5, "modSeedDemo", "No question could be swapped onto exam " & examId & "'s " & code & " topic."
End Sub

Private Sub ApproveExamChain(ByVal examId As Long, ByVal makeDocs As Boolean)
    ApproveExamAs examId, "dreyes", makeDocs
    ApproveExamAs examId, "mhale", makeDocs
    ApproveExamAs examId, "tbrooks", makeDocs
End Sub

' Approve exam examId at its step as who; on the final approval make its
' Word document, as Approvals does after the SM approves. A document that
' fails is reported and the seed goes on: the approval stands either way.
Private Sub ApproveExamAs(ByVal examId As Long, ByVal who As String, ByVal makeDocs As Boolean)
    Dim r As String, done As String, isFinal As Boolean, docNote As String, num As Long, desc As String
    ActAs who
    r = ApproveExam(examId, done, isFinal)
    If Len(r) > 0 Then Err.Raise 5, "modSeedDemo", "Approve exam " & examId & " as " & who & ": " & r
    If Not (isFinal And makeDocs) Then Exit Sub
    On Error Resume Next
    r = MakeExamDocument(examId, docNote)
    num = Err.Number: desc = Err.Description
    On Error GoTo 0
    If num <> 0 Then r = "Word stopped: " & num & " " & desc
    If Len(r) > 0 Then
        m_docsFailed = m_docsFailed + 1
        Note "FAILED: the Word document for E" & examId & ": " & r
    Else
        m_docsOK = m_docsOK + 1
        Note "Word document for E" & examId & ": " & Nz(DLookup("OutputFile", "tblExams", "ExamID=" & examId), "?") & _
            IIf(Len(docNote) > 0, " (" & docNote & ")", "")
    End If
End Sub

' ---------------------------------------------------------------------------
' Dates
' ---------------------------------------------------------------------------

' Everything above happened a moment ago. Move it back so the Statistics
' page and the histories read like weeks of use, keeping the order: the
' references first, the bank over the weeks after, the exams in turn, the
' questions still in the chain in the last few days.
Private Sub SpreadDates()
    Dim db As DAO.Database, rs As DAO.Recordset, i As Long, n As Long, k As Variant, days As Long
    Set db = CurrentDb
    db.Execute "UPDATE tblAuditTrail SET ActionOn=DateAdd('d',-80,ActionOn) WHERE ItemType=" & Q(ITEM_REFERENCE), dbFailOnError
    ' The bank: 75 days back for the first question to 40 for the last;
    ' every version of a question moves with it.
    n = m_bank.Count
    For Each k In m_bank
        days = 75 - Int(35 * i / IIf(n > 1, n - 1, 1))
        Set rs = db.OpenRecordset("SELECT QuestionID FROM tblQuestions WHERE QuestionKey=" & k & " AND StatusID IN (" & _
            QStatusID(QS_APPROVED) & "," & QStatusID(QS_RETIRED) & ")", dbOpenSnapshot)
        Do While Not rs.EOF
            ShiftQuestion rs!QuestionID, days
            rs.MoveNext
        Loop
        rs.Close
        i = i + 1
    Next
    ' Questions still being written or approved: the last few days.
    Set rs = db.OpenRecordset("SELECT QuestionID FROM tblQuestions WHERE StatusID NOT IN (" & QStatusID(QS_APPROVED) & "," & _
        QStatusID(QS_RETIRED) & ") ORDER BY QuestionID", dbOpenSnapshot)
    i = 0
    Do While Not rs.EOF
        ShiftQuestion rs!QuestionID, 4 - (i Mod 4)
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    ' The exams, oldest first: 38 days back for the first to 1 for the last, after the bank.
    n = DCount("*", "tblExams")
    Set rs = db.OpenRecordset("SELECT ExamID FROM tblExams ORDER BY ExamID", dbOpenSnapshot)
    i = 0
    Do While Not rs.EOF
        ShiftExam rs!ExamID, 38 - Int(37 * i / IIf(n > 1, n - 1, 1))
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Note "Dates: spread over the last 80 days."
End Sub

Private Sub ShiftQuestion(ByVal qid As Long, ByVal days As Long)
    Dim db As DAO.Database
    Set db = CurrentDb
    db.Execute "UPDATE tblQuestions SET CreatedOn=DateAdd('d'," & -days & ",CreatedOn), RetiredOn=DateAdd('d'," & -days & _
        ",RetiredOn) WHERE QuestionID=" & qid, dbFailOnError
    db.Execute "UPDATE tblAuditTrail SET ActionOn=DateAdd('d'," & -days & ",ActionOn) WHERE ItemType=" & Q(ITEM_QUESTION) & _
        " AND ItemID=" & qid, dbFailOnError
End Sub

Private Sub ShiftExam(ByVal examId As Long, ByVal days As Long)
    Dim db As DAO.Database, d As String
    Set db = CurrentDb
    d = "DateAdd('d'," & -days & ","
    db.Execute "UPDATE tblExams SET CreatedOn=" & d & "CreatedOn), FinalApprovedOn=" & d & "FinalApprovedOn) WHERE ExamID=" & examId, _
        dbFailOnError
    db.Execute "UPDATE tblExamTopics SET SignedOffOn=" & d & "SignedOffOn) WHERE ExamID=" & examId, dbFailOnError
    db.Execute "UPDATE tblTasks SET CreatedOn=" & d & "CreatedOn), CompletedOn=" & d & "CompletedOn) WHERE ItemType=" & Q(ITEM_EXAM) & _
        " AND ItemID=" & examId, dbFailOnError
    db.Execute "UPDATE tblAuditTrail SET ActionOn=" & d & "ActionOn) WHERE ItemType=" & Q(ITEM_EXAM) & " AND ItemID=" & examId, _
        dbFailOnError
End Sub

' ---------------------------------------------------------------------------
' Summary
' ---------------------------------------------------------------------------

Private Sub Summarize()
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT s.StatusName, Count(*) AS n FROM tblQuestions AS q INNER JOIN tblStatuses AS s " & _
        "ON q.StatusID=s.StatusID GROUP BY s.StatusName ORDER BY s.StatusName", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & IIf(Len(s) > 0, ", ", "") & rs!StatusName & " " & rs!n
        rs.MoveNext
    Loop
    rs.Close
    Note "Questions by status: " & s & "."
    s = ""
    Set rs = CurrentDb.OpenRecordset("SELECT s.StatusName, Count(*) AS n FROM tblExams AS e INNER JOIN tblStatuses AS s " & _
        "ON e.StatusID=s.StatusID GROUP BY s.StatusName ORDER BY s.StatusName", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & IIf(Len(s) > 0, ", ", "") & rs!StatusName & " " & rs!n
        rs.MoveNext
    Loop
    rs.Close
    Note "Exams by status: " & s & "."
    Note "Waiting by step: questions " & ByStep("tblQuestions", QStatusID(QS_REVIEW)) & "; exams " & _
        ByStep("tblExams", ExamStatusID(ES_REVIEW)) & "."
    Note "Open SME reviews: " & DCount("*", "tblTasks", "TaskStatus=" & Q(TS_OPEN)) & ". Audit rows: " & DCount("*", "tblAuditTrail") & _
        ". Error log rows: " & DCount("*", "error_log") & "."
End Sub

' How many rows of tbl in status statusId wait at each step, as "DH 5, XO 3".
Private Function ByStep(ByVal tbl As String, ByVal statusId As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT r.RoleCode, r.ChainOrder, Count(*) AS n FROM " & tbl & " AS x INNER JOIN tblRoles AS r " & _
        "ON x.CurrentStepRoleID=r.RoleID WHERE x.StatusID=" & statusId & " GROUP BY r.RoleCode, r.ChainOrder ORDER BY r.ChainOrder", _
        dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & IIf(Len(s) > 0, ", ", "") & rs!RoleCode & " " & rs!n
        rs.MoveNext
    Loop
    rs.Close
    ByStep = IIf(Len(s) > 0, s, "none")
End Function
