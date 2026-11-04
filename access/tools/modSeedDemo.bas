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
'   topics      4 in use and 1 turned off; templates 2 in use
'   questions   36 approved in the bank, one with an approved version 2
'               (version 1 retired), one with a version 2 in draft, two
'               drafts, one waiting at each of the DH, XO and SM, one rejected
'   exams       E1 and E2 approved with their Word documents (E2 had a
'               question swapped), E3 waiting on the XO, E4 sent back to its
'               SMEs by the DH, E5 still in SME review with one topic signed off
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
' none). makeDocs False skips Word.
' Returns "OK" and the report, one line each, or "ERROR ..." and the report
' so far.
Public Function SeedDemoRun(ByVal bePath As String, ByVal ownerUser As String, ByVal ownerFirst As String, _
        ByVal ownerLast As String, ByVal ownerRank As String, ByVal demoPassword As String, ByVal makeDocs As Boolean) As String
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
    m_step = "the Personnel file": owner = SetUpOwner(ownerUser, ownerFirst, ownerLast, ownerRank, demoPassword)
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
    Note "Topics: 4 in use and 1 turned off. Templates: Basic Qualification Board (9 questions) and Damage Control Refresher (4)."
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
    Note "References: 5 made-up references."
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
    Set BankSpecs = c
End Function

' Questions in every other state, and one question's history of versions.
Private Sub AddWorkflowQuestions()
    Dim qid As Long, baseId As Long, r As String
    ' Drafts, not yet submitted.
    WriteQuestion "NAV|MC|M|2|U|Which chart projection shows a rhumb line as a straight line?||*Mercator;Gnomonic;Polar;Lambert|NFH~1.5|", "jpatel"
    WriteQuestion "COM|SA|E|3|U|Describe how a garbled message is handled on receipt.|Log the time and circuit, ask for a repeat of the garbled part, and note it in the log.||CPG~2.3|", "ckim"
    ' Waiting at the DH, the XO and the SM.
    qid = WriteQuestion("NAV|TF|E|1|U|Soundings on most U.S. charts show depths at mean lower low water.|True||NFH~6.4|", "jpatel")
    SubmitAs qid, "jpatel"
    qid = WriteQuestion("COM|MC|M|2|U|Which of these makes the strongest password?||password1;Summer2026;*A long passphrase of unrelated words;The ship's name|CPG~4.3|", "ckim")
    SubmitAs qid, "ckim"
    ApproveAs qid, "dreyes"
    qid = WriteQuestion("ENG|SA|M|4|U|Explain the purpose of a relief valve.|It opens at a set pressure to protect the system from too much pressure.||EPB~3.5|", "srivera")
    SubmitAs qid, "srivera"
    ApproveAs qid, "dreyes"
    ApproveAs qid, "mhale"
    ' Rejected by the XO, back with its author.
    qid = WriteQuestion("DC|MC|M|2|U|Which extinguishing agent is best for a Class D fire?||Water;Foam;*Dry powder made for metal fires;Carbon dioxide|DCTM~3.5|", "achen")
    SubmitAs qid, "achen"
    ApproveAs qid, "dreyes"
    ActAs "mhale"
    r = RejectQuestion(qid, "The correct option names metal fires, which gives the answer away. Reword the options so they read alike.")
    If Len(r) > 0 Then Err.Raise 5, "modSeedDemo", "Reject: " & r
    ' A version 2, approved, which retires version 1 (the water question).
    baseId = BankQuestionId("DC", "Water is the preferred agent")
    qid = WriteNewVersion(baseId, "srivera", "Water from a fire hose is the preferred agent for a fire in energized electrical equipment.")
    SubmitAs qid, "srivera"
    ApproveChain qid
    ' A version 2 in draft: version 1 stays in use until it is approved.
    baseId = BankQuestionId("NAV", "How many minutes of latitude")
    WriteNewVersion baseId, "jpatel", "How many minutes of latitude make one nautical mile at any latitude?"
    Note "Other questions: 2 drafts, 1 waiting on each of the DH, XO and SM, 1 rejected by the XO, " & _
        "1 with an approved version 2 (version 1 retired), 1 with a version 2 in draft."
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
    Dim rs As DAO.Recordset, topicId As Long, typeName As String, diffId As Long, pts As Double, cls As Variant
    Dim ans As Variant, problems As String, r As String, qid As Long
    ActAs author
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblQuestions WHERE QuestionID=" & baseId, dbOpenSnapshot)
    topicId = rs!TopicID: typeName = QTypeName(rs!QuestionTypeID): diffId = rs!DifficultyID
    pts = rs!TotalPoints: cls = rs!ClassLevelID: ans = rs!ReferenceAnswer
    rs.Close
    LoadWork baseId
    problems = QuestionProblems(topicId, typeName, diffId, pts, cls, newText, True)
    If Len(problems) > 0 Then Err.Raise 5, "modSeedDemo", "New version: " & problems
    r = SaveQuestion(0, baseId, topicId, typeName, diffId, pts, cls, newText, ans, qid)
    If Len(r) > 0 Then Err.Raise 5, "modSeedDemo", "New version: " & r
    WriteNewVersion = qid
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
    Dim basic As Long, refresher As Long, e1 As Long, e2 As Long, e3 As Long, e4 As Long, e5 As Long, r As String
    basic = DLookup("TemplateID", "tblTemplates", "TemplateName=" & Q("Basic Qualification Board"))
    refresher = DLookup("TemplateID", "tblTemplates", "TemplateName=" & Q("Damage Control Refresher"))

    ' E1: the whole way, to its Word document.
    e1 = Generate(basic, "USS Harborview (demo)")
    SignOffAll e1, ""
    ApproveExamChain e1, makeDocs

    ' E2: an SME swaps a question first.
    e2 = Generate(basic, "USS Ridgeline (demo)")
    SwapOne e2, "DC"
    SignOffAll e2, ""
    ApproveExamChain e2, makeDocs

    ' E3: approved by the DH, waiting on the XO.
    e3 = Generate(refresher, "Damage Control School (demo)")
    SignOffAll e3, ""
    ApproveExamAs e3, "dreyes", makeDocs

    ' E4: the DH sends its damage control topic back to the SMEs.
    e4 = Generate(basic, "USS Cedar Point (demo)")
    SignOffAll e4, ""
    ActAs "dreyes"
    r = RejectExam(e4, "Two damage control questions cover the same fire class. Swap one for a question on flooding or " & _
        "material conditions.", ExamTopicId(e4, "DC"))
    If Len(r) > 0 Then Err.Raise 5, "modSeedDemo", "Send back exam " & e4 & ": " & r

    ' E5: still with its SMEs, the engineering topic signed off.
    e5 = Generate(refresher, "Engineering Training Unit (demo)")
    SignOffAll e5, "DC"

    Note "Exams: E" & e1 & " and E" & e2 & " approved (E" & e2 & " had a question swapped), E" & e3 & " waiting on the XO, E" & _
        e4 & " sent back to its SMEs by the DH, E" & e5 & " in SME review with one topic signed off."
    If makeDocs Then
        Note "Word documents: " & m_docsOK & " made" & IIf(m_docsFailed > 0, ", " & m_docsFailed & " FAILED (see above)", "") & _
            ", in " & ExamDocFolder()
    Else
        Note "Word documents: skipped (-NoDocs). Download a copy on the Exams page makes them."
    End If
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
    ' The exams, oldest first.
    Set rs = db.OpenRecordset("SELECT ExamID FROM tblExams ORDER BY ExamID", dbOpenSnapshot)
    i = 0
    Do While Not rs.EOF
        ShiftExam rs!ExamID, Nz(Choose(i + 1, 30, 21, 12, 6, 2), 1)
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
    Note "Open SME reviews: " & DCount("*", "tblTasks", "TaskStatus=" & Q(TS_OPEN)) & ". Audit rows: " & DCount("*", "tblAuditTrail") & _
        ". Error log rows: " & DCount("*", "error_log") & "."
End Sub
