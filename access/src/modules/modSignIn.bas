Option Compare Database
Option Explicit

' Exam Generator (Access edition) - start-up and sign-in (build phase 2).
'
' frmMain's Form_Load runs AppStartUp under the splash:
'   1. the shared data file: KitEnsureBackend finds it next to the front end or
'      asks for it when it has moved;
'   2. its version (tblSettings SchemaVersion) against DATA_VERSION, so an old
'      copy stops against a data file a newer copy has changed;
'   3. sign-in: KitPersonnelStart finds or sets up the shared Personnel file and
'      signs the Windows user in (TempVars PersonID, PersonName, PersonIsAdmin);
'   4. the person's Exam Generator roles (LoadRoles), from tblPeople and
'      tblPersonRoles, into TempVars EGRoles (",SME,DH,"), EGIsAdmin,
'      EGNotifyDash and EGNotifyEmail. A person with no role, or turned off,
'      is told why and who to ask, and the app closes;
'   5. the help (KitHelpSync): topics new in this build's help.txt go into the
'      data file's help_topic table, then that table, with the admin's edits,
'      replaces the build's copy, so F1 shows the edited text.
' Then ShowPagesForRoles hides the sidebar pages they can't use, and
' frmMain.RoutePage sends any other way in (Ctrl+K) to the dashboard.
'
' My settings' Sign out (AppSignOut) clears the session and shows the
' sign-in screen, so someone else can use this copy without closing Access;
' signing in reopens frmMain, whose AppStartUp skips step 3 and reads the
' new person's roles.
'
' The first Personnel admin to open a new data file (tblPeople still empty)
' becomes its Exam Generator admin, so it can be set up from the app. That
' happens once: tblSettings AdminBootstrapped records it, and a data file
' that later has no admin is not handed to whichever Personnel admin opens
' it next (an admin is then restored in the data file itself).
'
' A front end built with no data file (build.ps1 with no -BackEnd, for UI work)
' has no tblSettings: it skips all of this and shows every page. The build
' records which kind it is (the HasDataFile property), so a copy built with a
' data file whose tblSettings link has gone stops at start-up instead of
' opening every page unsigned.
'
' The TempVars above only decide what the sidebar shows. Every write that
' needs a right checks it again in the data file at the moment of the write
' (RequireExamAdmin here, modQuestions.RequireQuestionWriter), so a right
' taken away since sign-in counts at once, and a way in that skips the page
' is still checked. They raise ERR_NOT_ALLOWED, which HandleError shows as
' the plain reason, not as a fault.
'
' Owner: start-up and sign-in (build phase 2). Depends on: modAdmin,
' modErrors, modGlobals; the kit's modKitBackend, modKitHelp, modKitLocks,
' modKitShell, modKitUI, modPersonnel, modPersonnelUI.

Public Const DATA_VERSION As Long = 4            ' the tblSettings SchemaVersion this copy reads

' Raised by the checks made at the point of a write when the signed-in person
' may not make it. Err.Description is the reason, written for the user.
Public Const ERR_NOT_ALLOWED As Long = vbObjectError + 70

Private Const PROBE_TABLE As String = "tblSettings"

' Whether the signed-in person may open a page, from the roles read at
' sign-in (TempVars): admin pages for the admin, the rest by role. It only
' decides what the sidebar shows and where Ctrl+K may go; every write checks
' the right again in the data file.
'   key   a sidebar key, or settings or qedit (keys only PageForm knows)
' Returns True for every page in a UI-only build (no data file).
' Errors: none.
Public Function PageAllowed(ByVal key As String) As Boolean
    If Not DataLinked() Then PageAllowed = True: Exit Function
    Select Case key
        Case "dash", "exams", "questions", "stats", "settings": PageAllowed = True
        Case "review", "qedit": PageAllowed = HasRole("SME")
        Case "approvals": PageAllowed = HasRole("DH") Or HasRole("XO") Or HasRole("SM")
        Case "topics", "templates", "people", "rules", "helptopics", "logs": PageAllowed = IsExamAdmin()
    End Select
End Function

' Start-up, run by frmMain.Form_Load under the splash: the data file and its
' version, sign-in, the person's roles and the help (see the top of the
' module).
' Returns True when the app may open, False when it should close (the user
' cancelled, or can't use it and was told why).
' Errors: the kit's and the engine's pass up to frmMain's handler.
' Side effects: may relink the data file (KitEnsureBackend), sign the user in
'   (TempVars), make the first admin of a new data file (LoadRoles) and
'   update the help table (KitHelpSync).
Public Function AppStartUp() As Boolean
    Dim problem As String
    If Not DataLinked() Then AppStartUp = True: Exit Function
    If Not TableExists(PROBE_TABLE) Then
        KitMsgBox "This copy of " & APP_NAME & " has lost its link to the data file, so it can't open. " & _
            "Get a fresh copy from your administrator.", vbExclamation, APP_NAME
        Exit Function
    End If
    SplashStatus "Opening the data file" & ChrW(8230)
    If Not KitEnsureBackend(PROBE_TABLE, BE_FILE) Then Exit Function
    problem = DataVersionProblem()
    If Len(problem) > 0 Then
        KitMsgBox problem, vbExclamation, APP_NAME
        Exit Function
    End If
    If SignedInPersonID() = 0 Then          ' a palette switch reopens the shell already signed in
        If Not KitPersonnelStart(PROBE_TABLE, APP_CODE) Then Exit Function
    End If
    SplashStatus "Loading your roles" & ChrW(8230)
    problem = LoadRoles()
    If Len(problem) > 0 Then
        KitMsgBox problem, vbExclamation, APP_NAME
        Exit Function
    End If
    SplashStatus "Loading help" & ChrW(8230)
    SyncHelp
    AppStartUp = True
End Function

' Start-up step 5. Help only helps: when it can't be read the error is logged
' and the app opens with the help the build put in it.
Private Sub SyncHelp()
    On Error GoTo Fail
    KitHelpSync
    Exit Sub
Fail:
    LogError "modSignIn.SyncHelp"
End Sub

' Hide the sidebar pages the signed-in person can't use (PageAllowed).
'   frm   the shell, frmMain
' Errors: the kit's (ShellShowItem).
Public Sub ShowPagesForRoles(frm As Form)
    Dim k As Variant
    For Each k In Array("review", "approvals", "topics", "templates", "people", "rules", "helptopics", "logs")
        ShellShowItem frm, CStr(k), PageAllowed(CStr(k))
    Next
End Sub

' Sign out and show the sign-in screen (My settings' Sign out). Only with a
' data file, only while the current page can be left, and only when the
' next person can't take over this Windows account's owner
' (OwnerPasswordReady).
' Errors: none (shown and logged by HandleError).
' Side effects: clears the session, the roles and the pages' hand-off
'   TempVars. Closing the sign-in screen closes the app, as at start-up.
'   Signing in reopens frmMain on the Dashboard, whose Form_Load reads the
'   new person's roles (AppStartUp); a person without a role is told why and
'   the app closes, as at start-up.
Public Function AppSignOut()
    Dim v As Variant
    On Error GoTo Err_Handler
    If Not DataLinked() Then Exit Function
    If IsFormOpen("frmMain") Then
        If Not Forms("frmMain").CanLeavePage() Then Exit Function
    End If
    If Not OwnerPasswordReady() Then Exit Function
    ClearRoles
    ' The pages' hand-offs name the last person's items: drop them.
    On Error Resume Next                    ' one that was never set
    For Each v In Array("BankQuestionID", "EditQuestionID", "EditMode", "ExamsExamID", _
            "ApprovalsExamID", "ApprovalsQuestionID", "ReviewTaskID")
        TempVars.Remove v
    Next
    On Error GoTo Err_Handler
    If Not KitPersonnelSignOut(APP_CODE) Then
        DoCmd.Quit acQuitSaveNone
        Exit Function
    End If
    ' A toast still showing would be left at the top left by the reopen (as
    ' for a palette switch, modDarkMode).
    If IsFormOpen("dlgToast") Then DoCmd.Close acForm, "dlgToast", acSaveNo
    If IsFormOpen("frmMain") Then Forms("frmMain").Reopen "dash"
    Exit Function
Err_Handler:
    HandleError "modSignIn.AppSignOut"
End Function

' The sign-in screen lets whoever runs this Windows account choose its
' owner's first app password (the kit's PersonnelCanSetFirstPassword), which
' after Sign out could be someone else using the owner's session. True when
' that can't happen: no person has this Windows account, theirs has an app
' password already, or the owner is the one signing out and chooses one now.
' False, having said why, otherwise.
' Errors: the engine's and the kit's.
Private Function OwnerPasswordReady() As Boolean
    Dim owner As Long, user As String, p1 As Variant, p2 As Variant, problem As String
    owner = PersonnelWindowsPersonId()
    If owner = 0 Then OwnerPasswordReady = True: Exit Function
    user = Nz(DLookup("username", PERSON_TABLE, "person_id=" & owner), "")
    If Not PersonnelCanSetFirstPassword(user) Then
        OwnerPasswordReady = True
        Exit Function
    End If
    If owner <> SignedInPersonID() Then
        KitMsgBox "Sign out isn't available on this Windows account yet: its owner has no app password, so the next " & _
            "person could choose one for them. Close " & APP_NAME & " instead.", vbExclamation, "Sign out"
        Exit Function
    End If
    p1 = KitInputBox("Choose an app password first (8 or more characters), so only you can sign in as you " & _
        "after you sign out.", "Sign out", , True)
    If IsNull(p1) Then Exit Function
    p2 = KitInputBox("Type the same password again.", "Sign out", , True)
    If IsNull(p2) Then Exit Function
    If StrComp(p1, p2, vbBinaryCompare) <> 0 Then
        KitMsgBox "The two passwords don't match. Click Sign out to try again.", vbExclamation, "Sign out"
        Exit Function
    End If
    problem = PersonnelSetPassword(owner, CStr(p1))
    If Len(problem) > 0 Then KitMsgBox problem, vbExclamation, "Sign out": Exit Function
    OwnerPasswordReady = True
End Function

' True when this front end belongs with a data file: it was built with one
' (BuiltWithDataFile), or it is linked to (or holds) the data file's tables.
' False only for a UI-only build. A copy built with a data file stays True
' when its tblSettings link is removed, so AppStartUp refuses it rather than
' skipping sign-in.
Public Function DataLinked() As Boolean
    DataLinked = BuiltWithDataFile() Or TableExists(PROBE_TABLE)
End Function

' The HasDataFile property BuildAll set: True when this front end was built
' linked to a data file. False for a UI-only build and for a copy built
' before the property existed (DataLinked then goes by the link alone).
Private Function BuiltWithDataFile() As Boolean
    On Error Resume Next                    ' error 3270: no such property
    BuiltWithDataFile = CurrentDb.Properties("HasDataFile")
End Function

' Why this copy can't use its data file (KitVersionProblem), or "". A data
' file of an older version is brought up to date first (UpgradeData), and
' the links are checked on every copy: the build links to the data file as
' it was then, so a copy that didn't run the upgrade still has the old
' columns until its link is refreshed.
' Errors: the engine's, when tblSettings can't be read.
Public Function DataVersionProblem() As String
    Dim v As Long, n As Long
    v = DataVersion()
    If v = 0 Then Exit Function             ' no version recorded: the per-feature checks say what is missing
    If v < DATA_VERSION Then
        SplashStatus "Updating the data file" & ChrW(8230)
        n = UpgradeData(v)
        v = DataVersion()
    End If
    If v < DATA_VERSION Then
        If n = 0 Or IsOpenElsewhere(n) Then
            ' A table another copy holds open: closing the other copies clears it.
            DataVersionProblem = "The shared data file needs an update that this copy of " & APP_NAME & " couldn't make, most " & _
                "likely because " & APP_NAME & " is open on another computer." & vbCrLf & vbCrLf & _
                "Close " & APP_NAME & " on every other computer, then open it here again."
        Else
            ' Anything else happens again at every start: only the admin can clear it.
            DataVersionProblem = "The shared data file needs an update that this copy of " & APP_NAME & " couldn't make " & _
                "(error " & n & ")." & vbCrLf & vbCrLf & "The details are in the error log inside the data file. Tell the " & _
                APP_NAME & " admin, who can read it there without the app (see the README's Release section)."
        End If
        Exit Function
    End If
    If v >= 2 Then EnsureLinkHas "tblExamTopics", "ChangeNo"
    DataVersionProblem = KitVersionProblem(v, DATA_VERSION, "", "")
End Function

' Refresh linked table tbl when its link doesn't show column col yet (a
' column the data file gained after this copy was linked). A failed refresh
' is logged; the first use of the column then reports it.
Private Sub EnsureLinkHas(ByVal tbl As String, ByVal col As String)
    Dim td As DAO.TableDef, fld As DAO.Field
    On Error GoTo Fail
    Set td = CurrentDb.TableDefs(tbl)
    If Len(td.Connect) = 0 Then Exit Sub                 ' a local table: nothing to refresh
    For Each fld In td.Fields
        If StrComp(fld.Name, col, vbTextCompare) = 0 Then Exit Sub
    Next
    td.RefreshLink
    Exit Sub
Fail:
    LogError "modSignIn.EnsureLinkHas", tbl & "." & col
End Sub

' The link's options after its type and path (";PWD=..."), so the data
' file is opened as the links open it (as modKitBackend's ConnectRest).
Private Function LinkOptions(ByVal cn As String) As String
    Dim parts() As String, i As Long
    parts = Split(cn, ";")
    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 And UCase$(Left$(parts(i), 9)) <> "DATABASE=" Then LinkOptions = LinkOptions & ";" & parts(i)
    Next
End Function

' The data file's version (tblSettings SchemaVersion), 0 when none is recorded.
Private Function DataVersion() As Long
    DataVersion = CLng(Nz(DLookup("SettingValue", PROBE_TABLE, "SettingKey='SchemaVersion'"), "0"))
End Function

' Bring a data file of version v up to DATA_VERSION, one step per version, in
' the data file itself (modBuildBackEnd builds a new one at DATA_VERSION). The
' first copy of a newer build to open it does this. Each step only adds and
' can run again. When it can't (a table is open in another copy, or a data
' file changed by hand), the error is logged and the version stays, so
' DataVersionProblem says it couldn't be updated.
' Returns 0 when it finished, else the number of the error that stopped it.
'   2: tblExamTopics.ChangeNo (0 for the topics already there).
'   3: the unique indexes and relationships new data files have (UpgradeTo3).
'   4: ixAuditOn, the audit trail by date (Admin: Logs searches it by day),
'      and ixAuditPerson, by person and item (the Dashboard's recent activity).
Private Function UpgradeData(ByVal v As Long) As Long
    Dim db As DAO.Database, td As DAO.TableDef, fld As DAO.Field, has As Boolean, n As Long, d As String
    On Error GoTo Fail
    Set db = DBEngine.OpenDatabase(KitBackendPath(PROBE_TABLE), False, False, LinkOptions(CurrentDb.TableDefs(PROBE_TABLE).Connect))
    If v < 2 Then
        Set td = db.TableDefs("tblExamTopics")
        For Each fld In td.Fields
            If fld.Name = "ChangeNo" Then has = True
        Next
        If Not has Then
            ' Another copy upgrading at the same moment may add it first (3380).
            On Error Resume Next
            db.Execute "ALTER TABLE tblExamTopics ADD COLUMN ChangeNo LONG", dbFailOnError
            n = Err.Number: d = Err.Description
            On Error GoTo Fail
            If n <> 0 And n <> 3380 Then Err.Raise n, "modSignIn.UpgradeData", d
        End If
        db.Execute "UPDATE tblExamTopics SET ChangeNo=0 WHERE ChangeNo Is Null", dbFailOnError
        CurrentDb.TableDefs("tblExamTopics").RefreshLink        ' the link shows the new column
    End If
    If v < 3 Then UpgradeTo3 db
    If v < 4 Then
        AddIndexIfMissing db, "ixAuditOn", "tblAuditTrail", "ActionOn", False
        AddIndexIfMissing db, "ixAuditPerson", "tblAuditTrail", "PersonID, ItemType, ItemID", False
    End If
    db.Execute "UPDATE tblSettings SET SettingValue='" & DATA_VERSION & "' WHERE SettingKey='SchemaVersion'", dbFailOnError
    db.Close
    Exit Function
Fail:
    UpgradeData = Err.Number
    If UpgradeData = 0 Then UpgradeData = -1
    LogError "modSignIn.UpgradeData", "from version " & v
    On Error Resume Next                    ' the error is logged; a failed close changes nothing
    If Not db Is Nothing Then db.Close
End Function

' True for an error that means another copy has the data file or one of its
' tables open (closing the other copies clears it): the kit's lock errors,
' and 3006, 3045, 3050, 3051 and 3734 (file in use, locked, opened
' exclusively; 3051 can also mean no permission to the file).
Private Function IsOpenElsewhere(ByVal n As Long) As Boolean
    Select Case n
        Case 3006, 3045, 3050, 3051, 3734: IsOpenElsewhere = True
        Case Else: IsOpenElsewhere = IsLockError(n)
    End Select
End Function

' Step 3: the unique indexes and relationships a new data file is built
' with (modBuildBackEnd, from uxLevelName and rQ_Step on). Columns that held
' 0 for "none" become Null first, since a relationship can't point at row 0.
' One the existing rows break (two levels with one name, a row pointing at
' nothing) is left out and logged, so the app still opens and its own checks
' still apply; the admin finds it in the error log.
' Errors: all but those bad-row ones pass up (UpgradeSkip), a table another
'   copy holds open among them.
Private Sub UpgradeTo3(db As DAO.Database)
    Dim v As Variant, p() As String
    For Each v In Array("tblQuestions|CurrentStepRoleID", "tblQuestions|ReplacesQuestionID", "tblExams|CurrentStepRoleID", _
            "tblExams|OverallClassLevelID", "tblExamTopics|SignedOffByPersonID", "tblTasks|CompletedByPersonID")
        p = Split(v, "|")
        db.Execute "UPDATE " & p(0) & " SET " & p(1) & "=Null WHERE " & p(1) & "=0", dbFailOnError
    Next
    AddIndexIfMissing db, "uxLevelName", "tblDifficultyLevels", "LevelName"
    AddIndexIfMissing db, "uxMarking", "tblClassLevels", "Marking"
    AddIndexIfMissing db, "uxPortionMark", "tblClassLevels", "PortionMark"
    AddIndexIfMissing db, "uxRoleCode", "tblRoles", "RoleCode"
    AddIndexIfMissing db, "uxRefTitleNumber", "tblSourceRefs", "RefTitle, RefNumber"
    AddRelationIfMissing db, "rQ_Step", "tblRoles", "RoleID", "tblQuestions", "CurrentStepRoleID"
    AddRelationIfMissing db, "rE_Step", "tblRoles", "RoleID", "tblExams", "CurrentStepRoleID"
    AddRelationIfMissing db, "rE_C", "tblClassLevels", "ClassLevelID", "tblExams", "OverallClassLevelID"
    AddRelationIfMissing db, "rET_P", "tblPeople", "PersonID", "tblExamTopics", "SignedOffByPersonID"
    AddRelationIfMissing db, "rTk_CP", "tblPeople", "PersonID", "tblTasks", "CompletedByPersonID"
    AddRelationIfMissing db, "rQ_Rep", "tblQuestions", "QuestionID", "tblQuestions", "ReplacesQuestionID"
End Sub

' Add index nm on tbl (cols), unique unless isUnique is False, unless the
' table has it already.
Private Sub AddIndexIfMissing(db As DAO.Database, ByVal nm As String, ByVal tbl As String, ByVal cols As String, _
        Optional ByVal isUnique As Boolean = True)
    Dim ix As DAO.Index, n As Long, d As String
    For Each ix In db.TableDefs(tbl).Indexes
        If ix.Name = nm Then Exit Sub
    Next
    On Error Resume Next
    db.Execute "CREATE " & IIf(isUnique, "UNIQUE ", "") & "INDEX " & nm & " ON " & tbl & " (" & cols & ")", dbFailOnError
    n = Err.Number: d = Err.Description
    On Error GoTo 0
    UpgradeSkip n, d, nm
End Sub

' Add relationship nm (ct.fk refers to pt.pk, enforced) unless the data file
' has it already.
Private Sub AddRelationIfMissing(db As DAO.Database, ByVal nm As String, ByVal pt As String, ByVal pk As String, _
        ByVal ct As String, ByVal fk As String)
    Dim r As DAO.Relation, n As Long, d As String
    For Each r In db.Relations
        If r.Name = nm Then Exit Sub
    Next
    ' Stop at the first error, so the one passed on is the cause, not a later
    ' error 91 from a relation that was never made.
    On Error Resume Next
    Set r = db.CreateRelation(nm, pt, ct, 0)
    If Err.Number = 0 Then r.Fields.Append r.CreateField(pk)
    If Err.Number = 0 Then r.Fields(pk).ForeignName = fk
    If Err.Number = 0 Then db.Relations.Append r
    n = Err.Number: d = Err.Description
    On Error GoTo 0
    UpgradeSkip n, d, nm
End Sub

' After adding index or relationship nm ended in error n (0: it didn't).
' Only the errors known to mean the existing rows break it are logged and
' nm left out: 3022 (a unique index over duplicate values) and 3379
' ("Existing data in table ... violates referential integrity rules", as
' Microsoft's DAO error reference lists it for a relationship the rows
' break; check it once in Access). 3375, 3012, 3367 and 3378 ("already
' exists") mean another copy upgrading at the same moment made nm first, so
' it counts as done. Any other error, a table another copy holds open
' (IsOpenElsewhere) or anything unexpected, is raised, so the version stays
' and the step runs again at the next start.
Private Sub UpgradeSkip(ByVal n As Long, ByVal d As String, ByVal nm As String)
    Select Case n
        Case 0, 3375, 3012, 3367, 3378
        Case 3022, 3379
            On Error Resume Next
            Err.Raise n, "modSignIn.UpgradeData", d
            LogError "modSignIn.UpgradeData", nm & " left out: rows in the data file break it (a duplicate, or a row " & _
                "pointing at nothing). The app's own checks still apply."
            Err.Clear
        Case Else
            Err.Raise n, "modSignIn.UpgradeData", d
    End Select
End Sub

' Read the signed-in person's Exam Generator access into TempVars: EGRoles
' (",SME,DH,"), EGIsAdmin, EGNotifyDash and EGNotifyEmail (all cleared
' first).
' Returns "" when they may use the app, else what to tell them (no role,
' access turned off, nobody signed in).
' Errors: the engine's, and MakeFirstAdmin's other than 3022.
' Side effects: on a new data file (tblPeople empty, no AdminBootstrapped),
'   a Personnel admin is made its first admin (MakeFirstAdmin), once.
Public Function LoadRoles() As String
    Dim pid As Long, rs As DAO.Recordset, roles As String
    pid = SignedInPersonID()
    ClearRoles
    If pid = 0 Then LoadRoles = "Nobody is signed in.": Exit Function
    ' A new data file has nobody yet: the first Personnel admin to open the app becomes its admin, once.
    If DCount("*", "tblPeople") = 0 And SettingText(SET_ADMIN_BOOTSTRAPPED) <> "1" Then
        If PersonnelIsAdminNow() Then MakeFirstAdmin pid
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT IsAdmin, IsActive, NotifyDashboard, NotifyEmail FROM tblPeople WHERE PersonID=" & pid, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        LoadRoles = NoAccessText("you don't have an Exam Generator role yet")
        Exit Function
    End If
    If Not Nz(rs!IsActive, False) Then
        rs.Close
        LoadRoles = NoAccessText("your Exam Generator access is turned off")
        Exit Function
    End If
    TempVars("EGIsAdmin") = CBool(Nz(rs!IsAdmin, False))
    TempVars("EGNotifyDash") = CBool(Nz(rs!NotifyDashboard, False))
    TempVars("EGNotifyEmail") = CBool(Nz(rs!NotifyEmail, False))
    rs.Close
    roles = PersonRoleCodes(pid)
    If Len(roles) = 0 And Not TempVars("EGIsAdmin") Then
        ClearRoles
        LoadRoles = NoAccessText("you don't have an Exam Generator role yet")
        Exit Function
    End If
    TempVars("EGRoles") = "," & roles & IIf(Len(roles) > 0, ",", "")
End Function

' True when the signed-in person held a role at sign-in (TempVars EGRoles).
' For what the pages show; a write checks the role again in the data file.
'   code   role code: SME, DH, XO or SM
' Errors: none (False when the TempVars aren't set).
Public Function HasRole(ByVal code As String) As Boolean
    On Error Resume Next                    ' TempVars may be unset
    HasRole = (InStr(Nz(TempVars("EGRoles"), ""), "," & code & ",") > 0)
End Function

' Raise ERR_NOT_ALLOWED unless the signed-in person is an active Exam
' Generator admin now: their tblPeople row is admin and active, and the
' Personnel file still lets them use the app. Reads the data file, not the
' TempVars set at sign-in, so admin rights taken away since then count at
' once. Every admin write goes through this, as the kit's changes to people
' go through PersonnelRequireAdmin: the modAdmin writers call it, and so does
' each admin page before its own writes. A front end with no data file has
' nothing to protect and passes.
' Errors: ERR_NOT_ALLOWED with the reason for the user.
' Side effects: sets TempVars EGIsAdmin to the answer, so the sidebar
'   follows on the next page change.
Public Sub RequireExamAdmin()
    Dim pid As Long, ok As Boolean
    If Not DataLinked() Then Exit Sub
    pid = SignedInPersonID()
    If pid <> 0 Then
        ok = DCount("*", "tblPeople", "PersonID=" & pid & " AND IsAdmin AND IsActive") > 0
        If ok Then ok = PersonCanUse(pid, APP_CODE)
    End If
    TempVars("EGIsAdmin") = ok
    If Not ok Then Err.Raise ERR_NOT_ALLOWED, "modSignIn.RequireExamAdmin", _
        "Only an active Exam Generator admin can make this change. Your admin rights may have been taken away since you signed in."
End Sub

' True when the signed-in person was an admin when last checked (TempVars
' EGIsAdmin, set at sign-in and by RequireExamAdmin). For what the pages
' show; a write calls RequireExamAdmin instead.
' Errors: none (False when the TempVar isn't set).
Public Function IsExamAdmin() As Boolean
    On Error Resume Next
    IsExamAdmin = Nz(TempVars("EGIsAdmin"), False)
End Function

Private Sub ClearRoles()
    Dim v As Variant
    On Error Resume Next                    ' one that was never set
    For Each v In Array("EGRoles", "EGIsAdmin", "EGNotifyDash", "EGNotifyEmail")
        TempVars.Remove v
    Next
End Sub

' Add pid as the data file's first Exam Generator admin and record that the
' first admin was made (tblSettings AdminBootstrapped = 1), in one
' transaction. When another Personnel admin did it at the same moment (error
' 3022 on either row) this one is left out and the app opens as for anyone
' else.
Private Sub MakeFirstAdmin(ByVal pid As Long)
    Dim ws As DAO.Workspace, inTrans As Boolean
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    KitExecute "INSERT INTO tblPeople (PersonID, IsAdmin, IsActive, NotifyDashboard, NotifyEmail) " & _
        "VALUES (" & pid & ", True, True, True, False)"
    If DCount("*", "tblSettings", "SettingKey=" & Q(SET_ADMIN_BOOTSTRAPPED)) = 0 Then
        KitExecute "INSERT INTO tblSettings (SettingKey, SettingValue, Description) VALUES (" & Q(SET_ADMIN_BOOTSTRAPPED) & _
            ", '1', 'Set once the first admin was made from the Personnel admins')"
    Else
        KitExecute "UPDATE tblSettings SET SettingValue='1' WHERE SettingKey=" & Q(SET_ADMIN_BOOTSTRAPPED)
    End If
    ws.CommitTrans: inTrans = False
    Exit Sub
Fail:
    If inTrans Then RollbackKeepingError ws
    If Err.Number = 3022 Then Exit Sub
    RaiseAgain "modSignIn.MakeFirstAdmin", pid
End Sub

' "Signed in as <name>, but <why>, so ..." with who to ask.
Private Function NoAccessText(ByVal why As String) As String
    Dim who As String
    who = AdminNames()
    NoAccessText = "You're signed in as " & Nz(TempVars("PersonName"), "you") & ", but " & why & _
        ", so there's nothing here you can open." & vbCrLf & vbCrLf & _
        "Ask the Exam Generator admin" & IIf(Len(who) > 0, " (" & who & ")", "") & _
        " to add you under Admin: People and routing."
End Function

' Up to three active Exam Generator admins' names from the Personnel file,
' "A, B or C", or "".
Private Function AdminNames() As String
    Dim rs As DAO.Recordset, names As Collection, i As Long, s As String
    On Error GoTo Done                      ' the names only help: never stop the message for them
    Set names = New Collection
    Set rs = CurrentDb.OpenRecordset("SELECT Nz(p.display_name, p.username) AS nm FROM tblPeople AS e INNER JOIN " & _
        PERSON_TABLE & " AS p ON e.PersonID=p.person_id WHERE e.IsAdmin AND e.IsActive ORDER BY p.last_name, p.first_name", dbOpenSnapshot)
    Do While Not rs.EOF And names.Count < 3
        names.Add CStr(rs!nm)
        rs.MoveNext
    Loop
    rs.Close
    For i = 1 To names.Count
        If i > 1 Then s = s & IIf(i = names.Count, " or ", ", ")
        s = s & names(i)
    Next
    AdminNames = s
Done:
End Function
