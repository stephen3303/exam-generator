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
' The first Personnel admin to open the app while it has no Exam Generator
' admin becomes its admin, so a new data file can be set up from the app.
'
' A front end built with no data file (build.ps1 with no -BackEnd, for UI work)
' has no tblSettings: it skips all of this and shows every page.
'
' The TempVars above only decide what the sidebar shows. Every write that
' needs a right checks it again in the data file at the moment of the write
' (RequireExamAdmin here, modQuestions.RequireQuestionWriter), so a right
' taken away since sign-in counts at once, and a way in that skips the page
' is still checked. They raise ERR_NOT_ALLOWED, which HandleError shows as
' the plain reason, not as a fault.

Public Const DATA_VERSION As Long = 1            ' the tblSettings SchemaVersion this copy reads

' Raised by the checks made at the point of a write when the signed-in person
' may not make it. Err.Description is the reason, written for the user.
Public Const ERR_NOT_ALLOWED As Long = vbObjectError + 70

Private Const PROBE_TABLE As String = "tblSettings"

' Who may open which page: the sidebar keys and the two keys only PageForm
' knows (settings, qedit). Admin pages for the admin; the rest by role.
Public Function PageAllowed(ByVal key As String) As Boolean
    If Not DataLinked() Then PageAllowed = True: Exit Function
    Select Case key
        Case "dash", "exams", "questions", "stats", "settings": PageAllowed = True
        Case "review", "qedit": PageAllowed = HasRole("SME")
        Case "approvals": PageAllowed = HasRole("DH") Or HasRole("XO") Or HasRole("SM")
        Case "topics", "templates", "people", "rules", "helptopics", "logs": PageAllowed = IsExamAdmin()
    End Select
End Function

' Start-up (frmMain.Form_Load): True when the app may open, False when it
' should close (the user cancelled, or can't use it and was told why).
Public Function AppStartUp() As Boolean
    Dim problem As String
    If Not DataLinked() Then AppStartUp = True: Exit Function
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

' Hide the sidebar pages the signed-in person can't use.
Public Sub ShowPagesForRoles(frm As Form)
    Dim k As Variant
    For Each k In Array("review", "approvals", "topics", "templates", "people", "rules", "helptopics", "logs")
        ShellShowItem frm, CStr(k), PageAllowed(CStr(k))
    Next
End Sub

' True when this front end is linked to (or holds) the data file's tables.
Public Function DataLinked() As Boolean
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        If td.Name = PROBE_TABLE Then DataLinked = True: Exit Function
    Next
End Function

' Why this copy can't use its data file (KitVersionProblem), or "".
Public Function DataVersionProblem() As String
    Dim v As Long
    v = CLng(Nz(DLookup("SettingValue", PROBE_TABLE, "SettingKey='SchemaVersion'"), "0"))
    If v = 0 Then Exit Function             ' no version recorded: the per-feature checks say what is missing
    DataVersionProblem = KitVersionProblem(v, DATA_VERSION, "", "")
End Function

' Read the signed-in person's Exam Generator roles into TempVars. Returns ""
' when they may use the app, else what to tell them.
Public Function LoadRoles() As String
    Dim pid As Long, rs As DAO.Recordset, roles As String
    pid = SignedInPersonID()
    ClearRoles
    If pid = 0 Then LoadRoles = "Nobody is signed in.": Exit Function
    ' A new data file has no admin yet: the first Personnel admin to open the app becomes it.
    If DCount("*", "tblPeople", "IsAdmin AND IsActive") = 0 Then
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
    Set rs = CurrentDb.OpenRecordset("SELECT r.RoleCode FROM tblPersonRoles AS pr INNER JOIN tblRoles AS r " & _
        "ON pr.RoleID=r.RoleID WHERE pr.PersonID=" & pid & " ORDER BY r.ChainOrder", dbOpenSnapshot)
    Do While Not rs.EOF
        roles = roles & "," & rs!RoleCode
        rs.MoveNext
    Loop
    rs.Close
    If Len(roles) = 0 And Not TempVars("EGIsAdmin") Then
        ClearRoles
        LoadRoles = NoAccessText("you don't have an Exam Generator role yet")
        Exit Function
    End If
    TempVars("EGRoles") = roles & ","
End Function

' True when the signed-in person holds role code (SME, DH, XO, SM).
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
' each admin page before its own writes.
' Errors: ERR_NOT_ALLOWED with the reason for the user.
' Side effects: sets TempVars EGIsAdmin to the answer, so the sidebar
'   follows on the next page change.
' A front end with no data file has nothing to protect and passes.
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

' True when the signed-in person is the Exam Generator admin.
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

' Add pid as the Exam Generator admin (a row for them, or their row made admin).
Private Sub MakeFirstAdmin(ByVal pid As Long)
    If DCount("*", "tblPeople", "PersonID=" & pid) = 0 Then
        CurrentDb.Execute "INSERT INTO tblPeople (PersonID, IsAdmin, IsActive, NotifyDashboard, NotifyEmail) " & _
            "VALUES (" & pid & ", True, True, True, False)", dbFailOnError
    Else
        CurrentDb.Execute "UPDATE tblPeople SET IsAdmin=True, IsActive=True WHERE PersonID=" & pid, dbFailOnError
    End If
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
