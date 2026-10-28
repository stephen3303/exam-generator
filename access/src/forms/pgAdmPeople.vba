Option Compare Database
Option Explicit

' Admin: People and routing (see Build_pgAdmPeople). The people with Exam
' Generator access on the left; the open person's switches on the right.
' A switch saves when it is flipped: access (tblPeople IsActive), admin
' (IsAdmin) and the four roles (tblPersonRoles). The cards across the top
' name who holds DH, XO and SM, the steps an exam's approval goes through.
'
' Guards: nobody can turn off their own access or their own admin, and the
' app always keeps one active admin. Taking SME away from someone who is
' an SME of topics takes them off those topics too, once the admin agrees.
' A change to the signed-in person's own roles shows in the sidebar at once.

Private Const MSG_TITLE As String = "People and routing"

Private m_pid As Long           ' the open person (0 when none)
Private m_isLoading As Boolean  ' ShowPerson is setting the switches

Private Sub Form_Load()
    On Error GoTo Err_Handler
    If Not DataLinked() Then
        ShowEmpty Me, "emptyList", True, "No data file linked", "This copy was built without the data file, so there are no people to show."
        SetVisible Me.sfList, False
        SetEnabled Me.btnAdd, False
        ShowPerson 0
        Exit Sub
    End If
    Me.sfList.Form.RecordSource = ListSql()
    ShowRoute
    RefreshList SignedInPersonID()
    Exit Sub
Err_Handler:
    HandleError "pgAdmPeople.Form_Load"
End Sub

Private Sub txtPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    KitPickKey Me.txtPerson, KeyCode
End Sub

' Personnel: the kit's Personnel page (Personnel admins only), where people
' are added and their names and ranks changed. The list shows any change.
Private Sub btnPersonnel_Click()
    On Error GoTo Err_Handler
    KitPersonnelManage APP_CODE
    If DataLinked() Then
        ShowRoute
        RefreshList m_pid
    End If
    Exit Sub
Err_Handler:
    HandleError "pgAdmPeople.btnPersonnel_Click"
End Sub

' Add person: give the picked person access and a row here, then open them
' so the admin can give them a role.
Private Sub btnAdd_Click()
    Dim pid As Long
    On Error GoTo Err_Handler
    pid = Nz(Me.txtPersonKey, 0)
    If pid = 0 Then
        KitMsgBox "Type part of a name in the box, then pick the person from the list under it.", vbInformation, MSG_TITLE
        Me.txtPerson.SetFocus
        Exit Sub
    End If
    If DCount("*", "tblPeople", "PersonID=" & pid) > 0 Then
        KitToast "They're already listed.", "info"
    Else
        EnsurePerson pid
        KitToast "Added. Give them a role on the right.", "success"
    End If
    Me.txtPersonKey = Null
    KitPickShow Me.txtPerson
    RefreshList pid
    Exit Sub
Err_Handler:
    HandleError "pgAdmPeople.btnAdd_Click", pid
End Sub

' Row click in the list (subGrid): open that person.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    ShowPerson Nz(Me.sfList.Form.txtSel, 0)
    Exit Sub
Err_Handler:
    HandleError "pgAdmPeople.GridClick"
End Sub

' ---------------------------------------------------------------------------
' The switches (Toggle's onChange): each saves at once
' ---------------------------------------------------------------------------

Public Sub ActiveChanged()
    Dim v As Boolean
    On Error GoTo Err_Handler
    If m_isLoading Or m_pid = 0 Then Exit Sub
    RequireExamAdmin                        ' the right is checked now, not as of sign-in
    v = IsSwitchOn("tgActive")
    If Not v Then
        If m_pid = SignedInPersonID() Then
            Refuse "tgActive", "You can't turn off your own access. Another admin can."
            Exit Sub
        End If
    End If
    If Not SavePersonFlag("IsActive", v) Then
        Refuse "tgActive", "They're the only admin who can sign in, so their access stays on. Make someone else an admin first."
        Exit Sub
    End If
    Changed
    Exit Sub
Err_Handler:
    HandleError "pgAdmPeople.ActiveChanged", m_pid
    ShowSaved
End Sub

Public Sub AdminChanged()
    Dim v As Boolean
    On Error GoTo Err_Handler
    If m_isLoading Or m_pid = 0 Then Exit Sub
    RequireExamAdmin                        ' the right is checked now, not as of sign-in
    v = IsSwitchOn("tgAdmin")
    If Not v Then
        If m_pid = SignedInPersonID() Then
            Refuse "tgAdmin", "You can't take away your own admin. Another admin can."
            Exit Sub
        End If
    End If
    If Not SavePersonFlag("IsAdmin", v) Then
        Refuse "tgAdmin", "They're the only admin who can sign in. Make someone else an admin first."
        Exit Sub
    End If
    Changed
    Exit Sub
Err_Handler:
    HandleError "pgAdmPeople.AdminChanged", m_pid
    ShowSaved
End Sub

' Save m_pid's tblPeople flag fld (IsActive or IsAdmin) as v. Turning one off
' is checked inside the same transaction, after the write: when it leaves
' the app with no admin who can sign in (AdminCount) it is rolled back and
' False returned, so two admins turning each other off at once can't both
' succeed. Turning access on also gives them access in the Personnel file
' (EnsurePerson), in the same transaction; turning it off here leaves that
' as it is.
Private Function SavePersonFlag(ByVal fld As String, ByVal v As Boolean) As Boolean
    Dim ws As DAO.Workspace, inTrans As Boolean, before As Long
    On Error GoTo Fail
    before = AdminCount()
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    KitExecute "UPDATE tblPeople SET " & fld & "=" & IIf(v, "True", "False") & " WHERE PersonID=" & m_pid
    If Not v And before > 0 Then
        If AdminCount() = 0 Then
            ws.Rollback: inTrans = False
            Exit Function
        End If
    End If
    If v And fld = "IsActive" Then EnsurePerson m_pid
    ws.CommitTrans: inTrans = False
    SavePersonFlag = True
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "pgAdmPeople.SavePersonFlag", fld & " " & m_pid
End Function

Public Sub SMEChanged()
    RoleChanged "SME", "tgSME"
End Sub

Public Sub DHChanged()
    RoleChanged "DH", "tgDH"
End Sub

Public Sub XOChanged()
    RoleChanged "XO", "tgXO"
End Sub

Public Sub SMChanged()
    RoleChanged "SM", "tgSM"
End Sub

' A role switch: give the role or take it away. Taking SME from someone who
' is an SME of topics takes them off those topics too, once the admin says so.
Private Sub RoleChanged(ByVal code As String, ByVal nm As String)
    Dim v As Boolean, topics As String, ws As DAO.Workspace, inTrans As Boolean
    On Error GoTo Err_Handler
    If m_isLoading Or m_pid = 0 Then Exit Sub
    RequireExamAdmin                        ' the right is checked now, not as of sign-in
    v = IsSwitchOn(nm)
    If code = "SME" And Not v Then
        topics = PersonTopicsText(m_pid)
        If Len(topics) > 0 Then
            If KitMsgBox("They're an SME of " & topics & ". Taking SME away takes them off " & _
                    IIf(InStr(topics, ",") > 0, "those topics", "that topic") & " too." & vbCrLf & vbCrLf & _
                    "Take SME away?", vbYesNo + vbQuestion + vbDefaultButton2, MSG_TITLE) <> vbYes Then
                SetSwitch nm, True
                Exit Sub
            End If
        End If
    End If
    ' Off their topics and the role itself go together, or neither.
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    If Len(topics) > 0 Then KitExecute "DELETE FROM tblTopicSMEs WHERE PersonID=" & m_pid
    SetPersonRole m_pid, code, v
    ws.CommitTrans: inTrans = False
    Changed
    Exit Sub
Err_Handler:
    If inTrans Then RollbackKeepingError ws
    HandleError "pgAdmPeople.RoleChanged", code & " " & m_pid
    ShowSaved
End Sub

' ---------------------------------------------------------------------------

' A switch saved: the list's row, the route cards, and, when it was the
' signed-in person, their own roles and sidebar.
Private Sub Changed()
    Dim problem As String
    RefreshList m_pid
    ShowRoute
    If m_pid = SignedInPersonID() Then
        problem = LoadRoles()
        If Len(problem) = 0 Then ShowPagesForRoles Forms("frmMain")
    End If
End Sub

' Put a switch back and say why.
' After a switch's save failed or was refused: show the person as the data
' file has them, so no switch is left showing a change that wasn't saved.
Private Sub ShowSaved()
    On Error Resume Next                    ' already in a handler: the error has been shown
    If m_pid <> 0 Then ShowPerson m_pid
End Sub

Private Sub Refuse(ByVal nm As String, ByVal why As String)
    SetSwitch nm, True
    KitMsgBox why, vbInformation, MSG_TITLE
End Sub

Private Function IsSwitchOn(ByVal nm As String) As Boolean
    IsSwitchOn = CBool(Nz(Me.Controls(nm & "_val").Value, False))
End Function

' Set a switch in code (no onChange).
Private Sub SetSwitch(ByVal nm As String, ByVal v As Boolean)
    Me.Controls(nm & "_val").Value = v
    Me.Controls(nm & "_bg").Requery
End Sub

' The route cards: who holds DH, XO and SM.
Private Sub ShowRoute()
    Dim code As Variant, s As String
    For Each code In Array("DH", "XO", "SM")
        s = RoleHolders(CStr(code))
        If Len(s) = 0 Then s = "Nobody yet. Exams can't pass this step until someone holds it."
        Me.Controls("lblRoute" & code).Caption = s
    Next
End Sub

' Reload the list and open pid (0, or not listed: the first person).
Private Sub RefreshList(ByVal pid As Long)
    Dim g As Form, n As Long
    Set g = Me.sfList.Form
    g.Requery
    n = g.RecordsetClone.RecordCount
    ShowEmpty Me, "emptyList", n = 0
    SetVisible Me.sfList, n > 0
    If GridSelect(Me.sfList, pid) Then
        ShowPerson pid
    ElseIf IsNull(SelectFirstRow(Me.sfList)) Then
        ShowPerson 0
    End If
End Sub

' Show person pid on the right (0: nobody, and the hint instead).
Private Sub ShowPerson(ByVal pid As Long)
    Dim rs As DAO.Recordset, nm As Variant, isOpen As Boolean, s As String
    On Error GoTo Fail
    m_pid = pid
    isOpen = (pid <> 0)
    If isOpen Then
        Set rs = CurrentDb.OpenRecordset("SELECT p.first_name, p.last_name, p.rank, p.username, e.IsActive, e.IsAdmin " & _
            "FROM tblPeople AS e INNER JOIN " & PERSON_TABLE & " AS p ON e.PersonID=p.person_id WHERE e.PersonID=" & pid, dbOpenSnapshot)
        If rs.EOF Then
            isOpen = False: m_pid = 0
        Else
            m_isLoading = True
            Me.lblName.Caption = Trim$(Nz(rs!rank, "") & " " & Nz(rs!first_name, "") & " " & Nz(rs!last_name, ""))
            s = "Signs in as " & Nz(rs!username, "?")
            If pid = SignedInPersonID() Then s = s & " (you)"
            Me.lblWho.Caption = s
            SetSwitch "tgActive", Nz(rs!IsActive, False)
            SetSwitch "tgAdmin", Nz(rs!IsAdmin, False)
            SetSwitch "tgSME", PersonHasRole(pid, "SME")
            SetSwitch "tgDH", PersonHasRole(pid, "DH")
            SetSwitch "tgXO", PersonHasRole(pid, "XO")
            SetSwitch "tgSM", PersonHasRole(pid, "SM")
            s = PersonTopicsText(pid)
            Me.lblTopics.Caption = IIf(Len(s) > 0, s, "No topics yet. Add SMEs to a topic on the Topics page.")
            m_isLoading = False
        End If
        rs.Close: Set rs = Nothing
    End If
    For Each nm In Array("lblName", "lblWho", "tgActive", "tgAdmin", "tgSME", "tgDH", "tgXO", "tgSM", _
            "lbltgActive", "lbltgAdmin", "lbltgSME", "lbltgDH", "lbltgXO", "lbltgSM", "lblSMEFor", "lblTopics")
        SetVisible Me.Controls(nm), isOpen
    Next
    SetVisible Me.lblDetailEmpty, Not isOpen
    Exit Sub
Fail:
    ' A failed look-up must not leave the switches ignoring every flip.
    m_isLoading = False
    If Not rs Is Nothing Then rs.Close
    RaiseAgain "pgAdmPeople.ShowPerson", pid
End Sub

' Everyone with a row in tblPeople, by name, with their roles.
Private Function ListSql() As String
    ListSql = "SELECT e.PersonID AS id, " & PersonNameSql() & " AS person, PersonRolesText(e.PersonID) AS roles, " & _
        "IIf(e.IsActive,'Active','Off') AS state, p.last_name, p.first_name FROM tblPeople AS e INNER JOIN " & _
        PERSON_TABLE & " AS p ON e.PersonID=p.person_id ORDER BY p.last_name, p.first_name"
End Function
