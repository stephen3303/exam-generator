Option Compare Database
Option Explicit

' My settings (see Build_pgMySettings), a sidebar link for everyone. The
' signed-in person's two notice switches, each saved at once to their
' tblPeople row and to this session's TempVars (EGNotifyDash,
' EGNotifyEmail), so the next notice goes by it:
'   Dashboard list  the Dashboard's queue tiles and list (pgDash)
'   Emails          an Outlook email when something waits on them (modNotify)
' Under them the address emails go to (the Personnel file's) and Send me a
' test email. Then Colors, the palette picker. With no data file linked,
' nobody signed in, or no row for them in tblPeople, the switches are off
' and locked, and the line under them says which.
'
' Depends on: modNotify, modSignIn; the kit's modKitLocks, modKitUI,
' modKitUtil, modPersonnel.

Private Const MSG_TITLE As String = "My settings"

Private m_pid As Long           ' the signed-in person (0: none, or no data file)

Public Property Get HelpTopic() As String
    HelpTopic = "settings"
End Property

Private Sub Form_Load()
    Dim rs As DAO.Recordset, why As String
    On Error GoTo Err_Handler
    m_pid = 0
    If Not DataLinked() Then
        why = "This copy was built without the data file, so there are no notices to set."
    Else
        m_pid = SignedInPersonID()
        If m_pid = 0 Then
            why = "Nobody is signed in, so there are no notices to set."
        Else
            Set rs = CurrentDb.OpenRecordset("SELECT NotifyDashboard, NotifyEmail FROM tblPeople WHERE PersonID=" & m_pid, dbOpenSnapshot)
            If rs.EOF Then
                m_pid = 0
                why = "You aren't on the app's list of people yet, so there are no notices to set. Ask the admin to add you."
            Else
                SetSwitch "tgDash", Nz(rs!NotifyDashboard, False): SetSwitch "tgEmail", Nz(rs!NotifyEmail, False)
            End If
            rs.Close
        End If
    End If
    If m_pid = 0 Then
        SetSwitch "tgDash", False: SetSwitch "tgEmail", False
        SetEnabled Me.tgDash, False: SetEnabled Me.tgEmail, False
        Me.tgDash_val.Locked = True: Me.tgEmail_val.Locked = True    ' the links beside them toggle the value too
        SetEnabled Me.btnTestEmail, False
        Me.lblEmail.Caption = why
        Exit Sub
    End If
    ShowAddress
    Exit Sub
Err_Handler:
    HandleError "pgMySettings.Form_Load"
End Sub

' ---------------------------------------------------------------------------
' The switches (Toggle's onChange): each saves at once
' ---------------------------------------------------------------------------

Public Sub DashChanged()
    On Error GoTo Err_Handler
    SaveSwitch "tgDash", "NotifyDashboard", "EGNotifyDash"
    Exit Sub
Err_Handler:
    HandleError "pgMySettings.DashChanged", m_pid
    ReloadSwitch "tgDash", "NotifyDashboard"
End Sub

Public Sub EmailChanged()
    On Error GoTo Err_Handler
    SaveSwitch "tgEmail", "NotifyEmail", "EGNotifyEmail"
    If IsSwitchOn("tgEmail") And Len(PersonEmail(m_pid)) = 0 Then
        KitMsgBox "Emails are on, but the Personnel file has no email address for you, so none can reach you yet. " & _
            "Ask the admin to add it.", vbInformation, MSG_TITLE
    End If
    Exit Sub
Err_Handler:
    HandleError "pgMySettings.EmailChanged", m_pid
    ReloadSwitch "tgEmail", "NotifyEmail"
End Sub

' Save switch nm to tblPeople column col and TempVars tv.
Private Sub SaveSwitch(ByVal nm As String, ByVal col As String, ByVal tv As String)
    Dim v As Boolean
    If m_pid = 0 Then Exit Sub
    v = IsSwitchOn(nm)
    KitExecute "UPDATE tblPeople SET " & col & "=" & IIf(v, "True", "False") & " WHERE PersonID=" & m_pid
    TempVars(tv) = v
    KitToast IIf(nm = "tgDash", "Dashboard list ", "Emails ") & IIf(v, "on", "off"), "success"
End Sub

' After a failed save: show the switch as the data file has it.
Private Sub ReloadSwitch(ByVal nm As String, ByVal col As String)
    On Error Resume Next                    ' the save already failed and was reported
    SetSwitch nm, Nz(DLookup(col, "tblPeople", "PersonID=" & m_pid), False)
End Sub

Private Function IsSwitchOn(ByVal nm As String) As Boolean
    IsSwitchOn = CBool(Nz(Me.Controls(nm & "_val").Value, False))
End Function

' Set a switch in code (no onChange).
Private Sub SetSwitch(ByVal nm As String, ByVal v As Boolean)
    Me.Controls(nm & "_val").Value = v
    Me.Controls(nm & "_bg").Requery
End Sub

' ---------------------------------------------------------------------------
' The address and the test email
' ---------------------------------------------------------------------------

Private Sub ShowAddress()
    Dim addr As String
    addr = PersonEmail(m_pid)
    If Len(addr) > 0 Then
        Me.lblEmail.Caption = "Emails go to " & Replace(addr, "&", "&&") & ", the address in the Personnel file."
        Me.lblEmail.ForeColor = Th(CLR_MUTED_FG)
    Else
        Me.lblEmail.Caption = "The Personnel file has no email address for you. Ask the admin to add it."
        Me.lblEmail.ForeColor = Th(CLR_WARN_TEXT)
    End If
End Sub

Private Sub btnTestEmail_Click()
    Dim why As String
    On Error GoTo Err_Handler
    StartNotices                            ' its note says only whether this test may wait in the Outbox
    DoCmd.Hourglass True
    why = SendTestEmail()
    DoCmd.Hourglass False
    If Len(why) = 0 Then
        KitToast WithNoticeNote("Test email sent to " & PersonEmail(m_pid) & "."), "success"
    Else
        KitMsgBox why, vbInformation, MSG_TITLE
    End If
    Exit Sub
Err_Handler:
    DoCmd.Hourglass False
    HandleError "pgMySettings.btnTestEmail_Click", m_pid
End Sub
