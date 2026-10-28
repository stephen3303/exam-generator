Option Compare Database
Option Explicit

' Dashboard (see Build_pgDash). Four tiles: the signed-in person's open
' reviews (modExams.ReviewCount), the questions and exams they can approve
' now (modNotify.DashApprovalCount), their questions sent back
' (ReturnedCount), and the topics short of approved questions for the
' templates (modAdmin.TopicsShortText, the same rule as Admin: Topics), with
' those topics in the alert under the tiles. Then everything waiting on
' them, the longest waiting first (modNotify.WaitingSql), and the last 20
' audit rows the person may see (RecentActivitySql: an admin sees the whole
' app's, anyone else what they had a hand in).
'
' With the person's Dashboard switch off (My settings) the three queue tiles
' read "Off" and the list says so; the short topics and the recent activity
' still show. Open goes to the picked item: a review on My reviews, a
' question or exam on Approvals (TempVars ReviewTaskID, ApprovalsQuestionID,
' ApprovalsExamID tell the page which), a question sent back in the editor.
' The list reads roles from the data file, but pages open by the roles read
' at sign-in, so Open first checks the page is open to them (PageAllowed).
'
' Depends on: modAdmin, modExams, modNotify, modQuestions, modSignIn, modUI;
' the kit's modKitUI.

Public Property Get HelpTopic() As String
    HelpTopic = "dash"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    If Not DataLinked() Then
        ShowTiles False
        ShowAlert Me, "alShort", "This copy was built without the data file, so there are no figures to show.", "warning"
        ShowEmpty Me, "emptyWaiting", True, "No data file linked", "This copy was built without the data file, so nothing waits on you."
        SetVisible Me.emptyWaiting_b, False
        SetVisible Me.sfWaiting, False
        SetEnabled Me.btnOpen, False
        ShowEmpty Me, "emptyActivity", True
        SetVisible Me.sfActivity, False
        Exit Sub
    End If
    ShowShortTopics
    ShowTiles DashboardOn()
    RefreshWaiting
    RefreshActivity
    Exit Sub
Err_Handler:
    HandleError "pgDash.Form_Load"
End Sub

' ---------------------------------------------------------------------------
' Buttons and rows
' ---------------------------------------------------------------------------

Private Sub btnOpen_Click()
    On Error GoTo Err_Handler
    OpenPicked
    Exit Sub
Err_Handler:
    HandleError "pgDash.btnOpen_Click"
End Sub

' The empty state's button, shown while the Dashboard switch is off.
Private Sub emptyWaiting_b_Click()
    On Error GoTo Err_Handler
    NavTo "settings"
    Exit Sub
Err_Handler:
    HandleError "pgDash.emptyWaiting_b_Click"
End Sub

' A row clicked; a double-click on a waiting item opens it.
Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    If grid = "sfWaiting" And dbl Then OpenPicked
    Exit Sub
Err_Handler:
    HandleError "pgDash.GridClick", grid
End Sub

' Go to the picked waiting item, by its id's letter (WaitingSql). A role
' given or taken away since sign-in can list an item whose page isn't open
' to the person yet: say so, rather than land back on the Dashboard with
' the page's TempVar left set.
Private Sub OpenPicked()
    Dim id As String, n As Long, key As String, tv As String
    id = Nz(Me.sfWaiting.Form.txtSel, "") & ""
    If Len(id) < 2 Then Exit Sub
    n = CLng(Mid$(id, 2))
    Select Case Left$(id, 1)
        Case "t": key = "review": tv = "ReviewTaskID"
        Case "q": key = "approvals": tv = "ApprovalsQuestionID"
        Case "x": key = "approvals": tv = "ApprovalsExamID"
        Case "r": key = "qedit"
        Case Else: Exit Sub
    End Select
    If Not PageAllowed(key) Then
        If Len(tv) > 0 Then TempVars(tv) = Null
        KitMsgBox "Your roles changed since you signed in. Sign out and sign in again.", vbInformation, "Dashboard"
        Exit Sub
    End If
    If Len(tv) > 0 Then
        TempVars(tv) = n
        NavTo key
    Else
        OpenQuestion n
    End If
End Sub

' ---------------------------------------------------------------------------
' The tiles, the alert and the tables
' ---------------------------------------------------------------------------

' The tiles: the queue counts when dashOn, else "Off" (or "-" with no data
' file); the short topics' count with data.
Private Sub ShowTiles(ByVal dashOn As Boolean)
    Dim linked As Boolean
    linked = DataLinked()
    If dashOn And linked Then
        SetTile "tileReviews", ReviewCount(False), "exam topics for you to sign off"
        SetTile "tileApprovals", DashApprovalCount(), "questions and exams at your step"
        SetTile "tileReturned", ReturnedCount(), "your questions to fix and resubmit"
    Else
        SetTile "tileReviews", IIf(linked, "Off", "-"), IIf(linked, "turned off in My settings", "")
        SetTile "tileApprovals", IIf(linked, "Off", "-"), IIf(linked, "turned off in My settings", "")
        SetTile "tileReturned", IIf(linked, "Off", "-"), IIf(linked, "turned off in My settings", "")
    End If
    If Not linked Then SetTile "tileShort", "-", ""
End Sub

Private Sub SetTile(ByVal nm As String, ByVal v As Variant, ByVal cap As String)
    Me.Controls(nm).Value = v
    Me.Controls(nm & "_cap").Caption = Replace(cap, "&", "&&")
End Sub

' The topics short of approved questions: their count on the tile and their
' lines in the alert, or that every topic has enough.
Private Sub ShowShortTopics()
    Dim n As Long, s As String
    s = TopicsShortText(OverlapPct(), OverlapWindow(), n, 2, 32)    ' two lines, long names cut short, fit the titled alert's one line
    If n = 0 Then
        SetTile "tileShort", 0, "every active topic has enough"
        ShowAlert Me, "alShort", "Every active topic has enough approved questions for the templates that use it.", "success"
    Else
        SetTile "tileShort", n, IIf(n = 1, "topic needs", "topics need") & " more approved questions"
        ShowAlert Me, "alShort", Replace(Replace(s, "&", "&&"), vbCrLf, "; "), "warning", "Topics short of approved questions"
    End If
End Sub

' Everything waiting on the person, or why the list is empty.
Private Sub RefreshWaiting()
    Dim n As Long
    If Not DashboardOn() Then
        ShowEmpty Me, "emptyWaiting", True, "Your Dashboard list is off", _
            "Turn it on under My settings to see the reviews, approvals and questions sent back that wait on you."
        Me.emptyWaiting_b.Caption = "My settings"
        SetVisible Me.emptyWaiting_b, True
        SetVisible Me.sfWaiting, False
        SetEnabled Me.btnOpen, False
        Exit Sub
    End If
    Me.sfWaiting.Form.RecordSource = WaitingSql()
    n = Me.sfWaiting.Form.RecordsetClone.RecordCount
    ShowEmpty Me, "emptyWaiting", n = 0, "Nothing waits on you", "Reviews, approvals and questions sent back to you show here."
    SetVisible Me.emptyWaiting_b, False
    SetVisible Me.sfWaiting, n > 0
    SetEnabled Me.btnOpen, n > 0
    If n > 0 Then SelectFirstRow Me.sfWaiting
End Sub

Private Sub RefreshActivity()
    Dim n As Long
    Me.sfActivity.Form.RecordSource = RecentActivitySql(20)
    n = Me.sfActivity.Form.RecordsetClone.RecordCount
    If IsExamAdmin() Then
        ShowEmpty Me, "emptyActivity", n = 0
    Else                                    ' theirs covers only their own work (RecentActivitySql)
        ShowEmpty Me, "emptyActivity", n = 0, "Nothing about your work yet", _
            "What you and others do to your questions, reviews and exams shows here."
    End If
    SetVisible Me.sfActivity, n > 0
End Sub
