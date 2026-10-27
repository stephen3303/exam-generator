Option Compare Database
Option Explicit

' Exam Generator shell: the sidebar and the page host sfPage (ShellBuild).
' Its events go to the kit's modKitShell, which swaps pages from the timer
' (QueueNav), never while one of their own events is running.

Private Sub Form_Load()
    On Error GoTo Err_Handler
    DoCmd.Maximize
    SplashStatus "Opening the dashboard" & ChrW(8230)
    ShellLoad Me
    Exit Sub
Err_Handler:
    HandleError "frmMain.Form_Load"
End Sub

' A resize doesn't fire a table's OnCurrent, so recheck which tables need a
' scroll bar on every page shown (GridScrollAll).
Private Sub Form_Resize()
    On Error GoTo Err_Handler
    GridScrollAll Me
    Exit Sub
Err_Handler:
    HandleError "frmMain.Form_Resize"
End Sub

' Key of the page being shown (SetPalette reopens on it; the error log and
' KitHelp read it).
Public Property Get PageKey() As String
    PageKey = ShellPageKey()
End Property

' Go to page key once the calling event has finished. Not while the current
' page refuses to be left (CanLeavePage).
Public Sub QueueNav(ByVal key As String)
    ShellQueue Me, key
End Sub

' Reopen the window on page pg (after a palette switch).
Public Sub Reopen(ByVal pg As String)
    ShellReopen Me, pg
End Sub

' True when the current page can be swapped or closed.
Public Function CanLeavePage() As Boolean
    CanLeavePage = ShellCanLeave(Me)
End Function

' The form for a page no sidebar item opens (the shell's PageForm hook):
' My settings (a bottom link) and the question editor (opened from the
' question bank).
Public Function PageForm(ByVal key As String) As String
    On Error GoTo Err_Handler
    Select Case key
        Case "settings": PageForm = "pgMySettings"
        Case "qedit": PageForm = "pgQuestionEdit"
    End Select
    Exit Function
Err_Handler:
    HandleError "frmMain.PageForm", key
End Function

' Closing the window: not while the page refuses to be left, so Access
' doesn't drop the page's unsaved record or ask "close anyway?".
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Err_Handler
    Cancel = Not ShellCanLeave(Me)
    Exit Sub
Err_Handler:
    HandleError "frmMain.Form_Unload"
End Sub

Private Sub Form_Timer()
    On Error GoTo Err_Handler
    ShellTimer Me
    Exit Sub
Err_Handler:
    HandleError "frmMain.Form_Timer"
End Sub
