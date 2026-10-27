Option Compare Database
Option Explicit

' Exam Generator (Access edition) - BuildAll, the entry point build.ps1 calls,
' and the Build_<form> procedures for the shell and its pages.
'
' The front end is a build output: every form is drawn here with the kit's
' builders (modUIKit) and rebuilt by build.ps1, never laid out by hand. The
' shared data file is made separately (the back-end script); BuildAll links
' this front end to it when build.ps1 is given its path, so UI work never
' touches data.
'
' Steps, in order: link, user interface (subforms, then pages, then dialogs,
' then the shell last), help, startup. A failed build returns
' "ERROR during <step>: <number> <description> [<last form traced>]".
'
' Phase 1 builds every page as a placeholder (its header and an empty state
' naming the build phase that fills it), so the shell can be clicked
' through. Each phase replaces its pages' placeholders with Build_<page>.

Public Const APP_VERSION As String = "0.1.0"

Public Const PAGE_W As Long = 15840             ' the page area, 1056 px
Public Const PAGE_H As Long = 9600              ' 640 px: the admin's sidebar (12 pages and 2 links) fits
Public Const PAGE_PAD As Long = 480                   ' page padding, 32 px

Private m_step As String

' Content width inside the page padding.
Public Function ContentW() As Long
    ContentW = PAGE_W - 2 * PAGE_PAD
End Function

' Build entry point (build.ps1). bePath is the shared data file to link to,
' "" for none (a front end with no data, for UI work). Returns "OK" or the
' step that failed.
Public Function BuildAll(ByVal srcRoot As String, ByVal outDir As String, ByVal uiOnly As Boolean, _
        ByVal bePath As String) As String
    On Error GoTo Fail
    g_Silent = True
    UseBuildPalette
    g_SrcRoot = srcRoot
    If Not uiOnly And Len(bePath) > 0 Then
        m_step = "link": LinkBackEnd bePath
    End If
    m_step = "user interface": BuildUserInterface
    m_step = "help": LoadHelpTopics srcRoot & "\help\help.txt"
    m_step = "startup": SetStartupOptions
    g_Silent = False
    BuildAll = "OK"
    Exit Function
Fail:
    BuildAll = "ERROR during " & m_step & ": " & Err.Number & " " & Err.Description & " [" & BuildTrace() & "]"
    g_Silent = False
End Function

' ---------------------------------------------------------------------------
' Link
' ---------------------------------------------------------------------------

' Link every table in the shared data file (all but Access's system tables),
' replacing any links of the same names.
Private Sub LinkBackEnd(ByVal bePath As String)
    Dim be As DAO.Database, td As DAO.TableDef, names As String
    KitCheckBackendPath bePath
    Set be = DBEngine.OpenDatabase(bePath, False, True)
    For Each td In be.TableDefs
        If Left$(td.Name, 4) <> "MSys" And Left$(td.Name, 4) <> "USys" And Left$(td.Name, 1) <> "~" Then
            names = names & IIf(Len(names) > 0, vbTab, "") & td.Name
        End If
    Next
    be.Close
    If Len(names) = 0 Then Err.Raise 5, , "The data file has no tables: " & bePath
    KitLinkTables bePath, Split(names, vbTab)
End Sub

' ---------------------------------------------------------------------------
' User interface
' ---------------------------------------------------------------------------

' Subforms before the pages that host them, the kit's dialogs, the splash,
' and the shell last.
Private Sub BuildUserInterface()
    BuildPages
    Build_dlgAppearance            ' palette picker (modBuildAppearance)
    BuildMessageDialog             ' the kit's dlgMessage, for KitMsgBox
    BuildInputDialog               ' dlgInput, for KitInputBox
    BuildConfirmDialog             ' dlgConfirm, for KitConfirmTyped
    BuildDatePicker                ' dlgDatePick, for DateBox
    BuildToast                     ' dlgToast, for KitToast
    BuildBusy                      ' dlgBusy, for KitBusy
    BuildColumnsDialog             ' dlgColumns, for ChooseGridColumns
    BuildFilterDialog              ' dlgFilter, for the header funnels (GridFilter)
    BuildMenuDialog                ' dlgMenu and dlgPopover, for KitMenu and KitPopover
    BuildPopoverDialog
    BuildPickDialog                ' dlgPick, for PickBox and PersonPickBox
    BuildHelpDialog                ' dlgHelp, for KitHelp (Help in the sidebar and F1)
    BuildPaletteDialog             ' dlgPalette, for Ctrl+K (KitPalette)
    BuildPersonnelDialogs          ' dlgPersonnelSetup, dlgSignIn, dlgPersonnel, dlgPerson (KitPersonnelStart)
    BuildSplash "frmMain", "Question bank and exams", ChrW(&HE82D)
    Build_frmMain
End Sub

' The pages, in sidebar order. Each phase swaps its placeholders for the
' real Build_<page>.
Private Sub BuildPages()
    Build_Stub "pgDash", "Dashboard", "Your open reviews, approvals waiting and topics short of questions.", 9
    Build_Stub "pgReview", "My reviews", "Each exam topic you review: check the drawn questions, swap any out and sign off.", 7
    Build_Stub "pgApprovals", "Approvals", "Questions and exams waiting for your approval.", 6
    Build_Stub "pgExams", "Exams", "Every exam and where it stands in review and approval.", 8
    Build_Stub "pgQuestions", "Question bank", "Browse and find questions, and the references they cite.", 5
    Build_Stub "pgQuestionEdit", "Question editor", "Enter and check a question, its parts, markings and references.", 5
    Build_Stub "pgStats", "Statistics", "Questions given per topic and per reference over a period or series.", 9
    Build_Stub "pgAdmTopics", "Admin: Topics", "Topics, their default question count and their SMEs.", 4
    Build_Stub "pgAdmTemplates", "Admin: Templates", "The exam templates: which topics each covers and how many questions each gets.", 4
    Build_Stub "pgAdmPeople", "Admin: People and routing", "Who holds each Exam Generator role.", 4
    Build_Stub "pgAdmRules", "Admin: Rules and lists", "The overlap settings and the lookup lists.", 4
    Build_Stub "pgAdmHelp", "Admin: Help topics", "Edit the help for any page or dialog without a rebuild.", 3
    Build_Stub "pgAdmLogs", "Admin: Logs", "The audit trail and the error log.", 10
    Build_pgMySettings
End Sub

' My settings: for now only the palette picker (Appearance); the
' notification switches come in phase 9.
Private Sub Build_pgMySettings()
    Dim c As Control
    NewForm "", 0, PAGE_W, PAGE_H
    PageHeader "My settings", "Your colors now; your notification switches come in build phase 9.", PAGE_PAD
    Card "Colors", "Colors", PAGE_PAD, px(84), px(480), px(120)
    Lbl "lblColors", "Choose the palette the whole app uses on this computer: light, dark and six more.", _
        PAGE_PAD + px(16), px(124), px(448), px(36), 9, False, CLR_MUTED_FG
    Set c = Btn("btnAppearance", "Appearance", PAGE_PAD + px(16), px(164), px(120), px(28), "outline")
    c.OnClick = "=AppAppearance()"
    SaveForm "pgMySettings"
End Sub

' A placeholder page: the page's header and an empty state naming the build
' phase that fills it in.
Private Sub Build_Stub(ByVal formName As String, ByVal title As String, ByVal tip As String, ByVal phase As Long)
    NewForm "", 0, PAGE_W, PAGE_H
    PageHeader title, tip, PAGE_PAD
    EmptyState "emptyStub", PAGE_PAD, px(140), ContentW(), "E946", "This page comes in build phase " & phase, _
        "It is a placeholder so the sidebar can be clicked through. " & tip, shown:=True
    SaveForm formName
End Sub

' The shell: the sidebar of pages and the page host sfPage. At sign-in the
' pages a person's roles don't allow are hidden (ShowPagesForRoles).
' My settings is a bottom link; it and the question editor, which no sidebar
' item opens, come from frmMain's PageForm.
Private Sub Build_frmMain()
    ShellBuild APP_NAME, "Question bank and exams", _
        Array("dash|Dashboard|pgDash|E80F", "review|My reviews|pgReview|E73E", _
            "approvals|Approvals|pgApprovals|E8FB", "exams|Exams|pgExams|E8A5", _
            "questions|Question bank|pgQuestions|E82D", "stats|Statistics|pgStats|E9D2", _
            "topics|Topics|pgAdmTopics|E8EC", "templates|Templates|pgAdmTemplates|E8F1", _
            "people|People and routing|pgAdmPeople|E716", "rules|Rules and lists|pgAdmRules|E9E9", _
            "helptopics|Help topics|pgAdmHelp|E70F", "logs|Logs|pgAdmLogs|E81C"), _
        Array("help|Help|E897|=KitHelp()", "settings|My settings|E713|=NavTo(""settings"")"), _
        PAGE_W, PAGE_H
    SaveForm "frmMain"
End Sub

' ---------------------------------------------------------------------------
' Startup
' ---------------------------------------------------------------------------

Private Sub SetProp(ByVal name As String, ByVal kind As Integer, ByVal value As Variant)
    Dim db As DAO.Database
    Set db = CurrentDb
    On Error Resume Next
    db.Properties(name) = value
    If Err.Number = 3270 Then
        Err.Clear
        db.Properties.Append db.CreateProperty(name, kind, value)
    End If
End Sub

' Open on the splash screen, which opens frmMain, without Access's own
' splash, the navigation pane or document tabs, and record when this front
' end was built (error-log entries and the splash show it). BuildTrustCheck
' opens the splash at startup (in place of StartUpForm) and shows the kit's
' trust page when Access would ask to enable content.
Private Sub SetStartupOptions()
    SetProp "AppTitle", dbText, APP_NAME
    SetProp "StartUpShowDBWindow", dbBoolean, False
    SetProp "ShowDocumentTabs", dbBoolean, False
    SetProp "Perform Name AutoCorrect", dbLong, 0
    SetProp "Track Name AutoCorrect Info", dbLong, 0
    SetProp "AppVersion", dbText, APP_VERSION
    SetProp "AppBuild", dbText, Format$(Now(), "yyyy-mm-dd hh:nn")
    BuildShortcutKeys               ' Ctrl+S, Ctrl+N, Ctrl+F, Ctrl+K and F1, the AutoKeys macro
    HideAccessSplash                ' it writes a file, so a failure leaves the settings above in place
    BuildTrustCheck "frmSplash"     ' after any StartUpForm setting: it removes StartUpForm
End Sub
