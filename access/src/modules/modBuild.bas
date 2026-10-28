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
' Steps, in order: link, local tables (the question editor's working copy
' and the Statistics page's figures),
' user interface (subforms, then pages, then dialogs, then the shell last),
' help, startup. A failed build returns
' "ERROR during <step>: <number> <description> [<last form traced>]".
'
' Each page has its Build_<page>, here or in a modBuild<area> module. (Up to
' phase 9, a page not built yet was a placeholder naming its phase; phase 10
' built the last one.)
'
' Owner: the shell and the build (build phase 1). Depends on: modBuildAdmin,
' modBuildAppearance, modBuildDash, modBuildExams, modBuildLogs, modBuildQuestions, modErrors, modGlobals,
' modQuestions, modStats; the kit's modDarkMode, modGridCols, modGridFilter,
' modKitBackend, modKitHelp, modKitKeys, modKitPalette, modKitShell,
' modMenu, modPersonnelUI, modTrust, modUIKit.

Public Const APP_VERSION As String = "0.1.0"

Public Const PAGE_W As Long = 15840             ' the page area, 1056 px
Public Const PAGE_H As Long = 9600              ' 640 px: the admin's sidebar (12 pages and 2 links) fits
Public Const PAGE_PAD As Long = 480                   ' page padding, 32 px
Private Const HELP_LIST_W As Long = 5100        ' Help topics: the list card, 340 px

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
    m_step = "local tables": EnsureWorkTables      ' the question editor's working copy (modQuestions)
    EnsureStatTables                                ' the Statistics page's figures (modStats)
    m_step = "user interface": BuildUserInterface
    m_step = "help": LoadHelpTopics srcRoot & "\help\help.txt"
    m_step = "startup": SetStartupOptions uiOnly, bePath
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
    Build_dlgTemplateTopic         ' a template's topic and count (modBuildAdmin)
    Build_dlgPart                  ' a question's part (modBuildQuestions)
    Build_dlgReference             ' cite or add a reference
    Build_dlgHistory               ' a question's history
    Build_dlgReject                ' send a question back to its author
    Build_dlgNewExam               ' generate an exam (modBuildExams)
    Build_dlgSwap                  ' swap a question on an exam topic
    Build_dlgRejectExam            ' send an exam back to its SMEs
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

' The pages, in sidebar order, after the tables they host.
Private Sub BuildPages()
    BuildGrid "subGridHelp", HelpCols(), HELP_LIST_W - px(32)
    BuildAdminGrids                ' the admin pages' tables and detail forms (modBuildAdmin)
    BuildQuestionGrids             ' the question pages' and dialogs' tables (modBuildQuestions)
    BuildExamGrids                 ' the exam pages' and dialogs' tables (modBuildExams)
    BuildDashGrids                 ' the Dashboard's and Statistics' tables (modBuildDash)
    BuildLogGrids                  ' Admin: Logs' tables (modBuildLogs)
    Build_pgDash
    Build_pgReview                 ' an SME's exam topic reviews (modBuildExams)
    Build_pgApprovals              ' question approvals (modBuildQuestions)
    Build_pgExams                  ' every exam, and the DH's New exam (modBuildExams)
    Build_pgQuestions
    Build_pgQuestionEdit
    Build_pgStats
    Build_pgAdmTopics
    Build_pgAdmTemplates
    Build_pgAdmPeople
    Build_pgAdmRules
    Build_pgAdmHelp
    Build_pgAdmLogs
    Build_pgMySettings             ' a sidebar link, not a page in the list (modBuildDash)
End Sub

' Help topics (pgAdmHelp): the topics on the left, the open topic's key,
' title and text on the right, with Save, Preview (the help window on the
' saved topic) and Undo changes. The fields are unbound: the page saves with
' KitHelpSave, which writes the data file's help_topic and this front end's
' zz_help at once.
Private Sub Build_pgAdmHelp()
    Dim c As Control, x As Long, y As Long, h As Long, dw As Long, fx As Long, fw As Long, btnY As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Admin: Help topics", "Edit the help for any page or dialog. A saved topic shows at once here, " & _
        "and for everyone else the next time they open the app.", PAGE_PAD
    Set c = Btn("btnNew", "New topic", PAGE_W - PAGE_PAD - px(112), px(24), px(112), px(32), "outline")
    c.ControlTipText = "New topic (Ctrl+N)": Anchor c, 1, 0
    SearchBox "txtSearch", PAGE_PAD, px(84), px(300), "SearchFor", "Search titles, keys and text...", px(28)
    y = px(124): h = PAGE_H - y - px(16)
    Card "List", "Topics", PAGE_PAD, y, HELP_LIST_W, h, , 0, 2
    GridHeaders "hdrHelp", HelpCols(), PAGE_PAD + px(16), y + px(40), w:=HELP_LIST_W - px(32)
    Set c = SubF("sfList", "subGridHelp", PAGE_PAD + px(16), y + px(62), HELP_LIST_W - px(32), h - px(78)): Anchor c, 0, 2
    EmptyState "emptyList", PAGE_PAD + px(16), y + px(90), HELP_LIST_W - px(32), "E721", "No topics match", _
        "No topic's title, key or text has those words.", "Clear search"
    ' The open topic.
    x = PAGE_PAD + HELP_LIST_W + px(12): dw = ContentW() - HELP_LIST_W - px(12)
    fx = x + px(16): fw = dw - px(32)
    Card "Detail", "Topic", x, y, dw, h, , 2, 2
    FieldLbl "lblKey", "KEY", fx, y + px(44), px(220)
    Set c = Txt("txtKey", "", fx, y + px(62), px(220), px(30))
    c.ControlTipText = "A page's key (exams), or a parent's key, a dot and a word (exams.new) for a topic under it"
    Ev c, "OnChange"
    FieldLbl "lblTitle", "TITLE", fx + px(236), y + px(44), fw - px(236)
    Set c = Txt("txtTitle", "", fx + px(236), y + px(62), fw - px(236), px(30)): Anchor c, 2, 0
    Ev c, "OnChange"
    FieldLbl "lblBody", "TEXT", fx, y + px(102), fw
    Set c = Txt("txtBody", "", fx, y + px(120), fw, h - px(220)): Anchor c, 2, 2
    c.EnterKeyBehavior = True               ' Enter starts a new line, as in the help file
    c.ScrollBars = 2
    c.TopMargin = px(6): c.BottomMargin = px(6)
    Ev c, "OnChange"
    btnY = y + h - px(48)
    Set c = Lbl("lblMarkup", "A line starting # is a heading and - a bullet; **text** is bold; " & _
        "a blank line starts a new paragraph.", fx, btnY - px(40), fw, px(32), 9, False, CLR_MUTED_FG)
    Anchor c, 2, 1
    Set c = Btn("btnSave", "Save", fx, btnY, px(88), px(32)): c.ControlTipText = "Save (Ctrl+S)": Anchor c, 0, 1
    Set c = Btn("btnPreview", "Preview", fx + px(96), btnY, px(96), px(32), "outline"): Anchor c, 0, 1
    c.ControlTipText = "Open the help window on this topic"
    Set c = Btn("btnUndo", "Undo changes", fx + px(200), btnY, px(120), px(32), "ghost"): Anchor c, 0, 1
    Set c = Lbl("lblStamp", "", fx + px(332), btnY, fw - px(332), px(32), 9, False, CLR_MUTED_FG)
    c.TextAlign = 3: Anchor c, 2, 1
    SaveForm "pgAdmHelp"
End Sub

' The topics table on pgAdmHelp: "field|caption|width px|options".
Private Function HelpCols() As Variant
    HelpCols = Array("topic|Topic|190|bs", "topic_key|Key|118|ms")
End Function

' Grouping the controls a page shows and hides together (one tab's): call
' NamesNow before building them and GroupSince after.

' The names of the form's controls now.
Public Function NamesNow() As Object
    Dim c As Control, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    For Each c In F.Controls
        d(c.Name) = True
    Next
    Set NamesNow = d
End Function

' List the controls built since before that show (Visible) in a hidden
' label nm, "a;b;c", for the page to show and hide as a group.
Public Sub GroupSince(before As Object, ByVal nm As String)
    Dim c As Control, s As String
    For Each c In F.Controls
        If Not before.Exists(c.Name) Then
            If c.Visible Then s = s & ";" & c.Name
        End If
    Next
    If Len(s) > 2048 Then Err.Raise 5, , "Group " & nm & " lists too many controls for a label caption"
    Set c = Lbl(nm, Mid$(s, 2), 0, 0, px(8), px(8))
    c.Visible = False
End Sub

' 8pt bold muted uppercase caption above a field.
Public Sub FieldLbl(ByVal nm As String, ByVal cap As String, ByVal x As Long, ByVal y As Long, ByVal w As Long)
    Lbl nm, cap, x, y, w, px(16), 8, True, CLR_MUTED_FG
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

' Set database property name, adding it when the front end hasn't got it
' yet (error 3270). Any other failure, and a failed add, stops the build.
Private Sub SetProp(ByVal name As String, ByVal kind As Integer, ByVal value As Variant)
    Dim db As DAO.Database
    Set db = CurrentDb
    On Error GoTo Missing
    db.Properties(name) = value
    Exit Sub
Missing:
    If Err.Number <> 3270 Then RaiseAgain "modBuild.SetProp", name
    db.Properties.Append db.CreateProperty(name, kind, value)   ' an error here goes up to BuildAll
End Sub

' Open on the splash screen, which opens frmMain, without Access's own
' splash, the navigation pane or document tabs, and record when this front
' end was built (error-log entries and the splash show it). BuildTrustCheck
' opens the splash at startup (in place of StartUpForm) and shows the kit's
' trust page when Access would ask to enable content.
' Shift at opening (AllowBypassKey) and F11 (AllowSpecialKeys) stay on in
' this .accdb, the developer's copy; build.ps1 -Release turns them off in
' the locked .accde users get.
' HasDataFile records whether this front end was built linked to a data file,
' so a copy that loses its tblSettings link isn't taken for a UI-only build
' (modSignIn.DataLinked). -UiOnly keeps the value the full build set.
Private Sub SetStartupOptions(ByVal uiOnly As Boolean, ByVal bePath As String)
    If Not uiOnly Then SetProp "HasDataFile", dbBoolean, (Len(bePath) > 0)
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
