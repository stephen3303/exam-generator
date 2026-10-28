Option Compare Database
Option Explicit

' Exam Generator (Access edition) - the pages of build phase 9: the
' Dashboard (pgDash), Statistics (pgStats) and My settings (pgMySettings).
' modBuild.BuildPages calls BuildDashGrids first, then the pages.
'
' Owner: build phase 9. Depends on: modBuild (PAGE_W, PAGE_H, PAGE_PAD,
' ContentW); the kit's modUIKit.

Private Const SIDE_W As Long = 5400             ' the right-hand cards (Recent activity, References), 360 px

Private Function MainW() As Long
    MainW = ContentW() - SIDE_W - px(12)
End Function

' Every grid the phase 9 pages show, before the forms that host them.
Public Sub BuildDashGrids()
    BuildGrid "subGridWaiting", WaitingCols(), MainW() - px(32)
    BuildGrid "subGridActivity", ActivityCols(), SIDE_W - px(32)
    BuildGrid "subGridStatTopics", StatTopicCols(), MainW() - px(32)
    BuildGrid "subGridStatRefs", StatRefCols(), SIDE_W - px(32)
End Sub

' "field|caption|width px|options[|sort field]"
Private Function WaitingCols() As Variant
    WaitingCols = Array("kind|Kind|88|p", "what|Item|324|b", "step|Step|60|m", "since_txt|Waiting since|116|m|since")
End Function

Private Function ActivityCols() As Variant
    ActivityCols = Array("when_txt|When|84|m|ActionOn", "who|Who|110|", "what|What|134|")
End Function

Private Function StatTopicCols() As Variant
    StatTopicCols = Array("TopicName|Topic|160|bs", "nq|Given|56|r", "nx|Exams|56|r", "nref|Refs|52|r", _
        "top_txt|Top share|72|r|top_pct", "top_ref_txt|Most cited reference|192|s")
End Function

Private Function StatRefCols() As Variant
    StatRefCols = Array("ref|Reference|200|s", "n|Qs|56|r", "pct_txt|Share|72|r|pct")
End Function

' ---------------------------------------------------------------------------
' Dashboard (page 1)
' ---------------------------------------------------------------------------

' Four tiles (reviews open, approvals waiting, questions sent back, topics
' short of approved questions), the short topics' alert under them, then
' everything waiting on the person, the longest waiting first, with Open,
' and the recent activity on the right. The tiles' values are set in code.
Public Sub Build_pgDash()
    Dim c As Control, y As Long, h As Long, tw As Long, i As Long, x As Long, mw As Long
    Dim tiles As Variant, labels As Variant, glyphs As Variant
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Dashboard", "What waits on you, and the topics short of approved questions. " & _
        "Turn the list off, or emails on, under My settings.", PAGE_PAD
    tiles = Array("tileReviews", "tileApprovals", "tileReturned", "tileShort")
    labels = Array("Reviews open", "Approvals waiting", "Sent back to you", "Topics short")
    glyphs = Array("E73E", "E8FB", "E7A7", "E7BA")
    tw = (ContentW() - 3 * px(12)) \ 4
    For i = 0 To 3
        x = PAGE_PAD + i * (tw + px(12))
        StatTile CStr(tiles(i)), CStr(labels(i)), "", x, px(84), IIf(i = 3, ContentW() - 3 * (tw + px(12)), tw), px(112), _
            caption:=" ", glyph:=CStr(glyphs(i))
    Next
    Alert "alShort", PAGE_PAD, px(208), ContentW(), px(56)

    y = px(276): h = PAGE_H - y - px(16): mw = MainW()
    Card "Waiting", "Waiting on you", PAGE_PAD, y, mw, h, , 2, 2
    GridHeaders "hdrWaiting", WaitingCols(), PAGE_PAD + px(16), y + px(40), w:=mw - px(32), sortSf:="sfWaiting"
    Set c = SubF("sfWaiting", "subGridWaiting", PAGE_PAD + px(16), y + px(62), mw - px(32), h - px(124)): Anchor c, 2, 2
    EmptyState "emptyWaiting", PAGE_PAD + px(16), y + px(70), mw - px(32), "E930", "Nothing waits on you", _
        "Reviews, approvals and questions sent back to you show here.", "My settings"
    Set c = Btn("btnOpen", "Open", PAGE_PAD + px(16), y + h - px(46), px(88), px(30)): Anchor c, 0, 1
    c.ControlTipText = "Go to the picked item (or double-click its row)"

    x = PAGE_PAD + mw + px(12)
    Card "Activity", "Recent activity", x, y, SIDE_W, h, , 1, 2
    GridHeaders "hdrActivity", ActivityCols(), x + px(16), y + px(40), w:=SIDE_W - px(32)
    Set c = SubF("sfActivity", "subGridActivity", x + px(16), y + px(62), SIDE_W - px(32), h - px(78)): Anchor c, 1, 2
    EmptyState "emptyActivity", x + px(16), y + px(70), SIDE_W - px(32), "E81C", "No activity yet", _
        "Saves, approvals and exams generated show here as they happen."
    SaveForm "pgDash"
End Sub

' ---------------------------------------------------------------------------
' Statistics (page 6)
' ---------------------------------------------------------------------------

' The period (the last 6 or 12 months, or two dates) and the template over
' three tiles (exams, questions given, references cited), the weak
' reference coverage alert, the topics with their most cited reference, and
' the picked topic's references on the right. Print and Export take the
' topics table.
Public Sub Build_pgStats()
    Dim c As Control, x As Long, y As Long, h As Long, tw As Long, i As Long, mw As Long
    Dim tiles As Variant, labels As Variant, glyphs As Variant
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Statistics", "How many questions each topic gave on approved exams over a period, and how much of each " & _
        "topic came from one reference. A topic where one reference gives 80% or more of the questions has weak coverage.", PAGE_PAD
    y = px(84)
    Segmented "segPeriod", "", PAGE_PAD, y, px(300), "m6|6 months;m12|12 months;dates|Dates", px(32), "PeriodChanged"
    x = PAGE_PAD + px(312)
    Lbl "lblFrom", "From", x, y, px(36), px(32), 9, False, CLR_MUTED_FG
    Set c = DateBox("dtFrom", "", x + px(40), y, px(32)): Ev c, "AfterUpdate"
    Ev CalButton("dtFrom"), "OnClick"
    x = x + px(40) + DATEFIELD_W + px(8)
    Lbl "lblTo", "to", x, y, px(20), px(32), 9, False, CLR_MUTED_FG
    Set c = DateBox("dtTo", "", x + px(28), y, px(32)): Ev c, "AfterUpdate"
    Ev CalButton("dtTo"), "OnClick"
    x = x + px(28) + DATEFIELD_W + px(16)
    Set c = Cbo("cboTemplate", "", "", x, y, PAGE_W - PAGE_PAD - x, px(32), 2, "0;" & (PAGE_W - PAGE_PAD - x))
    c.ControlTipText = "The template the exams were generated from (a template stands in for a series)"
    Ev c, "AfterUpdate"

    tiles = Array("tileExams", "tileGiven", "tileRefs")
    labels = Array("Exams approved", "Questions given", "References cited")
    glyphs = Array("E8A5", "E82D", "E8F1")
    tw = (ContentW() - 2 * px(12)) \ 3
    For i = 0 To 2
        StatTile CStr(tiles(i)), CStr(labels(i)), "", PAGE_PAD + i * (tw + px(12)), px(128), _
            IIf(i = 2, ContentW() - 2 * (tw + px(12)), tw), px(112), caption:=" ", glyph:=CStr(glyphs(i)), fmt:="#,##0"
    Next
    Alert "alWeak", PAGE_PAD, px(252), ContentW(), px(56)

    y = px(320): h = PAGE_H - y - px(16): mw = MainW()
    Card "Topics", "Topics", PAGE_PAD, y, mw, h, , 2, 2
    GridHeaders "hdrTopics", StatTopicCols(), PAGE_PAD + px(16), y + px(40), w:=mw - px(32), sortSf:="sfTopics"
    Set c = SubF("sfTopics", "subGridStatTopics", PAGE_PAD + px(16), y + px(62), mw - px(32), h - px(124)): Anchor c, 2, 2
    EmptyState "emptyTopics", PAGE_PAD + px(16), y + px(62), mw - px(32), "E9D2", "No approved exams in this period", _
        "Pick a longer period or every template. Only exams the SM gave final approval to count."
    Set c = Btn("btnPrint", "Print", PAGE_PAD + px(16), y + h - px(46), px(80), px(30), "outline"): Anchor c, 0, 1
    c.ControlTipText = "Print the topics table"
    Set c = Btn("btnExport", "Export", PAGE_PAD + px(104), y + h - px(46), px(88), px(30), "outline"): Anchor c, 0, 1
    c.ControlTipText = "Save the topics table as an Excel or CSV file"

    x = PAGE_PAD + mw + px(12)
    Card "Refs", "References", x, y, SIDE_W, h, , 1, 2
    GridHeaders "hdrRefs", StatRefCols(), x + px(16), y + px(40), w:=SIDE_W - px(32), sortSf:="sfRefs"
    Set c = SubF("sfRefs", "subGridStatRefs", x + px(16), y + px(62), SIDE_W - px(32), h - px(78)): Anchor c, 1, 2
    EmptyState "emptyRefs", x + px(16), y + px(62), SIDE_W - px(32), "E8F1", "No references", _
        "Pick a topic to see which references its questions cite."
    SaveForm "pgStats"
End Sub

' A DateBox's calendar button: the page handles its click, so that a picked
' date refreshes the figures (the kit's PickDate sets the box without firing
' its AfterUpdate).
Private Function CalButton(ByVal nm As String) As Control
    Set CalButton = F.Controls("cal_" & nm)
End Function

' ---------------------------------------------------------------------------
' My settings (sidebar link)
' ---------------------------------------------------------------------------

' The person's notice switches (the Dashboard's list, and emails), the
' address emails go to with a test send, and the palette picker.
Public Sub Build_pgMySettings()
    Dim c As Control, x As Long, y As Long, w As Long, fw As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "My settings", "Your notices and your colors. They change only your own settings.", PAGE_PAD
    x = PAGE_PAD: y = px(84): w = px(480): fw = w - px(32)
    Card "Notices", "Notices", x, y, w, px(248)
    Toggle "tgDash", "", x + px(20), y + px(52), px(36), px(20), "DashChanged"
    Link "lbltgDash", "Dashboard list", x + px(68), y + px(48), fw - px(52), px(22), "=ToggleSwitch([Form],""tgDash"")", 10, CLR_FG
    Lbl "lblDashSub", "What waits on you, and the counts over it, on the Dashboard.", x + px(68), y + px(70), fw - px(52), px(18), _
        9, False, CLR_MUTED_FG
    Toggle "tgEmail", "", x + px(20), y + px(108), px(36), px(20), "EmailChanged"
    Link "lbltgEmail", "Emails", x + px(68), y + px(104), fw - px(52), px(22), "=ToggleSwitch([Form],""tgEmail"")", 10, CLR_FG
    Lbl "lblEmailSub", "An Outlook email when a review, an approval or a question sent back waits on you.", _
        x + px(68), y + px(126), fw - px(52), px(32), 9, False, CLR_MUTED_FG
    Lbl "lblEmail", "", x + px(16), y + px(166), fw, px(32), 9, False, CLR_MUTED_FG   ' two lines: an address or why there is none
    Set c = Btn("btnTestEmail", "Send me a test email", x + px(16), y + px(200), px(168), px(30), "outline")
    c.ControlTipText = "Check that the app's emails reach you"

    y = y + px(248) + px(12)
    Card "Colors", "Colors", x, y, w, px(120)
    Lbl "lblColors", "Choose the palette the whole app uses on this computer: light, dark and six more.", _
        x + px(16), y + px(40), fw, px(36), 9, False, CLR_MUTED_FG
    Set c = Btn("btnAppearance", "Appearance", x + px(16), y + px(80), px(120), px(28), "outline")
    c.OnClick = "=AppAppearance()"

    ' Sign out (modSignIn.SignOutNow): the next person signs in on this
    ' computer, which is also how one computer walks the approval chain.
    y = y + px(120) + px(12)
    Card "Account", "Sign out", x, y, w, px(120)
    Lbl "lblSignOut", "Let someone else sign in on this computer with their sign-in name and app password.", _
        x + px(16), y + px(40), fw, px(36), 9, False, CLR_MUTED_FG
    Set c = Btn("btnSignOut", "Sign out", x + px(16), y + px(80), px(120), px(28), "outline")
    c.OnClick = "=SignOutNow()"
    SaveForm "pgMySettings"
End Sub
