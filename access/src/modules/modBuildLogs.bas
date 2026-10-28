Option Compare Database
Option Explicit

' Exam Generator (Access edition) - the Admin: Logs page (build phase 10).
' modBuild.BuildPages calls BuildLogGrids first, then Build_pgAdmLogs.
'
' Owner: build phase 10. Depends on: modBuild (PAGE_W, PAGE_H, PAGE_PAD,
' ContentW, NamesNow, GroupSince); the kit's modUIKit.

Private Const LOG_DETAIL_W As Long = 5400       ' the entry card on the right, 360 px

Private Function LogListW() As Long
    LogListW = ContentW() - LOG_DETAIL_W - px(12)
End Function

Public Sub BuildLogGrids()
    BuildGrid "subGridAudit", AuditCols(), LogListW() - px(32)
    BuildGrid "subGridErrors", ErrorCols(), LogListW() - px(32)
End Sub

' "field|caption|width px|options[|sort field]"
Private Function AuditCols() As Variant
    AuditCols = Array("when_txt|When|96|m|ActionOn", "who|Who|130|s", "ActionType|Action|90|", "item|Item|120|bs", _
        "step|As|44|m", "note|Note|108|s")
End Function

Private Function ErrorCols() As Variant
    ErrorCols = Array("when_txt|When|96|m|logged_at", "who|User|110|s", "page|Page|72|m", "source|Procedure|140|s", _
        "err_txt|Error|170|s")
End Function

' Two tabs, Audit trail and Errors, over one search box and one day range;
' each tab's list on the left with its column funnels, and the picked
' entry in full on the right with Copy. Export saves the open list; Clear
' log (Errors only) empties the error log.
Public Sub Build_pgAdmLogs()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, lw As Long, before As Object
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Admin: Logs", "The audit trail (who authored, changed, submitted, reviewed, approved, sent back, generated and " & _
        "downloaded what) and the error log. Both are read only; the error log can be cleared.", PAGE_PAD
    TabStrip PAGE_PAD, px(76), ContentW(), "TabPick", "a|Audit trail", "e|Errors"

    ' The filters both tabs share, and Export.
    y = px(120)
    SearchBox "txtSearch", PAGE_PAD, y, px(260), "SearchFor", "Search...", px(28)
    x = PAGE_PAD + px(272)
    Lbl "lblFrom", "From", x, y, px(36), px(28), 9, False, CLR_MUTED_FG
    Set c = DateBox("dtFrom", "", x + px(40), y, px(28)): Ev c, "AfterUpdate"
    Ev F.Controls("cal_dtFrom"), "OnClick"
    x = x + px(40) + DATEFIELD_W + px(8)
    Lbl "lblTo", "to", x, y, px(20), px(28), 9, False, CLR_MUTED_FG
    Set c = DateBox("dtTo", "", x + px(28), y, px(28)): Ev c, "AfterUpdate"
    Ev F.Controls("cal_dtTo"), "OnClick"
    Set c = Btn("btnExport", "Export", PAGE_W - PAGE_PAD - px(88), y, px(88), px(28), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Save the list as an Excel or CSV file"

    y = px(164): h = PAGE_H - y - px(16): lw = LogListW()
    ' The Audit trail tab.
    Set before = NamesNow()
    Card "List", "Audit trail", PAGE_PAD, y, lw, h, , 2, 2
    GridHeaders "hdrList", AuditCols(), PAGE_PAD + px(16), y + px(40), w:=lw - px(32), sortSf:="sfList", _
        filterSf:="sfList", onFilter:="ListFiltered"
    Set c = SubF("sfList", "subGridAudit", PAGE_PAD + px(16), y + px(62), lw - px(32), h - px(78)): Anchor c, 2, 2
    GroupSince before, "grpA"
    EmptyState "emptyList", PAGE_PAD + px(16), y + px(90), lw - px(32), "E81C", "Nothing in the audit trail", _
        "No entry fits the search, the dates and the column filters.", "Show all"

    ' The Errors tab.
    Set before = NamesNow()
    Card "EList", "Errors", PAGE_PAD, y, lw, h, , 2, 2
    GridHeaders "hdrEList", ErrorCols(), PAGE_PAD + px(16), y + px(40), w:=lw - px(32), sortSf:="sfEList", _
        filterSf:="sfEList", onFilter:="ListFiltered"
    Set c = SubF("sfEList", "subGridErrors", PAGE_PAD + px(16), y + px(62), lw - px(32), h - px(78)): Anchor c, 2, 2
    Set c = Btn("btnClear", "Clear log", PAGE_W - PAGE_PAD - px(88) - px(8) - px(96), px(120), px(96), px(28), "destructive")
    Anchor c, 1, 0
    c.ControlTipText = "Delete every entry in the error log"
    GroupSince before, "grpE"
    EmptyState "emptyEList", PAGE_PAD + px(16), y + px(90), lw - px(32), "E783", "No errors logged", _
        "No error fits the search, the dates and the column filters.", "Show all"

    ' The picked entry, either tab.
    x = PAGE_PAD + lw + px(12): fx = x + px(16): fw = LOG_DETAIL_W - px(32)
    Card "Detail", "Entry", x, y, LOG_DETAIL_W, h, , 1, 2
    Set c = RO("txtDetail", "", fx, y + px(40), fw, h - px(40) - px(62), 9)
    c.ScrollBars = 2: Anchor c, 1, 2
    Set c = Btn("btnCopy", "Copy", fx, y + h - px(46), px(80), px(30), "outline"): Anchor c, 1, 1
    c.ControlTipText = "Copy the whole entry, to paste into an email or a report"
    Set c = Lbl("lblDetailEmpty", "Pick an entry in the list to see all of it here.", fx, y + px(120), fw, px(20), 9, False, _
        CLR_MUTED_FG)
    c.TextAlign = 2
    SaveForm "pgAdmLogs"
End Sub
