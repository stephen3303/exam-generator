Option Compare Database
Option Explicit

' Exam Generator (Access edition) - the question bank of build phase 5: the
' Question bank page (pgQuestions), the Question editor (pgQuestionEdit) and
' its dialogs dlgPart, dlgReference and dlgHistory; and the question
' approvals of build phase 6: the Approvals page (pgApprovals) and dlgReject.
' modBuild.BuildPages calls BuildQuestionGrids first, then the pages;
' BuildUserInterface builds the dialogs.

Private Const DETAIL_W As Long = 5400            ' the bank's detail card, 360 px
Private Const BANK_Y As Long = 2400              ' the bank's cards' top, 160 px
Private Const FIELD_Y As Long = 1320             ' the editor's field captions, 88 px
Private Const TAB_Y As Long = 2280               ' the editor's tabs, 152 px
Private Const APPR_W As Long = 6300              ' the approvals' detail card, 420 px

Private Function ListW() As Long
    ListW = ContentW() - DETAIL_W - px(12)
End Function

Private Function DetailX() As Long
    DetailX = PAGE_PAD + ListW() + px(12)
End Function

' The editor's tab pages: inside width, and the top and bottom of their content.
Private Function TabInW() As Long
    TabInW = ContentW() - px(32)
End Function

Private Function TabH() As Long
    TabH = PAGE_H - TAB_Y - px(16)
End Function

Private Function TabTop() As Long
    TabTop = TAB_Y + px(44)
End Function

Private Function TabBottom() As Long
    TabBottom = TAB_Y + TabH() - px(16)
End Function

' Every grid the question pages and dialogs show, before the forms that host them.
Public Sub BuildQuestionGrids()
    BuildGrid "subGridQuestions", QuestionCols(), ListW() - px(32), 2
    BuildGrid "subGridRefs", RefCols(), ListW() - px(32)
    BuildGrid "subGridRefQs", RefQCols(), DETAIL_W - px(32)
    BuildGrid "subGridParts", PartCols(), TabInW(), 2
    BuildGrid "subGridOptions", OptionCols(), TabInW(), 2
    BuildGrid "subGridQRefs", QRefCols(), TabInW()
    BuildGrid "subGridHistory", HistoryCols(), TabInW()
    BuildGrid "subGridHistoryDlg", HistoryCols(), px(720)
    BuildGrid "subGridApprovals", ApprovalCols(), ApprListW() - px(32), 2
    BuildGrid "subGridExamQueue", ExamQueueCols(), ApprListW() - px(32)
End Sub

' "field|caption|width px|options[|sort field]"
Private Function QuestionCols() As Variant
    QuestionCols = Array("ref|#|60|m|QuestionKey", "QuestionText|Question|200|ws|qsort", "TopicName|Topic|104|s", _
        "type_short|Type|60|m", "LevelName|Difficulty|76|m", "status|Status|88|p")
End Function

Private Function RefCols() As Variant
    RefCols = Array("RefTitle|Reference|236|bs", "RefNumber|Number|120|ms", "Revision|Rev.|56|m", "cited|Questions|88|r", _
        "state|Status|88|p")
End Function

Private Function RefQCols() As Variant
    RefQCols = Array("ref|#|60|m", "QuestionText|Question|180|", "status|Status|88|p")
End Function

Private Function PartCols() As Variant
    PartCols = Array("PartLabel|Part|56|b", "PartText|Text|540|w", "pts|Points|72|r", "mark|Marking|140|", "ans|Ref. answer|152|m")
End Function

Private Function OptionCols() As Variant
    OptionCols = Array("part|Part|56|b", "letter|Option|64|b", "OptionText|Text|660|w", "correct|Correct|180|p")
End Function

Private Function QRefCols() As Variant
    QRefCols = Array("RefTitle|Reference|420|b", "RefNumber|Number|200|m", "Revision|Rev.|80|m", "Paragraph|Paragraph|260|")
End Function

Private Function ApprovalCols() As Variant
    ApprovalCols = Array("ref|#|60|m|QuestionKey", "QuestionText|Question|180|ws|qsort", "TopicName|Topic|88|s", _
        "author|Author|100|", "kind|Kind|100|p")
End Function

Private Function ExamQueueCols() As Variant
    ExamQueueCols = Array("ref|#|60|m|id", "exam|Exam for|172|bs", "TemplateName|Template|120|s", "step|Step|76|m", _
        "since_txt|Waiting since|100|m|since")
End Function

Private Function HistoryCols() As Variant
    HistoryCols = Array("when_txt|When|140|m|ActionOn", "ver|Version|64|", "ActionType|Action|110|b", "person|Who|190|", _
        "step|Step|56|m", "Comments|Comments|160|")
End Function

' ---------------------------------------------------------------------------
' Question bank (page 5)
' ---------------------------------------------------------------------------

' Two tabs. Questions: the search box and status chips over the list of
' questions (funnels on each column), and the picked question on the right
' with Edit, History and More. References: the list of references with how
' many questions cite each, and the picked one on the right, its fields and
' the questions citing it. The two tabs' controls are listed in the hidden
' labels grpQ and grpR (GroupSince), which the page shows and hides.
Public Sub Build_pgQuestions()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, btnY As Long, before As Object
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Question bank", "Every question in the bank and the references they cite. Pick a row to see it; " & _
        "SMEs open a question to change it.", PAGE_PAD
    Set c = Btn("btnNew", "New question", PAGE_W - PAGE_PAD - px(132), px(24), px(132), px(32), "outline")
    c.ControlTipText = "New question (Ctrl+N)": Anchor c, 1, 0
    Set c = Btn("btnExport", "Export", PAGE_W - PAGE_PAD - px(132) - px(88), px(24), px(80), px(32), "ghost")
    c.ControlTipText = "Save the rows shown as a CSV file or Excel workbook": Anchor c, 1, 0
    TabStrip PAGE_PAD, px(76), ContentW(), "TabPick", "q|Questions|0", "r|References|0"
    SearchBox "txtSearch", PAGE_PAD, px(120), px(260), "SearchFor", "Search...", px(28)
    y = BANK_Y: h = PAGE_H - y - px(16)
    x = DetailX(): fx = x + px(16): fw = DETAIL_W - px(32): btnY = y + h - px(16) - px(30)

    ' The Questions tab.
    Set before = NamesNow()
    ChipRow "chip", PAGE_PAD + px(272), px(121), "ChipPick", "All", "Draft", "In review", "Approved", "Rejected", "Retired"
    Card "List", "Questions", PAGE_PAD, y, ListW(), h, , 2, 2
    Set c = Link("lnkAllQs", "Show all questions", PAGE_PAD + ListW() - px(16) - px(150), y + px(10), px(150), px(20), _
        "=ClearRefFilter()", 9)
    c.TextAlign = 3: c.Visible = False: Anchor c, 1, 0
    GridHeaders "hdrList", QuestionCols(), PAGE_PAD + px(16), y + px(40), w:=ListW() - px(32), sortSf:="sfList", _
        filterSf:="sfList", onFilter:="ListFiltered"
    Set c = SubF("sfList", "subGridQuestions", PAGE_PAD + px(16), y + px(62), ListW() - px(32), h - px(78)): Anchor c, 2, 2
    EmptyState "emptyList", PAGE_PAD + px(16), y + px(90), ListW() - px(32), "E82D", "No questions match", _
        "No question here has those words.", "Clear search"
    Card "Detail", "Question", x, y, DETAIL_W, h, , 1, 2
    Lbl "lblQRef", "", fx, y + px(40), fw, px(24), 11, True
    Lbl "lblQState", "", fx, y + px(64), fw, px(32), 9, False, CLR_MUTED_FG
    Set c = RO("txtQText", "", fx, y + px(100), fw, px(108), 10)
    c.ScrollBars = 2
    DetailPair "Topic", "Type", fx, y + px(216), fw
    DetailPair "Diff", "Points", fx, y + px(262), fw
    DetailPair "Mark", "Parts", fx, y + px(308), fw
    FieldLbl "capQRefs", "REFERENCES", fx, y + px(354), fw
    Lbl "lblQRefs", "", fx, y + px(372), fw, px(36), 9, False, CLR_FG
    Set c = Btn("btnOpen", "Edit", fx, btnY, px(96), px(30))
    c.ControlTipText = "Open the question in the editor (or double-click its row)"
    Set c = Btn("btnHistory", "History", fx + px(104), btnY, px(96), px(30), "outline")
    c.ControlTipText = "Who wrote, changed, submitted and approved each version"
    Set c = Btn("btnMore", "More", fx + fw - px(72), btnY, px(72), px(30), "ghost")
    c.ControlTipText = "New version, history, delete a draft (or right-click the row)"
    Set c = Lbl("lblDetailEmpty", "Pick a question in the list to see it here.", fx, y + px(120), fw, px(20), 9, False, CLR_MUTED_FG)
    c.TextAlign = 2
    GroupSince before, "grpQ"

    ' The References tab.
    Set before = NamesNow()
    Card "RList", "References", PAGE_PAD, y, ListW(), h, , 2, 2
    GridHeaders "hdrRefs", RefCols(), PAGE_PAD + px(16), y + px(40), w:=ListW() - px(32), sortSf:="sfRefs"
    Set c = SubF("sfRefs", "subGridRefs", PAGE_PAD + px(16), y + px(62), ListW() - px(32), h - px(78)): Anchor c, 2, 2
    EmptyState "emptyRefs", PAGE_PAD + px(16), y + px(90), ListW() - px(32), "E8F1", "No references match", _
        "No reference's title or number has those words.", "Clear search"
    Card "RDetail", "Reference", x, y, DETAIL_W, h, , 1, 2
    Set c = Btn("btnRefMore", "More", x + DETAIL_W - px(16) - px(72), y + px(8), px(72), px(28), "ghost"): Anchor c, 1, 0
    c.ControlTipText = "Turn the reference off or on, or delete it (or right-click its row)"
    FieldLbl "capRefTitle", "TITLE", fx, y + px(44), fw
    Txt "txtRefTitle", "", fx, y + px(62), fw, px(30)
    FieldLbl "capRefNumber", "NUMBER", fx, y + px(100), px(200)
    Txt "txtRefNumber", "", fx, y + px(118), px(200), px(30)
    FieldLbl "capRefRev", "REVISION", fx + px(212), y + px(100), fw - px(212)
    Txt "txtRefRev", "", fx + px(212), y + px(118), fw - px(212), px(30)
    Set c = Btn("btnRefSave", "Save", fx, y + px(158), px(80), px(28), "outline")
    c.ControlTipText = "Save the reference's title, number and revision (Ctrl+S)"
    Lbl "lblRefState", "", fx + px(92), y + px(158), fw - px(92), px(28), 9, False, CLR_MUTED_FG
    FieldLbl "capRefQs", "QUESTIONS CITING IT", fx, y + px(200), fw
    GridHeaders "hdrRefQs", RefQCols(), fx, y + px(220), w:=fw
    SubF "sfRefQs", "subGridRefQs", fx, y + px(242), fw, btnY - px(8) - (y + px(242))
    Set c = Btn("btnShowCiting", "Show in the question list", fx, btnY, px(200), px(30), "outline")
    c.ControlTipText = "Go to the Questions tab showing only the questions citing this reference"
    Set c = Lbl("lblRefEmpty", "Pick a reference in the list to see it here.", fx, y + px(120), fw, px(20), 9, False, CLR_MUTED_FG)
    c.TextAlign = 2
    GroupSince before, "grpR"
    SaveForm "pgQuestions"
End Sub

' Two captions side by side with a value under each: cap<a>/lbl<a> on the
' left half, cap<b>/lbl<b> on the right.
Private Sub DetailPair(ByVal a As String, ByVal b As String, ByVal x As Long, ByVal y As Long, ByVal w As Long)
    Dim hw As Long
    hw = (w - px(12)) \ 2
    FieldLbl "capQ" & a, UCase$(PairCaption(a)), x, y, hw
    Lbl "lblQ" & a, "", x, y + px(18), hw, px(20), 10
    FieldLbl "capQ" & b, UCase$(PairCaption(b)), x + hw + px(12), y, hw
    Lbl "lblQ" & b, "", x + hw + px(12), y + px(18), hw, px(20), 10
End Sub

Private Function PairCaption(ByVal k As String) As String
    Select Case k
        Case "Diff": PairCaption = "Difficulty"
        Case "Mark": PairCaption = "Portion marking"
        Case Else: PairCaption = k
    End Select
End Function

' The names of the form's controls now.
Private Function NamesNow() As Object
    Dim c As Control, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    For Each c In F.Controls
        d(c.Name) = True
    Next
    Set NamesNow = d
End Function

' List the controls built since before that show (Visible) in a hidden
' label nm, "a;b;c", for the page to show and hide as a group.
Private Sub GroupSince(before As Object, ByVal nm As String)
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

' ---------------------------------------------------------------------------
' Question editor (page 6)
' ---------------------------------------------------------------------------

' The question's own fields across the top (topic, type, difficulty, points,
' portion marking) and six tabs under them: Question (its text), Parts,
' Answer options, References, Reference answer and History. Save, Submit
' for approval, More and Back in the header; the subtitle says which
' question and version it is and where it stands. Unbound: the page loads
' the question and keeps its parts, options and references in the working
' copy (modQuestions) until Save.
Public Sub Build_pgQuestionEdit()
    Dim c As Control, x As Long, y As Long, ix As Long, iw As Long, iy As Long, ib As Long, btnY As Long, bx As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Question editor", "Write a question with its parts, portion markings, answer options and references. " & _
        "Save keeps it as a draft; Submit sends it to the DH.", PAGE_PAD
    ' The header's buttons, right to left.
    bx = PAGE_W - PAGE_PAD - px(88)
    Set c = Btn("btnSave", "Save", bx, px(24), px(88), px(32)): Anchor c, 1, 0
    c.ControlTipText = "Save the question as a draft (Ctrl+S)"
    bx = bx - px(8) - px(164)
    Set c = Btn("btnSubmit", "Submit for approval", bx, px(24), px(164), px(32), "outline"): Anchor c, 1, 0
    c.ControlTipText = "Save it and send it to the department head"
    bx = bx - px(8) - px(72)
    Set c = Btn("btnMore", "More", bx, px(24), px(72), px(32), "ghost"): Anchor c, 1, 0
    c.ControlTipText = "Undo changes, new version, history, delete a draft"
    bx = bx - px(8) - px(72)
    Set c = Btn("btnBack", "Back", bx, px(24), px(72), px(32), "ghost"): Anchor c, 1, 0
    c.ControlTipText = "Back to the question bank"
    ' The question's own fields.
    x = PAGE_PAD: y = FIELD_Y
    FieldLbl "capTopic", "TOPIC", x, y, px(260)
    Set c = PickBox("txtTopic", "", "SELECT TopicID, TopicName FROM tblTopics WHERE IsActive AND TopicWritable(TopicID)", x, y + px(18), px(260), px(30), _
        "Type part of a topic's name")
    Ev c, "OnKeyDown"
    x = x + px(271)
    FieldLbl "capType", "TYPE", x, y, px(330)
    Segmented "segType", "", x, y + px(18), px(330), "mc|Multiple choice;tf|True/false;sa|Short answer", px(30), "TypeChanged"
    x = x + px(341)
    FieldLbl "capDiff", "DIFFICULTY", x, y, px(130)
    Set c = Cbo("cboDiff", "", "SELECT DifficultyID, LevelName FROM tblDifficultyLevels ORDER BY SortOrder", x, y + px(18), _
        px(130), px(30), 2, "0;" & px(130))
    Ev c, "AfterUpdate"
    x = x + px(141)
    FieldLbl "capPoints", "POINTS", x, y, px(80)
    Set c = Txt("txtPoints", "", x, y + px(18), px(80), px(30))
    c.TextAlign = 3
    Ev c, "AfterUpdate"
    x = x + px(91)
    FieldLbl "capMark", "PORTION MARKING", x, y, PAGE_PAD + ContentW() - x
    Set c = Cbo("cboMark", "", "SELECT ClassLevelID, Marking FROM tblClassLevels ORDER BY LevelRank", x, y + px(18), _
        PAGE_PAD + ContentW() - x, px(30), 2, "0;" & (PAGE_PAD + ContentW() - x))
    Ev c, "AfterUpdate"
    ' The tabs.
    TabCtl "tabQ", PAGE_PAD, TAB_Y, ContentW(), TabH(), "tbQuestion|Question", "tbParts|Parts", "tbOptions|Answer options", _
        "tbRefs|References", "tbAnswer|Reference answer", "tbHistory|History"
    ix = PAGE_PAD + px(16): iw = TabInW(): iy = TabTop(): ib = TabBottom(): btnY = ib - px(28)
    ' Question.
    TabLbl "capText", "QUESTION TEXT, AS IT WILL PRINT", ix, iy, iw, "tbQuestion"
    Set c = Txt("txtText", "", ix, iy + px(18), iw, ib - px(26) - (iy + px(18)), 10, acDetail, "tbQuestion")
    c.EnterKeyBehavior = True: c.ScrollBars = 2
    Ev c, "AfterUpdate"
    Lbl "lblTextHint", "Parts (A, B and so on) go on the Parts tab, answer choices on Answer options. Extra spaces and " & _
        "blank lines are tidied when you save.", ix, ib - px(20), iw, px(20), 9, False, CLR_MUTED_FG, acDetail, "tbQuestion"
    ' Parts.
    Alert "alPoints", ix, iy, iw, px(40), "info", "", "", False, acDetail, "tbParts"
    GridHeaders "hdrParts", PartCols(), ix, iy + px(52), "tbParts", w:=iw
    SubF "sfParts", "subGridParts", ix, iy + px(74), iw, btnY - px(8) - (iy + px(74)), acDetail, "tbParts"
    RowButtons "Part", "part", ix, btnY, "tbParts"
    ' Answer options.
    Alert "alOptions", ix, iy, iw, px(40), "info", "", "", False, acDetail, "tbOptions"
    GridHeaders "hdrOptions", OptionCols(), ix, iy + px(52), "tbOptions", w:=iw
    SubF "sfOptions", "subGridOptions", ix, iy + px(74), iw, btnY - px(8) - (iy + px(74)), acDetail, "tbOptions"
    RowButtons "Option", "option", ix, btnY, "tbOptions"
    Set c = Btn("btnCorrectOption", "Mark correct", ix + px(368), btnY, px(112), px(28), "outline", acDetail, "tbOptions")
    c.ControlTipText = "Make the selected option the correct answer for its question or part"
    ' Which part Add option adds to (shown only while the question has parts).
    Lbl "lblOptPart", "Add to part", ix + px(492), btnY + px(5), px(80), px(20), 9, False, CLR_MUTED_FG, acDetail, "tbOptions"
    Set c = Cbo("cboOptPart", "", "", ix + px(576), btnY, px(72), px(28), 2, "0;" & px(72), acDetail, "tbOptions")
    c.ControlTipText = "The part a new option is added to"
    ' References.
    Alert "alRefs", ix, iy, iw, px(40), "info", "", "", False, acDetail, "tbRefs"
    GridHeaders "hdrQRefs", QRefCols(), ix, iy + px(52), "tbRefs", w:=iw
    SubF "sfRefs", "subGridQRefs", ix, iy + px(74), iw, btnY - px(8) - (iy + px(74)), acDetail, "tbRefs"
    Set c = Btn("btnAddRef", "Cite a reference", ix, btnY, px(140), px(28), "outline", acDetail, "tbRefs")
    c.ControlTipText = "Cite a reference already in the list, or add a new one"
    Set c = Btn("btnEditRef", "Change paragraph", ix + px(148), btnY, px(144), px(28), "ghost", acDetail, "tbRefs")
    c.ControlTipText = "Change the paragraph cited (or double-click the row)"
    Set c = Btn("btnRemoveRef", "Remove", ix + px(300), btnY, px(80), px(28), "ghost", acDetail, "tbRefs")
    c.ControlTipText = "Stop citing the selected reference"
    ' Reference answer.
    Alert "alAnswer", ix, iy, iw, px(40), "info", "", "Reference answers help grade the command's answer key. They are " & _
        "never printed on the exam. Each part's own answer is entered with the part.", False, acDetail, "tbAnswer"
    TabLbl "capRefAns", "REFERENCE ANSWER (OPTIONAL)", ix, iy + px(52), iw, "tbAnswer"
    Set c = Txt("txtRefAns", "", ix, iy + px(70), iw, ib - (iy + px(70)), 10, acDetail, "tbAnswer")
    c.EnterKeyBehavior = True: c.ScrollBars = 2
    Ev c, "AfterUpdate"
    ' History.
    GridHeaders "hdrHistory", HistoryCols(), ix, iy, "tbHistory", w:=iw, sortSf:="sfHistory"
    SubF "sfHistory", "subGridHistory", ix, iy + px(22), iw, btnY - px(8) - (iy + px(22)), acDetail, "tbHistory"
    Set c = Btn("btnCopyHistory", "Copy", ix, btnY, px(80), px(28), "outline", acDetail, "tbHistory")
    c.ControlTipText = "Copy the history as text, to paste into an email or a report"
    Lbl "lblHistoryEmpty", "", ix + px(92), btnY, iw - px(92), px(28), 9, False, CLR_MUTED_FG, acDetail, "tbHistory"
    SaveForm "pgQuestionEdit"
End Sub

' An 8pt bold muted caption on a tab page (FieldLbl with a parent).
Private Sub TabLbl(ByVal nm As String, ByVal cap As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal pg As String)
    Lbl nm, cap, x, y, w, px(16), 8, True, CLR_MUTED_FG, acDetail, pg
End Sub

' The row buttons under a parts or options table: Add, Edit, up, down and
' Remove (btnAdd<what>, btnEdit<what>, btnUp<what>, btnDown<what>,
' btnRemove<what>).
Private Sub RowButtons(ByVal what As String, ByVal word As String, ByVal x As Long, ByVal y As Long, ByVal pg As String)
    Dim c As Control
    Set c = Btn("btnAdd" & what, "Add " & word, x, y, px(112), px(28), "outline", acDetail, pg)
    c.ControlTipText = "Add a " & word
    Set c = Btn("btnEdit" & what, "Edit", x + px(120), y, px(72), px(28), "ghost", acDetail, pg)
    c.ControlTipText = "Change the selected " & word & " (or double-click it)"
    TabGlyphBtn "btnUp" & what, &HE70E, "Move the selected " & word & " up", x + px(200), y, pg
    TabGlyphBtn "btnDown" & what, &HE70D, "Move the selected " & word & " down", x + px(236), y, pg
    Set c = Btn("btnRemove" & what, "Remove", x + px(276), y, px(80), px(28), "ghost", acDetail, pg)
    c.ControlTipText = "Remove the selected " & word
End Sub

Private Sub TabGlyphBtn(ByVal nm As String, ByVal glyph As Long, ByVal tip As String, ByVal x As Long, ByVal y As Long, _
        ByVal pg As String)
    Dim c As Control
    Set c = Btn(nm, ChrW(glyph), x, y, px(32), px(28), "ghost", acDetail, pg)
    c.FontName = "Segoe MDL2 Assets"
    c.ControlTipText = tip
End Sub

' ---------------------------------------------------------------------------
' Dialogs
' ---------------------------------------------------------------------------

' A part (A, B...) of the question in the editor: its label, points, portion
' marking, text and reference answer. Writes the working copy (wq_part).
Public Sub Build_dlgPart()
    Dim c As Control, w As Long, iw As Long
    w = px(520): iw = w - px(40)
    NewDialog "Part", w, px(448)
    Ev F, "OnLoad"
    Lbl "lblHead", "", px(20), px(20), iw, px(24), 11, True
    FieldLbl "capLabel", "LABEL", px(20), px(60), px(72)
    Set c = Txt("txtLabel", "", px(20), px(78), px(72), px(30))
    c.TextAlign = 2
    FieldLbl "capPoints", "POINTS", px(104), px(60), px(96)
    Set c = Txt("txtPartPoints", "", px(104), px(78), px(96), px(30))
    c.TextAlign = 3
    Ev c, "AfterUpdate"
    FieldLbl "capMark", "PORTION MARKING", px(212), px(60), w - px(232)
    Cbo "cboPartMark", "", "SELECT ClassLevelID, Marking FROM tblClassLevels ORDER BY LevelRank", px(212), px(78), _
        w - px(232), px(30), 2, "0;" & (w - px(232))
    FieldLbl "capPartText", "TEXT, AS IT WILL PRINT", px(20), px(120), iw
    Set c = Txt("txtPartText", "", px(20), px(138), iw, px(96))
    c.EnterKeyBehavior = True: c.ScrollBars = 2
    FieldLbl "capPartAns", "REFERENCE ANSWER (NEVER PRINTED)", px(20), px(246), iw
    Set c = Txt("txtPartAns", "", px(20), px(264), iw, px(72))
    c.EnterKeyBehavior = True: c.ScrollBars = 2
    Lbl "lblPointsHint", "", px(20), px(344), iw, px(36), 9, False, CLR_MUTED_FG
    Set c = Btn("btnOK", "OK", w - px(20) - 2 * px(88) - px(8), px(394), px(88), px(30))
    c.Default = True
    Set c = Btn("btnCancel", "Cancel", w - px(20) - px(88), px(394), px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgPart"
End Sub

' Cite a reference (from the editor) or add one (from the bank). Cite:
' pick a reference already listed, or add a new one, and give the
' paragraph. Changing a citation: only the paragraph.
Public Sub Build_dlgReference()
    Dim c As Control, w As Long, iw As Long
    w = px(500): iw = w - px(40)
    NewDialog "Reference", w, px(410)
    Ev F, "OnLoad"
    Lbl "lblHead", "", px(20), px(20), iw, px(24), 11, True
    Segmented "segSource", "", px(20), px(60), iw, "pick|Cite one already listed;new|Add a new reference", px(30), "SourceChanged"
    FieldLbl "capRef", "REFERENCE", px(20), px(104), iw
    Set c = PickBox("txtRef", "", "SELECT RefID, RefTitle, Nz(RefNumber,'') & IIf(Nz(Revision,'')='','',' rev ' & Revision) AS info " & _
        "FROM tblSourceRefs WHERE IsActive", px(20), px(122), iw, px(30), "Type part of a title or number")
    Ev c, "OnKeyDown"
    FieldLbl "capTitle", "TITLE", px(20), px(104), iw
    Txt "txtTitle", "", px(20), px(122), iw, px(30)
    FieldLbl "capNumber", "NUMBER", px(20), px(162), px(300)
    Txt "txtNumber", "", px(20), px(180), px(300), px(30)
    FieldLbl "capRev", "REVISION", px(332), px(162), iw - px(312)
    Txt "txtRev", "", px(332), px(180), iw - px(312), px(30)
    FieldLbl "capPara", "PARAGRAPH", px(20), px(222), px(220)
    Txt "txtPara", "", px(20), px(240), px(220), px(30)
    Lbl "lblParaHint", "Where in the reference the answer is, such as 3.2.1 or Table 4. Optional.", px(252), px(236), _
        iw - px(232), px(36), 9, False, CLR_MUTED_FG
    Lbl "lblNote", "", px(20), px(284), iw, px(54), 9, False, CLR_MUTED_FG
    Set c = Btn("btnOK", "OK", w - px(20) - 2 * px(88) - px(8), px(360), px(88), px(30))
    c.Default = True
    Set c = Btn("btnCancel", "Cancel", w - px(20) - px(88), px(360), px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgReference"
End Sub

' Every version of one question: who wrote, changed, submitted, approved,
' rejected and retired it, and when, with Copy.
Public Sub Build_dlgHistory()
    Dim c As Control, w As Long, iw As Long
    w = px(760): iw = w - px(40)
    NewDialog "History", w, px(500)
    Ev F, "OnLoad"
    Set c = Txt("txtAuthor", "", px(20), px(20), px(40), px(40))
    c.Visible = False: c.TabStop = False
    Avatar "avAuthor", "txtAuthor", px(20), px(20), px(40)
    Lbl "lblHead", "", px(72), px(20), iw - px(52), px(24), 11, True
    Lbl "lblSub", "", px(72), px(44), iw - px(52), px(20), 9, False, CLR_MUTED_FG
    GridHeaders "hdrHist", HistoryCols(), px(20), px(84), w:=iw, sortSf:="sfHist"
    SubF "sfHist", "subGridHistoryDlg", px(20), px(106), iw, px(330)
    Set c = Btn("btnCopy", "Copy", px(20), px(452), px(88), px(30), "outline")
    c.ControlTipText = "Copy the history as text"
    Set c = Btn("btnClose", "Close", w - px(20) - px(88), px(452), px(88), px(30))
    c.Default = True: c.Cancel = True
    SaveForm "dlgHistory"
End Sub

' ---------------------------------------------------------------------------
' Approvals (page 3)
' ---------------------------------------------------------------------------

Private Function ApprListW() As Long
    ApprListW = ContentW() - APPR_W - px(12)
End Function

' Two tabs. Questions: the questions waiting for the signed-in approver (the
' search box over the list, funnels and sorting on each column), and the
' picked one on the right: where it is in the chain (SME, DH, XO, SM), its
' author, a read-only preview of the whole question, and Approve, Reject and
' History. Exams: a placeholder until build phase 8. The Questions tab's
' controls are listed in the hidden label grpQ (GroupSince); the stepper and
' the avatar, which paint their own parts, are shown and hidden by the page.
Public Sub Build_pgApprovals()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, btnY As Long, before As Object
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Approvals", "Questions and exams waiting for your approval at your step of the chain. Approve sends one on; " & _
        "Reject sends it back with your reason: a question to its author, an exam to its SMEs.", PAGE_PAD
    TabStrip PAGE_PAD, px(76), ContentW(), "TabPick", "q|Questions|0", "x|Exams|0"
    y = BANK_Y: h = PAGE_H - y - px(16)
    x = PAGE_PAD + ApprListW() + px(12): fx = x + px(16): fw = APPR_W - px(32): btnY = y + h - px(16) - px(30)

    ' The Questions tab.
    Set before = NamesNow()
    SearchBox "txtSearch", PAGE_PAD, px(120), px(260), "SearchFor", "Search...", px(28)
    Card "List", "Waiting for you", PAGE_PAD, y, ApprListW(), h, , 2, 2
    GridHeaders "hdrList", ApprovalCols(), PAGE_PAD + px(16), y + px(40), w:=ApprListW() - px(32), sortSf:="sfList", _
        filterSf:="sfList", onFilter:="ListFiltered"
    Set c = SubF("sfList", "subGridApprovals", PAGE_PAD + px(16), y + px(62), ApprListW() - px(32), h - px(78)): Anchor c, 2, 2
    EmptyState "emptyList", PAGE_PAD + px(16), y + px(90), ApprListW() - px(32), "E8FB", "Nothing waiting for you", _
        "Questions sent to your step of the chain show here.", "Clear search"
    Card "Detail", "Question", x, y, APPR_W, h, , 1, 2
    Lbl "lblQRef", "", fx + px(44), y + px(112), fw - px(44), px(22), 11, True
    Lbl "lblQBy", "", fx + px(44), y + px(134), fw - px(44), px(20), 9, False, CLR_MUTED_FG
    Lbl "lblQState", "", fx, y + px(162), fw, px(36), 9, False, CLR_FG
    FieldLbl "capPreview", "THE QUESTION, AS SUBMITTED", fx, y + px(204), fw
    Set c = RO("txtPreview", "", fx, y + px(222), fw, btnY - px(12) - (y + px(222)), 10)
    c.ScrollBars = 2
    Set c = Btn("btnApprove", "Approve", fx, btnY, px(96), px(30))
    c.ControlTipText = "Approve it at your step and send it on"
    Set c = Btn("btnReject", "Reject", fx + px(104), btnY, px(88), px(30), "destructive")
    c.ControlTipText = "Send it back to its author with your reason"
    Set c = Btn("btnHistory", "History", fx + fw - px(88), btnY, px(88), px(30), "outline")
    c.ControlTipText = "Who wrote, changed, submitted and approved each version"
    Set c = Lbl("lblDetailEmpty", "Pick a question in the list to see it here.", fx, y + px(120), fw, px(20), 9, False, CLR_MUTED_FG)
    c.TextAlign = 2
    GroupSince before, "grpQ"
    ' Shown and hidden by the page (ShowQuestion), not with the group.
    Stepper "stpQ", fx, y + px(44), fw, 2, "", "SME", "DH", "XO", "SM"
    Set c = Txt("txtAuthor", "", fx, y + px(114), px(36), px(36))
    c.Visible = False: c.TabStop = False
    Avatar "avAuthor", "txtAuthor", fx, y + px(114), px(36)

    ' The Exams tab: the same layout, with its own search box.
    Set before = NamesNow()
    SearchBox "txtXSearch", PAGE_PAD, px(120), px(260), "SearchExams", "Search...", px(28)
    Card "XList", "Waiting for you", PAGE_PAD, y, ApprListW(), h, , 2, 2
    GridHeaders "hdrXList", ExamQueueCols(), PAGE_PAD + px(16), y + px(40), w:=ApprListW() - px(32), sortSf:="sfXList", _
        filterSf:="sfXList", onFilter:="ListFiltered"
    Set c = SubF("sfXList", "subGridExamQueue", PAGE_PAD + px(16), y + px(62), ApprListW() - px(32), h - px(78)): Anchor c, 2, 2
    EmptyState "emptyXList", PAGE_PAD + px(16), y + px(90), ApprListW() - px(32), "E8A5", "Nothing waiting for you", _
        "Exams sent to your step of the chain show here once their SMEs sign off every topic.", "Clear search"
    Card "XDetail", "Exam", x, y, APPR_W, h, , 1, 2
    Lbl "lblXRef", "", fx, y + px(104), fw, px(22), 11, True
    Lbl "lblXBy", "", fx, y + px(126), fw, px(32), 9, False, CLR_MUTED_FG
    Lbl "lblXState", "", fx, y + px(160), fw, px(36), 9, False, CLR_FG
    FieldLbl "capXPreview", "THE EXAM, WITH ITS ANSWERS", fx, y + px(204), fw
    Set c = RO("txtXPreview", "", fx, y + px(222), fw, btnY - px(12) - (y + px(222)), 10)
    c.ScrollBars = 2
    Set c = Btn("btnXApprove", "Approve", fx, btnY, px(96), px(30))
    c.ControlTipText = "Approve it at your step and send it on"
    Set c = Btn("btnXReject", "Reject", fx + px(104), btnY, px(88), px(30), "destructive")
    c.ControlTipText = "Send one topic or all of them back to their SMEs with your reason"
    Set c = Btn("btnXHistory", "History", fx + fw - px(88), btnY, px(88), px(30), "outline")
    c.ControlTipText = "Who generated, signed off, approved and sent back the exam, and when"
    Set c = Lbl("lblXDetailEmpty", "Pick an exam in the list to see it here.", fx, y + px(120), fw, px(20), 9, False, CLR_MUTED_FG)
    c.TextAlign = 2
    GroupSince before, "grpX"
    Stepper "stpX", fx, y + px(44), fw, 2, "", "SME", "DH", "XO", "SM"
    SaveForm "pgApprovals"
End Sub

' Reject a question: who it goes back to, and the reason (required), which
' goes in the audit trail for its author to read.
Public Sub Build_dlgReject()
    Dim c As Control, w As Long, iw As Long
    w = px(500): iw = w - px(40)
    NewDialog "Reject", w, px(400)
    Ev F, "OnLoad"
    Lbl "lblHead", "", px(20), px(20), iw, px(24), 11, True
    Lbl "lblSub", "", px(20), px(46), iw, px(20), 9, False, CLR_MUTED_FG
    FieldLbl "capTo", "GOES BACK TO", px(20), px(80), iw
    RO "txtTo", "", px(20), px(98), iw, px(48), 10
    FieldLbl "capReason", "REASON (REQUIRED)", px(20), px(156), iw
    Set c = Txt("txtReason", "", px(20), px(174), iw, px(128))
    c.EnterKeyBehavior = True: c.ScrollBars = 2
    Lbl "lblHint", "Say what to change. Its author reads this in the question's history and in the editor.", px(20), px(306), _
        iw, px(20), 9, False, CLR_MUTED_FG
    Set c = Btn("btnReject", "Reject", w - px(20) - 2 * px(88) - px(8), px(346), px(88), px(30), "destructive")
    Set c = Btn("btnCancel", "Cancel", w - px(20) - px(88), px(346), px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgReject"
End Sub
