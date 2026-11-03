Option Compare Database
Option Explicit

' Exam Generator (Access edition) - the exam pages of build phases 7 and 8:
' My reviews (pgReview), the Exams page (pgExams: every exam, where it
' stands, Download a copy, and the DH's New exam), and the dialogs
' dlgNewExam and dlgSwap. The Approvals page's Exams tab is built with its
' Questions tab (modBuildQuestions.Build_pgApprovals). modBuild.BuildPages
' calls BuildExamGrids first, then the pages; BuildUserInterface builds the
' dialogs.

Private Const REV_LIST_W As Long = 6600         ' My reviews' list card, 440 px
Private Const EXAM_DETAIL_W As Long = 6300      ' the Exams page's detail card, 420 px

Private Function ExamListW() As Long
    ExamListW = ContentW() - EXAM_DETAIL_W - px(12)
End Function

Private Function RevDetailW() As Long
    RevDetailW = ContentW() - REV_LIST_W - px(12)
End Function

' Every grid the exam pages and dialogs show, before the forms that host them.
Public Sub BuildExamGrids()
    BuildGrid "subGridReviews", ReviewCols(), REV_LIST_W - px(32)
    BuildGrid "subGridExamQs", ExamQCols(), RevDetailW() - px(32)
    BuildGrid "subGridPlan", PlanCols(), px(580)
    BuildGrid "subGridSwap", SwapCols(), px(680)
    BuildGrid "subGridExams", ExamCols(), ExamListW() - px(32)
    BuildGrid "subGridExamTopics", ExamTopicCols(), EXAM_DETAIL_W - px(32)
End Sub

' "field|caption|width px|options[|sort field]"
Private Function ReviewCols() As Variant
    ReviewCols = Array("exam|Exam for|132|bs", "TopicName|Topic|108|s", "since_txt|Since|72|m|CreatedOn", "status|Status|96|p")
End Function

Private Function ExamQCols() As Variant
    ' The note's pill can read "Swapped, given lately"; it is the last column, so it takes the rest of the width.
    ExamQCols = Array("ref|#|60|m", "QuestionText|Question|204|", "LevelName|Difficulty|76|m", "pts|Points|52|r", "note_txt|Note|116|p")
End Function

Private Function PlanCols() As Variant
    PlanCols = Array("TopicName|Topic|200|b", "per_exam|Questions|84|r", "on_hand|Approved|84|r", "fresh|Not given lately|116|r", _
        "state|State|96|p")
End Function

Private Function SwapCols() As Variant
    SwapCols = Array("ref|#|60|m|QuestionKey", "QuestionText|Question|380|s|qsort", "LevelName|Difficulty|80|m", "pts|Points|60|r", _
        "note_txt|Note|100|p")
End Function

Private Function ExamCols() As Variant
    ExamCols = Array("ref|#|60|m|id", "exam|Exam for|150|bs", "TemplateName|Template|110|s", "status|Status|96|p", _
        "step|Waits on|52|m", "created_txt|Created|60|m|CreatedOn")
End Function

Private Function ExamTopicCols() As Variant
    ExamTopicCols = Array("TopicName|Topic|140|b", "qn|Qs|44|r", "status|State|88|p", "signer|Signed off by|116|")
End Function

' ---------------------------------------------------------------------------
' My reviews (page 2)
' ---------------------------------------------------------------------------

' The search box and the Open and Done chips over the person's review tasks
' (one per exam topic), and the picked one on the right: the exam it's on,
' where the topic stands, the questions drawn for it with the picked one in
' full under them, Sign off and Swap.
Public Sub Build_pgReview()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, btnY As Long, lw As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad", "OnTimer"             ' OnTimer: the opening sweep's emails, sent once loaded
    PageHeader "My reviews", "Each exam topic you review: check the questions drawn for it, swap any out and sign off. " & _
        "One SME's sign-off completes the topic.", PAGE_PAD
    SearchBox "txtSearch", PAGE_PAD, px(84), px(260), "SearchFor", "Search...", px(28)
    ChipRow "chip", PAGE_PAD + px(272), px(85), "ChipPick", "Open", "Done"
    y = px(128): h = PAGE_H - y - px(16): lw = REV_LIST_W
    x = PAGE_PAD + lw + px(12): fx = x + px(16): fw = RevDetailW() - px(32): btnY = y + h - px(16) - px(30)

    Card "List", "Exam topics", PAGE_PAD, y, lw, h, , 2, 2
    GridHeaders "hdrList", ReviewCols(), PAGE_PAD + px(16), y + px(40), w:=lw - px(32), sortSf:="sfList", _
        filterSf:="sfList", onFilter:="ListFiltered"
    Set c = SubF("sfList", "subGridReviews", PAGE_PAD + px(16), y + px(62), lw - px(32), h - px(78)): Anchor c, 2, 2
    EmptyState "emptyList", PAGE_PAD + px(16), y + px(90), lw - px(32), "E73E", "Nothing to review", _
        "When the DH generates an exam, each topic you're an SME of shows here.", "Clear search"

    Card "Detail", "Review", x, y, RevDetailW(), h, , 1, 2
    Lbl "lblExam", "", fx, y + px(40), fw, px(24), 11, True
    Lbl "lblExamBy", "", fx, y + px(64), fw, px(32), 9, False, CLR_MUTED_FG
    Alert "alReview", fx, y + px(100), fw, px(56)
    FieldLbl "capQs", "QUESTIONS DRAWN", fx, y + px(164), fw
    GridHeaders "hdrQs", ExamQCols(), fx, y + px(182), w:=fw, sortSf:="sfQs"
    Set c = SubF("sfQs", "subGridExamQs", fx, y + px(204), fw, px(120))
    FieldLbl "capPreview", "THE PICKED QUESTION", fx, y + px(332), fw
    Set c = RO("txtPreview", "", fx, y + px(350), fw, btnY - px(8) - (y + px(350)), 10)
    c.ScrollBars = 2
    Set c = Btn("btnSignOff", "Sign off", fx, btnY, px(96), px(30))
    c.ControlTipText = "The questions are right: complete this topic for every SME on it"
    Set c = Btn("btnSwap", "Swap...", fx + px(104), btnY, px(88), px(30), "outline")
    c.ControlTipText = "Put another approved question of the topic in place of the picked one (or double-click its row)"
    Set c = Lbl("lblDetailEmpty", "Pick an exam topic in the list to see it here.", fx, y + px(120), fw, px(20), 9, False, _
        CLR_MUTED_FG)
    c.TextAlign = 2
    SaveForm "pgReview"
End Sub

' ---------------------------------------------------------------------------
' Exams (page 4)
' ---------------------------------------------------------------------------

' The search box and a chip per status over every exam, newest first, and
' the picked one on the right: where it is in the chain, who it is for and
' who generated it, where it stands (its overall marking once approved, the
' reason when it was sent back), its topics with who signed each off,
' Download a copy and History. The DH's New exam button is top right.
Public Sub Build_pgExams()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, btnY As Long, lw As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad", "OnTimer"             ' OnTimer: the opening sweep's emails, sent once loaded
    PageHeader "Exams", "Every exam and where it stands in review and approval. The department head starts a new exam here; " & _
        "once the SM approves one, anyone can download its Word document.", PAGE_PAD
    Set c = Btn("btnNew", "New exam", PAGE_W - PAGE_PAD - px(120), px(24), px(120), px(32), "outline")
    c.ControlTipText = "Pick a template and who the exam is for; the questions are drawn for you": Anchor c, 1, 0
    SearchBox "txtSearch", PAGE_PAD, px(84), px(260), "SearchFor", "Search...", px(28)
    ChipRow "chip", PAGE_PAD + px(272), px(85), "ChipPick", Split(EXAM_CHIPS, "|")
    y = px(128): h = PAGE_H - y - px(16): lw = ExamListW()
    x = PAGE_PAD + lw + px(12): fx = x + px(16): fw = EXAM_DETAIL_W - px(32): btnY = y + h - px(16) - px(30)

    Card "List", "Exams", PAGE_PAD, y, lw, h, , 2, 2
    GridHeaders "hdrList", ExamCols(), PAGE_PAD + px(16), y + px(40), w:=lw - px(32), sortSf:="sfList", _
        filterSf:="sfList", onFilter:="ListFiltered"
    Set c = SubF("sfList", "subGridExams", PAGE_PAD + px(16), y + px(62), lw - px(32), h - px(78)): Anchor c, 2, 2
    EmptyState "emptyList", PAGE_PAD + px(16), y + px(90), lw - px(32), "E8A5", "No exams yet", _
        "When the DH generates an exam it shows here, from SME review to final approval.", "Show all"

    Card "Detail", "Exam", x, y, EXAM_DETAIL_W, h, , 1, 2
    Stepper "stpExam", fx, y + px(44), fw, 1, "", "SME", "DH", "XO", "SM"
    Lbl "lblExam", "", fx, y + px(104), fw, px(22), 11, True
    Lbl "lblExamBy", "", fx, y + px(126), fw, px(32), 9, False, CLR_MUTED_FG
    Lbl "lblState", "", fx, y + px(160), fw, px(52), 9, False, CLR_FG
    FieldLbl "capTopics", "TOPICS", fx, y + px(218), fw
    Lbl "lblTopics", "", fx, y + px(234), px(140), px(20), 9, False, CLR_MUTED_FG
    Set c = Txt("txtPct", "", fx, y + px(234), px(40), px(20))
    c.Visible = False: c.TabStop = False
    Progress "prgTopics", "[txtPct]", fx + px(148), y + px(241), fw - px(148), px(6)
    GridHeaders "hdrTopics", ExamTopicCols(), fx, y + px(262), w:=fw
    Set c = SubF("sfTopics", "subGridExamTopics", fx, y + px(284), fw, btnY - px(12) - (y + px(284)))
    c.TabStop = False
    Set c = Btn("btnDownload", "Download a copy", fx, btnY, px(136), px(30))
    c.ControlTipText = "Save a copy of the approved exam's Word document"
    Set c = Btn("btnHistory", "History", fx + fw - px(88), btnY, px(88), px(30), "outline")
    c.ControlTipText = "Who generated, signed off, approved, sent back and downloaded it, and when"
    Set c = Lbl("lblDetailEmpty", "Pick an exam in the list to see it here.", fx, y + px(120), fw, px(20), 9, False, CLR_MUTED_FG)
    c.TextAlign = 2
    SaveForm "pgExams"
End Sub

' ---------------------------------------------------------------------------
' Dialogs
' ---------------------------------------------------------------------------

' New exam (the DH): a template and who it is for, the topics it draws from
' with how many approved questions each has, and what stops or warns.
Public Sub Build_dlgNewExam()
    Dim c As Control, w As Long, iw As Long
    w = px(620): iw = w - px(40)
    NewDialog "New exam", w, px(534)
    Ev F, "OnLoad"
    Lbl "lblHead", "New exam", px(20), px(20), iw, px(24), 11, True
    Lbl "lblSub", "Pick a template and type who the exam is for. Its questions are drawn at random from the approved bank.", _
        px(20), px(46), iw, px(20), 9, False, CLR_MUTED_FG
    FieldLbl "capTemplate", "TEMPLATE", px(20), px(84), px(280)
    Set c = Cbo("cboTemplate", "", "SELECT TemplateID, TemplateName FROM tblTemplates WHERE IsActive ORDER BY TemplateName", _
        px(20), px(102), px(280), px(30), 2, "0;" & px(280))
    Ev c, "AfterUpdate"
    FieldLbl "capFor", "FOR (BOAT, SCHOOLHOUSE...)", px(312), px(84), iw - px(292)
    Txt "txtRecipient", "", px(312), px(102), iw - px(292), px(30)
    FieldLbl "capPlan", "TOPICS AND QUESTIONS", px(20), px(146), iw
    GridHeaders "hdrPlan", PlanCols(), px(20), px(164), w:=iw
    Set c = SubF("sfPlan", "subGridPlan", px(20), px(186), iw, px(196))
    c.TabStop = False
    Set c = Lbl("lblPlanEmpty", "Pick a template to see the topics it draws from.", px(20), px(260), iw, px(20), 9, False, CLR_MUTED_FG)
    c.TextAlign = 2
    Alert "alPlan", px(20), px(392), iw, px(72)
    Set c = Btn("btnGenerate", "Generate", w - px(20) - px(104) - px(8) - px(88), px(484), px(104), px(30))
    c.ControlTipText = "Draw the questions and send each topic to its SMEs"
    Set c = Btn("btnCancel", "Cancel", w - px(20) - px(88), px(484), px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgNewExam"
End Sub

' Send an exam back (an approver, from the Approvals page): which topic, or
' every topic, who gets it back, and the reason (required).
Public Sub Build_dlgRejectExam()
    Dim c As Control, w As Long, iw As Long
    w = px(500): iw = w - px(40)
    NewDialog "Send back", w, px(462)
    Ev F, "OnLoad"
    Lbl "lblHead", "", px(20), px(20), iw, px(24), 11, True
    Lbl "lblSub", "", px(20), px(46), iw, px(32), 9, False, CLR_MUTED_FG
    FieldLbl "capTopic", "SEND BACK", px(20), px(88), iw
    Set c = Cbo("cboTopic", "", "", px(20), px(106), iw, px(30), 2, "0;" & iw)
    Ev c, "AfterUpdate"
    FieldLbl "capTo", "GOES BACK TO", px(20), px(146), iw
    RO "txtTo", "", px(20), px(164), iw, px(48), 10
    FieldLbl "capReason", "REASON (REQUIRED)", px(20), px(222), iw
    Set c = Txt("txtReason", "", px(20), px(240), iw, px(128))
    c.EnterKeyBehavior = True: c.ScrollBars = 2
    Lbl "lblHint", "Say what to change. Its SMEs read this under My reviews, and it stays in the exam's history.", px(20), px(372), _
        iw, px(20), 9, False, CLR_MUTED_FG
    Set c = Btn("btnReject", "Send back", w - px(20) - px(104) - px(8) - px(88), px(410), px(104), px(30), "destructive")
    Set c = Btn("btnCancel", "Cancel", w - px(20) - px(88), px(410), px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgRejectExam"
End Sub

' Swap a question on an exam topic: the approved questions of the topic not
' on the exam (searchable), the picked one in full, and the overlap rule's
' word on it.
Public Sub Build_dlgSwap()
    Dim c As Control, w As Long, iw As Long
    w = px(720): iw = w - px(40)
    NewDialog "Swap a question", w, px(574)
    Ev F, "OnLoad"
    Lbl "lblHead", "", px(20), px(20), iw, px(24), 11, True
    Lbl "lblSub", "", px(20), px(46), iw, px(20), 9, False, CLR_MUTED_FG
    SearchBox "txtSearch", px(20), px(80), px(260), "SearchFor", "Search...", px(28)
    GridHeaders "hdrCand", SwapCols(), px(20), px(120), w:=iw, sortSf:="sfCand"
    Set c = SubF("sfCand", "subGridSwap", px(20), px(142), iw, px(180))
    c.TabStop = False
    EmptyState "emptyCand", px(20), px(160), iw, "E82D", "No question to swap in", _
        "Every other approved question of this topic is on the exam already.", "Clear search"
    FieldLbl "capPreview", "THE PICKED QUESTION", px(20), px(330), iw
    Set c = RO("txtPreview", "", px(20), px(348), iw, px(96), 10)
    c.ScrollBars = 2
    Alert "alSwap", px(20), px(452), iw, px(56)
    Set c = Btn("btnSwap", "Swap", w - px(20) - 2 * px(88) - px(8), px(524), px(88), px(30))
    c.ControlTipText = "Put the picked question on the exam in place of the one you chose"
    Set c = Btn("btnCancel", "Cancel", w - px(20) - px(88), px(524), px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgSwap"
End Sub
