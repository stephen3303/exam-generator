Option Compare Database
Option Explicit

' Exam Generator (Access edition) - the exam kickoff and SME review of build
' phase 7: My reviews (pgReview), the Exams page's New exam button (pgExams,
' whose list comes in phase 8), and the dialogs dlgNewExam and dlgSwap.
' modBuild.BuildPages calls BuildExamGrids first, then the pages;
' BuildUserInterface builds the dialogs.

Private Const REV_LIST_W As Long = 6600         ' My reviews' list card, 440 px

Private Function RevDetailW() As Long
    RevDetailW = ContentW() - REV_LIST_W - px(12)
End Function

' Every grid the exam pages and dialogs show, before the forms that host them.
Public Sub BuildExamGrids()
    BuildGrid "subGridReviews", ReviewCols(), REV_LIST_W - px(32)
    BuildGrid "subGridExamQs", ExamQCols(), RevDetailW() - px(32)
    BuildGrid "subGridPlan", PlanCols(), px(580)
    BuildGrid "subGridSwap", SwapCols(), px(680)
End Sub

' "field|caption|width px|options[|sort field]"
Private Function ReviewCols() As Variant
    ReviewCols = Array("exam|Exam for|132|bs", "TopicName|Topic|108|s", "since_txt|Since|72|m|CreatedOn", "status|Status|96|p")
End Function

Private Function ExamQCols() As Variant
    ' The note's pill can read "Swapped, given lately"; it is the last column, so it takes the rest of the width.
    ExamQCols = Array("ref|#|60|m", "QuestionText|Question|204|", "LevelName|Difficulty|76|m", "pts|Points|52|r", "note|Note|116|p")
End Function

Private Function PlanCols() As Variant
    PlanCols = Array("TopicName|Topic|200|b", "per_exam|Questions|84|r", "on_hand|Approved|84|r", "fresh|Not given lately|116|r", _
        "state|State|96|p")
End Function

Private Function SwapCols() As Variant
    SwapCols = Array("ref|#|60|m|QuestionKey", "QuestionText|Question|380|s|qsort", "LevelName|Difficulty|80|m", "pts|Points|60|r", _
        "note|Note|100|p")
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
    Ev F, "OnLoad"
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
' Exams (page 4): the list comes in phase 8
' ---------------------------------------------------------------------------

' The DH's New exam button over a placeholder for the list of exams.
Public Sub Build_pgExams()
    Dim c As Control
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Exams", "Every exam and where it stands in review and approval. The department head starts a new exam here.", _
        PAGE_PAD
    Set c = Btn("btnNew", "New exam", PAGE_W - PAGE_PAD - px(120), px(24), px(120), px(32), "outline")
    c.ControlTipText = "Pick a template and who the exam is for; the questions are drawn for you": Anchor c, 1, 0
    EmptyState "emptyStub", PAGE_PAD, px(140), ContentW(), "E8A5", "The list of exams comes in build phase 8", _
        "New exam generates one now. Each topic's SMEs then find it under My reviews.", shown:=True
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
