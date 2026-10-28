Option Compare Database
Option Explicit

' Exam Generator (Access edition) - the admin pages of build phase 4:
' People and routing (pgAdmPeople), Rules and lists (pgAdmRules), Topics
' (pgAdmTopics, with its detail form subTopic) and Templates
' (pgAdmTemplates, with subTemplate and the dialog dlgTemplateTopic).
' modBuild.BuildPages calls BuildAdminGrids first, then the pages.
'
' The list pages follow the plan's pattern: the list card on the left and a
' 360px detail card on the right, both from y 124px under the search box.
'
' Owner: the admin pages (build phase 4). Depends on: modAdmin, modBuild;
' the kit's modMenu, modPersonnelUI, modUIKit.

Private Const DETAIL_W As Long = 5400            ' the detail card, 360 px
Private Const LIST_Y As Long = 1860              ' the list and detail cards' top, 124 px

' The list card's width beside the detail card.
Private Function ListW() As Long
    ListW = ContentW() - DETAIL_W - px(12)
End Function

' The detail card's left edge.
Private Function DetailX() As Long
    DetailX = PAGE_PAD + ListW() + px(12)
End Function

' Every grid the admin pages show, before the pages that host them.
Public Sub BuildAdminGrids()
    BuildGrid "subGridPeople", PeopleCols(), ListW() - px(32)
    BuildGrid "subGridTopics", TopicCols(), ListW() - px(32)
    BuildGrid "subGridSMEs", SMECols(), DETAIL_W - px(32)
    BuildGrid "subGridTemplates", TemplateCols(), ListW() - px(32)
    BuildGrid "subGridTplTopics", TplTopicCols(), DETAIL_W - px(32)
    BuildGrid "subGridDiff", DiffCols(), RulesRightW() - px(32)
    BuildGrid "subGridClass", ClassCols(), RulesRightW() - px(32)
    Build_subTopic
    Build_subTemplate
End Sub

' ---------------------------------------------------------------------------
' People and routing (page 10)
' ---------------------------------------------------------------------------

' Who holds each role. Three cards across the top name who each approval
' step goes to (DH, XO, SM). Under them a person from the Personnel file
' is added with the picker, the people with Exam Generator access are
' listed on the left, and the open person's switches (access, admin and the
' four roles) are on the right; each switch saves at once.
Public Sub Build_pgAdmPeople()
    Dim c As Control, i As Long, cw As Long, x As Long, y As Long, h As Long, fx As Long, fw As Long
    Dim codes As Variant, titles As Variant, tg As Variant, caps As Variant, subs As Variant
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Admin: People and routing", "Who holds each Exam Generator role, and who each approval step " & _
        "goes to. A switch saves when you flip it.", PAGE_PAD
    Set c = Btn("btnPersonnel", "Personnel", PAGE_W - PAGE_PAD - px(112), px(24), px(112), px(32), "outline")
    c.ControlTipText = "Add people to the Personnel file or change their names and ranks": Anchor c, 1, 0
    ' The approval route.
    codes = Array("DH", "XO", "SM")
    titles = Array("Department head (DH)", "Executive officer (XO)", "Senior member (SM)")
    cw = (ContentW() - 2 * px(12)) \ 3
    For i = 0 To 2
        x = PAGE_PAD + i * (cw + px(12))
        Card "Route" & codes(i), "Step " & (i + 1) & ": " & titles(i), x, px(84), cw, px(72)
        Lbl "lblRoute" & codes(i), "", x + px(16), px(116), cw - px(32), px(32), 9, False, CLR_MUTED_FG
    Next
    ' Add a person.
    y = px(168)
    Set c = PersonPickBox("txtPerson", "", PAGE_PAD, y, px(300), px(30))
    Ev c, "OnKeyDown"
    Set c = Btn("btnAdd", "Add person", PAGE_PAD + px(308), y, px(112), px(30), "outline")
    c.ControlTipText = "Give the picked person access to the Exam Generator and list them here"
    Lbl "lblAddHint", "Only people in the Personnel file can be picked. Add anyone missing there first.", _
        PAGE_PAD + px(432), y, ListW() - px(432), px(30), 9, False, CLR_MUTED_FG
    ' The list.
    y = px(212): h = PAGE_H - y - px(16)
    Card "List", "People", PAGE_PAD, y, ListW(), h, , 2, 2
    GridHeaders "hdrPeople", PeopleCols(), PAGE_PAD + px(16), y + px(40), w:=ListW() - px(32), sortSf:="sfList"
    Set c = SubF("sfList", "subGridPeople", PAGE_PAD + px(16), y + px(62), ListW() - px(32), h - px(78)): Anchor c, 2, 2
    EmptyState "emptyList", PAGE_PAD + px(16), y + px(90), ListW() - px(32), "E716", "Nobody here yet", _
        "Pick someone from the Personnel file in the box above, then Add person."
    ' The open person.
    x = DetailX(): fx = x + px(16): fw = DETAIL_W - px(32)
    Card "Detail", "Person", x, y, DETAIL_W, h, , 1, 2
    Lbl "lblName", "", fx, y + px(44), fw, px(24), 11, True
    Lbl "lblWho", "", fx, y + px(68), fw, px(18), 9, False, CLR_MUTED_FG
    tg = Array("tgActive", "tgAdmin", "tgSME", "tgDH", "tgXO", "tgSM")
    caps = Array("Can use the Exam Generator", "Exam Generator admin", "Subject matter expert (SME)", _
        "Department head (DH)", "Executive officer (XO)", "Senior member (SM)")
    subs = Array("ActiveChanged", "AdminChanged", "SMEChanged", "DHChanged", "XOChanged", "SMChanged")
    For i = 0 To 5
        Toggle CStr(tg(i)), "", fx + px(4), y + px(104) + i * px(36), px(36), px(20), CStr(subs(i))
        Link "lbl" & tg(i), CStr(caps(i)), fx + px(52), y + px(100) + i * px(36), fw - px(52), px(28), _
            "=ToggleSwitch([Form],""" & tg(i) & """)", 9, CLR_FG
    Next
    FieldLbl "lblSMEFor", "SME FOR", fx, y + px(324), fw
    Lbl "lblTopics", "", fx, y + px(342), fw, px(52), 9, False, CLR_MUTED_FG
    Set c = Lbl("lblDetailEmpty", "Pick someone in the list to see their roles.", fx, y + px(120), fw, px(20), _
        9, False, CLR_MUTED_FG)
    c.TextAlign = 2
    SaveForm "pgAdmPeople"
End Sub

' The people list: "field|caption|width px|options".
Private Function PeopleCols() As Variant
    PeopleCols = Array("person|Name|220|b", "roles|Roles|280|m", "state|Status|88|p")
End Function

' ---------------------------------------------------------------------------
' Rules and lists (page 11)
' ---------------------------------------------------------------------------

' The right-hand column's width (the two list cards).
Private Function RulesRightW() As Long
    RulesRightW = ContentW() - px(480) - px(12)
End Function

' The overlap rule on the left: two sliders and the default question count,
' with an alert listing the topics short of approved questions under the
' values as they are set, and Save and Undo changes. On the right the two
' lookup lists, difficulty levels and classification levels, each with Add
' and More (rename, move, turn off, delete) and the same menu on a
' right-click.
Public Sub Build_pgAdmRules()
    Dim c As Control, x As Long, y As Long, w As Long, h As Long, fx As Long, fw As Long, rw As Long, ch As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad", "OnKeyDown"
    F.KeyPreview = True                     ' the sliders' arrow keys (SliderStep)
    PageHeader "Admin: Rules and lists", "How far a new exam may repeat recent exams, and the lists questions " & _
        "pick their difficulty and classification from.", PAGE_PAD
    ' The overlap rule.
    y = px(84): w = px(480): h = PAGE_H - y - px(16)
    fx = PAGE_PAD + px(16): fw = w - px(32)
    Card "Rule", "Overlap rule", PAGE_PAD, y, w, h, , 0, 2
    Lbl "lblRuleIntro", "A new exam may reuse questions given in recent exams, up to a limit. These settings set " & _
        "that limit for every topic.", fx, y + px(40), fw, px(36), 9, False, CLR_MUTED_FG
    FieldLbl "lblPct", "MOST OF A TOPIC'S QUESTIONS THAT MAY REPEAT", fx, y + px(88), fw
    Slider "sldPct", "", fx, y + px(108), fw, 0, 100, 5, "RuleChanged", True, "0""%"""   ' a quoted %: Access drops the backslash of 0\%, and a bare % multiplies by 100
    Lbl "lblPctHint", "0% means every question must be fresh; 100% means repeats are never limited.", _
        fx, y + px(132), fw, px(18), 9, False, CLR_MUTED_FG
    FieldLbl "lblWindow", "RECENT EXAMS CHECKED", fx, y + px(168), fw
    Slider "sldWindow", "", fx, y + px(188), fw, 1, 40, 1, "RuleChanged", True, "0"
    Lbl "lblWindowHint", "A question counts as a repeat if it was on any of this many of the latest exams.", _
        fx, y + px(212), fw, px(18), 9, False, CLR_MUTED_FG
    FieldLbl "lblDefault", "QUESTIONS PER TOPIC TO START WITH", fx, y + px(248), fw
    Set c = Txt("txtDefault", "", fx, y + px(266), px(80), px(30))
    c.TextAlign = 3
    Ev c, "OnChange"
    Lbl "lblDefaultHint", "What new topics, and topics added to a template, start with (" & _
        MIN_QUESTIONS_PER_TOPIC & " to " & MAX_QUESTIONS_PER_TOPIC & ").", _
        fx + px(92), y + px(264), fw - px(92), px(38), 9, False, CLR_MUTED_FG
    Alert "alShort", fx, y + px(312), fw, px(160), "info", "Topics short of questions"
    Set c = Btn("btnSave", "Save", fx, y + h - px(48), px(88), px(32)): c.ControlTipText = "Save (Ctrl+S)": Anchor c, 0, 1
    Set c = Btn("btnUndo", "Undo changes", fx + px(96), y + h - px(48), px(120), px(32), "ghost"): Anchor c, 0, 1
    ' The two lists.
    x = PAGE_PAD + w + px(12): rw = RulesRightW()
    ch = (h - px(12)) \ 2
    RulesList "Diff", "Difficulty levels", DiffCols(), "subGridDiff", x, y, rw, ch
    RulesList "Class", "Classification levels", ClassCols(), "subGridClass", x, y + ch + px(12), rw, ch
    SaveForm "pgAdmRules"
End Sub

' One lookup list's card on pgAdmRules: title, Add and More in the header
' row, the column captions and the table sf<nm>.
Private Sub RulesList(ByVal nm As String, ByVal title As String, ByVal cols As Variant, ByVal grid As String, _
        ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim c As Control
    Card nm, title, x, y, w, h, , 2, 0
    Set c = Btn("btn" & nm & "More", "More", x + w - px(16) - px(72), y + px(8), px(72), px(28), "ghost"): Anchor c, 1, 0
    c.ControlTipText = "Rename, move, turn off or delete the selected row (or right-click the row)"
    Set c = Btn("btn" & nm & "Add", "Add", x + w - px(16) - px(152), y + px(8), px(72), px(28), "outline"): Anchor c, 1, 0
    GridHeaders "hdr" & nm, cols, x + px(16), y + px(44), w:=w - px(32)
    Set c = SubF("sf" & nm, grid, x + px(16), y + px(66), w - px(32), h - px(82)): Anchor c, 2, 0
End Sub

Private Function DiffCols() As Variant
    DiffCols = Array("LevelName|Level|340|b", "state|Status|88|p")
End Function

Private Function ClassCols() As Variant
    ClassCols = Array("Marking|Marking|220|b", "PortionMark|Portion mark|120|", "LevelRank|Rank|88|r")
End Function

' ---------------------------------------------------------------------------
' Topics (page 8)
' ---------------------------------------------------------------------------

' The topics on the left, with their SMEs, approved questions and status
' (Active, Short of approved questions for the overlap rule, or Off). On
' the right the open topic (subTopic: name, questions per exam, active,
' description), its SMEs with a picker to add one, and an alert with its
' approved questions against what the rule needs.
Public Sub Build_pgAdmTopics()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, sy As Long, ay As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Admin: Topics", "The subjects questions are written in, how many questions each gets on an exam " & _
        "and who reviews them.", PAGE_PAD
    Set c = Btn("btnNew", "New topic", PAGE_W - PAGE_PAD - px(112), px(24), px(112), px(32), "outline")
    c.ControlTipText = "New topic (Ctrl+N)": Anchor c, 1, 0
    SearchBox "txtSearch", PAGE_PAD, px(84), px(300), "SearchFor", "Search topics...", px(28)
    y = LIST_Y: h = PAGE_H - y - px(16)
    ListCard "Topics", TopicCols(), "subGridTopics", y, h, "E8EC", "No topics match", _
        "No topic's name or description has those words."
    ' The open topic.
    x = DetailX(): fx = x + px(16): fw = DETAIL_W - px(32)
    Card "Detail", "Topic", x, y, DETAIL_W, h, , 1, 2
    Set c = Btn("btnMore", "More", x + DETAIL_W - px(16) - px(72), y + px(8), px(72), px(28), "ghost"): Anchor c, 1, 0
    c.ControlTipText = "Turn the topic off or on, or delete it (or right-click its row)"
    Set c = SubF("sfDetail", "", fx, y + px(40), fw, px(206)): Anchor c, 1, 0
    c.Visible = False
    Set c = Lbl("lblDetailEmpty", "Pick a topic in the list to see it here.", fx, y + px(120), fw, px(20), 9, False, CLR_MUTED_FG)
    c.TextAlign = 2
    ' Its SMEs.
    sy = y + px(258): ay = y + h - px(16) - px(56)
    FieldLbl "lblSMEs", "SMES (THEY REVIEW THIS TOPIC'S QUESTIONS AND EXAMS)", fx, sy, fw
    Set c = PersonPickBox("txtSME", "", fx, sy + px(18), px(184), px(30))
    Ev c, "OnKeyDown"
    Set c = Btn("btnAddSME", "Add", fx + px(192), sy + px(18), px(64), px(30), "outline")
    c.ControlTipText = "Make the picked person an SME of this topic"
    Set c = Btn("btnRemoveSME", "Remove", fx + px(264), sy + px(18), px(64), px(30), "ghost")
    c.ControlTipText = "Take the selected SME off this topic"
    SubF "sfSMEs", "subGridSMEs", fx, sy + px(56), fw, ay - px(8) - (sy + px(56))
    Alert "alSupply", fx, ay, fw, px(56)
    SaveForm "pgAdmTopics"
End Sub

' A list page's list card: title, captions, the table sfList (sortable) and
' its empty state, with a Clear search button.
Private Sub ListCard(ByVal title As String, ByVal cols As Variant, ByVal grid As String, ByVal y As Long, ByVal h As Long, _
        ByVal glyph As String, ByVal emptyTitle As String, ByVal emptyText As String)
    Dim c As Control
    Card "List", title, PAGE_PAD, y, ListW(), h, , 2, 2
    GridHeaders "hdrList", cols, PAGE_PAD + px(16), y + px(40), w:=ListW() - px(32), sortSf:="sfList"
    Set c = SubF("sfList", grid, PAGE_PAD + px(16), y + px(62), ListW() - px(32), h - px(78)): Anchor c, 2, 2
    EmptyState "emptyList", PAGE_PAD + px(16), y + px(90), ListW() - px(32), glyph, emptyTitle, emptyText, "Clear search"
End Sub

Private Function TopicCols() As Variant
    TopicCols = Array("TopicName|Topic|220|bs", "smes|SMEs|170|m", "on_hand|Approved|90|r", "state|Status|88|p")
End Function

Private Function SMECols() As Variant
    SMECols = Array("person|SME|200|", "username|Sign-in|128|m")
End Function

' The open topic on pgAdmTopics, bound to tblTopics (code in subTopic.vba):
' its name, questions per exam, the Active switch and its description, with
' Save. 328 x 206px.
Private Sub Build_subTopic()
    Dim c As Control
    NewForm "tblTopics", 0, px(328), px(206)
    Ev F, "BeforeUpdate", "AfterUpdate", "OnError"
    F.Section(acDetail).BackColor = CLR_SURFACE
    F.Section(acDetail).AlternateBackColor = CLR_SURFACE
    FieldLbl "lblName", "TOPIC", 0, 0, px(328)
    Txt "txtName", "TopicName", 0, px(18), px(328), px(30)
    FieldLbl "lblCount", "QUESTIONS PER EXAM", 0, px(56), px(156)
    Set c = Txt("txtCount", "DefaultCount", 0, px(74), px(156), px(30))
    c.TextAlign = 3
    c.ControlTipText = "What this topic starts with when it is added to a template"
    FieldLbl "lblActive", "ACTIVE", px(172), px(56), px(156)
    Toggle "tglActive", "IsActive", px(176), px(79), px(36), px(20), "ActiveChanged"
    Link "lblActiveText", "In use", px(220), px(75), px(108), px(28), "=ToggleSwitch([Form],""tglActive"")", 9, CLR_FG
    FieldLbl "lblDesc", "DESCRIPTION", 0, px(112), px(328)
    Set c = Txt("txtDesc", "Description", 0, px(130), px(328), px(40))
    c.EnterKeyBehavior = True
    c.ScrollBars = 2
    Set c = Btn("btnSave", "Save", px(232), px(178), px(96), px(28))
    c.ControlTipText = "Save (Ctrl+S)"
    SaveForm "subTopic"
End Sub

' ---------------------------------------------------------------------------
' Templates (page 9)
' ---------------------------------------------------------------------------

' The templates on the left, with how many topics and questions each
' gives. On the right the open template (subTemplate: name, active,
' description) and its topics in exam order with their question counts:
' Add topic and Change open dlgTemplateTopic; the arrows move the selected
' topic; Remove takes it off.
Public Sub Build_pgAdmTemplates()
    Dim c As Control, x As Long, y As Long, h As Long, fx As Long, fw As Long, ty As Long, btnY As Long
    NewForm "", 0, PAGE_W, PAGE_H
    Ev F, "OnLoad"
    PageHeader "Admin: Templates", "Each kind of exam: the topics it covers, in order, and how many questions " & _
        "each gives.", PAGE_PAD
    Set c = Btn("btnNew", "New template", PAGE_W - PAGE_PAD - px(128), px(24), px(128), px(32), "outline")
    c.ControlTipText = "New template (Ctrl+N)": Anchor c, 1, 0
    SearchBox "txtSearch", PAGE_PAD, px(84), px(300), "SearchFor", "Search templates...", px(28)
    y = LIST_Y: h = PAGE_H - y - px(16)
    ListCard "Templates", TemplateCols(), "subGridTemplates", y, h, "E8F1", "No templates match", _
        "No template's name or description has those words."
    ' The open template.
    x = DetailX(): fx = x + px(16): fw = DETAIL_W - px(32)
    Card "Detail", "Template", x, y, DETAIL_W, h, , 1, 2
    Set c = Btn("btnMore", "More", x + DETAIL_W - px(16) - px(72), y + px(8), px(72), px(28), "ghost"): Anchor c, 1, 0
    c.ControlTipText = "Copy the template, or delete it (or right-click its row)"
    Set c = SubF("sfDetail", "", fx, y + px(40), fw, px(150)): Anchor c, 1, 0
    c.Visible = False
    Set c = Lbl("lblDetailEmpty", "Pick a template in the list to see it here.", fx, y + px(120), fw, px(20), _
        9, False, CLR_MUTED_FG)
    c.TextAlign = 2
    ' Its topics.
    ty = y + px(202): btnY = y + h - px(16) - px(28)
    FieldLbl "lblTopics", "TOPICS, IN EXAM ORDER", fx, ty, px(180)
    Set c = Lbl("lblTotal", "", fx + px(180), ty, fw - px(180), px(16), 9, False, CLR_MUTED_FG)
    c.TextAlign = 3
    GridHeaders "hdrTopics", TplTopicCols(), fx, ty + px(22), w:=fw
    SubF "sfTopics", "subGridTplTopics", fx, ty + px(44), fw, btnY - px(8) - (ty + px(44))
    Set c = Btn("btnAddTopic", "Add topic", fx, btnY, px(96), px(28), "outline")
    c.ControlTipText = "Add a topic to this template"
    Set c = Btn("btnChange", "Change", fx + px(104), btnY, px(72), px(28), "ghost")
    c.ControlTipText = "Change the selected topic's question count (or double-click it)"
    GlyphBtn "btnUp", &HE70E, "Move the selected topic up", fx + px(184), btnY
    GlyphBtn "btnDown", &HE70D, "Move the selected topic down", fx + px(220), btnY
    Set c = Btn("btnRemove", "Remove", fx + px(256), btnY, px(72), px(28), "ghost")
    c.ControlTipText = "Take the selected topic off this template"
    SaveForm "pgAdmTemplates"
End Sub

' A 32 x 28px ghost button showing one Segoe MDL2 glyph.
Private Sub GlyphBtn(ByVal nm As String, ByVal glyph As Long, ByVal tip As String, ByVal x As Long, ByVal y As Long)
    Dim c As Control
    Set c = Btn(nm, ChrW(glyph), x, y, px(32), px(28), "ghost")
    c.FontName = "Segoe MDL2 Assets"
    c.ControlTipText = tip
End Sub

Private Function TemplateCols() As Variant
    TemplateCols = Array("TemplateName|Template|260|bs", "topics|Topics|90|r", "questions|Questions|100|r", "state|Status|88|p")
End Function

Private Function TplTopicCols() As Variant
    TplTopicCols = Array("TopicName|Topic|220|", "QuestionCount|Questions|108|r")
End Function

' The open template on pgAdmTemplates, bound to tblTemplates (code in
' subTemplate.vba): its name, the Active switch and its description, with
' Save. 328 x 150px.
Private Sub Build_subTemplate()
    Dim c As Control
    NewForm "tblTemplates", 0, px(328), px(150)
    Ev F, "BeforeUpdate", "AfterUpdate", "OnError"
    F.Section(acDetail).BackColor = CLR_SURFACE
    F.Section(acDetail).AlternateBackColor = CLR_SURFACE
    FieldLbl "lblName", "TEMPLATE", 0, 0, px(240)
    Txt "txtName", "TemplateName", 0, px(18), px(240), px(30)
    FieldLbl "lblActive", "ACTIVE", px(256), 0, px(72)
    Toggle "tglActive", "IsActive", px(260), px(23), px(36), px(20), "ActiveChanged"
    FieldLbl "lblDesc", "DESCRIPTION", 0, px(56), px(328)
    Set c = Txt("txtDesc", "Description", 0, px(74), px(328), px(40))
    c.EnterKeyBehavior = True
    c.ScrollBars = 2
    Set c = Btn("btnSave", "Save", px(232), px(122), px(96), px(28))
    c.ControlTipText = "Save (Ctrl+S)"
    SaveForm "subTemplate"
End Sub

' Add a topic to a template, or change a topic's question count on it
' (code in dlgTemplateTopic.vba): the topic picker (locked when changing),
' a slider for the count, a line on what the overlap rule needs for that
' count, and OK and Cancel.
Public Sub Build_dlgTemplateTopic()
    Dim c As Control, w As Long
    w = px(420)
    NewDialog "Add a topic", w, px(282)
    Ev F, "OnLoad", "OnTimer", "OnKeyDown"
    F.KeyPreview = True                     ' the slider's arrow keys (SliderStep)
    Lbl "lblHead", "", px(20), px(20), w - px(40), px(24), 11, True
    FieldLbl "lblTopic", "TOPIC", px(20), px(60), w - px(40)
    Set c = PickBox("txtTopic", "", "SELECT TopicID, TopicName, Nz(DefaultCount," & DEFAULT_QUESTION_COUNT & ") & ' per exam' AS per_exam " & _
        "FROM tblTopics WHERE IsActive", px(20), px(78), w - px(40), px(30), "Type part of a topic's name")
    Ev c, "OnKeyDown"
    FieldLbl "lblCount", "QUESTIONS ON EACH EXAM", px(20), px(124), w - px(40)
    Slider "sldCount", "", px(20), px(144), w - px(40), MIN_QUESTIONS_PER_TOPIC, MAX_QUESTIONS_PER_TOPIC, 1, "CountChanged", True, "0"
    Lbl "lblNeed", "", px(20), px(172), w - px(40), px(36), 9, False, CLR_MUTED_FG
    Set c = Btn("btnOK", "Add", w - px(20) - 2 * px(88) - px(8), px(228), px(88), px(30))
    c.Default = True
    Set c = Btn("btnCancel", "Cancel", w - px(20) - px(88), px(228), px(88), px(30), "outline")
    c.Cancel = True
    SaveForm "dlgTemplateTopic"
End Sub
