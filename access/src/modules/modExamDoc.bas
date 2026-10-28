Option Compare Database
Option Explicit

' Exam Generator (Access edition) - the exam's Word document, from build
' phase 8. Made when the SM gives an exam final approval (the Approvals
' page calls MakeExamDocument) and kept in the "Exam documents" folder next
' to the shared data file, so anyone can download a copy (Exams page,
' modExamApprovals.DownloadExam). tblExams.OutputFile holds the file's
' name only, found in ExamDocFolder, since each PC may reach the shared
' folder by its own path; a missing file is made again from the exam
' (EnsureExamDocument), as is one made before the admin last reordered the
' classification levels. Every document made writes an audit row.
'
' The document holds a cover page, then the questions in the exam's order,
' numbered, each part and answer option indented under its question, with
' points and portion markings. It never holds an answer: no correct option,
' no reference answer. No topic headings either, so every line on the
' question pages is a marked portion.
'
' Markings (design section 8): each question and part starts with its
' portion marking ("(C)"); options carry their question's or part's. Each
' page's header and footer carry the highest marking on that page, and the
' cover the exam's overall marking (the highest on the exam). Page breaks
' are only known once Word has laid the document out, so the code lays it
' out first and finds the pages each portion falls on. The header's and
' footer's marking line is then one IF field per page ({ IF { PAGE } = 3
' "SECRET" "" }), so Word shows each page its own marking. That changes
' nothing in the body, so the layout doesn't move. If a page can't be read,
' Word fails while the fields go in, or the fields moved a portion to
' another page, every page carries the overall marking instead: never lower
' than its content.
'
' Word is driven by late binding (no reference to set), hidden, and quit
' when done, also on an error.

Private Const DOC_FOLDER As String = "Exam documents"

' Word's constants (late binding).
Private Const wdAlignParagraphLeft As Long = 0
Private Const wdAlignParagraphCenter As Long = 1
Private Const wdActiveEndAdjustedPageNumber As Long = 1
Private Const wdStatisticPages As Long = 2
Private Const wdPrintView As Long = 3
Private Const wdCollapseStart As Long = 1
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdFieldEmpty As Long = -1
Private Const wdFieldPage As Long = 33
Private Const wdFieldIf As Long = 7
Private Const wdFieldNumPages As Long = 26
Private Const wdFormatXMLDocument As Long = 12
Private Const wdDoNotSaveChanges As Long = 0
Private Const wdStyleNormal As Long = -1
Private Const wdAlertsNone As Long = 0

Private m_doc As Object             ' the document being made
Private m_used As Boolean           ' its first (empty) paragraph has been used
' The marked portions. While the document is written each is kept as its
' paragraph's first and last character positions, numbers that later
' paragraphs, added after them, can't change; the body isn't changed after
' that, so StampMarkings reads their pages from them.
Private m_pS() As Long              ' each portion's paragraph: its Start
Private m_pE() As Long              '   its End
Private m_pRank() As Long           '   and its LevelRank
Private m_pN As Long

' ---------------------------------------------------------------------------
' Where documents are kept
' ---------------------------------------------------------------------------

' The folder the exam documents are kept in: "Exam documents" beside the
' shared data file (beside this front end when the data is local). Made
' when missing. Errors: the folder can't be made (no right to the share).
Public Function ExamDocFolder() As String
    Dim be As String, p As String
    be = KitBackendPath("tblExams")
    If Len(be) > 0 Then p = Left$(be, InStrRev(be, "\") - 1) Else p = CurrentProject.Path
    p = p & "\" & DOC_FOLDER
    If Len(Dir(p, vbDirectory)) = 0 Then MkDir p
    ExamDocFolder = p
End Function

' The path of exam examId's document in that folder: "E12 v2 USS Example.docx".
Private Function DocPathFor(ByVal examId As Long) As String
    Dim rs As DAO.Recordset, nm As String
    Set rs = CurrentDb.OpenRecordset("SELECT VersionNo, RecipientName FROM tblExams WHERE ExamID=" & examId, dbOpenSnapshot)
    If Not rs.EOF Then nm = ExamRef(examId, rs!VersionNo) & " " & Nz(rs!RecipientName, "")
    rs.Close
    DocPathFor = ExamDocFolder() & "\" & KitSafeFileName(Left$(nm, 120)) & ".docx"
End Function

' The stored document's full path in ExamDocFolder, or "" when none was
' made. Older rows may hold a whole path; only its file name is used.
Private Function StoredDocPath(ByVal examId As Long) As String
    Dim p As String
    p = Nz(DLookup("OutputFile", "tblExams", "ExamID=" & examId), "")
    If Len(p) > 0 Then StoredDocPath = ExamDocFolder() & "\" & Mid$(p, InStrRev(p, "\") + 1)
End Function

' The name a downloaded copy is offered under: "Exam E12 v2 USS Example.docx".
Public Function ExamDocFileName(ByVal examId As Long) As String
    Dim p As String
    p = Nz(DLookup("OutputFile", "tblExams", "ExamID=" & examId), "")
    If Len(p) > 0 Then p = Mid$(p, InStrRev(p, "\") + 1) Else p = "Exam E" & examId & ".docx"
    If Left$(p, 5) <> "Exam " Then p = "Exam " & p
    ExamDocFileName = p
End Function

' The path of exam examId's document, made first when it is missing (an
' approval whose document failed, or a file someone moved), or once after
' the admin reordered the classification levels: the stored file was
' marked under the old order and may mark a page lower than its content.
' Returns "" and the path in docPath, or why there is none; note as
' MakeExamDocument's when it was made now.
Public Function EnsureExamDocument(ByVal examId As Long, ByRef docPath As String, Optional ByRef note As String) As String
    Dim p As String
    docPath = "": note = ""
    p = StoredDocPath(examId)
    If Len(p) > 0 Then
        If Len(Dir(p)) > 0 And Not MadeBeforeReorder(examId) Then docPath = p: Exit Function
    End If
    EnsureExamDocument = MakeExamDocumentNow(examId, note)
    If Len(EnsureExamDocument) = 0 Then docPath = StoredDocPath(examId)
End Function

' True when exam examId's stored document was made before the admin last
' reordered the classification levels: its newest "Document made" audit
' row is no newer than the audit row that was newest then (setting
' ClassOrderChangedAt, written by Admin: Rules). Audit ids, not times, so
' no PC's clock comes into it. No such row: made before that row was kept,
' so True.
Private Function MadeBeforeReorder(ByVal examId As Long) As Boolean
    Dim stamp As String, made As Variant
    stamp = SettingText(SET_CLASS_ORDER_AT)
    If Not IsNumeric(stamp) Then Exit Function              ' never reordered
    made = DMax("AuditID", "tblAuditTrail", "ItemType=" & Q(ITEM_EXAM) & " AND ItemID=" & examId & " AND ActionType=" & Q(AU_DOC_MADE))
    If IsNull(made) Then MadeBeforeReorder = True Else MadeBeforeReorder = (made <= CLng(stamp))
End Function

' Make exam examId's document now, under the busy veil: after its final
' approval (pgApprovals), and from EnsureExamDocument. Returns "" or why it
' wasn't made; note as MakeExamDocument's. An unexpected error (Word
' stopping halfway, a folder it can't write to) is logged and returned as a
' sentence instead of raised, since the exam's approval stands either way.
Public Function MakeExamDocumentNow(ByVal examId As Long, ByRef note As String) As String
    Dim desc As String
    note = ""
    KitBusy "Making the exam's Word document..."
    On Error GoTo Fail
    MakeExamDocumentNow = MakeExamDocument(examId, note)
    KitBusyEnd
    Exit Function
Fail:
    desc = Err.Description
    LogError "modExamDoc.MakeExamDocumentNow", "Exam " & examId
    KitBusyEnd
    MakeExamDocumentNow = "The Word document couldn't be made (" & desc & "). Download a copy on the Exams page tries again; " & _
        "the error log has the details."
End Function

' ---------------------------------------------------------------------------
' Making the document
' ---------------------------------------------------------------------------

' Make approved exam examId's Word document in ExamDocFolder, replacing
' any made before, and record its path (tblExams.OutputFile). Returns "", or
' why it can't be made (not approved, a question with no portion marking,
' Word missing). note returns a sentence when the pages had to carry the
' overall marking (see the top of the module), else "".
' Errors: Word's and the file system's pass up, after Word is closed.
Public Function MakeExamDocument(ByVal examId As Long, Optional ByRef note As String) As String
    Dim wd As Object, path As String, st As String
    Dim overall As Long, num As Long, desc As String, src As String, lvl As Long, cur As Long, sameOrder As Boolean
    note = ""
    st = ExamStatusName(DLookup("StatusID", "tblExams", "ExamID=" & examId))
    If st <> ES_APPROVED Then MakeExamDocument = DownloadProblem(examId): Exit Function
    MakeExamDocument = UnmarkedProblem(examId)
    If Len(MakeExamDocument) > 0 Then Exit Function
    ' The marking the SM approved (stored then), or the exam's highest level
    ' under today's order when that ranks higher. When the admin reordered the
    ' levels since, every page carries the higher of the two: the stored one
    ' alone could mark a page lower than its content under today's order, and
    ' a page must never be. The remake isn't refused, since the exam stays
    ' approved and its copies are still needed.
    lvl = Nz(DLookup("OverallClassLevelID", "tblExams", "ExamID=" & examId), 0)
    cur = OverallLevelID(examId)
    sameOrder = (lvl = 0 Or lvl = cur)
    overall = ClassRank(cur)
    If lvl <> 0 Then
        If ClassRank(lvl) > overall Then overall = ClassRank(lvl)
    End If
    path = DocPathFor(examId)
    On Error Resume Next
    If Len(Dir(path)) > 0 Then Kill path
    If Err.Number <> 0 Then
        On Error GoTo 0
        MakeExamDocument = "The document made before couldn't be replaced; someone may have it open (" & path & "). " & _
            "Close it and try again."
        Exit Function
    End If
    Set wd = CreateObject("Word.Application")
    If wd Is Nothing Then
        On Error GoTo 0
        MakeExamDocument = "Word couldn't be started on this computer, so the document wasn't made. Anyone with Word can make " & _
            "it: Download a copy on the Exams page makes it when it is missing."
        Exit Function
    End If
    On Error GoTo Fail
    wd.Visible = False
    wd.DisplayAlerts = wdAlertsNone
    Set m_doc = wd.Documents.Add
    m_used = False
    m_pN = 0
    ReDim m_pS(0 To 63): ReDim m_pE(0 To 63): ReDim m_pRank(0 To 63)
    SetUpPage overall
    WriteCover examId, overall
    WriteQuestions examId
    StampMarkings overall, sameOrder, note
    m_doc.SaveAs2 path, wdFormatXMLDocument
    m_doc.Close wdDoNotSaveChanges
    Set m_doc = Nothing
    wd.Quit wdDoNotSaveChanges
    Set wd = Nothing
    ' Only the file's name: each PC may reach the shared folder by another path.
    KitExecute "UPDATE tblExams SET OutputFile=" & Q(Mid$(path, InStrRev(path, "\") + 1)) & " WHERE ExamID=" & examId
    AddAudit ITEM_EXAM, examId, DLookup("VersionNo", "tblExams", "ExamID=" & examId), AU_DOC_MADE, 0, _
        Mid$(path, InStrRev(path, "\") + 1) & ", marked " & RankMarking(overall) & IIf(Len(note) > 0, ". " & note, "")
    Exit Function
Fail:
    num = Err.Number: desc = Err.Description: src = Err.Source
    On Error Resume Next
    If Not m_doc Is Nothing Then m_doc.Close wdDoNotSaveChanges
    Set m_doc = Nothing
    If Not wd Is Nothing Then wd.Quit wdDoNotSaveChanges
    Set wd = Nothing
    On Error GoTo 0
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modExamDoc.MakeExamDocument", "Exam " & examId
End Function

' Margins, the Normal style, and a one-line header and two-line footer
' (marking; "Page n of m") so the layout has its final height before the
' pages are counted.
Private Sub SetUpPage(ByVal overall As Long)
    Dim sec As Object
    With m_doc.PageSetup
        .TopMargin = 72: .BottomMargin = 72: .LeftMargin = 72: .RightMargin = 72
        .HeaderDistance = 30: .FooterDistance = 30
        .DifferentFirstPageHeaderFooter = False
        .OddAndEvenPagesHeaderFooter = False
    End With
    With m_doc.Styles(wdStyleNormal)
        .Font.Name = "Arial"
        .Font.Size = 11
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 4
    End With
    Set sec = m_doc.Sections(1)
    SetBanner sec, RankMarking(overall)
End Sub

' Put marking in section sec's header and footer, the footer's second line
' "Page n of m".
Private Sub SetBanner(sec As Object, ByVal marking As String)
    Dim h As Object, f As Object, r As Object, ln As Object
    Set h = sec.Headers(wdHeaderFooterPrimary)
    h.Range.Text = marking
    BannerFormat h.Range
    Set f = sec.Footers(wdHeaderFooterPrimary)
    f.Range.Text = marking & vbCr & "Page X of Y"
    BannerFormat f.Range
    Set ln = f.Range.Paragraphs(2).Range
    ln.Font.Bold = False
    ln.Font.Size = 8
    ' The fields replace Y, then X (the later one first, so X stays put).
    Set r = f.Range.Duplicate
    r.SetRange ln.Start + 10, ln.Start + 11
    f.Range.Fields.Add r, wdFieldNumPages
    Set r = f.Range.Duplicate
    r.SetRange ln.Start + 5, ln.Start + 6
    f.Range.Fields.Add r, wdFieldPage
End Sub

Private Sub BannerFormat(r As Object)
    r.Font.Name = "Arial"
    r.Font.Size = 12
    r.Font.Bold = True
    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
    r.ParagraphFormat.SpaceBefore = 0
    r.ParagraphFormat.SpaceAfter = 0
End Sub

' Change only the marking line of a section's header and footer (the
' footer's page line stays).
Private Sub SetMarking(sec As Object, ByVal marking As String)
    Dim r As Object
    Set r = sec.Headers(wdHeaderFooterPrimary).Range.Paragraphs(1).Range
    r.MoveEnd 1, -1                         ' wdCharacter: keep the paragraph mark
    r.Text = marking
    Set r = sec.Footers(wdHeaderFooterPrimary).Range.Paragraphs(1).Range
    r.MoveEnd 1, -1
    r.Text = marking
End Sub

' Add a paragraph at the end of the document with text t (line breaks kept
' inside it) and return its range, formatted: size points, bold, centred or
' not, space before and after (points), left indent (points), and kept with
' the next paragraph or not. Every paragraph is kept whole on one page.
Private Function AddPara(ByVal t As String, Optional ByVal size As Single = 11, Optional ByVal bold As Boolean = False, _
        Optional ByVal centred As Boolean = False, Optional ByVal before As Single = 0, Optional ByVal after As Single = 4, _
        Optional ByVal indent As Single = 0, Optional ByVal keepNext As Boolean = False) As Object
    Dim r As Object
    If m_used Then m_doc.Content.InsertParagraphAfter
    m_used = True
    Set r = m_doc.Paragraphs(m_doc.Paragraphs.Count).Range
    r.InsertBefore OneParagraph(t)
    r.Font.Reset
    r.Font.Size = size
    r.Font.Bold = bold
    With r.ParagraphFormat
        .Alignment = IIf(centred, wdAlignParagraphCenter, wdAlignParagraphLeft)
        .SpaceBefore = before
        .SpaceAfter = after
        .LeftIndent = indent
        .FirstLineIndent = 0
        .KeepTogether = True
        .KeepWithNext = keepNext
        .PageBreakBefore = False
        .WidowControl = True
    End With
    Set AddPara = r
End Function

' t with its line breaks as Word line breaks (Chr(11)), so it stays one
' paragraph, and no other control characters.
Private Function OneParagraph(ByVal t As String) As String
    Dim i As Long, ch As String, s As String
    t = Replace(Replace(Replace(t, vbCrLf, Chr(11)), vbCr, Chr(11)), vbLf, Chr(11))
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        ' AscW is signed: characters from U+8000 up come back negative.
        If (AscW(ch) And &HFFFF&) < 32 And ch <> Chr(11) And ch <> vbTab Then ch = " "
        s = s & ch
    Next
    OneParagraph = s
End Function

' Remember paragraph r, the document's last, as a portion marked at rank
' rk: where it starts and ends, as numbers (see m_pS).
Private Sub AddPortion(r As Object, ByVal rk As Long)
    If m_pN > UBound(m_pS) Then
        ReDim Preserve m_pS(0 To 2 * m_pN): ReDim Preserve m_pE(0 To 2 * m_pN): ReDim Preserve m_pRank(0 To 2 * m_pN)
    End If
    m_pS(m_pN) = r.Start
    m_pE(m_pN) = r.End
    m_pRank(m_pN) = rk
    m_pN = m_pN + 1
End Sub

' The cover page: what the exam is and who it is for. Its banner carries
' the overall marking.
Private Sub WriteCover(ByVal examId As Long, ByVal overall As Long)
    Dim rs As DAO.Recordset, nQ As Long, pts As Variant
    Set rs = CurrentDb.OpenRecordset("SELECT e.RecipientName, e.VersionNo, e.FinalApprovedOn, tm.TemplateName FROM tblExams AS e " & _
        "LEFT JOIN tblTemplates AS tm ON e.TemplateID=tm.TemplateID WHERE e.ExamID=" & examId, dbOpenSnapshot)
    nQ = DCount("*", "tblExamQuestions", "ExamID=" & examId)
    ' DSum can't take a join as its domain, so sum with a query.
    With CurrentDb.OpenRecordset("SELECT Sum(q.TotalPoints) AS pts FROM tblExamQuestions AS eq INNER JOIN tblQuestions AS q " & _
        "ON eq.QuestionID=q.QuestionID WHERE eq.ExamID=" & examId, dbOpenSnapshot)
        pts = .Fields("pts").Value
        .Close
    End With
    AddPara "INSPECTION EXAM", 26, True, True, 150, 24
    AddPara "for", 12, False, True, 0, 6
    AddPara Nz(rs!RecipientName, ""), 20, True, True, 0, 30
    If Not IsNull(rs!TemplateName) Then AddPara rs!TemplateName, 12, False, True, 0, 4
    AddPara "Exam " & ExamRef(examId, rs!VersionNo), 11, False, True, 0, 4
    If Not IsNull(rs!FinalApprovedOn) Then AddPara "Approved " & Format$(rs!FinalApprovedOn, "d mmmm yyyy"), 11, False, True, 0, 4
    AddPara nQ & IIf(nQ = 1, " question, ", " questions, ") & PointsText(Nz(pts, 0)) & IIf(Nz(pts, 0) = 1, " point", " points"), _
        11, False, True, 0, 48
    AddPara "Answer every question. The points for each question and part are shown after it in brackets. Each question and " & _
        "part starts with its portion marking.", 10, False, True, 0, 4
    rs.Close
End Sub

' The questions, in the exam's order, numbered from 1, each with its
' answer options and its parts (and theirs), points and portion markings.
' The first starts the second page.
Private Sub WriteQuestions(ByVal examId As Long)
    Dim rs As DAO.Recordset, pr As DAO.Recordset, i As Long, r As Object, first As Boolean, rk As Long
    Set rs = CurrentDb.OpenRecordset("SELECT q.QuestionID, q.QuestionText, q.TotalPoints, c.PortionMark, c.LevelRank " & _
        "FROM (tblExamQuestions AS eq INNER JOIN tblQuestions AS q ON eq.QuestionID=q.QuestionID) " & _
        "LEFT JOIN tblClassLevels AS c ON q.ClassLevelID=c.ClassLevelID WHERE eq.ExamID=" & examId & _
        " ORDER BY eq.SortOrder, eq.ExamQuestionID", dbOpenSnapshot)
    first = True
    Do While Not rs.EOF
        i = i + 1
        rk = Nz(rs!LevelRank, 0)
        Set r = AddPara(i & ". " & Nz(rs!PortionMark, "") & " " & Nz(rs!QuestionText, "") & " (" & PointsLabel(rs!TotalPoints) & ")", _
            11, False, False, IIf(first, 0, 14), 4, 0, True)
        If first Then r.ParagraphFormat.PageBreakBefore = True
        first = False
        AddPortion r, rk
        WriteOptions rs!QuestionID, 0, 36, rk
        Set pr = CurrentDb.OpenRecordset("SELECT pt.PartID, pt.PartLabel, pt.PartText, pt.Points, c.PortionMark, c.LevelRank " & _
            "FROM tblQuestionParts AS pt LEFT JOIN tblClassLevels AS c ON pt.ClassLevelID=c.ClassLevelID WHERE pt.QuestionID=" & _
            rs!QuestionID & " ORDER BY pt.SortOrder, pt.PartID", dbOpenSnapshot)
        Do While Not pr.EOF
            Set r = AddPara(LCase$(Nz(pr!PartLabel, "")) & ". " & Nz(pr!PortionMark, "") & " " & Nz(pr!PartText, "") & " (" & _
                PointsLabel(pr!Points) & ")", 11, False, False, 6, 4, 24, True)
            AddPortion r, Nz(pr!LevelRank, 0)
            WriteOptions rs!QuestionID, pr!PartID, 48, Nz(pr!LevelRank, 0)
            pr.MoveNext
        Loop
        pr.Close
        ' The question's last paragraph lets the next question start a page.
        m_doc.Paragraphs(m_doc.Paragraphs.Count).Range.ParagraphFormat.KeepWithNext = False
        rs.MoveNext
    Loop
    rs.Close
End Sub

' The answer options of question qid's part partId (0: of the question
' itself), indented, as portions of their owner's rank. Never which is
' correct.
Private Sub WriteOptions(ByVal qid As Long, ByVal partId As Long, ByVal indent As Single, ByVal rk As Long)
    Dim rs As DAO.Recordset, r As Object
    Set rs = CurrentDb.OpenRecordset("SELECT OptionLabel, OptionText FROM tblMCOptions WHERE QuestionID=" & qid & _
        IIf(partId = 0, " AND PartID Is Null", " AND PartID=" & partId) & " ORDER BY SortOrder, OptionID", dbOpenSnapshot)
    Do While Not rs.EOF
        Set r = AddPara(Nz(rs!OptionLabel, "") & ". " & Nz(rs!OptionText, ""), 11, False, False, 0, 2, indent, True)
        AddPortion r, rk
        rs.MoveNext
    Loop
    rs.Close
End Sub

' "4 points", "1 point", "1.5 points".
Private Function PointsLabel(ByVal v As Variant) As String
    PointsLabel = PointsText(Nz(v, 0)) & IIf(Nz(v, 0) = 1, " point", " points")
End Function

' ---------------------------------------------------------------------------
' Page markings
' ---------------------------------------------------------------------------

' Lay the document out, find the page each portion is on, and make the
' header's and footer's marking line show each page the highest marking of
' the portions on it (the cover: the overall one). Falls back to the overall
' marking on every page when a page can't be read, Word fails while the
' fields go in, the fields moved any portion to another page, or sameOrder
' is False (the classification order changed since the exam was approved);
' note then says why.
Private Sub StampMarkings(ByVal overall As Long, ByVal sameOrder As Boolean, ByRef note As String)
    Dim nPages As Long, n2 As Long, pgRk() As Long, chk() As Long, p As Long, ok As Boolean, sec As Object
    Set sec = m_doc.Sections(1)
    On Error Resume Next
    m_doc.ActiveWindow.View.Type = wdPrintView      ' page numbers as printed
    On Error GoTo NoPages
    ok = sameOrder
    If ok Then ok = PageRanks(overall, nPages, pgRk)
    If ok Then
        PageMarkings sec.Headers(wdHeaderFooterPrimary), pgRk, overall
        PageMarkings sec.Footers(wdHeaderFooterPrimary), pgRk, overall
        ' The fields must leave every portion on its page.
        ok = PageRanks(overall, n2, chk)
        If ok Then ok = (n2 = nPages)
        If ok Then
            For p = 1 To nPages
                If chk(p) <> pgRk(p) Then ok = False
            Next
        End If
    End If
Stamped:
    On Error GoTo 0
    If Not ok Then SetMarking sec, RankMarking(overall)   ' also replaces any fields put in
    If Not sameOrder Then
        note = "The classification order was changed after the SM approved this exam, so every page carries the higher of " & _
            "the overall marking approved then and the exam's highest marking now."
    ElseIf Not ok Then
        note = "Word couldn't place each page's own marking reliably, so every page carries the exam's overall marking."
    End If
    Exit Sub
NoPages:
    LogError "modExamDoc.StampMarkings"                   ' so a Word that always fails here shows up
    ok = False
    Resume Stamped
End Sub

' Lay the document out and read, for each of its nPages pages, the highest
' rank of the portions on it into pgRk(1 To nPages): the cover, and a page
' with no portion, get overall. A portion longer than a page counts on every
' page it touches. False when a portion's page can't be read.
Private Function PageRanks(ByVal overall As Long, ByRef nPages As Long, ByRef pgRk() As Long) As Boolean
    Dim i As Long, p As Long, p0 As Long, p1 As Long
    m_doc.Repaginate
    nPages = m_doc.ComputeStatistics(wdStatisticPages)
    If nPages < 1 Then Exit Function
    ReDim pgRk(1 To nPages)
    pgRk(1) = overall
    For i = 0 To m_pN - 1
        p0 = PageOf(m_pS(i)): p1 = PageOf(m_pE(i) - 1)
        If p0 < 1 Or p1 < p0 Or p1 > nPages Then Exit Function
        For p = p0 To p1
            If m_pRank(i) > pgRk(p) Then pgRk(p) = m_pRank(i)
        Next
    Next
    For p = 2 To nPages
        If pgRk(p) = 0 Then pgRk(p) = overall
    Next
    PageRanks = True
End Function

' Make the marking line (the first paragraph) of header or footer hf show,
' on each page p, the marking of rank pgRk(p): one { IF { PAGE } = p
' "marking" "" } field per page, of which only that page's shows anything.
' Word works the fields out per page as it lays the pages out. The results
' saved with them are the overall marking in the first and nothing in the
' rest, so a viewer that doesn't work fields out shows the overall marking.
Private Sub PageMarkings(hf As Object, pgRk() As Long, ByVal overall As Long)
    Dim r As Object, f As Object, c As Object, p As Long, k As Long, mk As String, rk As Long, first As Boolean
    Set r = hf.Range.Paragraphs(1).Range
    r.MoveEnd 1, -1                         ' wdCharacter: keep the paragraph mark
    r.Text = ""
    rk = -1
    For p = UBound(pgRk) To 1 Step -1       ' each at the line's start, so they end up in page order
        If pgRk(p) <> rk Then
            rk = pgRk(p)
            mk = Replace(Replace(RankMarking(rk), "\", "\\"), """", "\""")
        End If
        Set r = hf.Range.Paragraphs(1).Range
        r.Collapse wdCollapseStart
        Set f = hf.Range.Fields.Add(r, wdFieldEmpty, "IF PAGENO = " & p & " """ & mk & """ """"", False)
        ' The page number goes in as a field of its own, inside the IF.
        Set c = f.Code
        k = InStr(c.Text, "PAGENO")
        If k = 0 Then Err.Raise vbObjectError + 513, "modExamDoc.PageMarkings", "The page field couldn't be placed."
        c.SetRange c.Start + k - 1, c.Start + k + 5
        hf.Range.Fields.Add c, wdFieldPage
    Next
    hf.Range.Fields.Update
    first = True
    For Each f In hf.Range.Paragraphs(1).Range.Fields     ' the marking line's only
        If f.Type = wdFieldIf Then
            If first Then f.Result.Text = RankMarking(overall): first = False Else f.Result.Text = ""
        End If
    Next
    BannerFormat hf.Range.Paragraphs(1).Range
End Sub

' The page character position pos is on.
Private Function PageOf(ByVal pos As Long) As Long
    PageOf = m_doc.Range(pos, pos).Information(wdActiveEndAdjustedPageNumber)
End Function

' The marking of the classification level of rank rk ("SECRET").
Private Function RankMarking(ByVal rk As Long) As String
    RankMarking = Nz(DLookup("Marking", "tblClassLevels", "LevelRank=" & rk), "")
    If Len(RankMarking) = 0 Then RankMarking = "MARKING MISSING"
End Function
