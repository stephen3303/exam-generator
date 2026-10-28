Option Compare Database
Option Explicit

' Exam Generator (Access edition) - the question bank's export and import:
' the whole bank, or chosen topics, to one Excel workbook, and questions
' from such a workbook back into the bank as drafts.
'
' The workbook has one sheet per table, linked by the file's own numbers
' (Question ID, Reference ID): Read me (the format line, when and by whom,
' the highest marking in it), Questions, Parts, Options, References,
' Citations, and Lists (the topics, types, difficulties and markings, for
' the dropdowns; import ignores it). Every cell is text, so commas, quotes,
' line breaks, "0800" and text starting with "=" come back as typed. The
' layout is the sample in review\question-bank-sample.xlsx.
'
' Excel itself writes and reads the file (Excel.Application, each sheet in
' one block). TransferSpreadsheet can't be used: on import it guesses each
' column's type from its first rows, so a long text after short ones is cut
' to 255 characters and 0800 becomes 800, and it writes only one table per
' sheet name. Excel keeps a line break as a line feed; the bank keeps
' CR LF, so the text is converted both ways.
'
' Export writes the version of each question in use (IsCurrent and not
' retired): an approved question as approved, a new question at whatever
' step it has reached. An open new version of an approved question is
' left out; its approved version is the one written.
'
' Import is in two steps, so nothing is written until the whole file has
' been checked. CheckBankFile reads the file and checks every row against
' the bank's rules (the save rules of QuestionProblems; the submit rules are
' only warned about, since every imported question starts as a draft), and
' keeps what it read. BankImportSummary says what ImportCheckedBank would
' do; ImportCheckedBank writes it all in one transaction. Each question
' comes in as a new question (its own QuestionKey, version 1, Draft, the
' importing admin as its author), so it goes through the SMEs and the
' approval chain like any other. A question whose text is already in the
' bank in the same topic (any version that isn't retired, compared with
' spaces and case ignored) is skipped and listed, as is a second copy in the
' file. A reference with the title and number of one in the bank is that
' one; any other is added to the shared list. Each imported question, each
' reference added and the import itself get audit rows.
'
' Only an active admin may export or import, checked in the data file at
' the moment of each (RequireExamAdmin).
'
' Owner: the question bank's export and import. Depends on: modAdmin,
' modErrors, modGlobals, modQuestions, modSignIn; the kit's modKitLocks,
' modKitUI, modPersonnel.

Public Const BANK_FORMAT As String = "Exam Generator question bank 1"   ' Read me!B3: the file's layout
Public Const ITEM_BANK As String = "Bank"       ' tblAuditTrail.ItemType of an export or import (ItemID 0)
Public Const AU_EXPORTED As String = "Exported"
Public Const AU_IMPORTED As String = "Imported"

Private Const SH_README As String = "Read me"
Private Const SH_Q As String = "Questions"
Private Const SH_P As String = "Parts"
Private Const SH_O As String = "Options"
Private Const SH_R As String = "References"
Private Const SH_C As String = "Citations"
Private Const SH_L As String = "Lists"

Private Const XL_CELL_MAX As Long = 32767       ' most characters an Excel cell holds
Private Const MAX_PROBLEMS As Long = 200        ' problems listed before the check gives up
Private Const SHOWN_LINES As Long = 12          ' lines a message lists before "and n more"
Private Const MAX_PART_LABEL As Long = 5        ' tblQuestionParts.PartLabel TEXT(5)
Private Const MAX_PARAGRAPH As Long = 100       ' tblQuestionRefs.Paragraph TEXT(100)

' Excel's numbers (late bound, so the app needs no reference to Excel).
Private Const xlWBATWorksheet As Long = -4167
Private Const xlOpenXMLWorkbook As Long = 51
Private Const xlTop As Long = -4160
Private Const xlValidateList As Long = 3
Private Const xlValidAlertStop As Long = 1
Private Const xlBetween As Long = 1

' What CheckBankFile read and checked, for BankImportSummary and
' ImportCheckedBank: Scripting.Dictionary with
'   File      the file's name (no folder)
'   Qs        Collection of the questions, in file order (QDict below)
'   Refs      Scripting.Dictionary: file Reference ID -> reference
'   Warnings  Collection of lines
' Nothing until a check passes; a failed check or an import clears it.
Private m_plan As Object

' ---------------------------------------------------------------------------
' Export
' ---------------------------------------------------------------------------

' Every topic, for the export dialog's list: TopicID, the topic's name
' ("(off)" after one turned off), how many questions export would write
' from it, and is_on (True for a topic in use; import refuses new
' questions in a topic turned off). Builds the SQL only. Errors: none.
Public Function BankTopicsSql() As String
    BankTopicsSql = "SELECT t.TopicID, t.TopicName & IIf(t.IsActive, '', ' (off)') AS topic, " & _
        "(SELECT Count(*) FROM tblQuestions AS q WHERE q.TopicID=t.TopicID AND " & InUseSql("q") & ") AS questions, " & _
        "IIf(t.IsActive, 1, 0) AS is_on FROM tblTopics AS t ORDER BY t.TopicName"
End Function

' How many questions export would write from these topics.
'   topicIds   TopicIDs separated by commas ("" gives 0)
' Errors: the engine's.
Public Function BankExportCount(ByVal topicIds As String) As Long
    If Len(topicIds) = 0 Then Exit Function
    BankExportCount = DCount("*", "tblQuestions", InUseSql("tblQuestions") & " AND TopicID IN (" & topicIds & ")")
End Function

' Write the questions in use in these topics, with their parts, options,
' references and citations, to an Excel workbook at path (replaced if it is
' there), with an Exported audit row.
'   path       the .xlsx file (KitSaveAs has asked before replacing one)
'   topicIds   TopicIDs separated by commas
' Returns "" when written, else why not, for the user (Excel missing, a
' text too long for a cell, a cell Excel didn't keep as written, the file
' open elsewhere). n returns how many questions were written.
' Errors: ERR_NOT_ALLOWED unless the signed-in person is an active admin
'   now; others pass up through RaiseAgain, with Excel closed.
Public Function ExportBank(ByVal path As String, ByVal topicIds As String, ByRef n As Long) As String
    Dim xl As Object, wb As Object, db As DAO.Database, rs As DAO.Recordset, keys As Object
    Dim where As String, qd As Variant, pd As Variant, od As Variant, rd As Variant, cd As Variant
    Dim nQ As Long, nP As Long, nO As Long, nR As Long, nC As Long, i As Long, top As Long, problem As String
    Dim num As Long, desc As String, src As String, fileName As String, marking As String, saved As Boolean
    n = 0
    RequireExamAdmin
    If Len(topicIds) = 0 Then ExportBank = "Pick at least one topic to export.": Exit Function
    On Error GoTo Fail
    Set db = CurrentDb
    where = InUseSql("q") & " AND q.TopicID IN (" & topicIds & ")"
    Set keys = CreateObject("Scripting.Dictionary")     ' QuestionID -> QuestionKey, the file's Question ID

    ' Questions: one row each, topic by topic.
    Set rs = db.OpenRecordset("SELECT q.QuestionID, q.QuestionKey, q.VersionNo, t.TopicName, ty.TypeName, d.LevelName, " & _
        "q.TotalPoints, c.Marking, c.LevelRank, q.QuestionText, q.ReferenceAnswer, s.StatusName, " & _
        "(SELECT Count(*) FROM tblQuestionParts AS p WHERE p.QuestionID=q.QuestionID) AS nParts " & _
        "FROM ((((tblQuestions AS q INNER JOIN tblTopics AS t ON q.TopicID=t.TopicID) " & _
        "INNER JOIN tblQuestionTypes AS ty ON q.QuestionTypeID=ty.QuestionTypeID) " & _
        "LEFT JOIN tblDifficultyLevels AS d ON q.DifficultyID=d.DifficultyID) " & _
        "LEFT JOIN tblClassLevels AS c ON q.ClassLevelID=c.ClassLevelID) " & _
        "INNER JOIN tblStatuses AS s ON q.StatusID=s.StatusID WHERE " & where & " ORDER BY t.TopicName, q.QuestionKey", dbOpenSnapshot)
    nQ = RowCount(rs)
    If nQ = 0 Then
        rs.Close
        ExportBank = "There are no questions in those topics to export."
        Exit Function
    End If
    ReDim qd(1 To nQ, 1 To 10)
    i = 0
    Do While Not rs.EOF
        i = i + 1
        keys(CLng(rs!QuestionID)) = CLng(rs!QuestionKey)
        qd(i, 1) = CStr(rs!QuestionKey)
        qd(i, 2) = rs!TopicName
        qd(i, 3) = rs!TypeName
        qd(i, 4) = rs!LevelName
        qd(i, 5) = PointsCell(rs!TotalPoints)
        If rs!nParts = 0 Then qd(i, 6) = rs!Marking   ' with parts, the marking is theirs (Read me)
        qd(i, 7) = rs!QuestionText
        qd(i, 8) = rs!ReferenceAnswer
        qd(i, 9) = rs!StatusName
        qd(i, 10) = QuestionRef(rs!QuestionKey, rs!VersionNo)
        If Nz(rs!LevelRank, 0) > top Then top = rs!LevelRank
        problem = CellProblem(qd, i, Array(7, 8), "Q" & rs!QuestionKey)
        If Len(problem) > 0 Then rs.Close: ExportBank = problem: Exit Function
        rs.MoveNext
    Loop
    rs.Close

    ' Parts, in their order.
    Set rs = db.OpenRecordset("SELECT p.QuestionID, p.PartLabel, p.Points, c.Marking, c.LevelRank, p.PartText, p.ReferenceAnswer " & _
        "FROM tblQuestionParts AS p LEFT JOIN tblClassLevels AS c ON p.ClassLevelID=c.ClassLevelID " & _
        "WHERE p.QuestionID IN (SELECT q.QuestionID FROM tblQuestions AS q WHERE " & where & ") " & _
        "ORDER BY p.QuestionID, p.SortOrder, p.PartLabel", dbOpenSnapshot)
    nP = RowCount(rs)
    If nP > 0 Then ReDim pd(1 To nP, 1 To 6)
    i = 0
    Do While Not rs.EOF
        i = i + 1
        pd(i, 1) = CStr(keys(CLng(rs!QuestionID)))
        pd(i, 2) = rs!PartLabel
        pd(i, 3) = PointsCell(rs!Points)
        pd(i, 4) = rs!Marking
        pd(i, 5) = rs!PartText
        pd(i, 6) = rs!ReferenceAnswer
        If Nz(rs!LevelRank, 0) > top Then top = rs!LevelRank
        problem = CellProblem(pd, i, Array(5, 6), "Q" & pd(i, 1) & " part " & rs!PartLabel)
        If Len(problem) > 0 Then rs.Close: ExportBank = problem: Exit Function
        rs.MoveNext
    Loop
    rs.Close

    ' Answer options, part by part, in their order.
    Set rs = db.OpenRecordset("SELECT o.QuestionID, p.PartLabel, o.OptionLabel, o.SortOrder, o.OptionText, o.IsCorrect " & _
        "FROM tblMCOptions AS o LEFT JOIN tblQuestionParts AS p ON o.PartID=p.PartID " & _
        "WHERE o.QuestionID IN (SELECT q.QuestionID FROM tblQuestions AS q WHERE " & where & ") " & _
        "ORDER BY o.QuestionID, Nz(p.SortOrder,0), o.SortOrder", dbOpenSnapshot)
    nO = RowCount(rs)
    If nO > 0 Then ReDim od(1 To nO, 1 To 5)
    i = 0
    Do While Not rs.EOF
        i = i + 1
        od(i, 1) = CStr(keys(CLng(rs!QuestionID)))
        od(i, 2) = rs!PartLabel
        od(i, 3) = Nz(rs!OptionLabel, OptionLetter(rs!SortOrder))
        od(i, 4) = rs!OptionText
        If Nz(rs!IsCorrect, False) Then od(i, 5) = "Yes"
        problem = CellProblem(od, i, Array(4), "Q" & od(i, 1) & " option " & od(i, 3))
        If Len(problem) > 0 Then rs.Close: ExportBank = problem: Exit Function
        rs.MoveNext
    Loop
    rs.Close

    ' The references they cite, and the citations.
    Set rs = db.OpenRecordset("SELECT r.RefID, r.RefTitle, r.RefNumber, r.Revision FROM tblSourceRefs AS r " & _
        "WHERE r.RefID IN (SELECT qr.RefID FROM tblQuestionRefs AS qr WHERE qr.QuestionID IN " & _
        "(SELECT q.QuestionID FROM tblQuestions AS q WHERE " & where & ")) ORDER BY r.RefTitle, r.RefNumber", dbOpenSnapshot)
    nR = RowCount(rs)
    If nR > 0 Then ReDim rd(1 To nR, 1 To 4)
    i = 0
    Do While Not rs.EOF
        i = i + 1
        rd(i, 1) = CStr(rs!RefID)
        rd(i, 2) = rs!RefTitle
        rd(i, 3) = rs!RefNumber
        rd(i, 4) = rs!Revision
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT qr.QuestionID, qr.RefID, qr.Paragraph FROM tblQuestionRefs AS qr " & _
        "WHERE qr.QuestionID IN (SELECT q.QuestionID FROM tblQuestions AS q WHERE " & where & ") " & _
        "ORDER BY qr.QuestionID, qr.RefID", dbOpenSnapshot)
    nC = RowCount(rs)
    If nC > 0 Then ReDim cd(1 To nC, 1 To 3)
    i = 0
    Do While Not rs.EOF
        i = i + 1
        cd(i, 1) = CStr(keys(CLng(rs!QuestionID)))
        cd(i, 2) = CStr(rs!RefID)
        cd(i, 3) = rs!Paragraph
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    ' The workbook.
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    On Error GoTo Fail
    If xl Is Nothing Then
        ExportBank = "Excel couldn't be started on this computer, so the bank wasn't exported. Export it from a computer with Excel."
        Exit Function
    End If
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    Set wb = xl.Workbooks.Add(xlWBATWorksheet)
    marking = Nz(DLookup("Marking", "tblClassLevels", "LevelRank=" & top), "")
    WriteReadMe wb.Worksheets(1), nQ, topicIds, marking
    problem = WriteSheet(wb, SH_Q, QuestionCols(), qd, nQ, Array(12, 20, 17, 11, 8, 16, 60, 50, 11, 10), Array(7, 8), 9)
    If Len(problem) = 0 Then problem = WriteSheet(wb, SH_P, PartCols(), pd, nP, Array(12, 6, 8, 16, 60, 50), Array(5, 6))
    If Len(problem) = 0 Then problem = WriteSheet(wb, SH_O, OptionCols(), od, nO, Array(12, 6, 8, 70, 9), Array(4))
    If Len(problem) = 0 Then problem = WriteSheet(wb, SH_R, RefCols(), rd, nR, Array(13, 45, 18, 10), Array(2))
    If Len(problem) = 0 Then problem = WriteSheet(wb, SH_C, CiteCols(), cd, nC, Array(12, 13, 24), Array())
    If Len(problem) = 0 Then problem = WriteLists(wb)
    If Len(problem) > 0 Then ExportBank = problem: GoTo CloseExcel
    AddDropdowns wb
    MarkSheets xl, wb, marking
    wb.Worksheets(1).Activate
    On Error Resume Next
    If Len(Dir(path)) > 0 Then Kill path
    If Err.Number <> 0 Then
        On Error GoTo Fail
        ExportBank = "The file " & path & " couldn't be replaced; someone may have it open. Close it, or pick another name, and export again."
        GoTo CloseExcel
    End If
    On Error GoTo Fail
    wb.SaveAs path, xlOpenXMLWorkbook
    saved = True
    wb.Close False
    Set wb = Nothing
    xl.Quit
    Set xl = Nothing
    ' The audit row last: a file whose export isn't recorded is deleted
    ' again (Fail).
    fileName = Mid$(path, InStrRev(path, "\") + 1)
    AddAudit ITEM_BANK, 0, Null, AU_EXPORTED, , nQ & IIf(nQ = 1, " question", " questions") & " in " & TopicNames(topicIds) & _
        " to " & fileName & IIf(Len(marking) > 0, ", highest marking " & marking, "")
    n = nQ
    Exit Function
CloseExcel:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Exit Function
Fail:
    num = Err.Number: desc = Err.Description: src = Err.Source
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not wb Is Nothing Then wb.Close False
    Set wb = Nothing
    If Not xl Is Nothing Then xl.Quit
    Set xl = Nothing
    If saved Then Kill path                 ' saved, but its audit row couldn't be written
    On Error GoTo 0
    n = 0
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modBank.ExportBank", path
End Function

' The where clause for a question version in use: current and not retired.
Private Function InUseSql(ByVal tbl As String) As String
    InUseSql = tbl & ".IsCurrent AND " & tbl & ".StatusID<>" & QStatusID(QS_RETIRED)
End Function

Private Function RowCount(rs As DAO.Recordset) As Long
    If rs.EOF Then Exit Function
    rs.MoveLast
    RowCount = rs.RecordCount
    rs.MoveFirst
End Function

' Why row i of data can't go in a workbook, or "": a text in one of cols
' longer than an Excel cell holds.
Private Function CellProblem(data As Variant, ByVal i As Long, ByVal cols As Variant, ByVal what As String) As String
    Dim c As Variant
    For Each c In cols
        If Len(Nz(data(i, c), "")) > XL_CELL_MAX Then
            CellProblem = what & " has a text of " & Len(data(i, c)) & " characters, more than an Excel cell holds (" & _
                XL_CELL_MAX & "), so the bank wasn't exported. Shorten it, or leave its topic out."
            Exit Function
        End If
    Next
End Function

' The topics' names, joined: "Damage Control, Electrical Safety and 3 more".
Private Function TopicNames(ByVal topicIds As String) As String
    Dim rs As DAO.Recordset, n As Long, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT TopicName FROM tblTopics WHERE TopicID IN (" & topicIds & ") ORDER BY TopicName", dbOpenSnapshot)
    Do While Not rs.EOF
        n = n + 1
        If n <= 3 Then s = s & IIf(n > 1, ", ", "") & rs!TopicName
        rs.MoveNext
    Loop
    rs.Close
    If n > 3 Then s = s & " and " & (n - 3) & " more"
    TopicNames = s
End Function

Private Function QuestionCols() As Variant
    QuestionCols = Array("Question ID", "Topic", "Type", "Difficulty", "Points", "Marking", "Question text", "Reference answer", _
        "Status", "Bank ref")
End Function

Private Function PartCols() As Variant
    PartCols = Array("Question ID", "Part", "Points", "Marking", "Part text", "Reference answer")
End Function

Private Function OptionCols() As Variant
    OptionCols = Array("Question ID", "Part", "Option", "Option text", "Correct")
End Function

Private Function RefCols() As Variant
    RefCols = Array("Reference ID", "Title", "Number", "Revision")
End Function

Private Function CiteCols() As Variant
    CiteCols = Array("Question ID", "Reference ID", "Paragraph")
End Function

' A value as the text it should read back as: "" for Null, line breaks as
' Excel's line feed.
Private Function CellExpected(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    CellExpected = Replace(Replace(CStr(v), vbCrLf, vbLf), vbCr, vbLf)
End Function

' A value as a workbook cell: Empty for Null or "", line breaks as Excel's
' line feed, a leading apostrophe doubled (Excel takes one as "this is
' text" and drops it).
Private Function CellText(ByVal v As Variant) As Variant
    Dim s As String
    s = CellExpected(v)
    If Len(s) = 0 Then CellText = Empty: Exit Function
    If Left$(s, 1) = "'" Then s = "'" & s
    CellText = s
End Function

' Points as a cell: 4, 1.5, with a dot whatever this computer's decimal
' separator (PointsText follows it), so the file reads the same anywhere.
Private Function PointsCell(ByVal v As Variant) As String
    If IsNull(v) Then Exit Function
    PointsCell = Trim$(Str$(CDbl(v)))
    If Left$(PointsCell, 1) = "." Then PointsCell = "0" & PointsCell     ' Str$(0.5) is " .5"
End Function

' Add a sheet after the others and fill it: the captions in row 1 (white on
' dark blue, grey for the columns from firstInfo on, which import ignores),
' then n rows of data, every column formatted as text first so Excel keeps
' each value as typed. The long text columns (longCols) go back to General
' once filled: Excel shows #### in a text cell over 255 characters, and the
' values already in them stay text. Frozen captions, a filter, wrapped text.
' Every cell is then read back and compared with what was meant; returns
' "" when they all match, else which cell didn't, for the user.
Private Function WriteSheet(wb As Object, ByVal nm As String, ByVal cols As Variant, data As Variant, ByVal n As Long, _
        ByVal widths As Variant, ByVal longCols As Variant, Optional ByVal firstInfo As Long = 0) As String
    Dim ws As Object, nc As Long, hdr() As Variant, i As Long, r As Long, out() As Variant, c As Variant, back As Variant
    Dim want As String, got As String
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = nm
    nc = UBound(cols) + 1
    ws.Range(ws.Cells(1, 1), ws.Cells(1, nc)).EntireColumn.NumberFormat = "@"
    ReDim hdr(1 To 1, 1 To nc)
    For i = 1 To nc
        hdr(1, i) = cols(i - 1)
        ws.Columns(i).ColumnWidth = widths(i - 1)
    Next
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nc))
        .Value = hdr
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With
    If firstInfo > 0 Then ws.Range(ws.Cells(1, firstInfo), ws.Cells(1, nc)).Interior.Color = RGB(127, 127, 127)
    If n > 0 Then
        ReDim out(1 To n, 1 To nc)
        For r = 1 To n
            For i = 1 To nc
                out(r, i) = CellText(data(r, i))
            Next
        Next
        ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, nc)).Value = out
        ' Read it all back, exactly (case and every character), while every
        ' cell is still text. A cell whose leading apostrophe Excel kept
        ' (instead of taking it as "this is text") is written again without
        ' the doubled one.
        back = ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, nc)).Value2
        For r = 1 To n
            For i = 1 To nc
                want = CellExpected(data(r, i))
                got = BackText(back(r, i))
                If StrComp(got, want, vbBinaryCompare) <> 0 And Left$(want, 1) = "'" And _
                        StrComp(got, "'" & want, vbBinaryCompare) = 0 Then
                    ws.Cells(r + 1, i).Value = want
                    got = BackText(ws.Cells(r + 1, i).Value2)
                End If
                If StrComp(got, want, vbBinaryCompare) <> 0 Then
                    WriteSheet = "Excel didn't keep the " & cols(i - 1) & " cell of " & nm & " row " & (r + 1) & " as written " & _
                        "(it should start """ & ClipText(Replace(want, vbLf, " "), 40) & """), so the bank wasn't exported. " & _
                        "The error log has no entry for this; tell whoever looks after the app."
                    Exit Function
                End If
            Next
        Next
        For Each c In longCols
            ws.Range(ws.Cells(2, c), ws.Cells(n + 1, c)).NumberFormat = "General"
        Next
    End If
    On Error Resume Next                    ' cosmetic: a failure here doesn't stop the export
    With ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, nc))
        .WrapText = True
        .VerticalAlignment = xlTop
        .AutoFilter
    End With
    On Error GoTo 0
    FreezeTopRow ws
End Function

' A cell as read back (Value2) for WriteSheet's comparison.
Private Function BackText(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then Exit Function
    BackText = CStr(v)
End Function

' Freeze row 1 of ws. Cosmetic: a failure is ignored.
Private Sub FreezeTopRow(ws As Object)
    On Error Resume Next
    ws.Activate
    With ws.Parent.Windows(1)
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub

' The highest marking in the file at the top and bottom of every sheet as
' it prints, and on the Read me sheet already. Cosmetic: a failure is
' ignored (the Read me still says it).
Private Sub MarkSheets(xl As Object, wb As Object, ByVal marking As String)
    Dim ws As Object
    If Len(marking) = 0 Then Exit Sub
    On Error Resume Next
    xl.PrintCommunication = False           ' Excel 2010 on: page setup without asking the printer each time
    For Each ws In wb.Worksheets
        ws.PageSetup.CenterHeader = "&B" & Replace(marking, "&", "&&")
        ws.PageSetup.CenterFooter = "&B" & Replace(marking, "&", "&&")
    Next
    xl.PrintCommunication = True
End Sub

' The Read me sheet: the format line (import checks it), when, by whom,
' from which topics, how many questions, the highest marking in the file,
' and how the sheets fit together.
Private Sub WriteReadMe(ws As Object, ByVal nQ As Long, ByVal topicIds As String, ByVal marking As String)
    Dim meta As Variant, notes As Variant, i As Long, r As Long
    ws.Name = SH_README
    ws.Columns(1).ColumnWidth = 16
    ws.Columns(2).ColumnWidth = 110
    ws.Columns(2).NumberFormat = "@"
    ws.Cells(1, 1).Value = "Exam Generator question bank"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    meta = Array("Format", BANK_FORMAT, "Exported", Format$(Now(), "yyyy-mm-dd hh:nn"), _
        "Exported by", SignedInName(), "Topics", TopicNames(topicIds), "Questions", CStr(nQ), _
        "Highest marking", marking)
    r = 3
    For i = 0 To UBound(meta) Step 2
        ws.Cells(r, 1).Value = meta(i)
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 2).Value = CellText(meta(i + 1))
        r = r + 1
    Next
    notes = Array( _
        "How the sheets fit together", _
        "Questions: one row per question. Question ID is this file's own number for the question; it only links the rows of the " & _
            "other sheets to it, and the bank gives an imported question a new number.", _
        "Parts: one row per part of a multipart question (A, B...). A question with parts needs at least two, the parts' points " & _
            "must add up to the question's points, and the question's marking is the highest of its parts' (its own Marking cell " & _
            "is left empty).", _
        "Options: the answer options of multiple choice questions only. Part is empty for a question without parts. Mark exactly " & _
            "one option Yes under Correct in each question or part.", _
        "References: one row per reference document. Reference ID is this file's own number. On import a reference with the same " & _
            "title and number as one in the bank is matched to it; any other is added to the bank's list.", _
        "Citations: which questions cite which references, with the paragraph (optional).", _
        "Lists: the topics, types, difficulties and markings the bank had when this file was made. Topic, Type, Difficulty and " & _
            "Marking must use these names exactly.", _
        "", _
        "Text", _
        "Commas, quotes and line breaks (Alt+Enter in Excel) are kept as typed. A question, part or option text can be at most " & _
            MAX_TEXT & " characters. Points use a dot for a half point (1.5). In the long text columns, type an apostrophe " & _
            "before a text that starts with = so Excel keeps it as text; Excel reads a long text again when you edit its cell, so " & _
            "the same goes for a text there that looks like a number (0800).", _
        "", _
        "Import", _
        "Import checks the whole file first and lists every problem with its sheet and row; nothing is added until the file has " & _
            "none. Each question comes in as a new draft, so it goes through the SMEs and the approval chain like any other. A " & _
            "question already in its topic (the same text, parts and options) is skipped and listed. Topics turned off take no " & _
            "new questions.", _
        "Grey columns (Status, Bank ref) are filled by export to show where a question came from; import ignores them.")
    r = r + 1
    For i = 0 To UBound(notes)
        ws.Cells(r, 1).Value = notes(i)
        If notes(i) = "How the sheets fit together" Or notes(i) = "Text" Or notes(i) = "Import" Then ws.Cells(r, 1).Font.Bold = True
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, 2))
            .Merge
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
        If Len(notes(i)) > 120 Then ws.Rows(r).RowHeight = 15 * (Len(notes(i)) \ 120 + 1)
        r = r + 1
    Next
    ws.Columns(2).NumberFormat = "General"  ' a long topics line would show as #### in a text cell
End Sub

' The signed-in person's name, for the Read me sheet ("" if not found).
Private Function SignedInName() As String
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT " & PersonNameSql() & " AS nm FROM " & PERSON_TABLE & " AS p WHERE p.person_id=" & _
        SignedInPersonID(), dbOpenSnapshot)
    If Not rs.EOF Then SignedInName = Nz(rs!nm, "")
    rs.Close
End Function

' The Lists sheet: every topic, the types, the difficulties and the
' markings, lowest first, with their portion marks. Returns as WriteSheet.
Private Function WriteLists(wb As Object) As String
    Dim lists(1 To 4) As Collection, rs As DAO.Recordset, i As Long, n As Long, r As Long, data() As Variant
    Dim portions As New Collection
    For i = 1 To 4
        Set lists(i) = New Collection
    Next
    ' Each value is added as text: rs!TopicName alone would add the Field
    ' object, which is gone once the recordset closes.
    Set rs = CurrentDb.OpenRecordset("SELECT TopicName FROM tblTopics ORDER BY TopicName", dbOpenSnapshot)
    Do While Not rs.EOF: lists(1).Add Nz(rs!TopicName.Value, ""): rs.MoveNext: Loop
    rs.Close
    Set rs = CurrentDb.OpenRecordset("SELECT TypeName FROM tblQuestionTypes ORDER BY QuestionTypeID", dbOpenSnapshot)
    Do While Not rs.EOF: lists(2).Add Nz(rs!TypeName.Value, ""): rs.MoveNext: Loop
    rs.Close
    Set rs = CurrentDb.OpenRecordset("SELECT LevelName FROM tblDifficultyLevels ORDER BY SortOrder", dbOpenSnapshot)
    Do While Not rs.EOF: lists(3).Add Nz(rs!LevelName.Value, ""): rs.MoveNext: Loop
    rs.Close
    Set rs = CurrentDb.OpenRecordset("SELECT Marking, PortionMark FROM tblClassLevels ORDER BY LevelRank", dbOpenSnapshot)
    Do While Not rs.EOF: lists(4).Add Nz(rs!Marking.Value, ""): portions.Add Nz(rs!PortionMark.Value, ""): rs.MoveNext: Loop
    rs.Close
    For i = 1 To 4
        If lists(i).Count > n Then n = lists(i).Count
    Next
    If n > 0 Then
        ReDim data(1 To n, 1 To 5)
        For r = 1 To n
            For i = 1 To 4
                If r <= lists(i).Count Then data(r, i) = lists(i)(r)
            Next
            If r <= portions.Count Then data(r, 5) = portions(r)
        Next
    End If
    WriteLists = WriteSheet(wb, SH_L, Array("Topic", "Type", "Difficulty", "Marking", "Portion mark"), data, n, _
        Array(22, 17, 11, 16, 13), Array())
End Function

' Dropdowns from the Lists sheet on the Questions and Parts columns that
' name a topic, type, difficulty or marking, for rows typed in Excel.
' Cosmetic: a failure is ignored.
Private Sub AddDropdowns(wb As Object)
    On Error Resume Next
    ListRule wb.Worksheets(SH_Q).Range("B2:B2000"), "=Lists!$A$2:$A$1000"
    ListRule wb.Worksheets(SH_Q).Range("C2:C2000"), "=Lists!$B$2:$B$20"
    ListRule wb.Worksheets(SH_Q).Range("D2:D2000"), "=Lists!$C$2:$C$50"
    ListRule wb.Worksheets(SH_Q).Range("F2:F2000"), "=Lists!$D$2:$D$50"
    ListRule wb.Worksheets(SH_P).Range("D2:D2000"), "=Lists!$D$2:$D$50"
    ListRule wb.Worksheets(SH_O).Range("E2:E5000"), "Yes"
End Sub

Private Sub ListRule(rng As Object, ByVal source As String)
    With rng.Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, source
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

' ---------------------------------------------------------------------------
' Import: check
' ---------------------------------------------------------------------------

' Read the workbook at path and check every row of it against the bank's
' rules. Nothing is written.
'   path   an .xlsx file in the layout export writes
' Returns "" when the file can be imported (what was read is kept for
' BankImportSummary and ImportCheckedBank), else every problem found, one
' per line, each with its sheet and row: "Questions row 4: Topic ...".
' Errors: ERR_NOT_ALLOWED unless the signed-in person is an active admin
'   now; others pass up through RaiseAgain, with Excel closed.
Public Function CheckBankFile(ByVal path As String) As String
    Dim xl As Object, wb As Object, sheets As Object, probs As Collection, warns As Collection
    Dim num As Long, desc As String, src As String
    Set m_plan = Nothing
    RequireExamAdmin
    Set probs = New Collection
    Set warns = New Collection
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    On Error GoTo Fail
    If xl Is Nothing Then
        CheckBankFile = "Excel couldn't be started on this computer, so the file couldn't be read. Import it from a computer with Excel."
        Exit Function
    End If
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3               ' msoAutomationSecurityForceDisable: no macros run from the file
    Set wb = xl.Workbooks.Open(path, 0, True)   ' no link updates, read only
    Set sheets = ReadSheets(wb, probs)
    wb.Close False
    Set wb = Nothing
    xl.Quit
    Set xl = Nothing
    If probs.Count = 0 Then CheckRows sheets, Mid$(path, InStrRev(path, "\") + 1), probs, warns
    If probs.Count > 0 Then
        Set m_plan = Nothing
        CheckBankFile = JoinLines(probs, MAX_PROBLEMS)
    Else
        Set m_plan("Warnings") = warns
    End If
    Exit Function
Fail:
    num = Err.Number: desc = Err.Description: src = Err.Source
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Set wb = Nothing
    If Not xl Is Nothing Then xl.Quit
    Set xl = Nothing
    Set m_plan = Nothing
    On Error GoTo 0
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modBank.CheckBankFile", path
End Function

' The format line and every sheet import reads, as Scripting.Dictionary:
' sheet name -> Collection of rows. A row is a Scripting.Dictionary of the
' sheet's columns by caption (HeaderKey) plus "row", its row in Excel. Blank
' rows are left out. A missing sheet, a missing column or a file of another
' format is a problem.
Private Function ReadSheets(wb As Object, probs As Collection) As Object
    Dim out As Object, nm As Variant, ws As Object, fmt As String
    Set out = CreateObject("Scripting.Dictionary")
    Set ws = FindSheet(wb, SH_README)
    If Not ws Is Nothing Then fmt = Trim$(CStr(Nz(ws.Range("B3").Value2, "")))
    If ws Is Nothing Or fmt <> BANK_FORMAT Then
        probs.Add "This isn't a question bank file in the layout this app exports: its Read me sheet must say """ & BANK_FORMAT & _
            """ next to Format" & IIf(Len(fmt) > 0, " (it says """ & ClipText(fmt, 60) & """)", "") & ". Export the bank to get one."
        Set ReadSheets = out
        Exit Function
    End If
    For Each nm In Array(SH_Q, SH_P, SH_O, SH_R, SH_C)
        Set ws = FindSheet(wb, CStr(nm))
        If ws Is Nothing Then
            probs.Add "The file has no " & nm & " sheet. Every sheet export makes must be there, even with no rows."
        Else
            Set out(CStr(nm)) = SheetRows(ws, ColsOf(CStr(nm)), probs)
        End If
    Next
    Set ReadSheets = out
End Function

Private Function ColsOf(ByVal nm As String) As Variant
    Select Case nm
        Case SH_Q: ColsOf = QuestionCols()
        Case SH_P: ColsOf = PartCols()
        Case SH_O: ColsOf = OptionCols()
        Case SH_R: ColsOf = RefCols()
        Case SH_C: ColsOf = CiteCols()
    End Select
End Function

Private Function FindSheet(wb As Object, ByVal nm As String) As Object
    Dim ws As Object
    For Each ws In wb.Worksheets
        If LCase$(Trim$(ws.Name)) = LCase$(nm) Then Set FindSheet = ws: Exit Function
    Next
End Function

' One sheet's rows (ReadSheets). Columns are found by their captions in row
' 1, in any order; "Status" and "Bank ref" may be missing.
Private Function SheetRows(ws As Object, ByVal cols As Variant, probs As Collection) As Collection
    Dim out As New Collection, used As Object, lastR As Long, lastC As Long, v As Variant, c As Long, r As Long
    Dim at As Object, k As Variant, d As Object, blank As Boolean, cell As String
    Set SheetRows = out
    Set used = ws.UsedRange
    lastR = used.Row + used.Rows.Count - 1
    lastC = used.Column + used.Columns.Count - 1
    If lastR < 1 Or lastC < 1 Then lastR = 1: lastC = 1
    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value2
    If Not IsArray(v) Then                  ' one cell: Value2 is the value itself
        Dim one(1 To 1, 1 To 1) As Variant
        one(1, 1) = v
        v = one
    End If
    Set at = CreateObject("Scripting.Dictionary")
    For c = 1 To lastC
        cell = HeaderKey(CellValue(v(1, c)))
        If Len(cell) > 0 And Not at.Exists(cell) Then at(cell) = c
    Next
    For Each k In cols
        If Not at.Exists(HeaderKey(k)) And k <> "Status" And k <> "Bank ref" Then
            probs.Add ws.Name & ": the " & k & " column is missing from row 1."
        End If
    Next
    If probs.Count > 0 Then Exit Function
    For r = 2 To lastR
        Set d = CreateObject("Scripting.Dictionary")
        blank = True
        For Each k In cols
            If at.Exists(HeaderKey(k)) Then
                cell = CellValue(v(r, at(HeaderKey(k))))
                If cell = "#ERROR" Then
                    probs.Add ws.Name & " row " & r & ": the " & k & " cell holds an Excel error (such as #N/A). Type the value in it."
                    cell = ""
                End If
            Else
                cell = ""
            End If
            d(HeaderKey(k)) = cell
            If Len(Trim$(cell)) > 0 Then blank = False
        Next
        d("row") = r
        If Not blank Then out.Add d
    Next
End Function

' A cell's value as text: "" for an empty cell, "#ERROR" for an Excel
' error, a number typed in Excel as its digits with a dot (1, 1.5) on any
' computer, line breaks as CR LF.
Private Function CellValue(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsError(v) Then CellValue = "#ERROR": Exit Function
    If VarType(v) = vbDouble Then
        CellValue = Trim$(Str$(v))
        If Left$(CellValue, 1) = "." Then CellValue = "0" & CellValue
        Exit Function
    End If
    CellValue = Replace(Replace(Replace(CStr(v), vbCrLf, vbLf), vbCr, vbLf), vbLf, vbCrLf)
End Function

' A caption as compared: lower case, without spaces, underscores, hyphens or
' dots ("Question ID" and "question_id" are the same).
Private Function HeaderKey(ByVal s As Variant) As String
    HeaderKey = LCase$(Replace(Replace(Replace(Replace(Trim$(Nz(s, "")), " ", ""), "_", ""), "-", ""), ".", ""))
End Function

' Check the rows read (ReadSheets) and, when no problem is found, keep them
' as m_plan. Every problem goes in probs, every warning in warns.
Private Sub CheckRows(sheets As Object, ByVal fileName As String, probs As Collection, warns As Collection)
    Dim qs As New Collection, byId As Object, refs As Object, d As Variant, qn As Object, p As Object, o As Object
    Dim id As String, s As String, k As Variant, g As Variant, seen As Object, bank As Object, key As String
    Dim plbl As String, n As Long, nCorrect As Long, sum As Double, top As Long, pt As String, refId As String, para As String
    Set byId = CreateObject("Scripting.Dictionary")
    Set refs = CreateObject("Scripting.Dictionary")

    ' Questions.
    For Each d In sheets(SH_Q)
        If probs.Count >= MAX_PROBLEMS Then Exit Sub
        s = SH_Q & " row " & d("row") & ": "
        id = IdText(d("questionid"))
        Set qn = CreateObject("Scripting.Dictionary")
        qn("row") = d("row")
        qn("id") = id
        If Len(id) = 0 Then
            probs.Add s & "Question ID is empty. Give each question its own number."
        ElseIf byId.Exists(id) Then
            probs.Add s & "Question ID " & id & " is also on row " & byId(id)("row") & ". Give each question its own number."
        Else
            byId.Add id, qn
        End If
        qn("topicId") = LookupID("TopicID", "tblTopics", "TopicName", d("topic"))
        If Len(Trim$(d("topic"))) = 0 Then
            probs.Add s & "Topic is empty."
        ElseIf qn("topicId") = 0 Then
            probs.Add s & "Topic """ & ClipText(Trim$(d("topic")), 60) & """ isn't in the bank. Use a topic name from the Lists " & _
                "sheet, or ask the admin to add the topic first."
        End If
        qn("topicOff") = False                  ' checked once the skips are known: a skipped question may be in one
        If qn("topicId") <> 0 Then qn("topicOff") = Not Nz(DLookup("IsActive", "tblTopics", "TopicID=" & qn("topicId")), False)
        qn("topicText") = Trim$(d("topic"))
        qn("type") = TypeFromText(d("type"))
        If Len(qn("type")) = 0 Then probs.Add s & "Type must be " & QT_MC & ", " & QT_TF & " or " & QT_SA & _
            IIf(Len(Trim$(d("type"))) > 0, " (it is """ & ClipText(Trim$(d("type")), 40) & """)", "") & "."
        qn("diffId") = LookupID("DifficultyID", "tblDifficultyLevels", "LevelName", d("difficulty"))
        If qn("diffId") = 0 Then probs.Add s & IIf(Len(Trim$(d("difficulty"))) = 0, "Difficulty is empty.", "Difficulty """ & _
            ClipText(Trim$(d("difficulty")), 40) & """ isn't one of the bank's (Lists sheet).")
        pt = PointsProblem(NumberOrNull(d("points")), "Points")
        If Len(pt) > 0 Then probs.Add s & pt Else qn("points") = CDbl(NumberOrNull(d("points")))
        qn("markId") = MarkingID(d("marking"))
        qn("markText") = Trim$(d("marking"))
        qn("text") = CleanText(d("questiontext"))
        If IsNull(qn("text")) Then
            probs.Add s & "Question text is empty."
        ElseIf Len(qn("text")) > MAX_TEXT Then
            probs.Add s & "Question text can be at most " & MAX_TEXT & " characters; it is " & Len(qn("text")) & "."
        End If
        qn("answer") = CleanText(d("referenceanswer"))
        Set qn("parts") = New Collection
        Set qn("partAt") = CreateObject("Scripting.Dictionary")
        Set qn("opts") = CreateObject("Scripting.Dictionary")   ' part label ("" for the question) -> Collection
        Set qn("cites") = New Collection
        Set qn("citeAt") = CreateObject("Scripting.Dictionary")
        qn("skip") = ""
        qs.Add qn
    Next

    ' Parts.
    For Each d In sheets(SH_P)
        If probs.Count >= MAX_PROBLEMS Then Exit Sub
        s = SH_P & " row " & d("row") & ": "
        Set qn = QuestionOf(byId, d("questionid"), s, probs)
        If Not qn Is Nothing Then
            plbl = UCase$(Trim$(d("part")))
            Set p = CreateObject("Scripting.Dictionary")
            p("row") = d("row")
            p("label") = plbl
            If Len(plbl) = 0 Then
                probs.Add s & "Part is empty. Letter the parts A, B, C..."
            ElseIf Len(plbl) > MAX_PART_LABEL Then
                probs.Add s & "Part """ & ClipText(plbl, 20) & """ is too long; letter the parts A, B, C..."
            ElseIf qn("partAt").Exists(plbl) Then
                probs.Add s & "Question " & qn("id") & " already has a part " & plbl & " (row " & qn("partAt")(plbl)("row") & ")."
            Else
                qn("partAt").Add plbl, p
            End If
            pt = PointsProblem(NumberOrNull(d("points")), "Part " & plbl & "'s points")
            If Len(pt) > 0 Then probs.Add s & pt Else p("points") = CDbl(NumberOrNull(d("points")))
            p("markId") = MarkingID(d("marking"))
            If p("markId") = 0 Then probs.Add s & IIf(Len(Trim$(d("marking"))) = 0, "Marking is empty: every part needs its own.", _
                "Marking """ & ClipText(Trim$(d("marking")), 40) & """ isn't one of the bank's (Lists sheet).")
            p("text") = CleanText(d("parttext"))
            If IsNull(p("text")) Then
                probs.Add s & "Part text is empty."
            ElseIf Len(p("text")) > MAX_TEXT Then
                probs.Add s & "Part text can be at most " & MAX_TEXT & " characters; it is " & Len(p("text")) & "."
            End If
            p("answer") = CleanText(d("referenceanswer"))
            qn("parts").Add p
        End If
    Next

    ' Options.
    For Each d In sheets(SH_O)
        If probs.Count >= MAX_PROBLEMS Then Exit Sub
        s = SH_O & " row " & d("row") & ": "
        Set qn = QuestionOf(byId, d("questionid"), s, probs)
        If Not qn Is Nothing Then
            plbl = UCase$(Trim$(d("part")))
            Set o = CreateObject("Scripting.Dictionary")
            o("row") = d("row")
            o("label") = UCase$(Trim$(d("option")))
            o("text") = CleanText(d("optiontext"))
            o("correct") = YesText(d("correct"))
            If qn("type") <> QT_MC And Len(qn("type")) > 0 Then
                probs.Add s & "Question " & qn("id") & " is " & LCase$(qn("type")) & "; only multiple choice questions have options."
            ElseIf Len(plbl) > 0 And Not qn("partAt").Exists(plbl) Then
                probs.Add s & "Question " & qn("id") & " has no part " & ClipText(plbl, 10) & " on the Parts sheet."
            ElseIf Len(plbl) = 0 And qn("parts").Count > 0 Then
                probs.Add s & "Question " & qn("id") & " has parts, so each option needs its part (A, B...)."
            End If
            If IsNull(o("text")) Then
                probs.Add s & "Option text is empty."
            ElseIf Len(o("text")) > MAX_TEXT Then
                probs.Add s & "Option text can be at most " & MAX_TEXT & " characters; it is " & Len(o("text")) & "."
            End If
            If IsNull(o("correct")) Then probs.Add s & "Correct must be Yes or empty (it is """ & ClipText(Trim$(d("correct")), 20) & """)."
            If Not qn("opts").Exists(plbl) Then qn("opts").Add plbl, New Collection
            qn("opts")(plbl).Add o
        End If
    Next

    ' References.
    For Each d In sheets(SH_R)
        If probs.Count >= MAX_PROBLEMS Then Exit Sub
        s = SH_R & " row " & d("row") & ": "
        id = IdText(d("referenceid"))
        Set p = CreateObject("Scripting.Dictionary")
        p("row") = d("row")
        p("title") = Trim$(d("title"))
        p("number") = Trim$(d("number"))
        p("rev") = Trim$(d("revision"))
        If Len(id) = 0 Then
            probs.Add s & "Reference ID is empty. Give each reference its own number."
        ElseIf refs.Exists(id) Then
            probs.Add s & "Reference ID " & id & " is also on row " & refs(id)("row") & "."
        Else
            refs.Add id, p
        End If
        pt = RefFieldProblem(p("title"), p("number"), p("rev"))
        If Len(pt) > 0 Then probs.Add s & pt
        For Each k In refs.Keys
            If k <> id Then
                If LCase$(refs(k)("title")) = LCase$(p("title")) And LCase$(refs(k)("number")) = LCase$(p("number")) Then
                    probs.Add s & "The same title and number as reference " & k & " (row " & refs(k)("row") & "). List each reference once."
                    Exit For
                End If
            End If
        Next
        p("bankId") = Nz(DLookup("RefID", "tblSourceRefs", "RefTitle=" & Q(p("title")) & " AND Nz(RefNumber,'')=" & Q(p("number"))), 0)
        p("used") = False
    Next

    ' Citations.
    For Each d In sheets(SH_C)
        If probs.Count >= MAX_PROBLEMS Then Exit Sub
        s = SH_C & " row " & d("row") & ": "
        Set qn = QuestionOf(byId, d("questionid"), s, probs)
        refId = IdText(d("referenceid"))
        para = Trim$(d("paragraph"))
        If Not qn Is Nothing Then
            If Len(refId) = 0 Then
                probs.Add s & "Reference ID is empty."
            ElseIf Not refs.Exists(refId) Then
                probs.Add s & "Reference ID " & refId & " isn't on the References sheet."
            ElseIf qn("citeAt").Exists(refId) Then
                probs.Add s & "Question " & qn("id") & " already cites reference " & refId & " (row " & qn("citeAt")(refId) & "). Cite it once."
            Else
                qn("citeAt").Add refId, d("row")
                Set o = CreateObject("Scripting.Dictionary")
                o("ref") = refId
                o("para") = NullIfEmpty(para)
                qn("cites").Add o
            End If
            If Len(para) > MAX_PARAGRAPH Then probs.Add s & "Paragraph can be at most " & MAX_PARAGRAPH & " characters."
        End If
    Next

    ' Each question as a whole: its parts, its options, its marking.
    For Each qn In qs
        If probs.Count >= MAX_PROBLEMS Then Exit Sub
        s = SH_Q & " row " & qn("row") & " (question " & qn("id") & "): "
        If qn("parts").Count = 0 Then
            If qn("markId") = 0 Then probs.Add s & IIf(Len(qn("markText")) = 0, "Marking is empty.", "Marking """ & _
                ClipText(qn("markText"), 40) & """ isn't one of the bank's (Lists sheet).")
        Else
            If qn("parts").Count = 1 Then probs.Add s & "it has only part " & qn("parts")(1)("label") & "; a question with parts needs at least two."
            n = 0: sum = 0: top = 0
            For Each p In qn("parts")
                n = n + 1
                If p("label") <> OptionLetter(n) And n <= 26 Then
                    probs.Add s & "its parts must be lettered A, B, C... in order on the Parts sheet; part " & OptionLetter(n) & _
                        " is """ & p("label") & """ (Parts row " & p("row") & ")."
                    Exit For
                End If
                If p.Exists("points") Then sum = sum + p("points")
                If ClassRank(p("markId")) > top Then top = ClassRank(p("markId")): qn("markId") = p("markId")
            Next
            If qn.Exists("points") And PartsHavePoints(qn) Then
                If Abs(sum - qn("points")) > 0.001 Then probs.Add s & "the parts' points (" & PointsText(sum) & _
                    ") don't add up to the question's (" & PointsText(qn("points")) & ")."
            End If
        End If
        For Each g In qn("opts").Keys
            n = 0: nCorrect = 0
            For Each o In qn("opts")(g)
                n = n + 1
                If o("label") <> OptionLetter(n) Then
                    probs.Add s & "the options of " & IIf(Len(g) > 0, "part " & g, "the question") & " must be lettered A, B, C... in " & _
                        "order on the Options sheet; option " & OptionLetter(n) & " is """ & ClipText(o("label"), 10) & """ (Options row " & o("row") & ")."
                    Exit For
                End If
                If Nz(o("correct"), False) Then nCorrect = nCorrect + 1
            Next
            If nCorrect > 1 Then probs.Add s & IIf(Len(g) > 0, "part " & g, "the question") & " has " & nCorrect & _
                " options marked correct; mark only one."
        Next
    Next
    If probs.Count >= MAX_PROBLEMS Then Exit Sub

    ' What import will skip: a question already in the bank in its topic, or
    ' a second copy in the file (DupKey: the same text, parts and options).
    Set bank = BankTexts(qs)
    Set seen = CreateObject("Scripting.Dictionary")
    For Each qn In qs
        If qn("topicId") <> 0 And Not IsNull(qn("text")) Then
            key = FileDupKey(qn)
            If bank.Exists(key) Then
                qn("skip") = "already in the bank as " & bank(key)
            ElseIf seen.Exists(key) Then
                qn("skip") = "the same as question " & seen(key) & " in this file"
            Else
                seen.Add key, qn("id")
            End If
        End If
    Next

    ' A topic turned off, or a reference turned off in the bank, takes no
    ' new question (the editor offers neither); a skipped one is no problem.
    For Each qn In qs
        If Len(qn("skip")) = 0 Then
            s = SH_Q & " row " & qn("row") & " (question " & qn("id") & "): "
            If qn("topicOff") Then probs.Add s & "Topic """ & ClipText(qn("topicText"), 60) & """ is turned off. Turn it on " & _
                "(Topics page) or use another topic."
            For Each o In qn("cites")
                If refs(o("ref"))("bankId") <> 0 Then
                    If Not Nz(DLookup("IsActive", "tblSourceRefs", "RefID=" & refs(o("ref"))("bankId")), True) Then
                        probs.Add s & "it cites reference " & o("ref") & " (" & ClipText(refs(o("ref"))("title"), 60) & _
                            "), which is turned off in the bank, so a new question can't cite it. Turn it on (References tab) " & _
                            "or take the citation out."
                    End If
                End If
            Next
        End If
    Next
    If probs.Count > 0 Then Exit Sub

    ' What a draft will still need before it can be submitted.
    For Each qn In qs
        If Len(qn("skip")) = 0 Then
            s = "Question " & qn("id") & " (row " & qn("row") & "): "
            If qn("cites").Count = 0 Then warns.Add s & "cites no reference."
            If qn("type") = QT_MC Then
                If qn("parts").Count = 0 Then
                    WarnOptions qn, "", "", s, warns
                Else
                    For Each p In qn("parts")
                        WarnOptions qn, p("label"), "part " & p("label") & " ", s, warns
                    Next
                End If
            End If
            For Each o In qn("cites")
                refs(o("ref"))("used") = True
            Next
        End If
    Next
    For Each k In refs.Keys
        If refs(k)("used") And refs(k)("bankId") <> 0 Then
            If LCase$(Nz(DLookup("Revision", "tblSourceRefs", "RefID=" & refs(k)("bankId")), "")) <> LCase$(refs(k)("rev")) Then
                warns.Add "Reference " & k & " (" & ClipText(refs(k)("title"), 60) & ") is cited as the bank has it, revision " & _
                    Nz(DLookup("Revision", "tblSourceRefs", "RefID=" & refs(k)("bankId")), "(none)") & "; the file says " & _
                    IIf(Len(refs(k)("rev")) > 0, refs(k)("rev"), "(none)") & "."
            End If
        End If
    Next
    Set m_plan = CreateObject("Scripting.Dictionary")
    m_plan("File") = fileName
    Set m_plan("Qs") = qs
    Set m_plan("Refs") = refs
End Sub

' True when every part of qn has points that passed PointsProblem.
Private Function PartsHavePoints(qn As Object) As Boolean
    Dim p As Object
    For Each p In qn("parts")
        If Not p.Exists("points") Then Exit Function
    Next
    PartsHavePoints = True
End Function

' Warn when an option group of a multiple choice question (the question
' itself, part "") has fewer than two options or none marked correct.
Private Sub WarnOptions(qn As Object, ByVal part As String, ByVal what As String, ByVal s As String, warns As Collection)
    Dim n As Long, nCorrect As Long, o As Object
    If qn("opts").Exists(part) Then
        For Each o In qn("opts")(part)
            n = n + 1
            If Nz(o("correct"), False) Then nCorrect = nCorrect + 1
        Next
    End If
    If n < 2 Then
        warns.Add s & what & "has " & n & IIf(n = 1, " option", " options") & "; it needs at least two."
    ElseIf nCorrect = 0 Then
        warns.Add s & what & "has no option marked correct."
    End If
End Sub

' The question a Parts, Options or Citations row points at, or Nothing
' (with a problem) when the Questions sheet has no such Question ID.
Private Function QuestionOf(byId As Object, ByVal v As String, ByVal s As String, probs As Collection) As Object
    Dim id As String
    id = IdText(v)
    If Len(id) = 0 Then
        probs.Add s & "Question ID is empty."
    ElseIf Not byId.Exists(id) Then
        probs.Add s & "Question ID " & ClipText(id, 20) & " isn't on the Questions sheet."
    Else
        Set QuestionOf = byId(id)
    End If
End Function

' A Question ID or Reference ID as compared: trimmed, and digits without
' their leading zeros (007 and 7 are the same ID). Any other text is
' compared as it is.
Private Function IdText(ByVal v As String) As String
    v = Trim$(v)
    If Len(v) > 0 And Len(v) <= 9 And Not v Like "*[!0-9]*" Then v = CStr(CLng(v))
    IdText = v
End Function

' The ID of the row of tbl whose field is nm (case and outer spaces
' ignored), or 0.
Private Function LookupID(ByVal idField As String, ByVal tbl As String, ByVal fld As String, ByVal nm As String) As Long
    nm = Trim$(nm)
    If Len(nm) = 0 Then Exit Function
    LookupID = Nz(DLookup(idField, tbl, fld & "=" & Q(nm)), 0)
End Function

' A marking's ClassLevelID, by its name (SECRET) or its portion mark ((S)),
' or 0.
Private Function MarkingID(ByVal nm As String) As Long
    MarkingID = LookupID("ClassLevelID", "tblClassLevels", "Marking", nm)
    If MarkingID = 0 Then MarkingID = LookupID("ClassLevelID", "tblClassLevels", "PortionMark", nm)
End Function

' The type name a cell gives (case ignored; "True/False" also as "True or
' false" or "TF", and "MC", "SA"), or "".
Private Function TypeFromText(ByVal v As String) As String
    Select Case LCase$(Replace(Replace(Trim$(v), " ", ""), "-", ""))
        Case LCase$(Replace(QT_MC, " ", "")), "mc": TypeFromText = QT_MC
        Case "true/false", "trueorfalse", "tf", "t/f": TypeFromText = QT_TF
        Case LCase$(Replace(QT_SA, " ", "")), "sa": TypeFromText = QT_SA
    End Select
End Function

' Correct as True (Yes, Y, True, X, 1), False (empty, No, N, False, 0), or
' Null for anything else.
Private Function YesText(ByVal v As String) As Variant
    Select Case LCase$(Trim$(v))
        Case "yes", "y", "true", "x", "1": YesText = True
        Case "", "no", "n", "false", "0": YesText = False
        Case Else: YesText = Null
    End Select
End Function

' A points cell for PointsProblem: Null when empty; the number when it is
' digits with at most one dot, or a comma followed by one or two digits,
' for the decimals (4, 1.5, 1,5), read the same on any computer; else the
' text, which PointsProblem names as not a number ("1e400", "$5", "&H10"
' and "1,000" are text here).
Private Function NumberOrNull(ByVal v As String) As Variant
    Dim t As String
    v = Trim$(v)
    t = v
    If t Like "*,#" Or t Like "*,##" Then t = Replace(t, ",", ".")
    If Len(v) = 0 Then
        NumberOrNull = Null
    ElseIf Len(t) <= 12 And Not t Like "*[!0-9.]*" And Len(t) - Len(Replace(t, ".", "")) <= 1 And t Like "*[0-9]*" Then
        NumberOrNull = Val(t)
    Else
        NumberOrNull = v
    End If
End Function

' A reference's fields' sizes and its title, as RefProblem, without its
' look in the bank for the same reference (import matches that one).
Private Function RefFieldProblem(ByVal title As String, ByVal number As String, ByVal rev As String) As String
    If Len(title) = 0 Then
        RefFieldProblem = "Title is empty."
    ElseIf Len(title) > 200 Then
        RefFieldProblem = "Title can be at most 200 characters; it is " & Len(title) & "."
    ElseIf Len(number) > 100 Then
        RefFieldProblem = "Number can be at most 100 characters; it is " & Len(number) & "."
    ElseIf Len(rev) > 50 Then
        RefFieldProblem = "Revision can be at most 50 characters."
    End If
End Function

' Text as compared for duplicates: lower case, every run of spaces and line
' breaks as one space.
Private Function TextKey(ByVal v As Variant) As String
    Dim s As String
    s = LCase$(Replace(Replace(Nz(v, ""), vbCr, " "), vbLf, " "))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    TextKey = Trim$(s)
End Function

' A question's duplicate key: its topic, its text, its parts' texts in
' order and its options' texts (sorted, so their order doesn't matter),
' each as TextKey. Two questions with the generic text "Answer each part."
' but different parts are not the same question.
Private Function DupKey(ByVal topicId As Long, ByVal text As Variant, parts As Collection, opts As Collection) As String
    Dim s As String, v As Variant, a() As String, i As Long, j As Long, t As String
    s = topicId & "|" & TextKey(text)
    For Each v In parts
        s = s & Chr$(1) & "P" & TextKey(v)
    Next
    If opts.Count > 0 Then
        ReDim a(1 To opts.Count)
        For i = 1 To opts.Count
            a(i) = TextKey(opts(i))
        Next
        For i = 2 To opts.Count               ' a few options: insertion sort
            t = a(i): j = i - 1
            Do While j >= 1
                If a(j) <= t Then Exit Do
                a(j + 1) = a(j): j = j - 1
            Loop
            a(j + 1) = t
        Next
        For i = 1 To opts.Count
            s = s & Chr$(1) & "O" & a(i)
        Next
    End If
    DupKey = s
End Function

' DupKey of a question read from the file (CheckRows).
Private Function FileDupKey(qn As Variant) As String
    Dim parts As New Collection, opts As New Collection, p As Variant, g As Variant, o As Variant
    For Each p In qn("parts")
        parts.Add p("text")
    Next
    For Each g In qn("opts").Keys
        For Each o In qn("opts")(g)
            opts.Add o("text")
        Next
    Next
    FileDupKey = DupKey(qn("topicId"), qn("text"), parts, opts)
End Function

' The bank's questions in the topics qs use, every version that isn't
' retired: DupKey -> "Q12 v2".
Private Function BankTexts(qs As Collection) As Object
    Dim out As Object, ids As Object, qn As Variant, rs As DAO.Recordset, where As String, id As Variant
    Dim parts As Object, opts As Object, info As Object, v As Variant, k As String
    Set out = CreateObject("Scripting.Dictionary")
    Set ids = CreateObject("Scripting.Dictionary")
    For Each qn In qs
        If qn("topicId") <> 0 Then ids(CStr(qn("topicId"))) = True
    Next
    If ids.Count = 0 Then Set BankTexts = out: Exit Function
    where = "q.StatusID<>" & QStatusID(QS_RETIRED) & " AND q.TopicID IN (" & Join(ids.Keys, ",") & ")"
    Set info = CreateObject("Scripting.Dictionary")     ' QuestionID -> Array(TopicID, text, "Q12 v2")
    Set parts = CreateObject("Scripting.Dictionary")    ' QuestionID -> Collection of part texts
    Set opts = CreateObject("Scripting.Dictionary")     ' QuestionID -> Collection of option texts
    Set rs = CurrentDb.OpenRecordset("SELECT q.QuestionID, q.TopicID, q.QuestionKey, q.VersionNo, q.QuestionText " & _
        "FROM tblQuestions AS q WHERE " & where, dbOpenSnapshot)
    Do While Not rs.EOF
        info(CLng(rs!QuestionID)) = Array(CLng(rs!TopicID), rs!QuestionText.Value, QuestionRef(rs!QuestionKey, rs!VersionNo))
        parts.Add CLng(rs!QuestionID), New Collection
        opts.Add CLng(rs!QuestionID), New Collection
        rs.MoveNext
    Loop
    rs.Close
    Set rs = CurrentDb.OpenRecordset("SELECT p.QuestionID, p.PartText FROM tblQuestionParts AS p INNER JOIN tblQuestions AS q " & _
        "ON p.QuestionID=q.QuestionID WHERE " & where & " ORDER BY p.QuestionID, p.SortOrder, p.PartLabel", dbOpenSnapshot)
    Do While Not rs.EOF                     ' a question saved since the first query isn't in info: left out
        If parts.Exists(CLng(rs!QuestionID)) Then parts(CLng(rs!QuestionID)).Add rs!PartText.Value
        rs.MoveNext
    Loop
    rs.Close
    Set rs = CurrentDb.OpenRecordset("SELECT o.QuestionID, o.OptionText FROM tblMCOptions AS o INNER JOIN tblQuestions AS q " & _
        "ON o.QuestionID=q.QuestionID WHERE " & where, dbOpenSnapshot)
    Do While Not rs.EOF
        If opts.Exists(CLng(rs!QuestionID)) Then opts(CLng(rs!QuestionID)).Add rs!OptionText.Value
        rs.MoveNext
    Loop
    rs.Close
    For Each id In info.Keys
        v = info(id)
        k = DupKey(v(0), v(1), parts(id), opts(id))
        If Not out.Exists(k) Then out.Add k, v(2)
    Next
    Set BankTexts = out
End Function

' Lines joined by CR LF, at most maxLines of them, then "and n more".
Private Function JoinLines(lines As Collection, ByVal maxLines As Long) As String
    Dim s As String, i As Long
    For i = 1 To lines.Count
        If i > maxLines Then
            s = s & vbCrLf & "and " & (lines.Count - maxLines) & " more"
            Exit For
        End If
        s = s & IIf(i > 1, vbCrLf, "") & Replace(Replace(Replace(lines(i), vbCrLf, " "), vbCr, " "), vbLf, " ")
    Next
    JoinLines = s
End Function

' The first SHOWN_LINES lines of a list from CheckBankFile or
' BankImportSummary, for a message, with "and n more" for the rest.
'   s   lines joined by CR LF
' Errors: none.
Public Function FirstLines(ByVal s As String) As String
    Dim lines() As String, i As Long, out As String
    If Len(s) = 0 Then Exit Function
    lines = Split(s, vbCrLf)
    For i = 0 To UBound(lines)
        If i >= SHOWN_LINES Then
            out = out & vbCrLf & "and " & (UBound(lines) - i + 1) & " more"
            Exit For
        End If
        out = out & IIf(i > 0, vbCrLf, "") & "- " & lines(i)
    Next
    FirstLines = out
End Function

' How many lines a list from CheckBankFile has. Errors: none.
Public Function LineCount(ByVal s As String) As Long
    If Len(s) > 0 Then LineCount = UBound(Split(s, vbCrLf)) + 1
End Function

' ---------------------------------------------------------------------------
' Import: what it will do, and doing it
' ---------------------------------------------------------------------------

' What ImportCheckedBank would do with the file CheckBankFile last passed:
' how many questions and in which topics, the references matched and added,
' the questions skipped and why, and what the drafts still need. "" when
' there is nothing checked.
' Errors: the engine's.
Public Function BankImportSummary() As String
    Dim qn As Variant, k As Variant, nNew As Long, nSkip As Long, topics As Object, s As String, skipped As String
    Dim matched As Long, added As Long, refs As Object, warns As Collection, t As Variant, i As Long
    If m_plan Is Nothing Then Exit Function
    Set topics = CreateObject("Scripting.Dictionary")
    For Each qn In m_plan("Qs")
        If Len(qn("skip")) = 0 Then
            nNew = nNew + 1
            t = DLookup("TopicName", "tblTopics", "TopicID=" & qn("topicId"))
            topics(t) = topics(t) + 1
        Else
            nSkip = nSkip + 1
            If nSkip <= SHOWN_LINES Then skipped = skipped & vbCrLf & "- Question " & qn("id") & " (row " & qn("row") & "): " & qn("skip")
        End If
    Next
    If nSkip > SHOWN_LINES Then skipped = skipped & vbCrLf & "and " & (nSkip - SHOWN_LINES) & " more"
    Set refs = m_plan("Refs")
    For Each k In refs.Keys
        If refs(k)("used") Then
            If refs(k)("bankId") <> 0 Then matched = matched + 1 Else added = added + 1
        End If
    Next
    If nNew = 0 Then
        s = "Every question in " & m_plan("File") & " is already in the bank, so there is nothing to import." & skipped
        BankImportSummary = s
        Exit Function
    End If
    s = "Import " & nNew & IIf(nNew = 1, " question", " questions") & " from " & m_plan("File") & " as drafts?" & vbCrLf
    For Each t In topics.Keys
        i = i + 1
        If i > SHOWN_LINES Then s = s & vbCrLf & "and " & (topics.Count - SHOWN_LINES) & " more topics": Exit For
        s = s & vbCrLf & "- " & t & ": " & topics(t)
    Next
    If matched + added > 0 Then
        s = s & vbCrLf & vbCrLf & "References: " & IIf(matched > 0, matched & " already in the bank", "") & _
            IIf(matched > 0 And added > 0, ", ", "") & IIf(added > 0, added & " new, added to the shared list", "") & "."
    End If
    If nSkip > 0 Then s = s & vbCrLf & vbCrLf & "Skipped, " & nSkip & ":" & skipped
    Set warns = m_plan("Warnings")
    If warns.Count > 0 Then
        s = s & vbCrLf & vbCrLf & "Before they can be submitted, " & warns.Count & IIf(warns.Count = 1, " thing needs", " things need") & _
            " fixing in the editor:" & vbCrLf & FirstLines(JoinLines(warns, warns.Count))
    End If
    s = s & vbCrLf & vbCrLf & "They wait as drafts until an SME on their topic submits them."
    BankImportSummary = s
End Function

' True when a checked file waits to be imported (CheckBankFile passed and
' it has a question that isn't skipped). Errors: none.
Public Function BankImportReady() As Boolean
    Dim qn As Variant
    If m_plan Is Nothing Then Exit Function
    For Each qn In m_plan("Qs")
        If Len(qn("skip")) = 0 Then BankImportReady = True: Exit Function
    Next
End Function

' Forget the checked file (the user cancelled). Errors: none.
Public Sub ClearCheckedBank()
    Set m_plan = Nothing
End Sub

' Write the file CheckBankFile last passed into the bank, in one
' transaction: each question that isn't skipped as a new draft (version 1,
' the signed-in admin its author) with its parts, options and citations,
' the references it cites that the bank doesn't have added to the shared
' list, an Imported audit row per question, an Added row per reference and
' one Imported row for the file. Duplicates are looked for again inside the
' transaction, so a question saved meanwhile is skipped too; a topic or a
' reference turned off meanwhile stops the whole import.
'   skippedNow   returns how many were skipped only now (saved meanwhile)
'   problem      returns "" or, when nothing was imported because something
'                changed since the check, why, for the user
' Returns how many questions were imported (0 when nothing was checked).
' Errors: ERR_NOT_ALLOWED unless the signed-in person is an active admin
'   now; others pass up through RaiseAgain after the rollback. The checked
'   file is forgotten either way.
Public Function ImportCheckedBank(ByRef skippedNow As Long, ByRef problem As String) As Long
    Dim ws As DAO.Workspace, db As DAO.Database, rs As DAO.Recordset, inTrans As Boolean, plan As Object
    Dim qn As Variant, p As Variant, o As Variant, g As Variant, refs As Object, k As Variant, bank As Object
    Dim qid As Long, partIds As Object, n As Long, pos As Long, pid As Long, fileName As String, key As String
    Dim num As Long, desc As String, src As String
    skippedNow = 0
    problem = ""
    Set plan = m_plan
    Set m_plan = Nothing
    If plan Is Nothing Then Exit Function
    RequireExamAdmin
    pid = SignedInPersonID()
    fileName = plan("File")
    Set refs = plan("Refs")
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    On Error GoTo Fail
    ' A whole bank is thousands of rows in one transaction, more than the
    ' engine's usual 9,500 locks; this lasts until Access closes.
    DBEngine.SetOption dbMaxLocksPerFile, 500000
    ws.BeginTrans: inTrans = True
    Set bank = BankTexts(plan("Qs"))
    ' The references cited: the bank's own when it has them (looked up again),
    ' else added now.
    For Each k In refs.Keys
        If refs(k)("used") Then
            refs(k)("bankId") = Nz(DLookup("RefID", "tblSourceRefs", "RefTitle=" & Q(refs(k)("title")) & " AND Nz(RefNumber,'')=" & _
                Q(refs(k)("number"))), 0)
        End If
    Next
    For Each qn In plan("Qs")
        If Len(qn("skip")) = 0 Then
            key = FileDupKey(qn)
            If bank.Exists(key) Then
                skippedNow = skippedNow + 1
            ElseIf Not Nz(DLookup("IsActive", "tblTopics", "TopicID=" & qn("topicId")), False) Then
                problem = "Topic """ & qn("topicText") & """ was turned off after the file was checked, so nothing was imported. " & _
                    "Turn it on (Topics page) and import the file again."
                ws.Rollback: inTrans = False
                ImportCheckedBank = 0: skippedNow = 0
                Exit Function
            Else
                bank(key) = "imported"
                For Each o In qn("cites")
                    If refs(o("ref"))("bankId") <> 0 Then
                        If Not Nz(DLookup("IsActive", "tblSourceRefs", "RefID=" & refs(o("ref"))("bankId")), True) Then
                            problem = "Reference """ & refs(o("ref"))("title") & """ was turned off after the file was checked, so " & _
                                "nothing was imported. Turn it on (References tab) and import the file again."
                            ws.Rollback: inTrans = False
                            ImportCheckedBank = 0: skippedNow = 0
                            Exit Function
                        End If
                    End If
                Next
                For Each o In qn("cites")
                    If refs(o("ref"))("bankId") = 0 Then refs(o("ref"))("bankId") = AddImportedRef(db, refs(o("ref")), fileName)
                Next
                qid = WriteQuestion(db, qn, pid)
                ' Parts, then options lettered in their order, then citations.
                Set partIds = CreateObject("Scripting.Dictionary")
                n = 0
                Set rs = db.OpenRecordset("SELECT * FROM tblQuestionParts WHERE False", dbOpenDynaset)
                For Each p In qn("parts")
                    n = n + 1
                    rs.AddNew
                    rs!QuestionID = qid
                    rs!PartLabel = p("label")
                    rs!PartText = p("text")
                    rs!Points = p("points")
                    rs!ClassLevelID = p("markId")
                    rs!ReferenceAnswer = p("answer")
                    rs!SortOrder = n
                    rs.Update
                    rs.Bookmark = rs.LastModified
                    partIds(p("label")) = CLng(rs!PartID)
                Next
                rs.Close
                Set rs = db.OpenRecordset("SELECT * FROM tblMCOptions WHERE False", dbOpenDynaset)
                For Each g In qn("opts").Keys
                    pos = 0
                    For Each o In qn("opts")(g)
                        pos = pos + 1
                        rs.AddNew
                        rs!QuestionID = qid
                        If Len(g) > 0 Then rs!PartID = partIds(g)
                        rs!OptionLabel = OptionLetter(pos)
                        rs!OptionText = o("text")
                        rs!IsCorrect = Nz(o("correct"), False)
                        rs!SortOrder = pos
                        rs.Update
                    Next
                Next
                rs.Close
                Set rs = Nothing
                For Each o In qn("cites")
                    db.Execute "INSERT INTO tblQuestionRefs (QuestionID, RefID, Paragraph) VALUES (" & qid & ", " & _
                        refs(o("ref"))("bankId") & ", " & IIf(IsNull(o("para")), "Null", Q(o("para"))) & ")", dbFailOnError
                Next
                AddAudit ITEM_QUESTION, qid, 1, AU_IMPORTED, , "From " & fileName & ", question " & qn("id") & " (Questions row " & qn("row") & ")"
                ImportCheckedBank = ImportCheckedBank + 1
            End If
        End If
    Next
    If ImportCheckedBank > 0 Then
        AddAudit ITEM_BANK, 0, Null, AU_IMPORTED, , ImportCheckedBank & IIf(ImportCheckedBank = 1, " question", " questions") & _
            " from " & fileName & " as drafts"
    End If
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    num = Err.Number: desc = Err.Description: src = Err.Source
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If inTrans Then ws.Rollback
    On Error GoTo 0
    ImportCheckedBank = 0
    Err.Number = num: Err.Description = desc: Err.Source = src
    RaiseAgain "modBank.ImportCheckedBank", fileName
End Function

' A checked question as a new draft (inside ImportCheckedBank's
' transaction): version 1 of its own key, current, by pid now. Returns its
' QuestionID.
Private Function WriteQuestion(db As DAO.Database, qn As Variant, ByVal pid As Long) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM tblQuestions WHERE False", dbOpenDynaset)
    rs.AddNew
    rs!VersionNo = 1
    rs!StatusID = QStatusID(QS_DRAFT)
    rs!AuthorPersonID = pid
    rs!CreatedOn = Now()
    rs!IsCurrent = True
    rs!TopicID = qn("topicId")
    rs!QuestionTypeID = QTypeID(qn("type"))
    rs!DifficultyID = qn("diffId")
    rs!ClassLevelID = qn("markId")           ' with parts, CheckRows set it to the highest part's
    rs!QuestionText = qn("text")
    rs!TotalPoints = qn("points")
    rs!ReferenceAnswer = qn("answer")
    rs.Update
    rs.Bookmark = rs.LastModified
    WriteQuestion = rs!QuestionID
    rs.Edit
    rs!QuestionKey = rs!QuestionID
    rs.Update
    rs.Close
End Function

' Add a reference the file cites that the bank doesn't have, turned on, with
' an Added audit row (inside ImportCheckedBank's transaction). Returns its
' RefID.
Private Function AddImportedRef(db As DAO.Database, r As Object, ByVal fileName As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM tblSourceRefs WHERE False", dbOpenDynaset)
    rs.AddNew
    rs!RefTitle = r("title")
    rs!RefNumber = NullIfEmpty(r("number"))
    rs!Revision = NullIfEmpty(r("rev"))
    rs!IsActive = True
    rs.Update
    rs.Bookmark = rs.LastModified
    AddImportedRef = rs!RefID
    rs.Close
    AddAudit ITEM_REFERENCE, AddImportedRef, Null, AU_ADDED, , RefText(AddImportedRef) & " (imported from " & fileName & ")"
End Function
