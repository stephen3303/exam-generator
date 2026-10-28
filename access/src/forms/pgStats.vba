Option Compare Database
Option Explicit

' Statistics (see Build_pgStats). The figures for the exams the SM gave final
' approval to over a period (the last 6 or 12 months, or two dates) from
' every template or one (modStats.RunStats, which works them out into this
' front end's local tables): the exams, questions given and references cited
' on the tiles, the topics leaning on one reference for WEAK_SHARE of their
' questions in the alert, each topic with its most cited reference, and the
' picked topic's references on the right. Any change to the period or the
' template works the figures out again. Print and Export take the topics
' table, labelled with the period and template it was worked out for. While
' the dates are missing or the wrong way round, no figures show and Print
' and Export are off.
'
' Depends on: modSignIn, modStats, modUI; the kit's modGridOut,
' modKitFiles, modKitUI.

Private m_from As Date          ' the period shown, whole days, both included (as RunStats last ran)
Private m_to As Date
Private m_tplName As String     ' the template it ran for ("": every template)
Private m_topic As Long         ' the topic whose references show (0: none)

Public Property Get HelpTopic() As String
    HelpTopic = "stats"
End Property

Private Sub Form_Load()
    Dim c As Variant
    On Error GoTo Err_Handler
    Me.dtFrom = DateAdd("m", -12, Date): Me.dtTo = Date
    If Not DataLinked() Then
        ShowAlert Me, "alWeak", "This copy was built without the data file, so there are no figures to show.", "warning"
        For Each c In Array("tileExams", "tileGiven", "tileRefs")
            Me.Controls(CStr(c)).Value = Null
            Me.Controls(c & "_cap").Caption = ""
        Next
        For Each c In Array("cboTemplate", "dtFrom", "dtTo", "cal_dtFrom", "cal_dtTo", "btnPrint", "btnExport")
            SetEnabled Me.Controls(CStr(c)), False
        Next
        Me.segPeriod_val.Locked = True      ' SegmentClick ignores a locked value, so no period can run RunStats
        ShowEmpty Me, "emptyTopics", True, "No data file linked", "There are no exams to count."
        SetVisible Me.sfTopics, False
        SetVisible Me.sfRefs, False
        ShowEmpty Me, "emptyRefs", True
        Exit Sub
    End If
    Me.cboTemplate.RowSource = StatTemplatesSql()
    Me.cboTemplate = 0
    SetSegment Me, "segPeriod", "m6"
    ShowDates
    RefreshStats
    Exit Sub
Err_Handler:
    HandleError "pgStats.Form_Load"
End Sub

' ---------------------------------------------------------------------------
' The period and the template
' ---------------------------------------------------------------------------

' The period's Segmented control (onChange).
Public Sub PeriodChanged()
    On Error GoTo Err_Handler
    ShowDates
    RefreshStats
    Exit Sub
Err_Handler:
    HandleError "pgStats.PeriodChanged"
End Sub

Private Sub dtFrom_AfterUpdate()
    On Error GoTo Err_Handler
    RefreshStats
    Exit Sub
Err_Handler:
    HandleError "pgStats.dtFrom_AfterUpdate"
End Sub

Private Sub dtTo_AfterUpdate()
    On Error GoTo Err_Handler
    RefreshStats
    Exit Sub
Err_Handler:
    HandleError "pgStats.dtTo_AfterUpdate"
End Sub

' The calendar buttons: the kit's picker sets the box without its
' AfterUpdate, so the figures are worked out again here.
Private Sub cal_dtFrom_Click()
    On Error GoTo Err_Handler
    PickDate Me, "dtFrom"
    RefreshStats
    Exit Sub
Err_Handler:
    HandleError "pgStats.cal_dtFrom_Click"
End Sub

Private Sub cal_dtTo_Click()
    On Error GoTo Err_Handler
    PickDate Me, "dtTo"
    RefreshStats
    Exit Sub
Err_Handler:
    HandleError "pgStats.cal_dtTo_Click"
End Sub

Private Sub cboTemplate_AfterUpdate()
    On Error GoTo Err_Handler
    RefreshStats
    Exit Sub
Err_Handler:
    HandleError "pgStats.cboTemplate_AfterUpdate"
End Sub

' The dates follow the period: set and locked for 6 or 12 months, open for
' Dates (keeping the last ones shown).
Private Sub ShowDates()
    Dim k As String, custom As Boolean, c As Variant
    k = SegmentValue(Me, "segPeriod")
    custom = (k = "dates")
    If Not custom Then
        Me.dtTo = Date
        Me.dtFrom = DateAdd("m", IIf(k = "m6", -6, -12), Date) + 1
    End If
    For Each c In Array("dtFrom", "dtTo", "cal_dtFrom", "cal_dtTo")
        SetEnabled Me.Controls(CStr(c)), custom
    Next
End Sub

' ---------------------------------------------------------------------------
' The figures
' ---------------------------------------------------------------------------

' Work the figures out for the period and template shown, or say why not
' and show none.
Private Sub RefreshStats()
    Dim n As Long, s As String, nq As Long, d1 As Date, d2 As Date
    If Not IsDate(Me.dtFrom) Or Not IsDate(Me.dtTo) Then
        ClearFigures "Enter both dates to see the figures.", "info"
        Exit Sub
    End If
    d1 = DateValue(Me.dtFrom): d2 = DateValue(Me.dtTo)
    If d1 > d2 Then
        ClearFigures "The first date is after the second. Swap them to see the figures.", "warning"
        Exit Sub
    End If
    ClearFigures "Working out the figures...", "info"  ' if RunStats fails, nothing stale shows and Print and Export stay off
    RunStats d1, d2, Nz(Me.cboTemplate, 0)
    m_from = d1: m_to = d2                  ' what the figures are for, for the tile, Print and Export
    m_tplName = IIf(Nz(Me.cboTemplate, 0) = 0, "", Nz(Me.cboTemplate.Column(1), "one template"))
    SetTile "tileExams", StatExams(), PeriodText()
    nq = StatQuestions()
    SetTile "tileGiven", nq, "a question on three exams counts three times"
    SetTile "tileRefs", StatReferences(), "different references the questions cite"
    s = WeakTopicsText(n, 2)                ' two lines, long names cut short, fit the titled alert's one line
    If nq = 0 Then
        ShowAlert Me, "alWeak", "No exam had final approval in this period" & IIf(Len(m_tplName) > 0, " from this template", "") & ".", "info"
    ElseIf n = 0 Then
        ShowAlert Me, "alWeak", "No topic takes " & Format$(WEAK_SHARE, "0%") & " or more of its questions from one reference.", "success"
    Else
        ShowAlert Me, "alWeak", Replace(Replace(s, "&", "&&"), vbCrLf, "; "), "warning", "Weak reference coverage"
    End If
    Me.sfTopics.Form.RecordSource = StatTopicsSql()
    n = Me.sfTopics.Form.RecordsetClone.RecordCount
    ShowEmpty Me, "emptyTopics", n = 0, "No approved exams in this period", _
        "Pick a longer period or every template. Only exams the SM gave final approval to count."   ' ClearFigures may have changed them
    SetVisible Me.sfTopics, n > 0
    SetEnabled Me.btnPrint, n > 0
    SetEnabled Me.btnExport, n > 0
    If GridSelect(Me.sfTopics, m_topic) Then
        ShowRefs m_topic
    Else
        ShowRefs Nz(SelectFirstRow(Me.sfTopics), 0)
    End If
End Sub

' No figures, with msg in the alert: the tiles empty, both tables hidden
' behind their empty states, and Print and Export off, so nothing from an
' earlier period shows or prints under the new dates.
Private Sub ClearFigures(ByVal msg As String, ByVal vrnt As String)
    Dim c As Variant
    ShowAlert Me, "alWeak", msg, vrnt
    For Each c In Array("tileExams", "tileGiven", "tileRefs")
        SetTile CStr(c), Null, ""
    Next
    ShowEmpty Me, "emptyTopics", True, "No figures", "Enter the period's dates, the earlier one first."
    SetVisible Me.sfTopics, False
    SetEnabled Me.btnPrint, False
    SetEnabled Me.btnExport, False
    Me.lblRefsHdr.Caption = "References"
    ShowEmpty Me, "emptyRefs", True, "No references", "Pick a topic to see which references its questions cite."
    SetVisible Me.sfRefs, False
End Sub

Private Sub SetTile(ByVal nm As String, ByVal v As Variant, ByVal cap As String)
    Me.Controls(nm).Value = v
    Me.Controls(nm & "_cap").Caption = Replace(cap, "&", "&&")
End Sub

' "1 May 2026 to 28 Oct 2026, every template" (or the template's name), as
' RunStats last ran: the combo may have moved on since.
Private Function PeriodText() As String
    PeriodText = Format$(m_from, "d mmm yyyy") & " to " & Format$(m_to, "d mmm yyyy") & ", " & _
        IIf(Len(m_tplName) = 0, "every template", m_tplName)
End Function

' The references of topic topicId on the right (0: none).
Private Sub ShowRefs(ByVal topicId As Long)
    Dim n As Long
    m_topic = topicId
    Me.sfRefs.Form.RecordSource = StatRefsSql(topicId)
    n = Me.sfRefs.Form.RecordsetClone.RecordCount
    Me.lblRefsHdr.Caption = IIf(topicId = 0, "References", "References: " & _
        Replace(Nz(DLookup("TopicName", "ws_topic", "id=" & topicId), ""), "&", "&&"))
    If topicId = 0 Then
        ShowEmpty Me, "emptyRefs", True, "No references", "Pick a topic to see which references its questions cite."
    Else
        ShowEmpty Me, "emptyRefs", n = 0, "No references cited", "None of this topic's questions given cites a reference."
    End If
    SetVisible Me.sfRefs, n > 0
End Sub

' ---------------------------------------------------------------------------
' Buttons and rows
' ---------------------------------------------------------------------------

Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    If grid = "sfTopics" Then ShowRefs Nz(Me.sfTopics.Form.txtSel, 0)
    Exit Sub
Err_Handler:
    HandleError "pgStats.GridClick", grid
End Sub

Private Sub btnPrint_Click()
    On Error GoTo Err_Handler
    PrintGrid Me.sfTopics.Form, "Questions given per topic", PeriodText()
    Exit Sub
Err_Handler:
    HandleError "pgStats.btnPrint_Click"
End Sub

Private Sub btnExport_Click()
    On Error GoTo Err_Handler
    ExportGridAs Me.sfTopics.Form, "Exam statistics"
    Exit Sub
Err_Handler:
    HandleError "pgStats.btnExport_Click"
End Sub
