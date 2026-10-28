Option Compare Database
Option Explicit

' Admin: Logs (see Build_pgAdmLogs), for the Exam Generator admin. Two tabs
' over one search box and one day range (the last 30 days to start):
'   Audit trail  tblAuditTrail, newest first (modLogs.AuditLogSql): who did
'                what to which question, exam or reference, and as which
'                role. Read only: the app never changes or deletes it.
'   Errors       the error log (modLogs.ErrorListSql); Clear log empties it
'                after asking (modErrors.ClearErrorLog).
' Each list shows the newest LOG_MAX entries that fit, and the subtitle says
' when there are more. Typing in the search box waits for a pause (the
' form's Timer) before the list is loaded again. The picked entry shows in
' full on the right (Copy puts it on the clipboard); Export saves the open
' list.
'
' Depends on: modErrors, modLogs, modSignIn, modUI; the kit's modGridFilter,
' modKitFiles, modKitUI, modKitUtil.

Private Const MSG_TITLE As String = "Logs"
Private Const SEARCH_WAIT_MS As Long = 300      ' the pause in typing before the list is loaded again

Private m_tab As String         ' "a" Audit trail, "e" Errors
Private m_search As String
Private m_aid As Long           ' the audit entry shown (0: none)
Private m_eid As Long           ' the error entry shown (0: none)
Private m_ok As Boolean         ' a data file is linked and the signed-in person is an admin
Private m_total As Long         ' the entries the search and dates match in all

Public Property Get HelpTopic() As String
    HelpTopic = "logs"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    m_tab = "a"
    SelectTab Me, "a"
    Me.dtFrom = DateAdd("d", -29, Date): Me.dtTo = Date        ' the last 30 days, today included
    m_ok = DataLinked()
    If m_ok Then m_ok = IsExamAdmin()
    If Not m_ok Then
        SetEnabled Me.btnExport, False
        SetEnabled Me.btnClear, False
    End If
    ShowGroup
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.Form_Load"
End Sub

' ---------------------------------------------------------------------------
' Tabs, search and dates (the kit's builders call these)
' ---------------------------------------------------------------------------

Public Function TabPick(ByVal key As String)
    On Error GoTo Err_Handler
    If key = m_tab Then Exit Function
    m_tab = key
    SelectTab Me, key
    ShowGroup
    Exit Function
Err_Handler:
    HandleError "pgAdmLogs.TabPick", key
End Function

Public Sub SearchFor(ByVal text As String)
    On Error GoTo Err_Handler
    m_search = Trim$(text)
    Me.TimerInterval = SEARCH_WAIT_MS       ' Form_Timer loads the list once the typing pauses
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.SearchFor", text
End Sub

Private Sub Form_Timer()
    On Error GoTo Err_Handler
    Me.TimerInterval = 0
    RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.Form_Timer"
End Sub

' A column funnel changed the rows shown (GridHeaders' onFilter).
Public Sub ListFiltered(ByVal sfName As String)
    On Error GoTo Err_Handler
    AfterFilter
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.ListFiltered", sfName
End Sub

Private Sub dtFrom_AfterUpdate()
    On Error GoTo Err_Handler
    RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.dtFrom_AfterUpdate"
End Sub

Private Sub dtTo_AfterUpdate()
    On Error GoTo Err_Handler
    RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.dtTo_AfterUpdate"
End Sub

' The calendar buttons: the kit's picker sets the box without its
' AfterUpdate, so the list is loaded again here when the date changed.
Private Sub cal_dtFrom_Click()
    Dim was As Variant
    On Error GoTo Err_Handler
    was = Me.dtFrom
    PickDate Me, "dtFrom"
    If Nz(Me.dtFrom, 0) <> Nz(was, 0) Then RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.cal_dtFrom_Click"
End Sub

Private Sub cal_dtTo_Click()
    Dim was As Variant
    On Error GoTo Err_Handler
    was = Me.dtTo
    PickDate Me, "dtTo"
    If Nz(Me.dtTo, 0) <> Nz(was, 0) Then RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.cal_dtTo_Click"
End Sub

' ---------------------------------------------------------------------------
' Buttons and rows
' ---------------------------------------------------------------------------

Public Sub GridClick(ByVal grid As String, ByVal dbl As Boolean)
    On Error GoTo Err_Handler
    ShowEntry Nz(Me.Controls(grid).Form.txtSel, 0)
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.GridClick", grid
End Sub

Private Sub btnCopy_Click()
    On Error GoTo Err_Handler
    If Len(Nz(Me.txtDetail, "")) = 0 Then Exit Sub
    If KitCopyText(Me.txtDetail) Then
        KitToast "Copied", "success"
    Else
        KitMsgBox "Another program is using the clipboard. Try Copy again.", vbInformation, MSG_TITLE
    End If
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.btnCopy_Click"
End Sub

' Admins only, checked again: the lists hold every user's names and computers.
Private Sub btnExport_Click()
    On Error GoTo Err_Handler
    RequireExamAdmin
    If m_tab = "e" Then ExportGridAs Me.sfEList.Form, "Error log" Else ExportGridAs Me.sfList.Form, "Audit trail"
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.btnExport_Click"
End Sub

' Empty the error log, after asking.
Private Sub btnClear_Click()
    Dim n As Long
    On Error GoTo Err_Handler
    If DCount("*", "error_log") = 0 Then KitToast "The error log is already empty", "info": Exit Sub
    If KitMsgBox("Delete every entry in the error log, for every user? This can't be undone. Export it first to keep a copy.", _
            vbYesNo + vbExclamation + vbDefaultButton2, MSG_TITLE) <> vbYes Then Exit Sub
    n = ClearErrorLog()
    KitToast n & IIf(n = 1, " entry", " entries") & " deleted", "success"
    RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.btnClear_Click"
    On Error Resume Next                    ' show what is left; a second failure was already reported
    RefreshList
End Sub

' The empty states' button: clear the search, the dates and the column filters.
Private Sub emptyList_b_Click()
    On Error GoTo Err_Handler
    ShowAll
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.emptyList_b_Click"
End Sub

Private Sub emptyEList_b_Click()
    On Error GoTo Err_Handler
    ShowAll
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.emptyEList_b_Click"
End Sub

Private Sub ShowAll()
    ClearGridFilters Me.Controls(ListName())
    Me.dtFrom = Null: Me.dtTo = Null
    SetSearch Me, "txtSearch", ""           ' calls SearchFor, which sets m_search
    Me.TimerInterval = 0
    RefreshList
End Sub

' ---------------------------------------------------------------------------
' The lists and the entry
' ---------------------------------------------------------------------------

' The open tab's list subform's name.
Private Function ListName() As String
    ListName = IIf(m_tab = "e", "sfEList", "sfList")
End Function

' Show the open tab's controls and hide the other's, then load its list.
Private Sub ShowGroup()
    Dim nm As Variant
    For Each nm In Split(Me.grpA.Caption, ";")
        If Len(nm) > 0 Then SetVisible Me.Controls(nm), (m_tab = "a")
    Next
    For Each nm In Split(Me.grpE.Caption, ";")
        If Len(nm) > 0 Then SetVisible Me.Controls(nm), (m_tab = "e")
    Next
    ShowEmpty Me, IIf(m_tab = "e", "emptyList", "emptyEList"), False
    RefreshList
End Sub

' Load the open tab's list for the search and the dates.
Private Sub RefreshList()
    Dim sf As Control, n As Long, dFrom As Variant, dTo As Variant, nm As Variant
    Set sf = Me.Controls(ListName())
    If Not m_ok Then
        ' Nothing to sort or filter: the open tab's column captions and funnels go too.
        For Each nm In Split(Me.Controls(IIf(m_tab = "e", "grpE", "grpA")).Caption, ";")
            If Left$(nm, 3) = "hdr" Then SetVisible Me.Controls(nm), False
        Next
        SetVisible sf, False
        ShowEmpty Me, IIf(m_tab = "e", "emptyEList", "emptyList"), True, IIf(DataLinked(), "For the admin only", "No data file linked"), _
            IIf(DataLinked(), "Only the Exam Generator admin can read the logs.", "This copy was built without the data file, so there are no logs.")
        SetVisible Me.Controls(IIf(m_tab = "e", "emptyEList_b", "emptyList_b")), False
        Me.lblSubtitle.Caption = ""
        ShowEntry 0
        Exit Sub
    End If
    dFrom = Me.dtFrom: dTo = Me.dtTo
    If IsDate(dFrom) And IsDate(dTo) Then
        If DateValue(dFrom) > DateValue(dTo) Then
            If m_tab = "e" Then
                sf.Form.RecordSource = ErrorListSql("", Date + 1, Date)
            Else
                sf.Form.RecordSource = AuditLogSql("", Date + 1, Date)
            End If
            m_total = -1                            ' AfterFilter says why the list is empty
            AfterFilter
            Exit Sub
        End If
    End If
    If m_tab = "e" Then
        sf.Form.RecordSource = ErrorListSql(m_search, dFrom, dTo)
    Else
        sf.Form.RecordSource = AuditLogSql(m_search, dFrom, dTo)
    End If
    SetGridFilter sf, ""                    ' the column funnels, on the new rows
    GridHighlight sf, m_search
    ' Count them all only when the list may have been cut: with no funnel on,
    ' fewer than LOG_MAX rows shown is the whole count.
    n = ShownCount(sf)
    If n < LOG_MAX And Len(GridFilterWhere(sf)) = 0 Then
        m_total = n
    ElseIf m_tab = "e" Then
        m_total = ErrorListCount(m_search, dFrom, dTo)
    Else
        m_total = AuditLogCount(m_search, dFrom, dTo)
    End If
    AfterFilter
End Sub

' The rows the list shows now.
Private Function ShownCount(sf As Control) As Long
    Dim rs As DAO.Recordset
    Set rs = sf.Form.RecordsetClone
    If Not rs.EOF Then rs.MoveLast
    ShownCount = rs.RecordCount
End Function

' After the rows shown changed: the count in the subtitle, the empty state,
' and the entry shown kept when still listed (else the first).
Private Sub AfterFilter()
    Dim sf As Control, n As Long, keep As Long, emp As String, funnels As Boolean
    Set sf = Me.Controls(ListName())
    n = ShownCount(sf)
    funnels = Len(GridFilterWhere(sf)) > 0
    If m_total < 0 Then
        Me.lblSubtitle.Caption = "The first date is after the second: swap them to see entries."
    ElseIf m_total > LOG_MAX Then
        Me.lblSubtitle.Caption = "Showing the newest " & Format$(LOG_MAX, "#,##0") & " of " & Format$(m_total, "#,##0") & _
            " entries that match the search and dates. Narrow them to see older ones."
    Else
        Me.lblSubtitle.Caption = Format$(n, "#,##0") & IIf(n = 1, " entry", " entries") & IIf(funnels, " (column filters on)", "")
    End If
    emp = IIf(m_tab = "e", "emptyEList", "emptyList")
    If n = 0 And m_total = 0 And Len(m_search) = 0 And IsNull(Me.dtFrom) And IsNull(Me.dtTo) And Not funnels Then
        ShowEmpty Me, emp, True, IIf(m_tab = "e", "The error log is empty", "The audit trail is empty"), "Nothing has been recorded yet."
        SetVisible Me.Controls(emp & "_b"), False
    ElseIf m_tab = "e" Then
        ShowEmpty Me, emp, n = 0, "No errors logged", "No error fits the search, the dates and the column filters."
    Else
        ShowEmpty Me, emp, n = 0, "Nothing in the audit trail", "No entry fits the search, the dates and the column filters."
    End If
    SetVisible sf, n > 0
    SetEnabled Me.btnExport, n > 0
    SetEnabled Me.btnClear, m_ok
    keep = IIf(m_tab = "e", m_eid, m_aid)
    If GridSelect(sf, keep) Then
        ShowEntry keep
    ElseIf IsNull(SelectFirstRow(sf)) Then  ' SelectFirstRow shows the first row's entry itself (GridClick)
        ShowEntry 0
    End If
End Sub

' Show entry id of the open tab in full (0: the hint instead).
Private Sub ShowEntry(ByVal id As Long)
    Dim s As String
    If id <> 0 Then
        If m_tab = "e" Then s = ErrorLogDetail(id) Else s = AuditDetail(id)
    End If
    If m_tab = "e" Then m_eid = id Else m_aid = id
    Me.txtDetail = IIf(Len(s) > 0, s, Null)
    SetVisible Me.lblDetailEmpty, Len(s) = 0
    SetEnabled Me.btnCopy, Len(s) > 0
End Sub
