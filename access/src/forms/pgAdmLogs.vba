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
' when there are more. The picked entry shows in full on the right (Copy
' puts it on the clipboard); Export saves the open list.
'
' Depends on: modErrors, modLogs, modSignIn, modUI; the kit's modGridFilter,
' modKitFiles, modKitUI, modKitUtil.

Private Const MSG_TITLE As String = "Logs"

Private m_tab As String         ' "a" Audit trail, "e" Errors
Private m_search As String
Private m_aid As Long           ' the audit entry shown (0: none)
Private m_eid As Long           ' the error entry shown (0: none)
Private m_ok As Boolean         ' a data file is linked and the signed-in person is an admin

Public Property Get HelpTopic() As String
    HelpTopic = "logs"
End Property

Private Sub Form_Load()
    On Error GoTo Err_Handler
    m_tab = "a"
    SelectTab Me, "a"
    Me.dtFrom = DateAdd("d", -30, Date): Me.dtTo = Date
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
    RefreshList
    Me.txtSearch.SetFocus
    Me.txtSearch.SelStart = Len(text)
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.SearchFor", text
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
' AfterUpdate, so the list is loaded again here.
Private Sub cal_dtFrom_Click()
    On Error GoTo Err_Handler
    PickDate Me, "dtFrom"
    RefreshList
    Exit Sub
Err_Handler:
    HandleError "pgAdmLogs.cal_dtFrom_Click"
End Sub

Private Sub cal_dtTo_Click()
    On Error GoTo Err_Handler
    PickDate Me, "dtTo"
    RefreshList
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

Private Sub btnExport_Click()
    On Error GoTo Err_Handler
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
    SetSearch Me, "txtSearch", ""           ' calls SearchFor, which loads the list
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
    Dim sf As Control, total As Long, dFrom As Variant, dTo As Variant
    Set sf = Me.Controls(ListName())
    If Not m_ok Then
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
            Me.lblSubtitle.Caption = "The first date is after the second: swap them to see entries."
            sf.Form.RecordSource = IIf(m_tab = "e", ErrorListSql("", Date + 1, Date), AuditLogSql("", Date + 1, Date))
            AfterFilter
            Exit Sub
        End If
    End If
    If m_tab = "e" Then
        sf.Form.RecordSource = ErrorListSql(m_search, dFrom, dTo)
        total = ErrorListCount(m_search, dFrom, dTo)
    Else
        sf.Form.RecordSource = AuditLogSql(m_search, dFrom, dTo)
        total = AuditLogCount(m_search, dFrom, dTo)
    End If
    SetGridFilter sf, ""                    ' the column funnels, on the new rows
    GridHighlight sf, m_search
    If total > LOG_MAX Then
        Me.lblSubtitle.Caption = "Showing the newest " & Format$(LOG_MAX, "#,##0") & " of " & Format$(total, "#,##0") & _
            " entries. Narrow the dates or search to see older ones."
    Else
        Me.lblSubtitle.Caption = Format$(total, "#,##0") & IIf(total = 1, " entry", " entries")
    End If
    AfterFilter
End Sub

' After the rows shown changed: the empty state, and the entry shown kept
' when still listed (else the first).
Private Sub AfterFilter()
    Dim sf As Control, n As Long, keep As Long
    Set sf = Me.Controls(ListName())
    n = sf.Form.RecordsetClone.RecordCount
    ShowEmpty Me, IIf(m_tab = "e", "emptyEList", "emptyList"), n = 0
    SetVisible sf, n > 0
    SetEnabled Me.btnExport, n > 0
    SetEnabled Me.btnClear, m_ok
    keep = IIf(m_tab = "e", m_eid, m_aid)
    If GridSelect(sf, keep) Then
        ShowEntry keep
    Else
        ShowEntry Nz(SelectFirstRow(sf), 0)
    End If
End Sub

' Show entry id of the open tab in full (0: the hint instead).
Private Sub ShowEntry(ByVal id As Long)
    Dim s As String
    If id <> 0 Then s = IIf(m_tab = "e", ErrorLogDetail(id), AuditDetail(id))
    If m_tab = "e" Then m_eid = id Else m_aid = id
    Me.txtDetail = IIf(Len(s) > 0, s, Null)
    SetVisible Me.lblDetailEmpty, Len(s) = 0
    SetEnabled Me.btnCopy, Len(s) > 0
End Sub
