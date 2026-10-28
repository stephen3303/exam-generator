Option Compare Database
Option Explicit

' Exam Generator (Access edition) - the run-time UI hooks the kit calls on
' the app: ShowTip for the "?" marks, BadgeColors and StatusVariant for the
' status pills in BuildGrid tables, and BandLabels/BandBack, which the kit's
' BandCF needs to compile (the app shows no bands).
'
' Status words are mapped to pill variants here, never to colors, so every
' palette keeps them readable. The data file's tblStatuses names them
' (questions: Draft, In Review, Approved, Rejected, Retired; exams: SME Review,
' In Review, Approved, Rejected); the plan's words (Returned, Invalid) are
' mapped too. Review tasks: Open, Done, Closed. Admin lists: Active, Off,
' Short (a topic short of approved questions). The approvals queue's kinds:
' New, New version. A new exam's topics: Ready, Short, No SME, Repeats; and
' the questions drawn: Replaced, Swapped, Swapped, given lately, Given lately. Matching ignores case
' (Option Compare Database).

' Click on a "?" help mark (TitleTip): show its full explanation.
Public Function ShowTip(frm As Form, ByVal nm As String)
    On Error GoTo Err_Handler
    Notify frm(nm).Tag, vbInformation
    Exit Function
Err_Handler:
    HandleError "modUI.ShowTip", nm
End Function

' Badge variant -> Array(back, fore).
Public Function BadgeColors(ByVal vrnt As String) As Variant
    Select Case vrnt
        Case "success": BadgeColors = Array(Th(CLR_SUCCESS_BG), Th(CLR_SUCCESS_FG))
        Case "warning": BadgeColors = Array(Th(CLR_WARNING_BG), Th(CLR_WARNING_FG))
        Case "info": BadgeColors = Array(Th(CLR_INFO_BG), Th(CLR_INFO_FG))
        Case "destructive": BadgeColors = Array(Th(CLR_DESTRUCTIVE), Th(CLR_WHITE))
        Case "default": BadgeColors = Array(Th(CLR_PRIMARY), Th(CLR_WHITE))
        Case "outline": BadgeColors = Array(Th(CLR_SURFACE), Th(CLR_FG))
        Case Else: BadgeColors = Array(Th(CLR_SECONDARY_BG), Th(CLR_SECONDARY_FG))
    End Select
End Function

' Badge variant for a status (pill columns in BuildGrid), or "" for no pill.
' It runs on every repaint, also for a row deleted since the list was last
' requeried, whose status can't be read: that row gets no pill.
Public Function StatusVariant(ByVal status As Variant) As String
    On Error GoTo Gone
    Select Case Nz(status, "")
        Case "Approved", "Done", "Active", "Correct", "Ready": StatusVariant = "success"
        Case "In review", "SME review", "Open", "New", "Swapped": StatusVariant = "info"
        Case "Returned", "Rejected", "Short", "Repeats", "Given lately", "Swapped, given lately", "Replaced": StatusVariant = "warning"
        Case "Invalid", "No SME": StatusVariant = "destructive"
        Case "Draft", "Closed", "Retired", "Off", "New version": StatusVariant = "secondary"
    End Select
    Exit Function
Gone:
    StatusVariant = ""
End Function

' ---------------------------------------------------------------------------
' Helpers the pages share
' ---------------------------------------------------------------------------

' Select row id of the table in subform control sf (its RecordsetClone's id
' field), set its txtSel and repaint. False, with nothing selected, when it
' isn't listed (or id is 0).
Public Function GridSelect(sf As Control, ByVal id As Long) As Boolean
    Dim g As Form
    Set g = sf.Form
    g.txtSel = Null
    If id <> 0 Then
        With g.RecordsetClone
            .FindFirst "id=" & id
            If Not .NoMatch Then
                g.Bookmark = .Bookmark
                g.txtSel = id
                GridSelect = True
            End If
        End With
    End If
    g.Recalc
End Function

' The number after name= in frm's OpenArgs ("eq=3;stamp=0"), or 0.
Public Function OpenArgLong(frm As Form, ByVal name As String) As Long
    Dim part As Variant, v As String
    For Each part In Split(Nz(frm.OpenArgs, ""), ";")
        If LCase$(Left$(part, Len(name) + 1)) = LCase$(name) & "=" Then v = Mid$(part, Len(name) + 2)
    Next
    If IsNumeric(v) Then OpenArgLong = CLng(v)
End Function

' Band pill colors (BandCF). The app shows no bands; these keep the kit's
' BandCF compiling.
Public Function BandLabels() As Variant
    BandLabels = Array("Above", "At", "Below")
End Function

Public Function BandBack(ByVal label As String) As Long
    Select Case label
        Case "Above": BandBack = CLR_BAND_ABOVE
        Case "At": BandBack = CLR_BAND_AT
        Case "Below": BandBack = CLR_BAND_BELOW
        Case Else: BandBack = CLR_BAND_NE
    End Select
End Function
