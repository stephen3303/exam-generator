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
' New, New version. Matching ignores case
' (Option Compare Database).
'
' At the end, the helpers every page shares: GridSelect (select a table's
' row by id), OpenArgText and OpenArgLong (a dialog's OpenArgs).

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
        Case "Approved", "Done", "Active", "Correct": StatusVariant = "success"
        Case "In review", "SME review", "Open", "New": StatusVariant = "info"
        Case "Returned", "Rejected", "Short": StatusVariant = "warning"
        Case "Invalid": StatusVariant = "destructive"
        Case "Draft", "Closed", "Retired", "Off", "New version": StatusVariant = "secondary"
    End Select
    Exit Function
Gone:
    StatusVariant = ""
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

' ---------------------------------------------------------------------------
' Shared page helpers
' ---------------------------------------------------------------------------

' Select the row of table sf (a subGrid subform) whose id column is id, as a
' click on it would: the row's Bookmark and txtSel, then a Recalc so the
' selection mark follows. Every page's list and table selects rows this way.
'   sf       the subform control on the page
'   id       the row's id: a number, or text (help topic keys, through Q);
'            Null, 0 or "" selects no row
'   requery  True to reload the table's rows first
' Returns True when the row is listed and now selected. False leaves no row
' selected (txtSel Null).
' Errors: none of its own; a subform without txtSel raises Access's error.
Public Function GridSelect(sf As Control, ByVal id As Variant, Optional ByVal requery As Boolean = False) As Boolean
    Dim g As Form, crit As String
    Set g = sf.Form
    If requery Then g.Requery
    g.txtSel = Null
    If IsNull(id) Then
        crit = ""
    ElseIf VarType(id) = vbString Then
        If Len(id) > 0 Then crit = "id=" & Q(id)
    ElseIf id <> 0 Then
        crit = "id=" & CLng(id)
    End If
    If Len(crit) > 0 Then
        With g.RecordsetClone
            .FindFirst crit
            If Not .NoMatch Then
                g.Bookmark = .Bookmark
                g.txtSel = id
                GridSelect = True
            End If
        End With
    End If
    g.Recalc
End Function

' The text after name= in frm's OpenArgs ("id=3;total=4"), "" when it isn't
' there. The name matches whatever its case.
Public Function OpenArgText(frm As Form, ByVal name As String) As String
    Dim part As Variant
    For Each part In Split(Nz(frm.OpenArgs, ""), ";")
        If LCase$(Left$(part, Len(name) + 1)) = LCase$(name) & "=" Then OpenArgText = Mid$(part, Len(name) + 2)
    Next
End Function

' The whole number after name= in frm's OpenArgs, 0 when it isn't there or
' isn't digits only (so "1e3", "&H10" or "1.5" aren't read as numbers).
Public Function OpenArgLong(frm As Form, ByVal name As String) As Long
    Dim v As String
    v = OpenArgText(frm, name)
    If Len(v) > 0 And Len(v) <= 9 And Not v Like "*[!0-9]*" Then OpenArgLong = CLng(v)
End Function
