Option Compare Database
Option Explicit

' Exam Generator (Access edition) - the run-time UI hooks the kit calls on
' the app: ShowTip for the "?" marks, BadgeColors and StatusVariant for the
' status pills in BuildGrid tables, and BandLabels/BandBack, which the kit's
' BandCF needs to compile (the app shows no bands).
'
' Status words are mapped to pill variants here, never to colors, so every
' palette keeps them readable. Questions: Draft, In review, Approved,
' Returned, Invalid. Exams: SME review, In review, Approved, Returned.
' Review tasks: Open, Done, Closed.

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
        Case "Approved", "Done": StatusVariant = "success"
        Case "In review", "SME review", "Open": StatusVariant = "info"
        Case "Returned": StatusVariant = "warning"
        Case "Invalid": StatusVariant = "destructive"
        Case "Draft", "Closed": StatusVariant = "secondary"
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
