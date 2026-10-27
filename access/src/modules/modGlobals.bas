Option Compare Database
Option Explicit

' Exam Generator (Access edition) - app constants, build-time silence and the
' small helpers every module uses: Notify (messages), Q (SQL text), NavTo.
' The names the kit calls on the app (README, "What the app must provide")
' are here and in modUI; MeasureText comes from the kit's modTextMeasure
' adapter (kit\adapters), which build.ps1 loads with the kit.

Public Const APP_NAME As String = "Exam Generator"
Public Const APP_CODE As String = "EXAMGEN"      ' this app's code in the shared Personnel file (kit_person_app)
Public Const BE_FILE As String = "ExamGenerator_be.accdb"   ' the shared data file's usual name
Public g_Silent As Boolean                       ' True during the build: nothing may block

' Message to the user. During the build it is printed instead; warnings and
' errors raise, and questions are answered Yes.
Public Function Notify(ByVal msg As String, Optional ByVal style As VbMsgBoxStyle = vbInformation) As VbMsgBoxResult
    If g_Silent Then
        Debug.Print msg
        If (style And vbExclamation) = vbExclamation Or (style And vbCritical) = vbCritical Then
            Err.Raise vbObjectError + 999, APP_NAME, msg
        End If
        If (style And vbYesNo) = vbYesNo Then Notify = vbYes Else Notify = vbOK
    Else
        Notify = KitMsgBox(msg, style, APP_NAME)    ' centred over Access (modKitUI)
    End If
End Function

' Text as a quoted SQL string literal.
Public Function Q(ByVal s As Variant) As String
    Q = "'" & Replace(Nz(s, ""), "'", "''") & "'"
End Function

' Text for a LIKE pattern typed by the user: its wildcard characters are
' bracketed so they match themselves, as GridHighlight marks them.
Public Function LikeLiteral(ByVal s As String) As String
    s = Replace(s, "[", "[[]")
    LikeLiteral = Replace(Replace(Replace(s, "*", "[*]"), "?", "[?]"), "#", "[#]")
End Function

' Go to a page once the calling event has finished (the kit's ShellGo).
Public Function NavTo(ByVal key As String)
    NavTo = ShellGo(key)
End Function

' Open the palette picker, only while the current page can be left, because
' a palette switch reopens the main window (frmMain.CanLeavePage).
Public Function AppAppearance()
    On Error GoTo Err_Handler
    If IsFormOpen("frmMain") Then
        If Not Forms("frmMain").CanLeavePage() Then Exit Function
    End If
    ShowAppearance
    Exit Function
Err_Handler:
    HandleError "modGlobals.AppAppearance"
End Function
