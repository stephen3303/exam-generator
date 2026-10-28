Option Compare Database
Option Explicit

' Exam Generator (Access edition) - app constants, build-time silence and the
' small helpers every module uses: Notify (messages), Q (SQL text), NavTo,
' and at the end FieldText, Plural and TableExists.
' The names the kit calls on the app (README, "What the app must provide")
' are here and in modUI; MeasureText comes from the kit's modTextMeasure
' adapter (kit\adapters), which build.ps1 loads with the kit.

Public Const APP_NAME As String = "Exam Generator"
Public Const APP_CODE As String = "EXAMGEN"      ' this app's code in the shared Personnel file (kit_person_app)
Public Const BE_FILE As String = "ExamGen_BE.accdb"   ' the shared data file's usual name
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

' A text box's text: what is typed so far while it has the focus (Ctrl+S
' saves without leaving it), else its value. frm is the form it is on.
Public Function FieldText(frm As Form, c As Control) As String
    Dim act As String
    On Error Resume Next
    act = frm.ActiveControl.Name            ' error 2474 when none has the focus
    On Error GoTo 0
    If act = c.Name Then FieldText = Nz(c.Text, "") Else FieldText = Nz(c.Value, "")
End Function

' word, with an s added unless n is 1: "1 question", "3 questions".
Public Function Plural(ByVal n As Long, ByVal word As String) As String
    Plural = word & IIf(n = 1, "", "s")
End Function

' True when this front end has a table, or a link to one, called nm (any
' case). localOnly: only a table of its own, not a link.
Public Function TableExists(ByVal nm As String, Optional ByVal localOnly As Boolean = False) As Boolean
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        If StrComp(td.Name, nm, vbTextCompare) = 0 Then
            TableExists = (Not localOnly) Or Len(td.Connect) = 0
            Exit Function
        End If
    Next
End Function
