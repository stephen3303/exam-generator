Option Compare Database
Option Explicit

' Error handling and the error log (example: copy it into your app's src\modules).
'
' Needs from the app: APP_NAME, g_Silent (True during the automated build), a
' TempVars("Username") for the signed-in user (optional), frmMain.PageKey
' (optional), and an error_log table, ideally in the back end:
'
'     CREATE TABLE error_log (id COUNTER PRIMARY KEY, logged_at DATETIME,
'         error_number LONG, description LONGTEXT, err_source TEXT(255),
'         source TEXT(255), call_trail LONGTEXT, context LONGTEXT, page TEXT(32),
'         active_form TEXT(64), active_control TEXT(64), username TEXT(150),
'         windows_user TEXT(64), computer TEXT(64), fe_build TEXT(32),
'         access_version TEXT(32))
'
' Add app-specific context columns (current user id, current record) to
' WriteLogRow, WriteLogFile and ErrorLogDetail as needed.
'
' Every entry point (an event procedure, or a function a control or Access
' calls directly) ends with the standard handler:
'
'         On Error GoTo Err_Handler
'         ...
'         Exit Sub
'     Err_Handler:
'         HandleError "pgAdmin.btnSave_Click"
'
' HandleError writes the error to the error_log table, which is in the back end
' so admins see errors from every user (an admin error-log page), and shows the user a
' short message with the entry number. Helpers without a handler pass their
' errors up to the entry point. Library procedures that know useful context
' (which record, which file) add it with RaiseAgain on the way up, so the entry
' shows where the error happened and what it was working on.
'
' If the back end can't be written (often the cause of the error), the entry
' goes to %LOCALAPPDATA%\<APP_NAME>\logs\errors.log instead.
'
' Each entry is also written to the PC's Windows Event Log (Application log,
' source APP_NAME, event 1001; modKitEventLog), so the site's IT can collect
' every app's errors on its log server. The event leaves out the context and
' call trail.
'
' The build numbers the lines of its copy of every module (tools\VbaLineNumbers.ps1;
' the source in git stays un-numbered), so Erl says which line failed. LogError
' and RaiseAgain add it to the procedure name, e.g. "pgAdmin.btnSave_Click line 42",
' which is line 42 of src\forms\pgAdmin.vba. A copy built without numbers logs
' the name alone. Erl is per procedure: in a caller's handler it is the line of
' that caller's call into the procedure that failed, so each trail entry and the
' entry point name their own line.

' Context added by RaiseAgain, innermost first. Kept with the error it belongs
' to, so a trail left by an error that was later ignored isn't attached to an
' unrelated one.
Private m_trail As String
Private m_trailNum As Long
Private m_trailDesc As String
Private m_busy As Boolean           ' an error while logging must not log again

' Record the current error and return its error_log id (0 when it went to the
' fallback file instead). Call it first thing in a handler: it reads the Err
' object, and it never raises.
'   source:  "<module or form>.<procedure>" that caught the error
'   context: what it was working on, e.g. "order 1234"
Public Function LogError(ByVal source As String, Optional ByVal context As String = "") As Long
    Dim num As Long, desc As String, errSrc As String, trail As String
    num = Err.Number: desc = Err.Description: errSrc = Err.Source
    source = SourceWithLine(source)
    If m_trailNum = num And m_trailDesc = desc Then trail = m_trail
    m_trail = "": m_trailNum = 0: m_trailDesc = ""
    If m_busy Then Exit Function
    m_busy = True
    LogError = WriteLogRow(num, desc, errSrc, source, trail, context)
    If LogError = 0 Then WriteLogFile num, desc, errSrc, source, trail, context
    KitEventLogError LogError, num, desc, source, CurrentPageKey(), ActiveFormName(), ActiveControlName(), _
        Nz(TempVars("Username"), ""), FrontEndBuild()        ' the Windows Event Log, for the site's log server (modKitEventLog)
    m_busy = False
End Function

' Standard entry-point handler: log the error, then tell the user. intro
' replaces the opening sentence of the message. During the automated build and
' tests (g_Silent) the error is raised again instead of shown.
Public Sub HandleError(ByVal source As String, Optional ByVal context As String = "", _
                       Optional ByVal intro As String = "")
    Dim num As Long, desc As String, errSrc As String, id As Long, msg As String
    num = Err.Number: desc = Err.Description: errSrc = Err.Source
    id = LogError(source, context)
    DoCmd.Hourglass False
    If g_Silent Then Err.Raise num, errSrc, desc
    If Len(intro) = 0 Then intro = "Something went wrong, so that action didn't finish."
    msg = intro & vbCrLf & vbCrLf & desc & " (error " & num & ")" & vbCrLf & vbCrLf
    If id > 0 Then
        msg = msg & "The details are saved in the error log as entry " & id & ". " & _
            "If this keeps happening, give that number to your administrator."
    Else
        msg = msg & "The details are saved in " & LogFilePath() & "."
    End If
    KitMsgBox msg, vbExclamation, APP_NAME          ' centred over Access (modKitUI)
End Sub

' Add context to the error being handled and pass it up to the caller. Use in
' a library procedure's handler:
'     Fail:
'         RaiseAgain "modOrders.CreateOrder", "customer " & customerId
Public Sub RaiseAgain(ByVal source As String, Optional ByVal context As String = "")
    Dim num As Long, desc As String, errSrc As String
    num = Err.Number: desc = Err.Description: errSrc = Err.Source
    source = SourceWithLine(source)
    If m_trailNum <> num Or m_trailDesc <> desc Then m_trail = ""
    m_trail = m_trail & IIf(Len(m_trail) > 0, vbCrLf, "") & source & IIf(Len(context) > 0, " (" & context & ")", "")
    m_trailNum = num: m_trailDesc = desc
    Err.Raise num, errSrc, desc
End Sub

' source with the line of the error being handled added (" line 42"), from Erl.
' Erl is 0 when the failing procedure has no line numbers, and then source is
' returned as is. Call it while the error is still current (before any On Error
' or Resume, which clear it).
Public Function SourceWithLine(ByVal source As String) As String
    Dim ln As Long
    ln = Erl
    SourceWithLine = source & IIf(ln > 0, " line " & ln, "")
End Function

Private Function WriteLogRow(ByVal num As Long, ByVal desc As String, ByVal errSrc As String, ByVal source As String, _
                             ByVal trail As String, ByVal context As String) As Long
    Dim rs As DAO.Recordset, id As Long
    On Error GoTo Fail
    Set rs = CurrentDb.OpenRecordset("error_log", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!logged_at = Now()
    rs!error_number = num
    rs!description = Left$(desc, 4000)
    rs!err_source = Left$(errSrc, 255)
    rs!source = Left$(source, 255)
    rs!call_trail = trail
    rs!context = context
    rs!page = Left$(CurrentPageKey(), 32)
    rs!active_form = Left$(ActiveFormName(), 64)
    rs!active_control = Left$(ActiveControlName(), 64)
    rs!username = Left$(Nz(TempVars("Username"), ""), 150)
    rs!windows_user = Left$(Environ$("USERNAME"), 64)
    rs!computer = Left$(Environ$("COMPUTERNAME"), 64)
    rs!fe_build = Left$(FrontEndBuild(), 32)
    rs!access_version = Left$(Application.Version & "." & Application.Build, 32)
    id = rs!id
    rs.Update
    rs.Close
    WriteLogRow = id
    Exit Function
Fail:
    WriteLogRow = 0
End Function

Private Sub WriteLogFile(ByVal num As Long, ByVal desc As String, ByVal errSrc As String, ByVal source As String, _
                         ByVal trail As String, ByVal context As String)
    Dim f As Integer
    On Error GoTo Fail
    f = FreeFile
    Open LogFilePath() For Append As #f
    Print #f, String$(72, "-")
    Print #f, Format$(Now(), "yyyy-mm-dd hh:nn:ss") & "  error " & num & ": " & desc
    Print #f, "Procedure:   " & source & IIf(Len(context) > 0, "  (" & context & ")", "")
    If Len(trail) > 0 Then Print #f, "Called from: " & Replace(trail, vbCrLf, vbCrLf & "             ")
    Print #f, "Page:        " & CurrentPageKey() & "  form " & ActiveFormName() & "  control " & ActiveControlName()
    Print #f, "User:        " & Nz(TempVars("Username"), "(not signed in)") & _
        "  Windows " & Environ$("USERNAME") & " on " & Environ$("COMPUTERNAME")
    Print #f, "Build:       " & FrontEndBuild() & "  Access " & Application.Version & "." & Application.Build & _
        "  source " & errSrc
    Close #f
    Exit Sub
Fail:
    On Error Resume Next
    Close #f
End Sub

' %LOCALAPPDATA%\<APP_NAME>\logs\errors.log (the folder is created on first use).
Public Function LogFilePath() As String
    Dim base As String
    On Error Resume Next
    base = Environ$("LOCALAPPDATA")
    If Len(base) = 0 Then base = Environ$("TEMP")
    EnsureFolder base & "\" & APP_NAME & "\logs"
    LogFilePath = base & "\" & APP_NAME & "\logs\errors.log"
End Function

Private Function CurrentPageKey() As String
    On Error Resume Next
    If IsFormOpen("frmMain") Then CurrentPageKey = Forms("frmMain").PageKey
End Function

Private Function ActiveFormName() As String
    On Error Resume Next
    ActiveFormName = Screen.ActiveForm.Name
End Function

Private Function ActiveControlName() As String
    On Error Resume Next
    ActiveControlName = Screen.ActiveControl.Name
End Function

' Date and time the front end was built (BuildAll stores it in the AppBuild
' database property).
Public Function FrontEndBuild() As String
    On Error Resume Next
    FrontEndBuild = CurrentDb.Properties("AppBuild")
End Function

' ---------------------------------------------------------------------------
' Error log page
' ---------------------------------------------------------------------------

' Row source for the error list, newest first.
Public Function ErrorLogSql() As String
    ' Columns are qualified (l.) because an alias may not repeat a column name.
    ErrorLogSql = "SELECT TOP 500 l.id, Format(l.logged_at,'dd mmm yy hh:nn') AS [When], " & _
        "IIf(Len(l.username & '')>0, l.username, '(' & l.windows_user & ')') AS [User], l.page AS [Page], l.source AS [Procedure], " & _
        "l.error_number & ': ' & Left(Replace(l.description & '', Chr(13) & Chr(10), ' '), 120) AS [Error] " & _
        "FROM error_log AS l ORDER BY l.id DESC"
End Function

' Everything recorded for one entry, as text for the detail box.
Public Function ErrorLogDetail(ByVal id As Long) As String
    Dim rs As DAO.Recordset, s As String
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM error_log WHERE id=" & id, dbOpenSnapshot)
    If rs.EOF Then Exit Function
    s = "Entry " & rs!id & "  -  " & Format$(rs!logged_at, "dd mmm yyyy hh:nn:ss") & vbCrLf & _
        "Error " & rs!error_number & ": " & NzStr(rs!description) & vbCrLf & _
        "Procedure: " & NzStr(rs!source) & IIf(Len(NzStr(rs!context)) > 0, "  (" & NzStr(rs!context) & ")", "") & vbCrLf
    If Len(NzStr(rs!call_trail)) > 0 Then s = s & "Called from: " & Replace(NzStr(rs!call_trail), vbCrLf, ";  ") & vbCrLf
    s = s & "Page: " & NzStr(rs!page) & "   Form: " & NzStr(rs!active_form) & "   Control: " & NzStr(rs!active_control) & vbCrLf & _
        "User: " & NzStr(rs!username) & vbCrLf & _
        "Windows: " & NzStr(rs!windows_user) & " on " & NzStr(rs!computer) & vbCrLf & _
        "Front end built: " & NzStr(rs!fe_build) & "   Access: " & NzStr(rs!access_version) & "   Source: " & NzStr(rs!err_source)
    rs.Close
    ErrorLogDetail = s
End Function

' Write the whole log to CSV. Returns the number of entries.
Public Function ExportErrorLogCsv(ByVal path As String) As Long
    ExportErrorLogCsv = ExportSqlToCsv("SELECT * FROM error_log ORDER BY id", path)
End Function

' Delete every entry. Returns how many were deleted.
Public Function ClearErrorLog() As Long
    ClearErrorLog = DCount("*", "error_log")
    CurrentDb.Execute "DELETE FROM error_log", dbFailOnError
End Function

Private Function NzStr(ByVal v As Variant) As String
    If Not IsNull(v) Then NzStr = CStr(v)
End Function
