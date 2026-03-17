Option Compare Database
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public conn As ADODB.Connection
Public wdb As Access.Application

Function onTime()
On Error Resume Next

'---THIS IS THE PRIMARY FUNCTION---
'---WHEN THIS DB OPENS, AN AUTOEXEC MACRO CALLS THIS FUNCTION---

If CurrentProject.Path <> "C:\workingdb" Then
    Exit Function
End If

startOver:

Sleep 30000 '30000 Milliseconds = 30 second delay
runGhost

GoTo startOver

End Function

Function runGhost()
On Error Resume Next

Dim closeIt As Boolean
closeIt = False

Set conn = CurrentProject.Connection

Call grabSessionID

If grabWDB Then
    'if WDB found, then track the forms
    Dim openForms As String
    openForms = ""
    
    'find all open forms
    Dim obj, sForm As Control
    For Each obj In wdb.Application.forms
        openForms = openForms & obj.name & "["
        For Each sForm In obj.Controls
            If sForm.ControlType = acSubform Then
                openForms = openForms & "" & sForm.Form.name & ","
            End If
        Next sForm
        If Right(openForms, 1) = "," Then openForms = Left(openForms, Len(openForms) - 1)
        openForms = openForms & "]" & vbNewLine
nextOne:
    Next obj
    
    conn.Execute "UPDATE tblWdbSessions SET wdbVersion = '" & Nz(wdb.TempVars!wdbVersion, "") & "' AND openForms = '" & openForms & "' AND lastCheck = GETDATE()"
    
    checkCommands
Else
    'if no WDB found, unregister all open sessions, close Ghost DB
    closeAllMySessions
    closeIt = True
End If

'cleanup
Set wdb = Nothing
Set conn = Nothing

If closeIt Then Application.Quit

End Function

Function checkCommands()
On Error Resume Next

Dim rsGhostCommands As New ADODB.Recordset
rsGhostCommands.Open "SELECT * FROM tblGhostCommands WHERE actionStart is not null", conn, adOpenForwardOnly, adLockReadOnly 'actionStart means this function is ON

Dim doAction As Boolean

Do While Not rsGhostCommands.EOF
    With rsGhostCommands
    
        If Nz(!specificUser, "") <> "" And !specificUser <> Environ("username") Then Exit Function 'this is meant for a specific user
        If !actionStart <= DateAdd("n", 5, Now) Then
            '5 minute warning
            If IsNull(TempVars!min5Warn) Then TempVars.Add "min5Warn", "True"
        End If
        If !actionStart <= DateAdd("n", 2, Now) Then
            '2 minute warning
            If IsNull(TempVars!min2Warn) Then TempVars.Add "min2Warn", "True"
        End If
        
        doAction = !actionStart < Now
        
        Select Case !Action
            Case "closeWorkingDB"
                If doAction Then
                    wdb.TempVars.Add "forceClose", "True"
                    wdb.Application.Quit
                    closeAllMySessions
                    Application.Quit
                End If
                
                If TempVars!min2Warn = True Then
                    Call wdb.Run("snackBox", "error", "Maintenance Required", "Wdb will auto-close in 2 minutes due to " & !actionDetails, "DASHBOARD", True, False)
                    TempVars.Add "min2Warn", "SENT"
                    GoTo nextCommand
                End If
                If TempVars!min5Warn = True Then
                    Call wdb.Run("snackBox", "error", "Maintenance Required", "Wdb will auto-close in 5 minutes due to " & !actionDetails, "DASHBOARD", True, False)
                    TempVars.Add "min5Warn", "SENT"
                    GoTo nextCommand
                End If
                
            Case "message"
                If Nz(TempVars!messageSent, "") <> !actionDetails Then
                    Call wdb.Run("snackBox", "info", "Notice", !actionDetails, "DASHBOARD", True, False)
                    TempVars.Add "messageSent", CStr(!actionDetails)
                End If
            
            Case "openWorkingDB"
        End Select
nextCommand:
        .MoveNext
    
    End With
Loop

If rsGhostCommands.State = adStateOpen Then rsGhostCommands.Close
Set rsGhostCommands = Nothing

End Function

Function closeAllMySessions()
On Error Resume Next

conn.Execute "UPDATE tblWdbSessions SET openForms = '', sessionEnd = '" & Now() & "' WHERE username = '" & Environ("username") & "' AND sessionEnd is null"

End Function

Function grabSessionID()
On Error Resume Next

If IsNull(TempVars!SessionID) Then
    'current session is not registered.
    'unregister all old sessions and start new
    closeAllMySessions
    conn.Execute "INSERT INTO tblWdbSessions(username,sessionStart,lastCheck,machine) VALUES('" & Environ("username") & "',GETDATE(),GETDATE(),'" & Environ("COMPUTERNAME") & "')"
    TempVars.Add "SessionID", conn.Execute("SELECT @@identity")(0)
End If

End Function

Function grabWDB() As Boolean
On Error GoTo exitThis

grabWDB = False

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists("C:\workingdb\WorkingDB.laccdb") Then 'if the file exists, try to delete it.
    On Error GoTo errCat
    fso.Deletefile "C:\workingdb\WorkingDB.laccdb" 'if it does not let you delete it, that means the database is active and in use.
    On Error GoTo exitThis
End If

exitThis:
Set fso = Nothing
Exit Function

errCat:
If Err.Number = 70 Then
    Set wdb = GetObject("C:\workingdb\WorkingDB.accdb")
    grabWDB = True
End If

End Function