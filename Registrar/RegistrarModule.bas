Attribute VB_Name = "RegistrarModule"
'EXEL MONTESSORI ENROLLMENT SYSTEM

'Program Settings
Public ipaddress As String
Public localip As String
Public regadmin As administrator

'Current Student Index selected
Public currentStudentID As Integer

'Main method of the System
'This contains the methods called to be executed by the program before opening
Sub Main()
    Set regadmin = New administrator
    Call LoadSettings
    frmSplash.Show
End Sub
'Logs events on file
Sub LogEvent(eventLog As String)
On Error GoTo ErrHandler
  Dim nUnit As Integer
  nUnit = FreeFile
  ' This assumes write access to the directory containing the program '
  ' You will need to choose another directory if this is not possible '
  Open App.Path & "\log.txt" For Append As nUnit
  Print #nUnit, Format$(Now) & ":" & eventLog
  Close nUnit
  Exit Sub

ErrHandler:
  'Failed to write log for some reason.'
  'Show MsgBox so error does not go unreported '
  MsgBox "Error in " & ProcName & vbNewLine & _
    ErrNum & ", " & ErrorMsg
End Sub
'Loads the settings saved in registrar.ini
Sub LoadSettings()
On Error GoTo ProcError
    If Dir(App.Path & "\registrar.ini") = "" Then
        WriteIniValue App.Path & "\registrar.ini", "Default", "username", ""
        WriteIniValue App.Path & "\registrar.ini", "Default", "ipaddress", ""
    End If
    regadmin.usrn = ReadIniValue(App.Path & "\registrar.ini", "Default", "username")
    ipaddress = ReadIniValue(App.Path & "\registrar.ini", "Default", "ipaddress")
ProcExit:
    Exit Sub
    
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Sub

'Saves the user's settings in registrar.ini
Public Sub SaveSettings(usrn As String)
On Error GoTo ProcError
    WriteIniValue App.Path & "\registrar.ini", "Default", "username", usrn
    WriteIniValue App.Path & "\registrar.ini", "Default", "ipaddress", ipaddress
ProcExit:
    Exit Sub
    
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Sub

Public Function grade(grd As String, ByVal sender As Object)
    Select Case grd
        Case "preschool"
            grade = IIf(sender Is frmRegistrar, "Nursery", 0)
        Case "grade1"
            grade = IIf(sender Is frmRegistrar, "Grade I", 1)
        Case "grade2"
            grade = IIf(sender Is frmRegistrar, "Grade II", 2)
        Case "grade3"
            grade = IIf(sender Is frmRegistrar, "Grade III", 3)
        Case "grade4"
            grade = IIf(sender Is frmRegistrar, "Grade IV", 4)
        Case "grade5"
            grade = IIf(sender Is frmRegistrar, "Grade V", 5)
        Case "grade6"
            grade = IIf(sender Is frmRegistrar, "Grade VI", 6)
    End Select
End Function

Public Sub Logout()
    Set regadmin = New administrator
    Call frmRegistrar.ClearBoxes
    Unload frmRegistrar
    frmLogin.Show
End Sub
