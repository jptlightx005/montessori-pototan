Attribute VB_Name = "RegistrationFormModule"
'EXEL MONTESSORI ENROLLMENT SYSTEM
'Registration Form
Option Explicit

'Program Settings
Public ipaddress As String
Public localip As String
Public admin As administrator


'Main method of the System
'This contains the methods called to be executed by the program before opening
Sub Main()
    Set admin = New administrator
    Call LoadSettings
    frmSplash.Show
End Sub

'Loads the settings saved in registration_form.ini
Sub LoadSettings()
On Error GoTo ProcError
    If Dir(App.Path & "\registration_form.ini") = "" Then
        WriteIniValue App.Path & "\registration_form.ini", "Default", "username", ""
        WriteIniValue App.Path & "\registration_form.ini", "Default", "ipaddress", ""
    End If
    admin.usrn = ReadIniValue(App.Path & "\registration_form.ini", "Default", "username")
    ipaddress = ReadIniValue(App.Path & "\registration_form.ini", "Default", "ipaddress")
ProcExit:
    Exit Sub
    
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Sub

'Saves the user's settings in registration_form.ini
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

Public Sub Logout()
    Set admin = New administrator
    Call frmMain.ClearBoxes
    Unload frmMain
    frmLogin.Show
End Sub
