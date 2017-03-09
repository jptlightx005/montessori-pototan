Attribute VB_Name = "AccountantModule"
'EXEL MONTESSORI ENROLLMENT SYSTEM
Option Explicit
'ADODB variables
Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Public cmd As ADODB.Command
Public subcmd As ADODB.Command

'Program Settings
Public ipaddress As String
Public localip As String
Public acctadmin As administrator
Public password As String

'Main method of the System
'This contains the methods called to be executed by the program before opening
Sub Main()
    Set acctadmin = New administrator
    Call LoadSettings
    frmSplash.Show
End Sub

'Loads the settings saved in the ini file
Sub LoadSettings()
'On Error GoTo ProcError
    Dim settingsFile As String
    settingsFile = App.Path & "\accountant.ini"
    If Dir(settingsFile) = "" Then
        WriteIniValue settingsFile, "Default", "username", ""
        WriteIniValue settingsFile, "Default", "ipaddress", ""
    End If
    acctadmin.usrn = ReadIniValue(App.Path & "\accountant.ini", "Default", "username")
    ipaddress = ReadIniValue(App.Path & "\accountant.ini", "Default", "ipaddress")
    
ProcExit:
    Exit Sub
    
ProcError:
   MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Sub

'Saves the settings in the ini file
'Saves the user's settings in accountant.ini
Public Sub SaveSettings(usrn As String, ip As String)
On Error GoTo ProcError
    WriteIniValue App.Path & "\accountant.ini", "Default", "username", usrn
    WriteIniValue App.Path & "\accountant.ini", "Default", "ipaddress", ip
ProcExit:
    Exit Sub
    
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Sub

Public Function grade(grd As String) As String
    Select Case grd
        Case "preschool"
            grade = "Nursery"
        Case "grade1"
            grade = "Grade I"
        Case "grade2"
            grade = "Grade II"
        Case "grade3"
            grade = "Grade III"
        Case "grade4"
            grade = "Grade IV"
        Case "grade5"
            grade = "Grade V"
        Case "grade6"
            grade = "Grade VI"
    End Select
End Function

Public Sub Logout()
    Set acctadmin = New administrator
    frmAccountant.resetBoxes
    Unload frmAccountant
    frmLogin.Show
End Sub
