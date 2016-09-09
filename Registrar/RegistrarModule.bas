Attribute VB_Name = "RegistrarModule"
'EXEL MONTESSORI ENROLLMENT SYSTEM

'ADODB variables
Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Public cmd As ADODB.Command
Public subcmd As ADODB.Command

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
Public Sub SaveSettings()
On Error GoTo ProcError
    WriteIniValue App.Path & "\registrar.ini", "Default", "username", regadmin.usrn
    WriteIniValue App.Path & "\registrar.ini", "Default", "ipaddress", ipaddress
ProcExit:
    Exit Sub
    
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Sub

'Login Method
Public Sub LogIn(usrn As String, pssw As String, ip As String)
On Error GoTo ProcError 'If something goes wrong, skip to the Error message
    ipaddress = ip 'inserts the ip entered to the global variable
    
    'sets the Database Connection
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=" & ipaddress & ";Database=montessori-db; User=" & usrn & ";Password=" & pssw & ";"

    'sets the RecordSet for the log-in method
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    rs.Source = "SELECT * FROM montessori_admin WHERE usrn = '" + usrn + "'"
    'opens the recordset and scans the table
    'Exit Subs in this loop is used to skip the rest of the codes when conditions are met
    rs.Open
    Do Until rs.EOF
        If rs("role") = "registrar" Then 'if the admin's role is a registrar
            If rs("pssw") = pssw Then 'if the password entered is correct
                'increments the times the user has logged in
                rs("login_count") = rs("login_count") + 1
                rs.Update
                
                regadmin.usrn = usrn 'sets the current program's registrar admin to current user
                regadmin.pssw = pssw
                regadmin.role = rs("role").Value
                localip = frmLogin.sckMain.localip 'sets the program's local ip to the computer's network ip address
                
                'prompts the user has logged in successfully
                MsgBox "Logged in Successfully!", vbOKOnly + vbInformation 'prompts
                Unload frmLogin 'exits the current form
                'sets the registrar form's labels with the current entries
                frmRegistrar.lbladmin = regadmin.usrn
                frmRegistrar.lblIP = localip
                'shows the registrar form
                frmRegistrar.Show
                'closes the recordset
                rs.Close
                Exit Sub
            Else 'If the Password entered is wrong
                MsgBox "Wrong Password!", vbOKOnly + vbExclamation
                Exit Sub
            End If
        Else 'If the Admin role is not a registrar
            MsgBox "Must use registrar account!", vbOKOnly + vbExclamation
            Exit Sub
        End If
    Loop
    'If the scanning didn't match records
    MsgBox "Wrong username or username doesn't exist!", vbOKOnly + vbExclamation
    
ProcExit:
    Exit Sub
    
ProcError:
    MsgBox "Invalid credentials!", vbExclamation
    LogEvent (Err.Description)
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
