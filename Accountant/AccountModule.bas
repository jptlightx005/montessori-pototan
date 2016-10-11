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
Public Sub SaveSettings(usrn As String)
On Error GoTo ProcError
    WriteIniValue App.Path & "\accountant.ini", "Default", "username", usrn
    WriteIniValue App.Path & "\accountant.ini", "Default", "ipaddress", ipaddress
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
        If rs("role") = "accountant" Then 'if the admin's role is an accountant
            If rs("pssw") = pssw Then 'if the password entered is correct
                'increments the times the user has logged in
                rs("login_count") = rs("login_count") + 1
                rs.Update
                
                acctadmin.usrn = usrn 'sets the current program's registrar admin to current user
                acctadmin.pssw = pssw
                acctadmin.role = rs("role")
                localip = frmLogin.sckMain.localip 'sets the program's local ip to the computer's network ip address
                
                'prompts the user has logged in successfully
                MsgBox "Logged in Successfully!", vbOKOnly + vbInformation 'prompts
                Unload frmLogin 'exits the current form
                'sets the registrar form's labels with the current entries
                frmAccountant.lbladmin = acctadmin.usrn
                frmAccountant.lblIP = localip
                'shows the registrar form
                frmAccountant.Show
                'closes the recordset
                rs.Close
                Exit Sub
            Else 'If the Password entered is wrong
                MsgBox "Wrong Password!", vbOKOnly + vbExclamation
                Exit Sub
            End If
        Else 'If the Admin role is not a registrar
            MsgBox "Must use accountant account!", vbOKOnly + vbExclamation
            Exit Sub
        End If
    Loop
    'If the scanning didn't match records
    MsgBox "Wrong username or username doesn't exist!", vbOKOnly + vbExclamation
    
ProcExit:
    Exit Sub
    
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Sub

Public Function SearchStudent(searchStr As String) As student
On Error GoTo ProcError
'sets the RecordSet for the search method
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    rs.Source = "SELECT * FROM montessori_records WHERE Student_ID=" & searchStr
    'opens the recordset and scans the table
    'Exit Subs in this loop is used to skip the rest of the codes when conditions are met
    rs.Open
    Do Until rs.EOF
        Dim studentFound As student
        Set studentFound = New student
        studentFound.studentID = rs("Student_ID")
        studentFound.queueID = rs("Queue_ID")
        studentFound.firstName = rs("first_name").Value
        studentFound.middleName = rs("middle_name").Value
        studentFound.lastName = rs("last_name").Value
        
        studentFound.homeAddress = rs("home_address")
        studentFound.grade = grade(rs("current_grade"))
        studentFound.balancePaid = rs("balance_paid")
        studentFound.datePaid = rs("date_of_payment").Value
        Set SearchStudent = studentFound
        rs.Close
        Exit Function
    Loop
    Set SearchStudent = Nothing
ProcExit:
    Exit Function
ProcError:
    MsgBox Err.Description, vbExclamation
End Function

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
