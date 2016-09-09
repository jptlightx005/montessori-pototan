Attribute VB_Name = "RegistrationFormModule"
'EXEL MONTESSORI ENROLLMENT SYSTEM
'Registration Form
Option Explicit

'ADODB Variables
Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Public cmd As ADODB.Command
Public subcmd As ADODB.Command

'Program Settings
Public ipaddress As String
Public localip As String
Public admin As administrator
Public password As String

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
Public Sub SaveSettings()
On Error GoTo ProcError
    WriteIniValue App.Path & "\registration_form.ini", "Default", "username", admin.usrn
    WriteIniValue App.Path & "\registration_form.ini", "Default", "ipaddress", ipaddress
ProcExit:
    Exit Sub
    
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Sub

'Login Method
Public Sub LogIn(usrn As String, pssw As String, ip As String)
On Error GoTo ProcError
    ipaddress = ip
    
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
        If rs("role") = "admin" Then 'if the admin's role is an Administrator
            If rs("pssw") = pssw Then 'if the password entered is correct
                'increments the times the user has logged in
                rs("login_count") = rs("login_count") + 1
                rs.Update
                
                'saves the admin's information on the class
                admin.usrn = usrn
                admin.pssw = pssw
                admin.role = rs("role")
                'prompts the user has logged in successfully
                MsgBox "Logged in Successfully!", vbOKOnly + vbInformation
                Unload frmLogin
                
                localip = frmLogin.sckMain.localip
                frmMain.Show
                rs.Close
                
                Exit Sub
            Else
                MsgBox "Wrong Password!", vbOKOnly + vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "Must use admin account!", vbOKOnly + vbExclamation
            Exit Sub
        End If
    Loop
    MsgBox "Wrong username or username doesn't exist!", vbOKOnly + vbExclamation
    
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

Public Sub SubmitStudentInfo(inf As String)
On Error GoTo ProcError
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    rs.Source = "SELECT * FROM montessori_queue"
    rs.Open
    Do Until rs.EOF
        If rs("student_info") = inf Then
            MsgBox "It seems our system has detected an existing queue with the exact same information you just entered. The queue ID is " & rs("Queue_ID"), vbExclamation
            Exit Sub
        Else
            rs.MoveNext
        End If
    Loop
    
    rs.AddNew
    rs("rf_admin") = admin.usrn
    rs("rf_ip") = localip
    rs("student_info") = inf
    rs("status") = "onqueue"
    
    rs.Update
    rs.Close
    
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    rs.Source = "SELECT Queue_ID FROM montessori_queue WHERE student_info = '" & inf & "'"
    rs.Open
    
    Dim QueueID As Integer
    
    Do Until rs.EOF
        QueueID = rs("Queue_ID").Value
        MsgBox "Successfully registered! Your Queue ID is " & QueueID, vbInformation
        rs.Close
        Exit Sub
    Loop
ProcExit:
    Exit Sub
    
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Sub
