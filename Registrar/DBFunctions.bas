Attribute VB_Name = "DBFunctions"
Option Explicit

'ADODB variables
Public rs As ADODB.recordSet
Public cmd As ADODB.Command
Public subcmd As ADODB.Command

Public Function cn() As ADODB.Connection
    'sets the Database Connection
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=" & ipaddress & ";Database=montessori-db; User=" & regadmin.usrn & ";Password=" & regadmin.pssw & ";"
End Function

Public Function recordSet(tableName As String, conditionKey As String, conditionValue As String) As ADODB.recordSet
    Set recordSet = New ADODB.recordSet
    recordSet.ActiveConnection = cn
    recordSet.CursorLocation = adUseClient
    recordSet.CursorType = adOpenDynamic
    recordSet.LockType = adLockOptimistic
    recordSet.Source = "SELECT * FROM " + tableName
    If (Not (conditionKey = "" And conditionValue = "")) Then
        recordSet.Source = recordSet.Source + " WHERE " + conditionKey + " = '" + conditionValue + "'"
    End If
End Function

'Login Method
Public Function LogIn(usrn As String, pssw As String, ip As String, remember As Integer) As Boolean
On Error GoTo ProcError 'If something goes wrong, skip to the Error message
    ipaddress = ip 'inserts the ip entered to the global variable
                
    regadmin.usrn = usrn 'sets the current program's registrar admin to current user
    regadmin.pssw = pssw
    
    'sets the RecordSet for the log-in method
    Set rs = recordSet("montessori_admin", "usrn", usrn)
    'opens the recordset and scans the table
    'Exit Subs in this loop is used to skip the rest of the codes when conditions are met
    rs.Open
    Do Until rs.EOF
        If rs("role") = "registrar" Then 'if the admin's role is a registrar
            If rs("pssw") = pssw Then 'if the password entered is correct
                'increments the times the user has logged in
                rs("login_count") = rs("login_count") + 1
                rs.Update

                regadmin.role = rs("role").Value
                localip = frmLogin.sckMain.localip 'sets the program's local ip to the computer's network ip address
                SaveSettings (IIf(remember = 1, usrn, ""))
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
                LogIn = True
                Exit Function
            Else 'If the Password entered is wrong
                MsgBox "Wrong Password!", vbOKOnly + vbExclamation
                LogIn = False
                Exit Function
            End If
        Else 'If the Admin role is not a registrar
            MsgBox "Must use registrar account!", vbOKOnly + vbExclamation
            LogIn = False
            Exit Function
        End If
    Loop
    'If the scanning didn't match records
    MsgBox "Wrong username or username doesn't exist!", vbOKOnly + vbExclamation
    LogIn = False
ProcExit:
    Exit Function
    
ProcError:
    MsgBox "Invalid credentials!", vbExclamation
    LogEvent (Err.Description)
    LogIn = False
    Resume ProcExit
End Function

'Method used to count students on queue in the table of the database
Public Function EnrolleeCount() As Integer
    On Error GoTo ProcError 'If something goes wrong, skip to the Error message
    'sets the RecordSet for counting the enrollees
    Set rs = recordSet("montessori_queue", "status", "onqueue")
    'Opens the recordset
    rs.Open
    'Returns the value of the query
    Do Until rs.EOF
        EnrolleeCount = rs(0)
        rs.Close
        Exit Function
    Loop
ProcExit:
    Exit Function
    
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Function

'Method used to count students on process in the table of the database
Public Function OnProcessCount() As Integer
    On Error GoTo ProcError 'If something goes wrong, skip to the Error message
    'sets the RecordSet for counting the enrollees
    Set rs = recordSet("montessori_queue", "status", "onprocess")
    'Opens the recordset
    rs.Open
    'Returns the value of the query
    Do Until rs.EOF
        OnProcessCount = rs(0)
        rs.Close
        Exit Function
    Loop
ProcExit:
    Exit Function
    
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Function

Public Function Enroll(queueID As Integer) As Boolean
On Error GoTo ProcError 'If something goes wrong, skip to the Error message
    'sets the RecordSet for counting the enrollees
    Set rs = recordSet("montessori_queue", "Queue_ID", queueID)
    
    'Opens the recordset
    rs.Open
    Do Until rs.EOF
        rs("status").Value = "enrolled"
        rs.Update
        rs.Close
        MsgBox "The student has been successfully enrolled!", vbInformation
        Unload Me
        Enroll = True
        Exit Function
    Loop
    MsgBox "There has been a problem, contact your admin!", vbExclamation
    Enroll = False
ProcExit:
    Exit Function
ProcError:
    MsgBox Err.Description, vbExclamation
    Enroll = False
    Resume ProcExit
End Function
