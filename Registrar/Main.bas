Attribute VB_Name = "Main"
'EXEL MONTESSORI ENROLLMENT SYSTEM

'ADODB variables
Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Public cmd As ADODB.Command
Public subcmd As ADODB.Command

'Program Settings
Public ipaddress As String
Public localip As String
Public registraradmin As String

'Current Student selected
Public currentstudent As Integer

'Main method of the System
'This contains the methods called to be executed by the program after opening
Sub Main()
    frm_login.Show
End Sub

Public Sub LogIn(usrn As String, pssw As String, ip As String)
    On Error GoTo ProcError
    ipaddress = ip
    
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=" & ipaddress & ";Database=montessori-db; User=" & usrn & ";Password=" & pssw & ";"

    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    rs.Source = "SELECT * FROM montessori_admin WHERE usrn = '" + usrn + "'"
    rs.Open
    Do Until rs.EOF
        If rs("role") = "registrar" Then
            If rs("pssw") = pssw Then
                If rs("is_online") = 0 Then
                    MsgBox "Logged in Successfully!", vbOKOnly + vbInformation
                    rs("login_count") = rs("login_count") + 1
                    rs.Update
                    registraradmin = usrn
                    Unload frm_login
                    localip = frm_login.sckMain.localip
                    frmRegistrar.lbladmin = registraradmin
                    frmRegistrar.lblIP = localip
                    frmRegistrar.Show
                    rs.Close
                    Exit Sub
                Else
                    MsgBox "The account is Online!", vbOKOnly + vbExclamation
                    Exit Sub
                End If
            Else
                MsgBox "Wrong Password!", vbOKOnly + vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "Must use registrar account!", vbOKOnly + vbExclamation
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

Public Function EnrolleeCount() As Integer
    On Error GoTo ProcError
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    rs.Source = "SELECT count(*) FROM montessori_queue WHERE status = 'onqueue'"
    rs.Open
    Do Until rs.EOF
        EnrolleeCount = rs(0)
        Exit Function
    Loop
ProcExit:
    Exit Function
    
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Function



