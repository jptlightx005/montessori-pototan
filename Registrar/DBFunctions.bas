Attribute VB_Name = "DBFunctions"
Option Explicit

'Method used to count students on queue in the table of the database
Public Function EnrolleeCount() As Integer
    On Error GoTo ProcError 'If something goes wrong, skip to the Error message
    'sets the RecordSet for counting the enrollees
    Set rs = Recordset("montessori_queue", "status", "onqueue")
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
    Set rs = Recordset("montessori_queue", "status", "onprocess")
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
    Set rs = Recordset("montessori_queue", "Queue_ID", queueID)
    
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
