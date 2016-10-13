VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSearch 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Student"
   ClientHeight    =   5175
   ClientLeft      =   5805
   ClientTop       =   4140
   ClientWidth     =   6810
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6810
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   120
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid gridStudents 
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4895
      _Version        =   393216
   End
   Begin VB.ComboBox cmbFilter 
      Height          =   390
      ItemData        =   "frmSearch.frx":0000
      Left            =   1680
      List            =   "frmSearch.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Search:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const defaultHeight As Integer = 2055
Private Const expandedHeight As Integer = 5655

Private Const expandHeight As Integer = 3600

Private searchResults As Collection


Private Sub SearchStudent(filter As Integer)
On Error GoTo ProcError
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    
    Select Case filter
        Case 0
        
    End Select
    
    rs.Source = "SELECT * FROM montessori_queue WHERE Queue_ID = " & currentStudentID
    rs.Open
    
    Do Until rs.EOF
        
        Exit Sub
    Loop
ProcExit:
    Exit Sub
    
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Sub

Private Sub cmdSearch_Click()
    Dim searchParams As Dictionary
    Set searchParams = New Dictionary
    searchParams.Add "usrn", regadmin.usrn
    searchParams.Add "pssw", regadmin.pssw
    searchParams.Add "role", regadmin.role
    searchParams.Add "action", aSEARCH_STUDENT
    Debug.Print (vbCrLf & cmbFilter.ListIndex & vbCrLf)
    searchParams.Add "filter_key", filterKeyFromIndex(cmbFilter.ListIndex)
    searchParams.Add "filter_value", txtSearch.Text
    
    blnConnected = False
    
    Call sendRequest(sckMain, hAPI_STUDENTS, searchParams, hPOST_METHOD)
End Sub

Private Function filterKeyFromIndex(index As Integer) As String
    Select Case index
        Case -1 To 0
            filterKeyFromIndex = ""
        Case 1
            filterKeyFromIndex = "Queue_ID"
        Case 2
            filterKeyFromIndex = "first_name"
        Case 3
            filterKeyFromIndex = "middle_name"
        Case 4
            filterKeyFromIndex = "last_name"
        Case 5
            filterKeyFromIndex = "home_address"
        Case 6
            filterKeyFromIndex = "current_grade"
    End Select
End Function

Private Sub RefreshTableView()
    gridStudents.Cols = 6
    gridStudents.rows = searchResults.Count + 1
    gridStudents.TextMatrix(0, 1) = "First Name"
    gridStudents.TextMatrix(0, 2) = "Middle Name"
    gridStudents.TextMatrix(0, 3) = "Last Name"
    gridStudents.TextMatrix(0, 4) = "Gender"
    gridStudents.TextMatrix(0, 5) = "Grade"
    
    
    Dim i As Integer
    For i = 1 To searchResults.Count
        Dim studentInfo As Dictionary
        Set studentInfo = searchResults(i)
        gridStudents.TextMatrix(i, 0) = Format(i, String(4, "0"))
        gridStudents.TextMatrix(i, 1) = studentInfo("first_name")
        gridStudents.TextMatrix(i, 2) = studentInfo("middle_name")
        gridStudents.TextMatrix(i, 3) = studentInfo("last_name")
        gridStudents.TextMatrix(i, 4) = studentInfo("gender")
        gridStudents.TextMatrix(i, 5) = grade(studentInfo("current_grade"))
    Next
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

Private Sub sckMain_Connect()
    blnConnected = True
End Sub

' this event occurs when data is arriving via winsock
Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)
    Dim strResponse As String
    
    sckMain.GetData strResponse, vbString, bytesTotal
    
    Dim p As Object
    Set p = JSON.parse(getJSONFromResponse(strResponse))
    Debug.Print (JSON.toString(p))
    Dim message As Dictionary

    If p.Item("response") = 1 Then
        Set searchResults = p.Item("message")
        
        Call RefreshTableView
    Else
        Set searchResults = New Collection
        Call RefreshTableView
        MsgBox p.Item("message"), vbExclamation
    End If
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbExclamation, "Connection Error"
    MsgBox "Is Called"
    sckMain.Close
End Sub

Private Sub sckMain_Close()
    blnConnected = False
    sckMain.Close
End Sub
