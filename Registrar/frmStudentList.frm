VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmStudentList 
   BackColor       =   &H00C0E0FF&
   Caption         =   "List of Students"
   ClientHeight    =   6540
   ClientLeft      =   3495
   ClientTop       =   2310
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   7890
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cmbGrade 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmStudentList.frx":0000
      Left            =   240
      List            =   "frmStudentList.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   240
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid gridStudents 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9975
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmStudentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private searchResults As Collection

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If searchResults.Count >= 0 Then
        Set frmStudentListPrint.studentList = searchResults
        frmStudentListPrint.listGrade = cmbGrade.Text
        frmStudentListPrint.Show vbModal
    End If
End Sub

Private Sub cmdView_Click()
    If cmbGrade.ListIndex >= 0 Then
        cmdView.enabled = False
        Dim searchParams As Dictionary
        Set searchParams = New Dictionary
        searchParams.Add "usrn", regadmin.usrn
        searchParams.Add "pssw", regadmin.pssw
        searchParams.Add "role", regadmin.role
        searchParams.Add "action", aSEARCH_STUDENT
        searchParams.Add "filter_key", "current_grade"
        searchParams.Add "filter_value", setgrade(cmbGrade.ListIndex)
    
        blnConnected = False
    
        Call sendRequest(sckMain, hAPI_STUDENTS, searchParams, hPOST_METHOD)
        
    Else
        MsgBox "Please select a grade!", vbInformation
    End If
End Sub

Private Function setgrade(gradeindex As Integer) As String
    Select Case gradeindex
        Case 0:
            setgrade = "preschool"
        Case 1:
            setgrade = "grade1"
        Case 2:
            setgrade = "grade2"
        Case 3:
            setgrade = "grade3"
        Case 4:
            setgrade = "grade4"
        Case 5:
            setgrade = "grade5"
        Case 6:
            setgrade = "grade6"
    End Select
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

Private Sub RefreshTableView()
    gridStudents.Cols = 6
    gridStudents.rows = searchResults.Count + 1
    gridStudents.TextMatrix(0, 1) = "First Name"
    gridStudents.TextMatrix(0, 2) = "Middle Name"
    gridStudents.TextMatrix(0, 3) = "Last Name"
    gridStudents.TextMatrix(0, 4) = "Gender"
    gridStudents.TextMatrix(0, 5) = "Grade"
    
    Dim totalWidth As Integer
    totalWidth = 0
        
    Dim i As Integer
    For i = 0 To 4
        gridStudents.ColWidth(i) = TextWidth(gridStudents.TextMatrix(0, i))
    Next
    
    For i = 1 To searchResults.Count
        Dim studentInfo As Dictionary
        Set studentInfo = searchResults(i)
        gridStudents.TextMatrix(i, 0) = Format(i, String(4, "0"))
        gridStudents.TextMatrix(i, 1) = studentInfo("first_name")
        gridStudents.TextMatrix(i, 2) = studentInfo("middle_name")
        gridStudents.TextMatrix(i, 3) = studentInfo("last_name")
        gridStudents.TextMatrix(i, 4) = studentInfo("gender")
        gridStudents.TextMatrix(i, 5) = grade(studentInfo("current_grade"))
        
        Dim j As Integer
        
        For j = 0 To 5
            If TextWidth(gridStudents.TextMatrix(i, j)) > gridStudents.ColWidth(j) Then
                gridStudents.ColWidth(j) = TextWidth(gridStudents.TextMatrix(i, j))
            End If
        Next
    Next
    For i = 0 To 5
        totalWidth = totalWidth + gridStudents.ColWidth(i)
    Next
    Me.Width = totalWidth + 680
End Sub

Private Sub Form_Load()
    Set searchResults = New Collection
    RefreshTableView
End Sub

Private Sub Form_Resize()
    Debug.Print Me.Width & "=="
    gridStudents.Width = Me.Width - 556
    gridStudents.Height = Me.Height - 1455
    cmdClose.Left = Me.Width - 1650
    cmdPrint.Left = Me.Width - 2970
    cmdView.Left = Me.Width - 4290
    cmbGrade.Width = Me.Width - 4635
End Sub

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
        cmdPrint.enabled = True
        Call RefreshTableView
    Else
        Set searchResults = New Collection
        Call RefreshTableView
        MsgBox p.Item("message"), vbExclamation
    End If
    cmdView.enabled = True
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbExclamation, "Connection Error"
    MsgBox "Is Called"
    cmdView.enabled = True
    sckMain.Close
End Sub

Private Sub sckMain_Close()
    blnConnected = False
    cmdView.enabled = True
    sckMain.Close
End Sub
