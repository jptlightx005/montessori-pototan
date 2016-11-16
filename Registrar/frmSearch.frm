VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmSearch 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Search Student"
   ClientHeight    =   5175
   ClientLeft      =   5880
   ClientTop       =   4215
   ClientWidth     =   6810
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
      TabIndex        =   5
      Top             =   1680
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4895
      _Version        =   393216
      WordWrap        =   -1  'True
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
      Left            =   5400
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   5400
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
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
      Left            =   5400
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
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
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Search Last Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const defaultHeight As Integer = 2055
Private Const expandedHeight As Integer = 5655

Private Const expandHeight As Integer = 3600

Private searchResults As Collection

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Call searchStudent
End Sub

Public Sub searchStudent()
    Dim searchParams As Dictionary
    Set searchParams = New Dictionary
    searchParams.Add "usrn", regadmin.usrn
    searchParams.Add "pssw", regadmin.pssw
    searchParams.Add "role", regadmin.role
    searchParams.Add "action", aSEARCH_STUDENT
    searchParams.Add "filter_key", "last_name"
    searchParams.Add "filter_value", txtSearch.Text

    blnConnected = False

    Call sendRequest(sckMain, hAPI_STUDENTS, searchParams, hPOST_METHOD)
End Sub


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
    For i = 0 To 5
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
    Me.width = totalWidth + 500
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

Private Sub cmdView_Click()
    Debug.Print ("ROWSEL IS " & gridStudents.RowSel)
    If gridStudents.RowSel > 0 Then
        Set frmViewStudent.studentInfo = searchResults(gridStudents.RowSel)
        frmViewStudent.Show vbModal
    End If
End Sub

Private Sub Form_Resize()
    txtSearch.width = Me.width - 2115
    cmdSearch.Left = Me.width - 1650
    cmdCancel.Left = Me.width - 1650
    gridStudents.width = Me.width - 555
    cmdView.Left = Me.width - 1650
End Sub

Private Sub gridStudents_EnterCell()
    Debug.Print gridStudents.ColWidth(gridStudents.ColSel)
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
