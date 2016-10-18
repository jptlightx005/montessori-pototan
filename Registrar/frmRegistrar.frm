VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmRegistrar 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar"
   ClientHeight    =   8310
   ClientLeft      =   5655
   ClientTop       =   3405
   ClientWidth     =   7575
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
   ScaleHeight     =   8310
   ScaleWidth      =   7575
   Begin VB.CommandButton cmdStudentList 
      Caption         =   "Student List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   49
      Top             =   2160
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   120
      Top             =   7800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search Student"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   48
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Frame frameEnroll 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enroll Students"
      Height          =   1095
      Left            =   240
      TabIndex        =   44
      Top             =   7080
      Width           =   7215
      Begin VB.CommandButton cmdEnroll 
         Caption         =   "Enroll"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5400
         TabIndex        =   47
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblOnProcessCount 
         BackColor       =   &H00C0E0FF&
         Caption         =   "0"
         Height          =   375
         Left            =   1920
         TabIndex        =   46
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Students on Process:"
         Height          =   615
         Left            =   240
         TabIndex        =   45
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Logout"
      Height          =   495
      Left            =   6240
      TabIndex        =   43
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdDrop 
      Caption         =   "Drop"
      Height          =   495
      Left            =   6240
      TabIndex        =   42
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   495
      Left            =   4920
      TabIndex        =   41
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Timer tmr_update 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   6480
   End
   Begin VB.Frame frameQueue 
      BackColor       =   &H00C0E0FF&
      Height          =   4335
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   7215
      Begin VB.Line Line3 
         X1              =   4920
         X2              =   4920
         Y1              =   240
         Y2              =   4440
      End
      Begin VB.Label lblGrade 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   9
         Left            =   5160
         TabIndex        =   40
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label lblGrade 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   8
         Left            =   5160
         TabIndex        =   39
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label lblGrade 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   7
         Left            =   5160
         TabIndex        =   38
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label lblGrade 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   6
         Left            =   5160
         TabIndex        =   37
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label lblGrade 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   5
         Left            =   5160
         TabIndex        =   36
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblGrade 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   4
         Left            =   5160
         TabIndex        =   35
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lblGrade 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   3
         Left            =   5160
         TabIndex        =   34
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblGrade 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   33
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblGrade 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   32
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblGrade 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Pre-school"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5160
         TabIndex        =   31
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Grade"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   30
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   9
         Left            =   2040
         TabIndex        =   29
         Top             =   3840
         Width           =   2775
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   8
         Left            =   2040
         TabIndex        =   28
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   7
         Left            =   2040
         TabIndex        =   27
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   6
         Left            =   2040
         TabIndex        =   26
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   5
         Left            =   2040
         TabIndex        =   25
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   4
         Left            =   2040
         TabIndex        =   24
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   3
         Left            =   2040
         TabIndex        =   23
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   22
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------"
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   21
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Liza G. Soberano"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   20
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "10"
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   19
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "9"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "8"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "7"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "6"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "5"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "4"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "3"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "2"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.Line Line2 
         X1              =   1800
         X2              =   1800
         Y1              =   240
         Y2              =   4320
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Queue ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Admin:"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lbladmin 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0E0FF&
      Caption         =   "IP:"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblIP 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label lblEnrollees 
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enrollees Left:"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   7080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   240
      Picture         =   "frmRegistrar.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lbl_exel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "EXEL Montessori de Pototan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "frmRegistrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim queueCollection As Collection
Dim action As String

'Drops the current student
Private Sub cmdDrop_Click()
    Dim choice As Integer
    choice = MsgBox("Drop the enrollee?", vbYesNo + vbExclamation)
    Select Case choice
        Case vbYes
            Dim student As Dictionary
            Set student = queueCollection(1)
            
            Dim dropParams As Dictionary
            Set dropParams = New Dictionary
            dropParams.Add "usrn", regadmin.usrn
            dropParams.Add "pssw", regadmin.pssw
            dropParams.Add "role", regadmin.role
            dropParams.Add "action", aDROP_STUDENT
            dropParams.Add "queue_id", student("Queue_ID")
            blnConnected = False
            action = aDROP_STUDENT
            
            Call sendRequest(sckMain, hAPI_QUEUE, dropParams, hPOST_METHOD)
            tmr_update.enabled = False
    End Select
End Sub

Private Sub cmdEnroll_Click()
    Dim inputID As String
    inputID = InputBox("Enter student's ID")
    If inputID <> "" Then
        If IsNumeric(inputID) Then
            Dim enrollParams As Dictionary
            Set enrollParams = New Dictionary
            enrollParams.Add "usrn", regadmin.usrn
            enrollParams.Add "pssw", regadmin.pssw
            enrollParams.Add "role", regadmin.role
            enrollParams.Add "action", aSEARCH_STUDENT
            enrollParams.Add "student_id", inputID
            blnConnected = False
            
            action = aSEARCH_STUDENT
            
            Call sendRequest(sckMain, hAPI_ACCOUNT, enrollParams, hPOST_METHOD)
            tmr_update.enabled = False
        Else
            MsgBox "Invalid Input!", vbExclamation
        End If
    End If
End Sub

Private Function CheckStudentOnQueue(queueID As String) As String
'On Error GoTo ProcError
    'sets the RecordSet for counting the enrollees
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = cn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    'Looks for student with the specified studentID
    rs.Source = "SELECT * FROM montessori_queue WHERE Queue_ID =" & queueID
    'Opens the recordset
    rs.Open
    'if student with queue id is found
    Do Until rs.EOF
        CheckStudentOnQueue = rs("status").Value
        GoTo ProcExit
    Loop
    CheckStudentOnQueue = "Student is not found on queue! Contact Administrator!"
    'if not, just prompt the user
    MsgBox "Student is not found!", vbExclamation
ProcExit:
    rs.Close
    Exit Function
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Function

Private Sub cmdLogOut_Click()
    Call Logout
End Sub

Private Sub cmdSearch_Click()
    'tmr_update.Enabled = False
    frmSearch.Show vbModal
End Sub

Private Sub cmdStudentList_Click()
    frmStudentList.Show vbModal
End Sub

'Views the current student's information in the queue
Private Sub cmdView_Click()
    If queueCollection.Count > 0 Then
        Set frmVerification.selectedStudent = queueCollection(1)
        frmVerification.Show vbModal
    End If
End Sub

'The action that the window executes when loaded
Private Sub Form_Load()
    'lblEnrollees.Caption = EnrolleeCount
    Call ClearBoxes
    Call LoadQueue
End Sub

'Empties the labels
Sub ClearBoxes()
    Dim i As Integer
    For i = 0 To 9
        lblID(i).Caption = ""
        lblName(i).Caption = ""
        lblGrade(i).Caption = ""
    Next
End Sub

'The method that loads the lists of students
Sub LoadQueue()
    Dim listParams As Dictionary
    Set listParams = New Dictionary
    listParams.Add "usrn", regadmin.usrn
    listParams.Add "pssw", regadmin.pssw
    listParams.Add "role", regadmin.role
    listParams.Add "action", aQUEUE_LIST
    blnConnected = False
    
    action = aQUEUE_LIST
    
    Call sendRequest(sckMain, hAPI_QUEUE, listParams, hPOST_METHOD)

    tmr_update.enabled = False
End Sub


'Observes the database if enrollees keep increasing
Private Sub tmr_update_Timer()
    Call ClearBoxes
    Call LoadQueue
End Sub


Private Sub sckMain_Connect()
    blnConnected = True
End Sub

' this event occurs when data is arriving via winsock
Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)
    Dim strResponse As String
    
    sckMain.GetData strResponse, vbString, bytesTotal
    'Debug.Print (strResponse)
    Dim p As Object
    Set p = JSON.parse(getJSONFromResponse(strResponse))
    'Debug.Print (JSON.toString(p))
    Dim message As Dictionary

    If p.Item("response") = 1 Then
        If action = aQUEUE_LIST Then
            Set message = p.Item("message")
            Set queueCollection = message("list")
            Dim j As Integer
            For j = 1 To queueCollection.Count
                Dim record As Dictionary
                Set record = queueCollection(j)
                Dim i As Integer
                i = j - 1
    
                lblID(i).Caption = record("Queue_ID")
                Dim StudentInf() As String
                Dim MNameArray() As Byte
                StudentInf = Split(record("student_info"), "|")
                MNameArray = StrConv(StudentInf(3), vbFromUnicode)
                lblName(i).Caption = StudentInf(2) & " " & Chr(MNameArray(0)) & ". " & StudentInf(4)
                lblGrade(i).Caption = grade(StudentInf(1), Me)
            Next
            lblEnrollees.Caption = message("onqueue")
            lblOnProcessCount.Caption = message("onprocess")
            cmdView.enabled = message("onqueue") > 0
            cmdDrop.enabled = message("onqueue") > 0
            cmdEnroll.enabled = message("onprocess") > 0
            
        ElseIf action = aDROP_STUDENT Then
            MsgBox p.Item("message"), vbInformation
            
        ElseIf action = aSEARCH_STUDENT Then
            Set frmEnroll.student = p.Item("message")
            frmEnroll.Show vbModal
        End If
        
    Else
        If action = aQUEUE_LIST Then
            Set message = p.Item("message")
            lblEnrollees.Caption = message("onqueue")
            lblOnProcessCount.Caption = message("onprocess")
        ElseIf action = aSEARCH_STUDENT Then
            MsgBox p.Item("message"), vbExclamation
        End If
    End If
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbExclamation, "Connection Error"
    MsgBox "Is Called"
    sckMain.Close
End Sub

Private Sub sckMain_Close()
    blnConnected = False
    tmr_update.enabled = True
    sckMain.Close
End Sub
