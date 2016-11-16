VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmEnroll 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enroll Student"
   ClientHeight    =   4335
   ClientLeft      =   5805
   ClientTop       =   3255
   ClientWidth     =   6480
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
   ScaleHeight     =   4335
   ScaleWidth      =   6480
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   600
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrEnable 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   3840
   End
   Begin VB.CommandButton cmdEnroll 
      Caption         =   "Enroll"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   12
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame frameExtendedInfo 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Student Information"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Student ID Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblID 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N/A"
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Full Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Grade:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Paid Last:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label lblFullName 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N/A"
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N/A"
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label lblGrade 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N/A"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   2520
         Width           =   3375
      End
      Begin VB.Label lblPaidDate 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N/A"
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   3000
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmEnroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim countdown As Integer
Public student As Dictionary
Public status As String

Public Sub loadData()
    lblID.Caption = student("Student_ID")
    lblFullName.Caption = student("first_name") & " " & student("last_name")
    Dim studentAddress As String
    studentAddress = student("home_address_brgy")
    studentAddress = studentAddress & ", " & student("home_address_city")
    studentAddress = studentAddress & ", " & student("home_address_province")
    lblAddress.Caption = studentAddress

    lblGrade.Caption = grade(student("current_grade"))
    lblPaidDate.Caption = Format(student("latest_payment"), "mmmm dd, yyyy")

    cmdEnroll.enabled = False
    countdown = 3
    tmrEnable.enabled = True
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
Private Sub cmdReset_Click()

End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdEnroll_Click()
    Dim enrollParams As Dictionary
    Set enrollParams = New Dictionary
    enrollParams.Add "usrn", regadmin.usrn
    enrollParams.Add "pssw", regadmin.pssw
    enrollParams.Add "role", regadmin.role
    enrollParams.Add "action", aENROLL_STUDENT

    enrollParams.Add "student_id", student("ID")
    blnConnected = False

    Call sendRequest(sckMain, hAPI_ACCOUNT, enrollParams, hPOST_METHOD)
End Sub

Private Sub Form_Load()
    loadData
End Sub

Private Sub tmrEnable_Timer()
    countdown = countdown - 1
    cmdEnroll.Caption = str(countdown)
    If countdown < 0 Then
        cmdEnroll.Caption = "Enroll"
        cmdEnroll.enabled = True
        tmrEnable.enabled = False
    End If
End Sub


Private Sub sckMain_Connect()
    blnConnected = True
End Sub

' this event occurs when data is arriving via winsock
Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)
    Dim strResponse As String

    sckMain.GetData strResponse, vbString, bytesTotal
    Debug.Print strResponse
    Dim p As Object
    Set p = JSON.parse(getJSONFromResponse(strResponse))
    Debug.Print (JSON.toString(p))
    Dim message As Dictionary

    If p.Item("response") = 1 Then
        MsgBox p.Item("message"), vbInformation
        Unload Me
    Else
        MsgBox p.Item("message"), vbExclamation
        Unload Me
    End If
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbExclamation, "Connection Error"
    sckMain.Close
End Sub

Private Sub sckMain_Close()
    blnConnected = False
    tmrEnable.enabled = True
    sckMain.Close
End Sub

