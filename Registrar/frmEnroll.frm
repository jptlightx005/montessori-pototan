VERSION 5.00
Begin VB.Form frmEnroll 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enroll Student"
   ClientHeight    =   4860
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
   ScaleHeight     =   4860
   ScaleWidth      =   6480
   Begin VB.Timer tmrEnable 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   4320
   End
   Begin VB.CommandButton cmdEnroll 
      Caption         =   "Enroll"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   14
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   5040
      TabIndex        =   13
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame frameExtendedInfo 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Student Information"
      Height          =   3975
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
         TabIndex        =   12
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblID 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N/A"
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Full Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Grade:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Balance:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Paid Last:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label lblFullName 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N/A"
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N/A"
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label lblGrade 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N/A"
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   2520
         Width           =   3375
      End
      Begin VB.Label lblBalance 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N/A"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label lblPaidDate 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N/A"
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   3480
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
Public studentToEnroll As student

Public Sub loadData()
    lblID.Caption = studentToEnroll.studentID
    lblFullName.Caption = studentToEnroll.fullName
    lblAddress.Caption = studentToEnroll.homeAddress
    
    lblGrade.Caption = grade(studentToEnroll.grade)
    lblBalance.Caption = studentToEnroll.balancePaid
    lblPaidDate.Caption = Format(studentToEnroll.datePaid, "mmmm dd, yyyy")
    
    cmdEnroll.Enabled = False
    countdown = 3
    tmrEnable.Enabled = True
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
'On Error GoTo ProcError 'If something goes wrong, skip to the Error message
    'sets the RecordSet for counting the enrollees
    Set rs = New ADODB.recordSet
    rs.ActiveConnection = cn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    'Counts the number of students on queue in the table
    rs.Source = "SELECT * FROM montessori_queue WHERE Queue_ID =" & studentToEnroll.queueID
    'Opens the recordset
    rs.Open
    Do Until rs.EOF
        rs("status").Value = "enrolled"
        rs.Update
        rs.Close
        MsgBox "The student has been successfully enrolled!", vbInformation
        Unload Me
        Exit Sub
    Loop
    MsgBox "There has been a problem, contact your admin!", vbExclamation
ProcExit:
    Exit Sub
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Sub

Private Sub Form_Load()
    loadData
End Sub

Private Sub tmrEnable_Timer()
    countdown = countdown - 1
    cmdEnroll.Caption = Str(countdown)
    If countdown < 0 Then
        cmdEnroll.Caption = "Enroll"
        cmdEnroll.Enabled = True
        tmrEnable.Enabled = False
    End If
End Sub
