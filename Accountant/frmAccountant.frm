VERSION 5.00
Begin VB.Form frmAccountant 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accountant"
   ClientHeight    =   7545
   ClientLeft      =   5880
   ClientTop       =   3960
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
   ScaleHeight     =   7545
   ScaleWidth      =   7575
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Logout"
      Height          =   495
      Left            =   5880
      TabIndex        =   22
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   4800
      TabIndex        =   16
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      Height          =   435
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label lblPaidDate 
      BackColor       =   &H00C0E0FF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Label lblBalance 
      BackColor       =   &H00C0E0FF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   2760
      TabIndex        =   20
      Top             =   5400
      Width           =   3375
   End
   Begin VB.Label lblGrade 
      BackColor       =   &H00C0E0FF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H00C0E0FF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Label lblFullName 
      BackColor       =   &H00C0E0FF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   2760
      TabIndex        =   17
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Paid Last:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Balance:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Grade:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Address:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Full Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label lblID 
      BackColor       =   &H00C0E0FF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Student ID Number:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Search"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Admin:"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lbladmin 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0E0FF&
      Caption         =   "IP:"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblIP 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
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
      Picture         =   "frmAccountant.frx":0000
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
Attribute VB_Name = "frmAccountant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogOut_Click()
    Call Logout
End Sub

Public Sub resetBoxes()
    Set currentStudent = Nothing
    txtSearch.Text = ""
    cmdUpdate.Enabled = False
    lblID.Caption = "N/A"
    lblFullName.Caption = "N/A"
    lblAddress.Caption = "N/A"
    lblGrade.Caption = "N/A"
    lblBalance.Caption = "N/A"
    lblPaidDate.Caption = "N/A"
End Sub
Private Sub cmdReset_Click()
    Call resetBoxes
End Sub

Private Sub cmdSearch_Click()
    Set currentStudent = SearchStudent(txtSearch.Text)
    If Not currentStudent Is Nothing Then
        Dim fullName As String
        
        lblID.Caption = currentStudent.studentID
        lblFullName.Caption = currentStudent.fullName
        lblAddress.Caption = currentStudent.homeAddress
        lblGrade.Caption = currentStudent.grade
        lblBalance.Caption = Format(currentStudent.balancePaid, "P##,##0.00")
        lblPaidDate.Caption = Format(currentStudent.datePaid, "mmmm-dd-yyyy")
        cmdUpdate.Enabled = True
    Else
        MsgBox "Student not Found!", vbExclamation
        cmdUpdate.Enabled = False
    End If
End Sub

Private Sub cmdUpdate_Click()
    frmTransaction.Show vbModal
End Sub

Private Sub Form_Load()
    Call SaveSettings
End Sub


Public Sub ReloadData()
    Set currentStudent = SearchStudent(currentStudent.studentID)
    If Not currentStudent Is Nothing Then
        Dim fullName As String
        fullName = currentStudent.firstName & " " & Left$(CStr(currentStudent.middleName), 1) & " " & currentStudent.lastName
        lblID.Caption = currentStudent.studentID
        lblFullName.Caption = UCase(fullName)
        lblAddress.Caption = currentStudent.homeAddress
        lblGrade.Caption = currentStudent.grade
        lblBalance.Caption = Format(currentStudent.balancePaid, "P##,##0.00")
        lblPaidDate.Caption = Format(currentStudent.datePaid, "mmmm-dd-yyyy")
        cmdUpdate.Enabled = True
    End If
End Sub
