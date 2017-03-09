VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStatement 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statement of Account"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6735
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
   ScaleHeight     =   7650
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   6360
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblTeacher 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cashier"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Label lblTeacherName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "____________________"
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   6840
      Width           =   3255
   End
   Begin VB.Label lbl_exel 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "EXEL Montessori de Pototan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   240
      Width           =   6495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "F. Parcon St."
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Width           =   6495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pototan, Iloilo"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Student ID Number:"
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblID 
      BackColor       =   &H00FFFFFF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Full Name:"
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address:"
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Grade:"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Payment:"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Paid Last:"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label lblFullName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H00FFFFFF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label lblGrade 
      BackColor       =   &H00FFFFFF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label lblPayment 
      BackColor       =   &H00FFFFFF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Label lblPaidDate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   6000
      Width           =   3375
   End
   Begin VB.Label lblMatriculation 
      BackColor       =   &H00FFFFFF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   5400
      Width           =   3375
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Matriculation:"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Balance Left:"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label lblBalance 
      BackColor       =   &H00FFFFFF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label lblSchoolYear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "School Year:"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
End
Attribute VB_Name = "frmStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public selectedStudent As Dictionary

Private Sub lbladmin_Click()
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
        
    Dim BeginPage, EndPage, NumCopies, Orientation, i
    ' Set Cancel to True.
    cmnDlg.PrinterDefault = True
    cmnDlg.CancelError = True
    On Error GoTo ErrHandler
    ' Display the Print dialog box.
    cmnDlg.ShowPrinter
    
    ' Get user-selected values from the dialog box.
    BeginPage = cmnDlg.FromPage
    EndPage = cmnDlg.ToPage
    NumCopies = cmnDlg.Copies
    Orientation = cmnDlg.Orientation
    For i = 1 To NumCopies
        Set Printer.Font = lblFullName.Font
        Debug.Print Printer.FontName & " :: " & Printer.FontSize
        cmdPrint.Visible = False
        cmdClose.Visible = False
        PrintForm
        cmdPrint.Visible = True
        cmdClose.Visible = True
     'Printer.EndDoc
   Next
ErrHandler:
   ' User pressed Cancel button.
   Exit Sub
End Sub

Private Sub Form_Load()
    Dim fullName As String
        
    lblID.Caption = selectedStudent("Student_ID")
    lblFullName.Caption = selectedStudent("first_name") & " " & selectedStudent("last_name")
    Dim studentAddress As String
    studentAddress = selectedStudent("home_address_brgy")
    studentAddress = studentAddress & ", " & selectedStudent("home_address_city")
    studentAddress = studentAddress & ", " & selectedStudent("home_address_province")
    lblAddress.Caption = studentAddress
    lblSchoolYear.Caption = selectedStudent("school_year")
    lblGrade.Caption = grade(selectedStudent("current_grade"))
    lblPayment.Caption = Format(selectedStudent("total_payment"), "P##,##0.00")
    lblMatriculation.Caption = Format(selectedStudent("total_matriculation"), "P##,##0.00")
    Dim balanceLeft As Long
    balanceLeft = selectedStudent("total_matriculation") - selectedStudent("total_payment")
    lblBalance.Caption = Format(balanceLeft, "P##,##0.00")
    
    lblPaidDate.Caption = Format(selectedStudent("latest_payment"), "mmmm dd, yyyy")
End Sub

