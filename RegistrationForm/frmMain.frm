VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enrollment Form"
   ClientHeight    =   10410
   ClientLeft      =   3510
   ClientTop       =   1950
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   7680
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Logout"
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
      Left            =   5880
      TabIndex        =   51
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
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
      Left            =   1800
      TabIndex        =   50
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
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
      Left            =   480
      TabIndex        =   26
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CheckBox chkComm 
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Communion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   25
      Top             =   9480
      Width           =   2775
   End
   Begin VB.CheckBox chkBaptized 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Baptized"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   24
      Top             =   9480
      Width           =   2175
   End
   Begin VB.TextBox txtLast 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   22
      Top             =   9000
      Width           =   2655
   End
   Begin VB.TextBox txtReligion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   23
      Top             =   9000
      Width           =   2655
   End
   Begin VB.TextBox txtGAddress 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   8280
      Width           =   2655
   End
   Begin VB.TextBox txtGTelNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   21
      Top             =   8280
      Width           =   2655
   End
   Begin VB.TextBox txtGRelation 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   19
      Top             =   7560
      Width           =   2655
   End
   Begin VB.TextBox txtGuardian 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   7560
      Width           =   2655
   End
   Begin VB.TextBox txtTelNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   6840
      Width           =   2655
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   6840
      Width           =   2655
   End
   Begin VB.TextBox txtMocc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox txtMother 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox txtFocc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox txtFather 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox txtPlace 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4680
      TabIndex        =   11
      Top             =   4560
      Width           =   2535
   End
   Begin VB.ComboBox cmbYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox cmbDay 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox cmbMonth 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "frmMain.frx":0000
      Left            =   480
      List            =   "frmMain.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtLName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtMName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox txtFName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   2295
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
      Height          =   435
      ItemData        =   "frmMain.frx":0068
      Left            =   3720
      List            =   "frmMain.frx":0081
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CheckBox chkNew 
      BackColor       =   &H00C0E0FF&
      Caption         =   "New Student"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Grade*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   49
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblIP 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   48
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0E0FF&
      Caption         =   "IP:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   47
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lbladmin 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   46
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Admin:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   45
      Top             =   1200
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   120
      Picture         =   "frmMain.frx":00A5
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Last School Attended*"
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
      Left            =   480
      TabIndex        =   44
      Top             =   8640
      Width           =   2895
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Religion*"
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
      Left            =   3360
      TabIndex        =   43
      Top             =   8640
      Width           =   2655
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Address*"
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
      Left            =   480
      TabIndex        =   42
      Top             =   7920
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Telephone Number*"
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
      Left            =   3360
      TabIndex        =   41
      Top             =   7920
      Width           =   2655
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Relation*"
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
      Left            =   3360
      TabIndex        =   40
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Guardian*"
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
      Left            =   480
      TabIndex        =   39
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Telephone Number*"
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
      Left            =   3360
      TabIndex        =   38
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Address*"
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
      Left            =   480
      TabIndex        =   37
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Occupation*"
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
      Left            =   3360
      TabIndex        =   36
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Mother's Name*"
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
      Left            =   480
      TabIndex        =   35
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Occupation*"
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
      Left            =   3360
      TabIndex        =   34
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Father's Name*"
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
      Left            =   480
      TabIndex        =   33
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Place of Birth*"
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
      Left            =   4680
      TabIndex        =   32
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Date of Birth*"
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
      Left            =   480
      TabIndex        =   31
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Gender*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   30
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Last Name*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   29
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Middle Name*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   28
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Name*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   27
      Top             =   2880
      Width           =   1575
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
      Top             =   480
      Width           =   5055
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   7080
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'checks if the selected month is a
'31 day, 30 day or february and a leap year
Private Sub cmbMonth_Click()
    Dim i As Integer
    Select Case cmbMonth.ListIndex
        Case 0, 2, 4, 6, 7, 9, 11
            cmbDay.Clear
            For i = 1 To 31
                cmbDay.AddItem (i)
            Next
        Case 3, 5, 8, 10
            cmbDay.Clear
            For i = 1 To 30
                cmbDay.AddItem (i)
            Next
        Case 1
            cmbDay.Clear
            If cmbYear.ListIndex >= 0 Then
                If cmbYear.Text Mod 4 = 0 Then
                    For i = 1 To 29
                        cmbDay.AddItem (i)
                    Next
                Else
                    For i = 1 To 28
                        cmbDay.AddItem (i)
                    Next
                End If
            Else
                For i = 1 To 28
                    cmbDay.AddItem (i)
                Next
            End If
    End Select
End Sub

'checks if the selected month is February
'and the selected year is a Leap Year
'therefore, the number of days will change
Private Sub cmbYear_Click()
    If cmbMonth.ListIndex = 1 Then
        Dim i As Integer
        cmbDay.Clear
        If cmbYear.Text Mod 4 = 0 Then
            For i = 1 To 29
                cmbDay.AddItem (i)
            Next
        Else
            For i = 1 To 28
                cmbDay.AddItem (i)
            Next
        End If
    End If
End Sub

'clears the boxes
Public Sub ClearBoxes()
    chkNew = 0
    cmbGrade.ListIndex = -1
    txtFName.Text = ""
    txtMName.Text = ""
    txtLName.Text = ""
    optMale.Value = False
    optFemale.Value = False
    cmbMonth.ListIndex = -1
    cmbDay.ListIndex = -1
    cmbYear.ListIndex = -1
    txtPlace.Text = ""
    txtFather.Text = ""
    txtFocc.Text = ""
    txtMother.Text = ""
    txtMocc.Text = ""
    txtAddress.Text = ""
    txtTelNo.Text = ""
    txtGuardian.Text = ""
    txtGRelation.Text = ""
    txtGAddress.Text = ""
    txtGTelNo.Text = ""
    txtLast.Text = ""
    txtReligion.Text = ""
    chkBaptized = 0
    chkComm = 0
End Sub

Private Sub cmdLogOut_Click()
    Call Logout
End Sub

'Reset Button action method
Private Sub cmdReset_Click()
    Call ClearBoxes
End Sub

'submit button action
'nested codes that encodes information into a single string
'separated by the delimiter "|"
Private Sub cmdSubmit_Click()
    Dim studentInf As String
    'initializes the studentInf variable
    studentInf = ""
    'if student is new
    studentInf = studentInf & chkNew.Value & "|"
    'if grade is selected
    If cmbGrade.ListIndex >= 0 Then
        studentInf = studentInf & grade(cmbGrade.ListIndex) & "|"
        'if the name fields are not empty
        If txtFName.Text <> "" And txtMName.Text <> "" And txtLName.Text <> "" Then
            studentInf = studentInf & txtFName.Text & "|"
            studentInf = studentInf & txtMName.Text & "|"
            studentInf = studentInf & txtLName.Text & "|"
            'if a gender is selected
            If optMale.Value = True Or optFemale.Value = True Then
                studentInf = studentInf & gender() & "|"
                If cmbMonth.ListIndex >= 0 And _
                    cmbDay.ListIndex >= 0 And _
                    cmbYear.ListIndex >= 0 Then
                    studentInf = studentInf & DoB(cmbMonth.ListIndex, CInt(cmbDay.Text), CInt(cmbYear.Text)) & "|"
                    'if the place of birth field is not empty
                    If txtPlace.Text <> "" Then
                        studentInf = studentInf & txtPlace.Text & "|"
                        'if the father's name and occupation fields are not empty
                        If txtFather.Text <> "" And txtFocc.Text <> "" Then
                            studentInf = studentInf & txtFather.Text & "|"
                            studentInf = studentInf & txtFocc.Text & "|"
                            'if the mother's name and occupation fields are not empty
                            If txtMother.Text <> "" And txtMocc.Text <> "" Then
                                studentInf = studentInf & txtMother.Text & "|"
                                studentInf = studentInf & txtMocc.Text & "|"
                                'if the address and telephone number field is not empty
                                If txtAddress.Text <> "" And txtTelNo.Text <> "" Then
                                    studentInf = studentInf & txtAddress.Text & "|"
                                    studentInf = studentInf & txtTelNo.Text & "|"
                                    'if the guardian and relation field is not empty
                                    If txtGuardian.Text <> "" And txtGRelation.Text <> "" Then
                                        studentInf = studentInf & txtGuardian.Text & "|"
                                        studentInf = studentInf & txtGRelation.Text & "|"
                                        If txtGAddress.Text <> "" And txtGTelNo.Text <> "" Then
                                            studentInf = studentInf & txtGAddress.Text & "|"
                                            studentInf = studentInf & txtGTelNo.Text & "|"
                                            'if the last school attended and religion field are not empty
                                            If txtLast.Text <> "" And txtReligion.Text <> "" Then
                                                studentInf = studentInf & txtLast.Text & "|"
                                                studentInf = studentInf & txtReligion.Text & "|"
                                                studentInf = studentInf & chkBaptized.Value & "|"
                                                studentInf = studentInf & chkComm.Value
                                                'confirm data
                                                Dim choice As Integer
                                                choice = MsgBox("Submit student's info? (Please re-check)", vbYesNo + vbQuestion, "Submission")
                                                If choice = vbYes Then
                                                    Call SubmitStudentInfo(studentInf)
                                                    Call ClearBoxes
                                                    Exit Sub
                                                ElseIf choice = vbNo Then
                                                    Exit Sub
                                                End If
                                            Else
                                                MsgBox "Please enter the last school you attended and your religion!", vbExclamation
                                            End If
                                        Else
                                            MsgBox "Please enter your guardian's address and telephone number!", vbExclamation
                                        End If
                                    Else
                                        MsgBox "Please enter your guardian's name and relation!", vbExclamation
                                    End If
                                End If
                            Else
                                MsgBox "Please enter your father's name and occupation!", vbExclamation
                            End If
                        Else
                            MsgBox "Please enter your father's name and occupation!", vbExclamation
                        End If
                    Else
                        MsgBox "Please enter your place of birth!", vbExclamation
                    End If
                Else
                    MsgBox "Please enter your birth date!", vbExclamation
                End If
            Else
                MsgBox "Please select your gender!", vbExclamation
            End If
        Else
            MsgBox "Please enter your full name to each of the boxes provided!", vbExclamation
        End If
    Else
        MsgBox "Please select a grade!", vbExclamation
    End If
End Sub

'returns the grade as a grade code
Private Function grade(gradeindex As Integer) As String
    Select Case gradeindex
        Case 0:
            grade = "preschool"
        Case 1:
            grade = "grade1"
        Case 2:
            grade = "grade2"
        Case 3:
            grade = "grade3"
        Case 4:
            grade = "grade4"
        Case 5:
            grade = "grade5"
        Case 6:
            grade = "grade6"
    End Select
End Function

'returns the gender as a single character code
Private Function gender() As String
    If optMale.Value = True Then
        gender = "M"
    ElseIf optFemale.Value = True Then
        gender = "F"
    End If
End Function

'Returns the formatted date of birth combined from the combo boxes
Private Function DoB(bm As Integer, bd As Integer, by As Integer) As String
    DoB = Format$(CDate((bm + 1) & "-" & bd & "-" & by), "yyyy-mm-dd")
End Function

'this serves as a testing information
Private Sub cmdTester_Click()
    cmbGrade.ListIndex = 3
    txtFName.Text = "Liza"
    txtMName.Text = "Gil"
    txtLName.Text = "Soberano"
    optFemale.Value = True
    cmbMonth.ListIndex = 3
    cmbDay.ListIndex = 10
    cmbYear.ListIndex = 6
    txtPlace.Text = "Pototan, Iloilo"
    txtFather.Text = "Enrique T. Soberano"
    txtFocc.Text = "Teacher"
    txtMother.Text = "Sue G. Soberano"
    txtMocc.Text = "Teacher"
    txtAddress.Text = "Brgy. Cau-ayan Pototan, Iloilo"
    txtTelNo.Text = "022 329 3293"
    txtGuardian.Text = "Sue G. Soberano"
    txtGRelation.Text = "Mother"
    txtGAddress.Text = "Brgy. Cau-ayan Pototan, Iloilo"
    txtGTelNo.Text = "022 329 3293"
    txtLast.Text = "Rizal Elementary School"
    txtReligion.Text = "Roman Catholic"
    chkBaptized.Value = 1
    chkComm.Value = 1
End Sub

Private Sub Form_Load()
    'indicates the admin that logged into the system
    lbladmin = admin.usrn
    lblIP = localip
    'saves the current admin as default
    Call SaveSettings
    'empties the date combo boxes to renew the items inside them
    cmbYear.Clear
    cmbDay.Clear
    Dim i As Integer
    For i = 1 To 31
        cmbDay.AddItem (i)
    Next
    For i = Year(Now) - 2 To Year(Now) - 20 Step -1
        cmbYear.AddItem (i)
    Next
End Sub
