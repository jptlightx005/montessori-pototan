VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmVerification 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verification"
   ClientHeight    =   8220
   ClientLeft      =   3390
   ClientTop       =   1395
   ClientWidth     =   12405
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
   ScaleHeight     =   8220
   ScaleWidth      =   12405
   Begin VB.TextBox txtSchoolYear 
      Height          =   390
      Left            =   7560
      TabIndex        =   64
      Text            =   "2016-2017"
      Top             =   1065
      Width           =   1815
   End
   Begin VB.ComboBox cmbGender 
      Enabled         =   0   'False
      Height          =   390
      ItemData        =   "frmVerification.frx":0000
      Left            =   240
      List            =   "frmVerification.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CheckBox chkNew 
      BackColor       =   &H00C0E0FF&
      Caption         =   "New Student"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   52
      Top             =   240
      Width           =   2175
   End
   Begin VB.ComboBox cmbGrade 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmVerification.frx":001C
      Left            =   3480
      List            =   "frmVerification.frx":0035
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtFName 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   50
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtMName 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   49
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtLName 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   48
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox cmbMonth 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmVerification.frx":0059
      Left            =   3000
      List            =   "frmVerification.frx":0081
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox cmbDay 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox cmbYear 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtPlace 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   44
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox txtReligion 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   43
      Top             =   6000
      Width           =   2655
   End
   Begin VB.TextBox txtLast 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   42
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CheckBox chkBaptized 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Baptized"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9240
      TabIndex        =   41
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CheckBox chkComm 
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Communion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9240
      TabIndex        =   40
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Parents"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   31
      Top             =   2400
      Width           =   6015
      Begin VB.TextBox txtMocc 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   35
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtMother 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtFocc 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   33
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtFather 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Occupation"
         Height          =   495
         Left            =   3120
         TabIndex        =   39
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Mother's Name"
         Height          =   495
         Left            =   240
         TabIndex        =   38
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Occupation"
         Height          =   495
         Left            =   3120
         TabIndex        =   37
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Father's Name"
         Height          =   495
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Home Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   6240
      TabIndex        =   22
      Top             =   2400
      Width           =   6015
      Begin VB.TextBox txtCity 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   26
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtTelNo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   25
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtBrgy 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtProvince 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Telephone Number"
         Height          =   495
         Left            =   3120
         TabIndex        =   30
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Brgy/Street*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0E0FF&
         Caption         =   "City/Town*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Province*"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Guardian Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   8775
      Begin VB.TextBox txtGTelNo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   15
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtGRelation 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtGuardian 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtGCity 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   12
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtGBrgy 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtGProvince 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Telephone Number"
         Height          =   495
         Left            =   5880
         TabIndex        =   21
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Relation"
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Guardian"
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Brgy/Street"
         Height          =   495
         Left            =   3120
         TabIndex        =   18
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0E0FF&
         Caption         =   "City/Town"
         Height          =   495
         Left            =   5880
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Province"
         Height          =   495
         Left            =   3120
         TabIndex        =   16
         Top             =   1080
         Width           =   2055
      End
   End
   Begin VB.TextBox txtMatriculation 
      Height          =   390
      Left            =   4080
      TabIndex        =   7
      Text            =   "0"
      Top             =   7080
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   6480
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.Frame frameReq 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Requirements"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   6720
      Width           =   3615
      Begin VB.CheckBox chkBCert 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Birth Certificate"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.CheckBox chkReport 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Report Card"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox chkNoReport 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N/A"
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   11040
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   11040
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0E0FF&
      Caption         =   "School Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   63
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Name*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   61
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Middle Name*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   60
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Last Name*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   59
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Gender*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   58
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Date of Birth*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   57
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Place of Birth"
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
      Left            =   6960
      TabIndex        =   56
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Religion"
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
      Left            =   9240
      TabIndex        =   55
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Last School Attended"
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
      Left            =   9240
      TabIndex        =   54
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Grade*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   53
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Matriculation"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   6720
      Width           =   1500
   End
End
Attribute VB_Name = "frmVerification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Changeable As Boolean
Private expanded As Boolean

Const expandedFrameWidth As Integer = 8295
Const collapsedFrameWidth As Integer = 6855

Const expandHeight As Integer = 4560

Const defaultButtonX As Integer = 7200
Const defaultButtonY As Integer = 1680
Const movedButtonX As Integer = 5880
Const movedButtonY As Integer = 7680

'default form height is 495

Const collapsedWindowheight As Integer = 1
Const expandedWindowheight As Integer = 8790

Public selectedStudent As Dictionary

'Loads current student's information
Public Sub LoadStudentInfo()
'Dim StudentInf() As String

'StudentInf = Split(selectedStudent("student_info"), "|")
    chkNew.Value = selectedStudent("is_new")
    txtSchoolYear.Text = selectedStudent("temp_school_year")
    cmbGrade.ListIndex = grade(selectedStudent("current_grade"), Me)
    txtLName.Text = selectedStudent("last_name")
    txtFName.Text = selectedStudent("first_name")
    txtMName.Text = selectedStudent("middle_name")
    cmbGender.ListIndex = IIf(selectedStudent("gender") = "Male", 0, 1)
    Dim dateOfBirth As Date
    dateOfBirth = CDate(selectedStudent("date_of_birth"))
    cmbMonth.ListIndex = Month(dateOfBirth) - 1
    cmbDay.ListIndex = Day(dateOfBirth) - 1

    Dim i As Integer
    For i = 0 To cmbYear.ListCount - 1
        If cmbYear.List(i) = Year(dateOfBirth) Then
            cmbYear.ListIndex = i
            Exit For
        End If
    Next

    txtPlace.Text = selectedStudent("place_of_birth")
    txtFather.Text = selectedStudent("fathers_name")
    txtFocc.Text = selectedStudent("father_occupation")
    txtMother.Text = selectedStudent("mothers_name")
    txtMocc.Text = selectedStudent("mother_occupation")
    'txtAddress.Text = StudentInf(12)
    txtBrgy.Text = selectedStudent("home_address_brgy")
    txtCity.Text = selectedStudent("home_address_city")
    txtProvince.Text = selectedStudent("home_address_province")
    txtTelNo.Text = selectedStudent("home_number")

    txtGuardian.Text = selectedStudent("guardian_name")
    txtGRelation.Text = selectedStudent("guardian_relation")
    'txtGAddress.Text = StudentInf(12)
    txtGBrgy.Text = selectedStudent("guardian_address_brgy")
    txtGCity.Text = selectedStudent("guardian_address_city")
    txtGProvince.Text = selectedStudent("guardian_address_province")
    txtGTelNo.Text = selectedStudent("guardian_number")
    txtLast.Text = selectedStudent("last_school_attended")
    txtReligion.Text = selectedStudent("religion")
    chkBaptized.Value = selectedStudent("is_baptized")
    chkComm.Value = selectedStudent("first_communion")
End Sub

Private Sub chkBCert_Click()
    cmdRegister.enabled = (chkBCert.Value = 1 And (chkReport.Value = 1 Or chkNoReport.Value = 1))
End Sub

Private Sub chkNoReport_Click()
    If chkNoReport.Value = 1 Then
        chkReport.Value = 0
        chkReport.enabled = False
    Else
        chkReport.enabled = True
    End If
    cmdRegister.enabled = chkBCert.Value = 1 And (chkReport.Value = 1 Or chkNoReport.Value = 1)
End Sub

Private Sub chkReport_Click()
    cmdRegister.enabled = chkBCert.Value = 1 And (chkReport.Value = 1 Or chkNoReport.Value = 1)
End Sub

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

Sub EnableDisableControls()
    chkNew.enabled = Changeable
    cmbGrade.enabled = Changeable
    txtLName.enabled = Changeable
    txtFName.enabled = Changeable
    txtMName.enabled = Changeable
    cmbGender.enabled = Changeable
    cmbMonth.enabled = Changeable
    cmbDay.enabled = Changeable
    cmbYear.enabled = Changeable
    txtPlace.enabled = Changeable
    txtFather.enabled = Changeable
    txtFocc.enabled = Changeable
    'txtAddress.enabled = Changeable
    txtBrgy.enabled = Changeable
    txtCity.enabled = Changeable
    txtProvince.enabled = Changeable
    txtMother.enabled = Changeable
    txtMocc.enabled = Changeable
    txtTelNo.enabled = Changeable
    txtGuardian.enabled = Changeable
    'txtGAddress.enabled = Changeable
    txtGBrgy.enabled = Changeable
    txtGCity.enabled = Changeable
    txtGProvince.enabled = Changeable
    txtLast.enabled = Changeable
    txtGRelation.enabled = Changeable
    txtGTelNo.enabled = Changeable
    txtReligion.enabled = Changeable
    chkBaptized.enabled = Changeable
    chkComm.enabled = Changeable

    chkBCert.enabled = Not Changeable
    chkReport.enabled = Not Changeable
    chkNoReport.enabled = Not Changeable
    cmdRegister.enabled = IIf(Changeable, False, chkBCert.Value = 1 And (chkReport.Value = 1 Or chkNoReport.Value = 1))

    'Enables the save button
    cmdSave.enabled = Changeable
End Sub

Private Sub cmdEdit_Click()
    If Changeable Then
        'Disables submission controls for editing
        cmdEdit.Caption = "Edit"
        Changeable = False
        Call EnableDisableControls
    Else
        'Enables submission controls for submission
        cmdEdit.Caption = "Cancel"
        Changeable = True
        Call EnableDisableControls
    End If
End Sub

Private Sub cmdRegister_Click()
'On Error GoTo ProcError
    If chkBCert = 1 And (chkReport = 1 Or chkNoReport = 1) Then
        If CheckMatriculation() = False Then
            Exit Sub
        End If

        If Validation() Then
            Dim updateRecord As Dictionary
            Set updateRecord = New Dictionary
            updateRecord.Add "usrn", regadmin.usrn
            updateRecord.Add "pssw", regadmin.pssw
            updateRecord.Add "role", regadmin.role
            updateRecord.Add "action", aREGISTER_STUDENT

            updateRecord.Add "is_new", chkNew.Value
            updateRecord.Add "school_year", txtSchoolYear.Text
            updateRecord.Add "Student_ID", selectedStudent("ID")
            updateRecord.Add "current_grade", setgrade(cmbGrade.ListIndex)
            updateRecord.Add "last_name", Trim(txtLName.Text)
            updateRecord.Add "first_name", Trim(txtFName.Text)
            updateRecord.Add "middle_name", Trim(txtMName.Text)
            updateRecord.Add "gender", cmbGender.Text
            updateRecord.Add "date_of_birth", DoB(cmbMonth.ListIndex, CInt(cmbDay.Text), CInt(cmbYear.Text))
            updateRecord.Add "place_of_birth", Trim(txtPlace.Text)
            updateRecord.Add "fathers_name", Trim(txtFather.Text)

            updateRecord.Add "father_occupation", Trim(txtFocc.Text)
            updateRecord.Add "mothers_name", Trim(txtMother.Text)
            updateRecord.Add "mother_occupation", Trim(txtMocc.Text)
            'updateRecord.Add "home_address", Trim(txtAddress.Text)

            updateRecord.Add "home_address_brgy", Trim(txtBrgy.Text)
            updateRecord.Add "home_address_city", Trim(txtCity.Text)
            updateRecord.Add "home_address_province", Trim(txtProvince.Text)

            updateRecord.Add "home_number", Trim(txtTelNo.Text)
            updateRecord.Add "guardian_name", Trim(txtGuardian.Text)
            updateRecord.Add "guardian_relation", Trim(txtGRelation.Text)
            'updateRecord.Add "guardian_address", Trim(txtGAddress.Text)

            updateRecord.Add "guardian_address_brgy", Trim(txtGBrgy.Text)
            updateRecord.Add "guardian_address_city", Trim(txtGCity.Text)
            updateRecord.Add "guardian_address_province", Trim(txtGProvince.Text)

            updateRecord.Add "guardian_number", Trim(txtGTelNo.Text)
            updateRecord.Add "last_school_attended", Trim(txtLast.Text)
            updateRecord.Add "religion", Trim(txtReligion.Text)
            updateRecord.Add "is_baptized", chkBaptized.Value
            updateRecord.Add "first_communion", chkComm.Value
            updateRecord.Add "total_matriculation", txtMatriculation.Text

            blnConnected = False
            Call sendRequest(sckMain, hAPI_ACCOUNT, updateRecord, hPOST_METHOD)
        Else
            MsgBox "Please fill in required data!", vbExclamation
        End If
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

Private Sub cmdSave_Click()
    Changeable = False
    Call EnableDisableControls
    cmdEdit.Caption = "Edit"
End Sub

Private Sub Form_Load()
    expanded = False
    cmbYear.Clear
    cmbDay.Clear
    Dim i As Integer
    For i = 1 To 31
        cmbDay.AddItem (i)
    Next
    For i = Year(Now) - 4 To 1995 Step -1
        cmbYear.AddItem (i)
    Next
    Changeable = False
    Call LoadStudentInfo
End Sub
Private Sub sckMain_Connect()
    blnConnected = True
End Sub

' this event occurs when data is arriving via winsock
Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)
    Dim strResponse As String

    sckMain.GetData strResponse, vbString, bytesTotal
    Debug.Print (strResponse)
    Dim p As Object
    Set p = JSON.parse(getJSONFromResponse(strResponse))

    If p.Item("response") = 1 Then
        Dim message As String
        message = "The student has been registered!"

        MsgBox message, vbOKOnly + vbInformation

        frmStudentIDPrint.studentID = p.Item("message")
        frmStudentIDPrint.studentName = txtLName.Text
        frmStudentIDPrint.Show vbModal
        Unload Me
    Else
        MsgBox p.Item("message"), vbOKOnly + vbExclamation    'prompts
    End If
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbExclamation, "Connection Error"

    sckMain.Close
End Sub

Private Sub sckMain_Close()
    blnConnected = False

    sckMain.Close
End Sub

Private Function Validation() As Boolean
    Dim isValid As Boolean
    isValid = True
    isValid = isValid And cmbGrade.ListIndex >= 0
    isValid = isValid And txtFName.Text <> ""
    isValid = isValid And txtMName.Text <> ""
    isValid = isValid And txtLName.Text <> ""
    isValid = isValid And cmbGender.ListIndex >= 0
    isValid = isValid And cmbMonth.ListIndex >= 0
    isValid = isValid And cmbDay.ListIndex >= 0
    isValid = isValid And cmbYear.ListIndex >= 0
    isValid = isValid And txtBrgy.Text <> ""
    isValid = isValid And txtCity.Text <> ""
    isValid = isValid And txtProvince.Text <> ""

    Validation = isValid
End Function

Private Function CheckMatriculation() As Boolean
    Dim matriculation As Double
    matriculation = CDbl(txtMatriculation.Text)
    If matriculation < 10000 Then
        MsgBox "Matriculation should be 10000 and beyond", vbExclamation
        CheckMatriculation = False
    Else
        CheckMatriculation = True
    End If
End Function
Private Sub txtMatriculation_GotFocus()
    txtMatriculation.SelStart = 0
    txtMatriculation.SelLength = Len(txtMatriculation.Text)
End Sub

Private Sub txtMatriculation_KeyPress(KeyAscii As Integer)
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    ElseIf KeyAscii = vbKeyDecimal Or KeyAscii = Asc(".") Then
        If InStr(1, txtMatriculation.Text, ".") > 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtMatriculation_LostFocus()
    If txtMatriculation.Text = "" Then
        txtMatriculation.Text = "0"
    Else
        Dim num As Double
        num = CDbl(txtMatriculation.Text)
        txtMatriculation.Text = num
    End If
End Sub
