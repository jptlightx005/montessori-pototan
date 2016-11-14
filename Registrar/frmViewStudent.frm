VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmViewStudent 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   12525
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
   ScaleHeight     =   7005
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
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
      Left            =   240
      TabIndex        =   34
      Top             =   4080
      Width           =   8775
      Begin VB.TextBox txtGProvince 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   40
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtGBrgy 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   39
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtGCity 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   38
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtGuardian 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtGRelation 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtGTelNo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   35
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Province"
         Height          =   495
         Left            =   3120
         TabIndex        =   46
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0E0FF&
         Caption         =   "City/Town"
         Height          =   495
         Left            =   5880
         TabIndex        =   45
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Brgy/Street"
         Height          =   495
         Left            =   3120
         TabIndex        =   44
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Guardian"
         Height          =   495
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Relation"
         Height          =   495
         Left            =   240
         TabIndex        =   42
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Telephone Number"
         Height          =   495
         Left            =   5880
         TabIndex        =   41
         Top             =   1080
         Width           =   2655
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
      Left            =   6360
      TabIndex        =   25
      Top             =   1920
      Width           =   6015
      Begin VB.TextBox txtProvince 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtBrgy 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtTelNo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   27
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtCity 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   26
         Top             =   720
         Width           =   2655
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
         TabIndex        =   33
         Top             =   1080
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
         TabIndex        =   32
         Top             =   360
         Width           =   2055
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
         TabIndex        =   31
         Top             =   360
         Width           =   2055
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
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   6015
      Begin VB.TextBox txtFather 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtFocc 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   19
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtMother 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtMocc 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   17
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Father's Name"
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Occupation"
         Height          =   495
         Left            =   3120
         TabIndex        =   23
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Mother's Name"
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Occupation"
         Height          =   495
         Left            =   3120
         TabIndex        =   21
         Top             =   1080
         Width           =   2055
      End
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
      Left            =   9360
      TabIndex        =   15
      Top             =   6600
      Width           =   2295
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
      Left            =   9360
      TabIndex        =   14
      Top             =   6240
      Width           =   2175
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
      Left            =   9360
      TabIndex        =   13
      Top             =   4800
      Width           =   2655
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
      Left            =   9360
      TabIndex        =   12
      Top             =   5520
      Width           =   2655
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
      Left            =   7080
      TabIndex        =   11
      Top             =   1440
      Width           =   2535
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
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1440
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
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
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
      ItemData        =   "frmViewStudent.frx":0000
      Left            =   3120
      List            =   "frmViewStudent.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
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
      Left            =   7320
      TabIndex        =   7
      Top             =   600
      Width           =   2535
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
      Left            =   5160
      TabIndex        =   6
      Top             =   600
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
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Width           =   2295
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
      ItemData        =   "frmViewStudent.frx":0068
      Left            =   10320
      List            =   "frmViewStudent.frx":0081
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ComboBox cmbGender 
      Enabled         =   0   'False
      Height          =   390
      ItemData        =   "frmViewStudent.frx":00A5
      Left            =   360
      List            =   "frmViewStudent.frx":00AF
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   6240
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   4800
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
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
      Left            =   10320
      TabIndex        =   55
      Top             =   1080
      Width           =   1575
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
      Left            =   9360
      TabIndex        =   54
      Top             =   4440
      Width           =   2895
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
      Left            =   9360
      TabIndex        =   53
      Top             =   5160
      Width           =   2655
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
      Left            =   7080
      TabIndex        =   52
      Top             =   1080
      Width           =   2295
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
      Left            =   3120
      TabIndex        =   51
      Top             =   1080
      Width           =   2295
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
      Left            =   360
      TabIndex        =   50
      Top             =   1080
      Width           =   1215
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
      Left            =   7320
      TabIndex        =   49
      Top             =   240
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
      Left            =   5160
      TabIndex        =   48
      Top             =   240
      Width           =   1935
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
      Left            =   2760
      TabIndex        =   47
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmViewStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Changeable As Boolean
Private expanded As Boolean

Const expandHeight As Integer = 3600

'default form height is 495

Const collapsedWindowheight As Integer = 1
Const expandedWindowheight As Integer = 8790


Public studentInfo As Dictionary

Public Sub LoadStudentInfo()
    cmbGrade.ListIndex = grade(studentInfo("current_grade"), Me)
    txtLName.Text = studentInfo("last_name")
    txtFName.Text = studentInfo("first_name")
    txtMName.Text = studentInfo("last_name")
    cmbGender.ListIndex = IIf(studentInfo("gender") = "Male", 0, 1)
    cmbMonth.ListIndex = Month(CDate(studentInfo("date_of_birth"))) - 1
    cmbDay.ListIndex = Day(CDate(studentInfo("date_of_birth"))) - 1

    Dim i As Integer
    For i = 0 To cmbYear.ListCount - 1
        If cmbYear.List(i) = Year(CDate(studentInfo("date_of_birth"))) Then
            cmbYear.ListIndex = i
            Exit For
        End If
    Next

    txtPlace.Text = studentInfo("place_of_birth")
    txtFather.Text = studentInfo("fathers_name")
    txtFocc.Text = studentInfo("father_occupation")
    txtMother.Text = studentInfo("mothers_name")
    txtMocc.Text = studentInfo("mother_occupation")
    txtBrgy.Text = studentInfo("home_address_brgy")
    txtCity.Text = studentInfo("home_address_city")
    txtProvince.Text = studentInfo("home_address_province")
    txtTelNo.Text = studentInfo("home_number")
    txtGuardian.Text = studentInfo("guardian_name")
    txtGRelation.Text = studentInfo("guardian_relation")
    txtGBrgy.Text = studentInfo("guardian_address_brgy")
    txtGCity.Text = studentInfo("guardian_address_city")
    txtGProvince.Text = studentInfo("guardian_address_province")
    txtGTelNo.Text = studentInfo("guardian_number")
    txtLast.Text = studentInfo("last_school_attended")
    txtReligion.Text = studentInfo("religion")
    chkBaptized.Value = studentInfo("is_baptized")
    chkComm.Value = studentInfo("first_communion")
End Sub

Private Sub cmdClose_Click()
    Unload Me
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

Sub EnableDisableControls()
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
    txtBrgy.enabled = Changeable
    txtCity.enabled = Changeable
    txtProvince.enabled = Changeable
    txtMother.enabled = Changeable
    txtMocc.enabled = Changeable
    txtTelNo.enabled = Changeable
    txtGuardian.enabled = Changeable
    txtGBrgy.enabled = Changeable
    txtGCity.enabled = Changeable
    txtGProvince.enabled = Changeable
    txtLast.enabled = Changeable
    txtGRelation.enabled = Changeable
    txtGTelNo.enabled = Changeable
    txtReligion.enabled = Changeable
    chkBaptized.enabled = Changeable
    chkComm.enabled = Changeable

    'Enables the save button
    cmdSave.enabled = Changeable
End Sub

Private Sub cmdSave_Click()
    Changeable = False
    Call EnableDisableControls
    cmdEdit.Caption = "Edit"

    Dim updatedRecord As Dictionary
    Set updatedRecord = New Dictionary
    updatedRecord.Add "usrn", regadmin.usrn
    updatedRecord.Add "pssw", regadmin.pssw
    updatedRecord.Add "role", regadmin.role
    updatedRecord.Add "action", aUPDATE_STUDENT
    updatedRecord.Add "student_id", studentInfo("Student_ID")

    updatedRecord.Add "current_grade", setgrade(cmbGrade.ListIndex)
    updatedRecord.Add "last_name", txtLName.Text
    updatedRecord.Add "first_name", txtFName.Text
    updatedRecord.Add "middle_name", txtMName.Text
    updatedRecord.Add "gender", cmbGender.Text
    updatedRecord.Add "date_of_birth", DoB(cmbMonth.ListIndex, CInt(cmbDay.Text), CInt(cmbYear.Text))
    updatedRecord.Add "place_of_birth", txtPlace.Text
    updatedRecord.Add "fathers_name", txtFather.Text
    updatedRecord.Add "father_occupation", txtFocc.Text
    updatedRecord.Add "mothers_name", txtMother.Text
    updatedRecord.Add "mother_occupation", txtMocc.Text
    updatedRecord.Add "home_address_brgy", Trim(txtBrgy.Text)
    updatedRecord.Add "home_address_city", Trim(txtCity.Text)
    updatedRecord.Add "home_address_province", Trim(txtProvince.Text)

    updatedRecord.Add "home_number", txtTelNo.Text
    updatedRecord.Add "guardian_name", txtGuardian.Text
    updatedRecord.Add "guardian_relation", txtGRelation.Text
    updatedRecord.Add "guardian_address_brgy", Trim(txtGBrgy.Text)
    updatedRecord.Add "guardian_address_city", Trim(txtGCity.Text)
    updatedRecord.Add "guardian_address_province", Trim(txtGProvince.Text)
    updatedRecord.Add "guardian_number", txtGTelNo.Text
    updatedRecord.Add "last_school_attended", txtLast.Text
    updatedRecord.Add "religion", txtReligion.Text
    updatedRecord.Add "is_baptized", chkBaptized.Value
    updatedRecord.Add "first_communion", chkComm.Value

    blnConnected = False
    Call sendRequest(sckMain, hAPI_ACCOUNT, updatedRecord, hPOST_METHOD)
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call frmSearch.searchStudent
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
    If p.Item("response") = 1 Then
        MsgBox p.Item("message"), vbInformation
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

