VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enrollment Form"
   ClientHeight    =   9675
   ClientLeft      =   3150
   ClientTop       =   735
   ClientWidth     =   12780
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
   ScaleHeight     =   9675
   ScaleWidth      =   12780
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
      Left            =   480
      TabIndex        =   56
      Top             =   6720
      Width           =   8775
      Begin VB.TextBox txtGProvince 
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
         TabIndex        =   24
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtGBrgy 
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
         TabIndex        =   22
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtGCity 
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
         Left            =   5880
         TabIndex        =   23
         Top             =   720
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
         Left            =   240
         TabIndex        =   20
         Top             =   720
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
         Left            =   240
         TabIndex        =   21
         Top             =   1440
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
         Left            =   5880
         TabIndex        =   25
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Province"
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
         Left            =   3120
         TabIndex        =   62
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0E0FF&
         Caption         =   "City/Town"
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
         TabIndex        =   61
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Brgy/Street"
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
         Left            =   3120
         TabIndex        =   60
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Guardian"
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
         Left            =   240
         TabIndex        =   59
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Relation"
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
         Left            =   240
         TabIndex        =   58
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Telephone Number"
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
         TabIndex        =   57
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
      Left            =   6600
      TabIndex        =   51
      Top             =   4560
      Width           =   6015
      Begin VB.TextBox txtProvince 
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
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtBrgy 
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
         Left            =   240
         TabIndex        =   16
         Top             =   720
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
         Left            =   3120
         TabIndex        =   19
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtCity 
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
         TabIndex        =   17
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Telephone Number"
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
         Left            =   3120
         TabIndex        =   52
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
      Left            =   480
      TabIndex        =   46
      Top             =   4560
      Width           =   6015
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
         Left            =   240
         TabIndex        =   12
         Top             =   720
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
         Left            =   3120
         TabIndex        =   13
         Top             =   720
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
         Left            =   240
         TabIndex        =   14
         Top             =   1440
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
         Left            =   3120
         TabIndex        =   15
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Father's Name"
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
         Left            =   240
         TabIndex        =   50
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Occupation"
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
         Left            =   3120
         TabIndex        =   49
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Mother's Name"
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
         Left            =   240
         TabIndex        =   48
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Occupation"
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
         Left            =   3120
         TabIndex        =   47
         Top             =   1080
         Width           =   2055
      End
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   11760
      Top             =   9120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
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
      Left            =   8400
      TabIndex        =   45
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
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
      Left            =   7680
      TabIndex        =   44
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
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
      Left            =   6360
      TabIndex        =   30
      Top             =   9000
      Width           =   1215
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
      Left            =   3960
      TabIndex        =   29
      Top             =   9240
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
      Left            =   3960
      TabIndex        =   28
      Top             =   8880
      Width           =   2175
   End
   Begin VB.TextBox txtLast 
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
      TabIndex        =   26
      Top             =   7440
      Width           =   2655
   End
   Begin VB.TextBox txtReligion 
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
      TabIndex        =   27
      Top             =   8160
      Width           =   2655
   End
   Begin VB.TextBox txtPlace 
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
      Left            =   8760
      TabIndex        =   11
      Top             =   4080
      Width           =   2535
   End
   Begin VB.ComboBox cmbYear 
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
      Left            =   7440
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ComboBox cmbDay 
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ComboBox cmbMonth 
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
      ItemData        =   "frmMain.frx":0000
      Left            =   4800
      List            =   "frmMain.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Male"
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
      Left            =   2040
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
   End
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Female"
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
      Left            =   3240
      TabIndex        =   7
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtLName 
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
      Left            =   7440
      TabIndex        =   5
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox txtMName 
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
      Left            =   5280
      TabIndex        =   4
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox txtFName 
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
      Left            =   2880
      TabIndex        =   3
      Top             =   3240
      Width           =   2295
   End
   Begin VB.ComboBox cmbGrade 
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
      ItemData        =   "frmMain.frx":0068
      Left            =   6720
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
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      TabIndex        =   1
      Top             =   2400
      Width           =   2175
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
      Left            =   5880
      TabIndex        =   43
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
      Left            =   5640
      TabIndex        =   42
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
      Left            =   4920
      TabIndex        =   41
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
      Left            =   5640
      TabIndex        =   40
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
      Left            =   4560
      TabIndex        =   39
      Top             =   1200
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   2640
      Picture         =   "frmMain.frx":00A5
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1815
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
      TabIndex        =   38
      Top             =   7080
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
      TabIndex        =   37
      Top             =   7800
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
      Left            =   8760
      TabIndex        =   36
      Top             =   3720
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
      Left            =   4800
      TabIndex        =   35
      Top             =   3720
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
      Left            =   2040
      TabIndex        =   34
      Top             =   3720
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
      Left            =   7440
      TabIndex        =   33
      Top             =   2880
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
      Left            =   5280
      TabIndex        =   32
      Top             =   2880
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
      Left            =   2880
      TabIndex        =   31
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
      Left            =   4560
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
   Begin VB.Line Line1 
      X1              =   4560
      X2              =   9600
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
    'txtAddress.Text = ""
    txtBrgy.Text = ""
    txtCity.Text = ""
    txtProvince.Text = ""
    txtTelNo.Text = ""
    txtGuardian.Text = ""
    txtGRelation.Text = ""
    'txtGAddress.Text = ""
    txtGBrgy.Text = ""
    txtGCity.Text = ""
    txtGProvince.Text = ""
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
    If ValidateData() Then
        SubmitData
    Else
        MsgBox "Please fill in required data!", vbExclamation
    End If
End Sub

Private Function ValidateData() As Boolean
    Dim isValid As Boolean
    isValid = True
    isValid = isValid And cmbGrade.ListIndex >= 0
    isValid = isValid And txtFName.Text <> ""
    isValid = isValid And txtMName.Text <> ""
    isValid = isValid And txtLName.Text <> ""
    isValid = isValid And (optMale.Value Or optFemale.Value)
    isValid = isValid And cmbMonth.ListIndex >= 0
    isValid = isValid And cmbDay.ListIndex >= 0
    isValid = isValid And cmbYear.ListIndex >= 0
    isValid = isValid And txtBrgy.Text <> ""
    isValid = isValid And txtCity.Text <> ""
    isValid = isValid And txtProvince.Text <> ""
    
    ValidateData = isValid
End Function

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
    Debug.Print JSON.toString(p)
    If p.Item("response") = 1 Then
        Dim message As String
        message = "The student has been registered!"
        MsgBox message, vbOKOnly + vbInformation
        frmPriorityNumber.queueID = p.Item("message")
        frmPriorityNumber.studentName = Trim(txtLName.Text)
        frmPriorityNumber.Show vbModal
        
        Call ClearBoxes
    Else
        MsgBox p.Item("message"), vbOKOnly + vbExclamation 'prompts
    End If
    
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbExclamation, "Connection Error"
    
    sckMain.Close
End Sub

Private Sub sckMain_Close()
    blnConnected = False
    'MsgBox "Is Called"
    sckMain.Close
End Sub

Private Sub SubmitData()
    Dim newRecord As Dictionary
    Set newRecord = New Dictionary

    newRecord.Add "current_grade", grade(cmbGrade.ListIndex)
    newRecord.Add "last_name", Trim(txtLName.Text)
    newRecord.Add "first_name", Trim(txtFName.Text)
    newRecord.Add "middle_name", Trim(txtMName.Text)
    newRecord.Add "gender", gender()
    newRecord.Add "date_of_birth", DoB(cmbMonth.ListIndex, CInt(cmbDay.Text), CInt(cmbYear.Text))
    newRecord.Add "place_of_birth", Trim(txtPlace.Text)
    newRecord.Add "fathers_name", Trim(txtFather.Text)
    newRecord.Add "father_occupation", Trim(txtFocc.Text)
    newRecord.Add "mothers_name", Trim(txtMother.Text)
    newRecord.Add "mother_occupation", Trim(txtMocc.Text)
    'newRecord.Add "home_address", Trim(txtAddress.Text)
    Dim address As String
    address = Trim(txtBrgy.Text) & " "
    address = address & Trim(txtCity.Text) & " "
    address = address & Trim(txtProvince.Text)
    newRecord.Add "home_address", Trim(address)
    
    newRecord.Add "home_number", Trim(txtTelNo.Text)
    newRecord.Add "guardian_name", Trim(txtGuardian.Text)
    newRecord.Add "guardian_relation", Trim(txtGRelation.Text)
    'newRecord.Add "guardian_address", Trim(txtGAddress.Text)
    Dim gaddress As String
    address = Trim(txtGBrgy.Text) & " "
    address = address & Trim(txtGCity.Text) & " "
    address = address & Trim(txtGProvince.Text)
    newRecord.Add "guardian_address", Trim(gaddress)
    newRecord.Add "guardian_number", Trim(txtGTelNo.Text)
    newRecord.Add "last_school_attended", Trim(txtLast.Text)
    newRecord.Add "religion", Trim(txtReligion.Text)
    newRecord.Add "is_baptized", chkBaptized.Value
    newRecord.Add "first_communion", chkComm.Value

    Dim choice As Integer
    choice = MsgBox("Submit student's info? (Please re-check)", vbYesNo + vbQuestion, "Submission")
    
    If choice = vbYes Then
        newRecord.Add "usrn", admin.usrn
        newRecord.Add "pssw", admin.pssw
        newRecord.Add "role", admin.role
        newRecord.Add "action", aREGISTER_STUDENT
        newRecord.Add "registered_ip", localip
        blnConnected = False
        
        Call sendRequest(sckMain, hAPI_QUEUE, newRecord, hPOST_METHOD)
    End If
End Sub
