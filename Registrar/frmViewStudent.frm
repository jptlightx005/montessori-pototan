VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmViewStudent 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
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
   ScaleHeight     =   3855
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExpand 
      Caption         =   "Expand"
      Height          =   495
      Left            =   1440
      TabIndex        =   45
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   120
      TabIndex        =   44
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtLName 
      Enabled         =   0   'False
      Height          =   435
      Left            =   240
      TabIndex        =   40
      Text            =   "Soberano"
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtFName 
      Enabled         =   0   'False
      Height          =   435
      Left            =   2640
      TabIndex        =   39
      Text            =   "Liza"
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtMName 
      Enabled         =   0   'False
      Height          =   435
      Left            =   5520
      TabIndex        =   38
      Text            =   "Gil"
      Top             =   720
      Width           =   2775
   End
   Begin VB.ComboBox cmbGrade 
      Enabled         =   0   'False
      Height          =   390
      ItemData        =   "frmViewStudent.frx":0000
      Left            =   240
      List            =   "frmViewStudent.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   7320
      TabIndex        =   36
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6000
      TabIndex        =   35
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame frameBio 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Personal Info"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   8415
      Begin VB.CheckBox chkBaptized 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Baptized"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   4560
         Width           =   2175
      End
      Begin VB.CheckBox chkComm 
         BackColor       =   &H00C0E0FF&
         Caption         =   "First Communion"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   18
         Top             =   4560
         Width           =   2775
      End
      Begin VB.TextBox txtReligion 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   17
         Top             =   3720
         Width           =   2415
      End
      Begin VB.TextBox txtLast 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   16
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox txtGTelNo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   15
         Top             =   3720
         Width           =   2655
      End
      Begin VB.TextBox txtGAddress 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   14
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox txtGuardian 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox txtGRelation 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3720
         Width           =   2655
      End
      Begin VB.TextBox txtAddress 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   11
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtTelNo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   10
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtMother 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtMocc 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtFather 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtFocc 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtPlace 
         Enabled         =   0   'False
         Height          =   435
         Left            =   5640
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox cmbYear 
         Enabled         =   0   'False
         Height          =   390
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbDay 
         Enabled         =   0   'False
         Height          =   390
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbMonth 
         Enabled         =   0   'False
         Height          =   390
         ItemData        =   "frmViewStudent.frx":003D
         Left            =   1560
         List            =   "frmViewStudent.frx":0065
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbGender 
         Enabled         =   0   'False
         Height          =   390
         ItemData        =   "frmViewStudent.frx":00A5
         Left            =   120
         List            =   "frmViewStudent.frx":00AF
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Religion"
         Height          =   375
         Left            =   5760
         TabIndex        =   34
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Last School"
         Height          =   375
         Left            =   5760
         TabIndex        =   33
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Telephone Number"
         Height          =   375
         Left            =   3000
         TabIndex        =   32
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address"
         Height          =   375
         Left            =   3000
         TabIndex        =   31
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Relation"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Guardian"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Telephone Number"
         Height          =   375
         Left            =   5760
         TabIndex        =   28
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address"
         Height          =   375
         Left            =   5760
         TabIndex        =   27
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Occupation"
         Height          =   375
         Left            =   3000
         TabIndex        =   26
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Mother"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Occupation"
         Height          =   375
         Left            =   3000
         TabIndex        =   24
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Father"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Place of Birth"
         Height          =   255
         Left            =   5640
         TabIndex        =   22
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Date of Birth"
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Gender"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   6480
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Last Name"
      Height          =   255
      Left            =   240
      TabIndex        =   43
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Name"
      Height          =   255
      Left            =   2640
      TabIndex        =   42
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Middle Name"
      Height          =   255
      Left            =   5520
      TabIndex        =   41
      Top             =   1200
      Width           =   2775
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
    cmbGender.Text = studentInfo("gender")
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
    txtAddress.Text = studentInfo("home_address")
    txtTelNo.Text = studentInfo("home_number")
    txtGuardian.Text = studentInfo("guardian_name")
    txtGRelation.Text = studentInfo("guardian_relation")
    txtGAddress.Text = studentInfo("guardian_address")
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
    cmbGrade.Enabled = Changeable
    txtLName.Enabled = Changeable
    txtFName.Enabled = Changeable
    txtMName.Enabled = Changeable
    cmbGender.Enabled = Changeable
    cmbMonth.Enabled = Changeable
    cmbDay.Enabled = Changeable
    cmbYear.Enabled = Changeable
    txtPlace.Enabled = Changeable
    txtFather.Enabled = Changeable
    txtFocc.Enabled = Changeable
    txtAddress.Enabled = Changeable
    txtMother.Enabled = Changeable
    txtMocc.Enabled = Changeable
    txtTelNo.Enabled = Changeable
    txtGuardian.Enabled = Changeable
    txtGAddress.Enabled = Changeable
    txtLast.Enabled = Changeable
    txtGRelation.Enabled = Changeable
    txtGTelNo.Enabled = Changeable
    txtReligion.Enabled = Changeable
    chkBaptized.Enabled = Changeable
    chkComm.Enabled = Changeable

    'Enables the save button
    cmdSave.Enabled = Changeable
End Sub

Private Sub cmdExpand_Click()
    If Not expanded Then
        frmViewStudent.Height = frmViewStudent.Height + expandHeight
        frameBio.Height = frameBio.Height + expandHeight
        cmdEdit.Top = cmdEdit.Top + expandHeight
        cmdSave.Top = cmdSave.Top + expandHeight
        cmdExpand.Top = cmdExpand.Top + expandHeight
        cmdClose.Top = cmdClose.Top + expandHeight
        cmdExpand.Caption = "Collapse"
        expanded = True
    Else
        frmViewStudent.Height = frmViewStudent.Height - expandHeight
        frameBio.Height = frameBio.Height - expandHeight
        cmdEdit.Top = cmdEdit.Top - expandHeight
        cmdSave.Top = cmdSave.Top - expandHeight
        cmdExpand.Top = cmdExpand.Top - expandHeight
        cmdClose.Top = cmdClose.Top - expandHeight
        cmdExpand.Caption = "Expand"
        
        expanded = False
    End If
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
    updatedRecord.Add "student_id", studentInfo("student_id")
    
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
    updatedRecord.Add "home_address", txtAddress.Text
    updatedRecord.Add "home_number", txtTelNo.Text
    updatedRecord.Add "guardian_name", txtGuardian.Text
    updatedRecord.Add "guardian_relation", txtGRelation.Text
    updatedRecord.Add "guardian_address", txtGAddress.Text
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
        MsgBox p.Item("message"), vbOKOnly + vbExclamation 'prompts
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

