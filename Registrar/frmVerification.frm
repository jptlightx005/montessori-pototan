VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmVerification 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verification"
   ClientHeight    =   3750
   ClientLeft      =   4920
   ClientTop       =   4140
   ClientWidth     =   8595
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
   ScaleHeight     =   3750
   ScaleWidth      =   8595
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   6480
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.CommandButton cmdExpand 
      Caption         =   "Expand"
      Height          =   495
      Left            =   7200
      TabIndex        =   50
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame frameBio 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Personal Info"
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   6855
      Begin VB.ComboBox cmbGender 
         Enabled         =   0   'False
         Height          =   390
         ItemData        =   "frmVerification.frx":0000
         Left            =   120
         List            =   "frmVerification.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cmbMonth 
         Enabled         =   0   'False
         Height          =   390
         ItemData        =   "frmVerification.frx":0014
         Left            =   1560
         List            =   "frmVerification.frx":003C
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbDay 
         Enabled         =   0   'False
         Height          =   390
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbYear 
         Enabled         =   0   'False
         Height          =   390
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtPlace 
         Enabled         =   0   'False
         Height          =   435
         Left            =   5640
         TabIndex        =   30
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtFocc 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   29
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtFather 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtMocc 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   27
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtMother 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtTelNo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   25
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtAddress 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   24
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtGRelation 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   3720
         Width           =   2655
      End
      Begin VB.TextBox txtGuardian 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox txtGAddress 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   21
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox txtGTelNo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   20
         Top             =   3720
         Width           =   2655
      End
      Begin VB.TextBox txtLast 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   19
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox txtReligion 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5760
         TabIndex        =   18
         Top             =   3720
         Width           =   2415
      End
      Begin VB.CheckBox chkComm 
         BackColor       =   &H00C0E0FF&
         Caption         =   "First Communion"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   17
         Top             =   4560
         Width           =   2775
      End
      Begin VB.CheckBox chkBaptized 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Baptized"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Gender"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Date of Birth"
         Height          =   255
         Left            =   1560
         TabIndex        =   48
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Place of Birth"
         Height          =   255
         Left            =   5640
         TabIndex        =   47
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Father"
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Occupation"
         Height          =   375
         Left            =   3000
         TabIndex        =   45
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Mother"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Occupation"
         Height          =   375
         Left            =   3000
         TabIndex        =   43
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address"
         Height          =   375
         Left            =   5760
         TabIndex        =   42
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Telephone Number"
         Height          =   375
         Left            =   5760
         TabIndex        =   41
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Guardian"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Relation"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address"
         Height          =   375
         Left            =   3000
         TabIndex        =   38
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Telephone Number"
         Height          =   375
         Left            =   3000
         TabIndex        =   37
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Last School"
         Height          =   375
         Left            =   5760
         TabIndex        =   36
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Religion"
         Height          =   375
         Left            =   5760
         TabIndex        =   35
         Top             =   4080
         Width           =   2415
      End
   End
   Begin VB.Frame frameReq 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Requirements"
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   3615
      Begin VB.CheckBox chkBCert 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Birth Certificate"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2295
      End
      Begin VB.CheckBox chkReport 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Report Card"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox chkNoReport 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N/A"
         Height          =   315
         Left            =   2520
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7200
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   7200
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ComboBox cmbGrade 
      Enabled         =   0   'False
      Height          =   390
      ItemData        =   "frmVerification.frx":007C
      Left            =   2640
      List            =   "frmVerification.frx":0095
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
   Begin VB.CheckBox chkNew 
      BackColor       =   &H00C0E0FF&
      Caption         =   "New Student"
      BeginProperty DataFormat 
         Type            =   5
         Format          =   ""
         HaveTrueFalseNull=   1
         TrueValue       =   "True"
         FalseValue      =   "False"
         NullValue       =   ""
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   7
      EndProperty
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox txtMName 
      Enabled         =   0   'False
      Height          =   435
      Left            =   5520
      TabIndex        =   2
      Text            =   "Gil"
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtFName 
      Enabled         =   0   'False
      Height          =   435
      Left            =   2640
      TabIndex        =   1
      Text            =   "Liza"
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtLName 
      Enabled         =   0   'False
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Text            =   "Soberano"
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Middle Name"
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Name"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Last Name"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
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
Const movedButtonY As Integer = 6960

'default form height is 495

Const collapsedWindowheight As Integer = 1
Const expandedWindowheight As Integer = 8790

Public selectedStudent As Dictionary

'Loads current student's information
Public Sub LoadStudentInfo()
    Dim StudentInf() As String

    StudentInf = Split(selectedStudent("student_info"), "|")
    chkNew.Value = StudentInf(0)
    cmbGrade.ListIndex = grade(StudentInf(1), Me)
    txtLName.Text = StudentInf(4)
    txtFName.Text = StudentInf(2)
    txtMName.Text = StudentInf(3)
    cmbGender.Text = StudentInf(5)
    cmbMonth.ListIndex = Month(CDate(StudentInf(6))) - 1
    cmbDay.ListIndex = Day(CDate(StudentInf(6))) - 1
    
    Dim i As Integer
    For i = 0 To cmbYear.ListCount - 1
        If cmbYear.List(i) = Year(CDate(StudentInf(6))) Then
            cmbYear.ListIndex = i
            Exit For
        End If
    Next

    txtPlace.Text = StudentInf(7)
    txtFather.Text = StudentInf(8)
    txtFocc.Text = StudentInf(9)
    txtMother.Text = StudentInf(10)
    txtMocc.Text = StudentInf(11)
    txtAddress.Text = StudentInf(12)
    txtTelNo.Text = StudentInf(13)
    txtGuardian.Text = StudentInf(14)
    txtGRelation.Text = StudentInf(15)
    txtGAddress.Text = StudentInf(16)
    txtGTelNo.Text = StudentInf(17)
    txtLast.Text = StudentInf(18)
    txtReligion.Text = StudentInf(19)
    chkBaptized.Value = StudentInf(20)
    chkComm.Value = StudentInf(21)
End Sub

Private Sub chkBCert_Click()
    cmdRegister.Enabled = (chkBCert.Value = 1 And (chkReport.Value = 1 Or chkNoReport.Value = 1))
End Sub

Private Sub chkNoReport_Click()
    If chkNoReport.Value = 1 Then
        chkReport.Value = 0
        chkReport.Enabled = False
    Else
        chkReport.Enabled = True
    End If
    cmdRegister.Enabled = chkBCert.Value = 1 And (chkReport.Value = 1 Or chkNoReport.Value = 1)
End Sub

Private Sub chkReport_Click()
    cmdRegister.Enabled = chkBCert.Value = 1 And (chkReport.Value = 1 Or chkNoReport.Value = 1)
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
    chkNew.Enabled = Changeable
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
    
    chkBCert.Enabled = Not Changeable
    chkReport.Enabled = Not Changeable
    chkNoReport.Enabled = Not Changeable
    cmdRegister.Enabled = IIf(Changeable, False, chkBCert.Value = 1 And (chkReport.Value = 1 Or chkNoReport.Value = 1))
    
    'Enables the save button
    cmdSave.Enabled = Changeable
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

Private Sub cmdExpand_Click()
    If Not expanded Then
        frmVerification.Height = frmVerification.Height + expandHeight
        frameBio.Height = frameBio.Height + expandHeight
        frameBio.Width = expandedFrameWidth
        frameReq.Top = frameReq.Top + expandHeight
        cmdRegister.Top = cmdRegister.Top + expandHeight
        cmdEdit.Top = cmdEdit.Top + expandHeight
        cmdSave.Top = cmdSave.Top + expandHeight
        
        cmdExpand.Left = movedButtonX
        cmdExpand.Top = movedButtonY
        cmdExpand.Caption = "Collapse"
        expanded = True
    Else
        frmVerification.Height = frmVerification.Height - expandHeight
        frameBio.Height = frameBio.Height - expandHeight
        frameBio.Width = collapsedFrameWidth
        frameReq.Top = frameReq.Top - expandHeight
        cmdRegister.Top = cmdRegister.Top - expandHeight
        cmdEdit.Top = cmdEdit.Top - expandHeight
        cmdSave.Top = cmdSave.Top - expandHeight
        
        cmdExpand.Left = defaultButtonX
        cmdExpand.Top = defaultButtonY
        cmdExpand.Caption = "Expand"
        expanded = False
    End If
End Sub

Private Sub cmdRegister_Click()
'On Error GoTo ProcError
    If chkBCert = 1 And (chkReport = 1 Or chkNoReport = 1) Then
        Dim newRecord As Dictionary
        Set newRecord = New Dictionary
        newRecord.Add "usrn", regadmin.usrn
        newRecord.Add "pssw", regadmin.pssw
        newRecord.Add "role", regadmin.role
        newRecord.Add "action", "register_student"
    
        newRecord.Add "Queue_ID", selectedStudent("Queue_ID")
        newRecord.Add "current_grade", setgrade(cmbGrade.ListIndex)
        newRecord.Add "last_name", txtLName.Text
        newRecord.Add "first_name", txtFName.Text
        newRecord.Add "middle_name", txtMName.Text
        newRecord.Add "gender", cmbGender.Text
        newRecord.Add "date_of_birth", DoB(cmbMonth.ListIndex, CInt(cmbDay.Text), CInt(cmbYear.Text))
        newRecord.Add "place_of_birth", txtPlace.Text
        newRecord.Add "fathers_name", txtFather.Text
        newRecord.Add "father_occupation", txtFocc.Text
        newRecord.Add "mothers_name", txtMother.Text
        newRecord.Add "mother_occupation", txtMocc.Text
        newRecord.Add "home_address", txtAddress.Text
        newRecord.Add "home_number", txtTelNo.Text
        newRecord.Add "guardian_name", txtGuardian.Text
        newRecord.Add "guardian_relation", txtGRelation.Text
        newRecord.Add "guardian_address", txtGAddress.Text
        newRecord.Add "guardian_number", txtGTelNo.Text
        newRecord.Add "last_school_attended", txtLast.Text
        newRecord.Add "religion", txtReligion.Text
        newRecord.Add "is_baptized", chkBaptized.Value
        newRecord.Add "first_communion", chkComm.Value
    
        blnConnected = False
        Call sendRequest(sckMain, hAPI_ACCOUNT, newRecord, hPOST_METHOD)
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
Private Function DoB(bm As Integer, bd As Integer, by As Integer) As String
    
    DoB = Format$(CDate((bm + 1) & "-" & bd & "-" & by), "yyyy-mm-dd")
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
    
    Dim p As Object
    Set p = JSON.parse(getJSONFromResponse(strResponse))
    
    If p.Item("response") = 1 Then
        MsgBox "To Make sure: " & JSON.toString(p)
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
