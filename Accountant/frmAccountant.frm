VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAccountant 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accountant"
   ClientHeight    =   8790
   ClientLeft      =   5805
   ClientTop       =   735
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
   ScaleHeight     =   8790
   ScaleWidth      =   7575
   Begin VB.CommandButton cmdTransaction 
      Caption         =   "Transaction List"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   30
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   29
      Top             =   8160
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   120
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
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
      Left            =   4440
      TabIndex        =   15
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   8160
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      Height          =   435
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "School Year:"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label lblSchoolYear 
      BackColor       =   &H00C0E0FF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   2760
      TabIndex        =   27
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Label lblBalance 
      BackColor       =   &H00C0E0FF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   2760
      TabIndex        =   26
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Balance Left"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Total Matriculation"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label lblMatriculation 
      BackColor       =   &H00C0E0FF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   2760
      TabIndex        =   23
      Top             =   6960
      Width           =   3375
   End
   Begin VB.Label lblPaidDate 
      BackColor       =   &H00C0E0FF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   7560
      Width           =   3375
   End
   Begin VB.Label lblPayment 
      BackColor       =   &H00C0E0FF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   2760
      TabIndex        =   20
      Top             =   6000
      Width           =   3375
   End
   Begin VB.Label lblGrade 
      BackColor       =   &H00C0E0FF&
      Caption         =   "N/A"
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      Top             =   5520
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
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Total Payment"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Grade:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   5520
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
      Caption         =   "Student ID"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Search ID"
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

Dim selectedStudent As Dictionary
Dim studentAddress As String

Private Sub cmdLogOut_Click()
    Call Logout
End Sub

Public Sub resetBoxes()
    Set selectedStudent = Nothing
    txtSearch.Text = ""
    cmdUpdate.enabled = False
    cmdPrint.enabled = False
    lblID.Caption = "N/A"
    lblFullName.Caption = "N/A"
    lblAddress.Caption = "N/A"
    lblSchoolYear.Caption = "N/A"
    
    lblGrade.Caption = "N/A"
    lblMatriculation.Caption = "N/A"
    lblPayment.Caption = "N/A"
    lblBalance.Caption = "N/A"
    lblPaidDate.Caption = "N/A"
End Sub

Private Sub cmdPrint_Click()
    Set frmStatement.selectedStudent = selectedStudent
    frmStatement.Show vbModal
End Sub

Private Sub cmdReset_Click()
    Call resetBoxes
End Sub

Private Sub cmdSearch_Click()
    If txtSearch.Text <> "" Then
        Dim searchParams As Dictionary
        Set searchParams = New Dictionary
        searchParams.Add "usrn", acctadmin.usrn
        searchParams.Add "pssw", acctadmin.pssw
        searchParams.Add "role", acctadmin.role
        searchParams.Add "action", aSEARCH_STUDENT
        searchParams.Add "student_id", txtSearch.Text
        blnConnected = False

        Call sendRequest(sckMain, hAPI_ACCOUNT, searchParams, hPOST_METHOD)
    End If
End Sub

Private Sub cmdTransaction_Click()
    frmTotalTransaction.Show vbModal
End Sub

Private Sub cmdUpdate_Click()
    frmTransaction.currentBalance = selectedStudent("total_payment")
    frmTransaction.studentID = selectedStudent("Student_ID")
    frmTransaction.studentName = selectedStudent("first_name") & " " & selectedStudent("last_name")
    frmTransaction.studentAddress = studentAddress
    frmTransaction.Show vbModal
End Sub

Private Sub Form_Load()
    'Call SaveSettings
End Sub


Public Sub ReloadData()
    Dim searchParams As Dictionary
    Set searchParams = New Dictionary
    searchParams.Add "usrn", acctadmin.usrn
    searchParams.Add "pssw", acctadmin.pssw
    searchParams.Add "role", acctadmin.role
    searchParams.Add "action", aSEARCH_STUDENT
    searchParams.Add "student_id", selectedStudent("StudentID")
    blnConnected = False

    Call sendRequest(sckMain, hAPI_ACCOUNT, searchParams, hPOST_METHOD)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
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
    Dim message As Dictionary

    If p.Item("response") = 1 Then
        Set selectedStudent = p.Item("message")

        Dim fullName As String
        
        lblID.Caption = selectedStudent("StudentID")
        Dim MNameArray() As Byte
        MNameArray = StrConv(selectedStudent("middle_name"), vbFromUnicode)
        lblFullName.Caption = selectedStudent("first_name") & " " & Chr(MNameArray(0)) & ". " & selectedStudent("last_name")
        
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
        cmdUpdate.enabled = True
        cmdPrint.enabled = True
    Else
        MsgBox p.Item("message"), vbExclamation
    End If
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbExclamation, "Connection Error"
    MsgBox "Is Called"
    sckMain.Close
End Sub

Private Sub sckMain_Close()
    blnConnected = False
    sckMain.Close
End Sub

