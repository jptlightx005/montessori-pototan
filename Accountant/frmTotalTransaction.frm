VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTotalTransaction 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9000
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
   ScaleHeight     =   7455
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   120
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   6840
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   5640
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid gridStudents 
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   10398
      _Version        =   393216
      BackColorFixed  =   16777215
      BackColorSel    =   16777215
      BackColorBkg    =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date:"
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
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblDate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "October 29, 2016"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Transactions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmTotalTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim transRecord As Dictionary
Dim dateNow As Date

Private Sub Form_Load()

    Dim listParams As Dictionary
    Set listParams = New Dictionary
    listParams.Add "usrn", acctadmin.usrn
    listParams.Add "pssw", acctadmin.pssw
    listParams.Add "role", acctadmin.role
    listParams.Add "action", aTRANSACTION_LIST
    listParams.Add "filter_date", Format$(Now, "yyyy-mm-dd")
    blnConnected = False

    Call sendRequest(sckMain, hAPI_ACCOUNT, listParams, hPOST_METHOD)
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
        
        lblID.Caption = selectedStudent("Student_ID")
        lblFullName.Caption = selectedStudent("first_name") & " " & selectedStudent("last_name")
        lblAddress.Caption = selectedStudent("home_address")
        lblSchoolYear.Caption = selectedStudent("school_year")
        lblGrade.Caption = grade(selectedStudent("current_grade"))
        lblPayment.Caption = Format(selectedStudent("total_payment"), "P##,##0.00")
        lblMatriculation.Caption = Format(selectedStudent("total_matriculation"), "P##,##0.00")
        Dim balanceLeft As Long
        balanceLeft = selectedStudent("total_matriculation") - selectedStudent("total_payment")
        lblBalance.Caption = Format(balanceLeft, "P##,##0.00")
        
        lblPaidDate.Caption = Format(selectedStudent("date_of_payment"), "mmmm dd, yyyy")
        cmdUpdate.enabled = True
        cmdPrint.enabled = True
    Else
        MsgBox p.Item("message"), vbExclamation
        cmdUpdate.enabled = False
        cmdPrint.enabled = False
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
