VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log-in"
   ClientHeight    =   4065
   ClientLeft      =   5190
   ClientTop       =   5340
   ClientWidth     =   7380
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
   ScaleHeight     =   4065
   ScaleWidth      =   7380
   Begin VB.CheckBox chkRemember 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Remember Me"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2640
      TabIndex        =   9
      Top             =   3360
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   120
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtUsrn 
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
      Left            =   3360
      TabIndex        =   0
      Text            =   "accountantos"
      Top             =   1740
      Width           =   2655
   End
   Begin VB.TextBox txtPssw 
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
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "|"
      TabIndex        =   1
      Text            =   "regpssw"
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txtIP 
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
      Left            =   3360
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CommandButton cmdLogIn 
      Caption         =   "Log-in"
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
      Left            =   4800
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Accountant"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   240
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1815
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
      Left            =   2160
      TabIndex        =   7
      Top             =   480
      Width           =   5055
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   7200
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Username"
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
      Left            =   2160
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Password"
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
      Left            =   2160
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "IP Address"
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
      Left            =   2160
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogIn_Click()
    'Call LogIn(txtUsrn.Text, txtPssw.Text, txtIP.Text)
    
    Dim loginParams As Dictionary
    Set loginParams = New Dictionary
    loginParams.Add "usrn", txtUsrn.Text
    loginParams.Add "pssw", txtPssw.Text
    loginParams.Add "role", "accountant"
    ipaddress = txtIP.Text 'inserts the ip entered to the global variable
    
    acctadmin.usrn = txtUsrn.Text
    acctadmin.pssw = txtPssw.Text
    acctadmin.role = "accountant"
    
    EnableDisableControls (False)
    
    Call sendRequest(sckMain, hAPI_LOGIN, loginParams, hPOST_METHOD)
End Sub

Private Sub EnableDisableControls(enabled As Boolean)
    txtUsrn.enabled = enabled
    txtPssw.enabled = enabled
    txtIP.enabled = enabled
    
    chkRemember.enabled = enabled
    cmdLogIn.enabled = enabled
End Sub

Private Sub Form_Load()
    txtUsrn.Text = acctadmin.usrn
    txtPssw.Text = password
    chkRemember.Value = IIf(acctadmin.usrn = "", 0, 1)
    txtIP.Text = ipaddress
End Sub

Sub ifReturnKeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdLogIn_Click
    End If
End Sub

Private Sub txtIP_GotFocus()
    txtIP.SelStart = 0
    txtIP.SelLength = Len(txtIP.Text)
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
    Call ifReturnKeyPress(KeyAscii)
End Sub

Private Sub txtPssw_GotFocus()
    txtPssw.SelStart = 0
    txtPssw.SelLength = Len(txtPssw.Text)
End Sub

Private Sub txtPssw_KeyPress(KeyAscii As Integer)
    Call ifReturnKeyPress(KeyAscii)
End Sub

Private Sub txtUsrn_GotFocus()
    txtUsrn.SelStart = 0
    txtUsrn.SelLength = Len(txtUsrn.Text)
End Sub

Private Sub txtUsrn_KeyPress(KeyAscii As Integer)
    Call ifReturnKeyPress(KeyAscii)
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
        localip = sckMain.localip 'sets the program's local ip to the computer's network ip address
        
        Dim rememberValues As Boolean
        rememberValues = chkRemember.Value
        Dim usrname As String
        Dim ipadd As String
        usrname = IIf(rememberValues, txtUsrn.Text, "")
        ipadd = IIf(rememberValues, ipaddress, "")
        Call SaveSettings(usrname, ipadd)
        
        'prompts the user has logged in successfully
        MsgBox p.Item("message"), vbOKOnly + vbInformation 'prompts
        'sets the registrar form's labels with the current entries
        'increments the times the user has logged in

        Unload Me 'exits the current form
        'sets the registrar form's labels with the current entries
        frmAccountant.lbladmin = acctadmin.usrn
        frmAccountant.lblIP = localip
        'shows the registrar form
        frmAccountant.Show
    Else
        acctadmin.usrn = ""
        acctadmin.pssw = ""
        acctadmin.role = ""
        MsgBox p.Item("message"), vbOKOnly + vbExclamation 'prompts
    End If
    EnableDisableControls (True)
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbExclamation, "Connection Error"
    
    sckMain.Close
    EnableDisableControls (True)
End Sub

Private Sub sckMain_Close()
    blnConnected = False
    'MsgBox "Is Called"
    sckMain.Close
    
    EnableDisableControls (True)
End Sub

