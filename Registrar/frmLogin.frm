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
      Text            =   "registraros"
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
      Caption         =   "Registrar"
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

' we set this to true whil a connection is established
Private blnConnected As Boolean


'Calls the global login method
Private Sub cmdLogIn_Click()
    'Call LogIn(txtUsrn.Text, txtPssw.Text, txtIP.Text, chkRemember.Value)
    Dim eUrl As URL
    
    Dim strMethod As String
    Dim strData As String
    Dim strPostData As String
    Dim strHeaders As String
    
    Dim strHTTP As String
    Dim x As Integer
    
    strPostData = ""
    strHeaders = ""
    strMethod = "POST"
    
    If blnConnected Then Exit Sub
    
    ' get the url
    eUrl = ExtractUrl(txtIP.Text)
    
    If eUrl.Host = vbNullString Then
        MsgBox "Invalid Host", vbCritical, "ERROR"
    
        Exit Sub
    End If
    
     ' configure winsock
    sckMain.Protocol = sckTCPProtocol
    sckMain.RemoteHost = eUrl.Host
    
    If eUrl.Scheme = "http" Then
        If eUrl.Port > 0 Then
            sckMain.RemotePort = eUrl.Port
        Else
            sckMain.RemotePort = 80
        End If
    ElseIf eUrl.Scheme = vbNullString Then
        sckMain.RemotePort = 80
    Else
        MsgBox "Invalid protocol schema"
    End If
    
    ' build encoded data the data is url encoded in the form
    ' var1=value&var2=value
    strData = ""
    strData = strData & URLEncode("usrn") & "=" & txtUsrn.Text & "&" & _
                            URLEncode("pssw") & "=" & txtPssw.Text
                            
    If eUrl.Query <> vbNullString Then
        eUrl.URI = eUrl.URI & "?" & eUrl.Query
    End If
    
    ' check if any variables were supplied
    If strData <> vbNullString Then
        strData = Left(strData, Len(strData) - 1)
        
        
        If strMethod = "GET" Then
            ' if this is a GET request then the URL encoded data
            ' is appended to the URI with a ?
            If eUrl.Query <> vbNullString Then
                eUrl.URI = eUrl.URI & "&" & strData
            Else
                eUrl.URI = eUrl.URI & "?" & strData
            End If
        Else
            ' if it is a post request, the data is appended to the
            ' body of the HTTP request and the headers Content-Type
            ' and Content-Length added
            strPostData = strData
            strHeaders = "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
                         "Content-Length: " & Len(strPostData) & vbCrLf
                         
        End If
    End If
    
    ' clear the old HTTP response
    Dim response As String
    
    ' build the HTTP request in the form
    '
    ' {REQ METHOD} URI HTTP/1.0
    ' Host: {host}
    ' {headers}
    '
    ' {post data}
    strHTTP = strMethod & " " & eUrl.URI & " HTTP/1.0" & vbCrLf
    strHTTP = strHTTP & "Host: " & eUrl.Host & vbCrLf
    strHTTP = strHTTP & strHeaders
    strHTTP = strHTTP & vbCrLf
    strHTTP = strHTTP & strPostData

    response = strHTTP
    MsgBox response, vbInformation
    sckMain.Connect
    
    ' wait for a connection
    While Not blnConnected
        DoEvents
    Wend
    
    ' send the HTTP request
    sckMain.SendData strHTTP
End Sub

'The method called when newly loaded
Private Sub Form_Load()
    txtUsrn.Text = regadmin.usrn
    txtPssw.Text = regadmin.pssw
    chkRemember.Value = IIf(regadmin.usrn = "", 0, 1)
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

    ' we append this to the response box becuase data arrives
    ' in multiple packets
    MsgBox getJSONFromResponse(strResponse)
    
    Dim p As Object
        
   Set p = JSON.parse(getJSONFromResponse(strResponse))
   
   MsgBox "Parsed object output: " & JSON.toString(p)
   
  MsgBox "Get Bodystyle data: " & p.Item("status")
   
    MsgBox "Get Form Url data: " & p.Item("message")
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbExclamation, "ERROR"
    
    sckMain.Close
End Sub

Private Sub sckMain_Close()
    blnConnected = False
    
    sckMain.Close
End Sub
