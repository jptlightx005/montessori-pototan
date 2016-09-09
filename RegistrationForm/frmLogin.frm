VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log-in"
   ClientHeight    =   3990
   ClientLeft      =   5880
   ClientTop       =   4995
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7440
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   120
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLogIn 
      Caption         =   "Log-in"
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
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
      TabIndex        =   5
      Text            =   "127.0.0.1"
      Top             =   2280
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
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
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
      TabIndex        =   1
      Top             =   1260
      Width           =   2655
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
      TabIndex        =   6
      Top             =   2280
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
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
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
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   7200
      Y1              =   1080
      Y2              =   1080
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
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   240
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Log In Button Action
Private Sub cmdLogIn_Click()
    Call LogIn(txtUsrn.Text, txtPssw.Text, txtIP.Text)
End Sub

'text field at keypress actions
Sub ifReturnKeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdLogIn_Click
    End If
End Sub
Private Sub txtUsrn_KeyPress(KeyAscii As Integer)
    Call ifReturnKeyPress(KeyAscii)
End Sub
Private Sub txtPssw_KeyPress(KeyAscii As Integer)
    Call ifReturnKeyPress(KeyAscii)
End Sub
Private Sub txtIP_KeyPress(KeyAscii As Integer)
    Call ifReturnKeyPress(KeyAscii)
End Sub

'loads the user's default settings
Private Sub Form_Load()
    txtUsrn.Text = admin.usrn
    txtPssw.Text = password
    txtIP.Text = ipaddress
End Sub

'text field at focus actions
Private Sub txtIP_GotFocus()
    txtIP.SelStart = 0
    txtIP.SelLength = Len(txtIP.Text)
End Sub
Private Sub txtPssw_GotFocus()
    txtPssw.SelStart = 0
    txtPssw.SelLength = Len(txtPssw.Text)
End Sub
Private Sub txtUsrn_GotFocus()
    txtUsrn.SelStart = 0
    txtUsrn.SelLength = Len(txtUsrn.Text)
End Sub


