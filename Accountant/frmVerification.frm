VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTransaction 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction"
   ClientHeight    =   3150
   ClientLeft      =   6990
   ClientTop       =   5625
   ClientWidth     =   5220
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
   ScaleHeight     =   3150
   ScaleWidth      =   5220
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   120
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Distribution"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4935
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Payment:"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblPayment 
         BackColor       =   &H00C0E0FF&
         Caption         =   "P0.00"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Proceed"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Preview"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtPayment 
      Height          =   435
      Left            =   2520
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Payment (in Pesos)"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public currentBalance As Double
Public studentID As String
Public studentName As String
Public studentAddress As String
Public cashPaid As Double
Public cashSet As Boolean
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdProceed_Click()
    If cashSet Then
        Dim vbChoice As Integer
        vbChoice = MsgBox("Update the student's balance?", vbYesNoCancel)
        If vbChoice = vbYes Then
            Dim paymentParams As Dictionary
            Set paymentParams = New Dictionary
            paymentParams.Add "usrn", acctadmin.usrn
            paymentParams.Add "pssw", acctadmin.pssw
            paymentParams.Add "role", acctadmin.role
            paymentParams.Add "action", aSTUDENT_PAYMENT
            paymentParams.Add "student_id", studentID
            paymentParams.Add "balance_paid", cashPaid
            blnConnected = False
            Call sendRequest(sckMain, hAPI_ACCOUNT, paymentParams, hPOST_METHOD)
        End If
    End If
End Sub

Private Sub cmdShow_Click()
    cashPaid = CDbl(txtPayment.Text)
    cashSet = True
    lblPayment.Caption = Format(cashPaid, "P##,##0.00")
End Sub

Private Sub txtPayment_Change()
    If Len(txtPayment.Text) <= 0 Then
        txtPayment.Text = "0"
    End If
End Sub

Private Sub txtPayment_GotFocus()
    txtPayment.SelStart = 0
    txtPayment.SelLength = Len(txtPayment.Text)
End Sub

Private Sub txtPayment_KeyPress(KeyAscii As Integer)
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Then
        Exit Sub
    End If
    KeyAscii = 0
End Sub

Private Sub lblPayment_Change()
    cmdProceed.enabled = cashSet And cashPaid > 0
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

        MsgBox p.Item("message"), vbInformation
        cmdProceed.enabled = False
        
        frmReceiptPrint.fName = studentName
        frmReceiptPrint.fAddress = studentAddress
        frmReceiptPrint.pAmount = cashPaid
        frmReceiptPrint.Show vbModal
        
        frmAccountant.ReloadData

        Unload Me
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

