VERSION 5.00
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

Public cashPaid As Double
Public cashSet As Boolean
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdProceed_Click()
    On Error GoTo ProcError
    If cashSet Then
        'sets the RecordSet for the search method
        Set rs = New ADODB.Recordset
        rs.ActiveConnection = cn
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenDynamic
        rs.LockType = adLockOptimistic
        rs.Source = "SELECT * FROM montessori_records WHERE Student_ID=" & frmAccountant.currentStudent.studentID
        'opens the recordset and scans the table
        'Exit Subs in this loop is used to skip the rest of the codes when conditions are met
        rs.Open
        Do Until rs.EOF
            Dim vbChoice As Integer
            vbChoice = MsgBox("Update the student's balance?", vbYesNoCancel)
            If vbChoice = vbYes Then
                Dim currentBalance As Double
                currentBalance = rs("balance_paid").Value
                rs("balance_paid").Value = currentBalance + cashPaid
                rs("date_of_payment").Value = Now
                rs.Update
                rs.Close
                MsgBox "Transaction is now complete!", vbInformation
            End If
            cmdProceed.Enabled = False
            frmAccountant.ReloadData
            rs.Close
            Unload Me
            Exit Sub
        Loop
        MsgBox "There's an error occured"
    End If
ProcExit:
    Exit Sub
ProcError:
    MsgBox Err.Description, vbExclamation
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
    cmdProceed.Enabled = cashSet And cashPaid > 0
End Sub

