VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPriorityNumber 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Number"
   ClientHeight    =   2310
   ClientLeft      =   8085
   ClientTop       =   3135
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   3255
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   0
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblLastName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Trump"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblQueueID 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "0004"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Student ID"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmPriorityNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public queueID As String
Public studentName As String

Private Sub Form_Activate()
    Dim BeginPage, EndPage, NumCopies, Orientation, i
    ' Set Cancel to True.
    cmnDlg.PrinterDefault = True
    cmnDlg.CancelError = True
    On Error GoTo errHandler
    ' Display the Print dialog box.
    cmnDlg.ShowPrinter
    
    ' Get user-selected values from the dialog box.
    BeginPage = cmnDlg.FromPage
    EndPage = cmnDlg.ToPage
    NumCopies = cmnDlg.Copies
    Orientation = cmnDlg.Orientation
    For i = 1 To NumCopies

     PrintForm

     'Printer.EndDoc
   Next
   Exit Sub
errHandler:
   ' User pressed Cancel button.
   Exit Sub
End Sub

Private Sub Form_Load()
    lblQueueID.Caption = queueID
    lblLastName.Caption = studentName
End Sub

