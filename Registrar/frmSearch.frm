VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSearch 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Student"
   ClientHeight    =   5175
   ClientLeft      =   5805
   ClientTop       =   4140
   ClientWidth     =   6810
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
   ScaleHeight     =   5175
   ScaleWidth      =   6810
   Begin VB.ComboBox cmbFilter 
      Height          =   390
      ItemData        =   "frmSearch.frx":0000
      Left            =   1680
      List            =   "frmSearch.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid gridStudents 
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   5
      Enabled         =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Search:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const defaultHeight As Integer = 2055
Private Const expandedHeight As Integer = 5655

Private Const expandHeight As Integer = 3600

Private searchResults As Integer

Private Sub SearchStudent(filter As Integer)
On Error GoTo ProcError
    Set rs = New ADODB.recordSet
    rs.ActiveConnection = cn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    
    Select Case filter
        Case 0
        
    End Select
    
    rs.Source = "SELECT * FROM montessori_queue WHERE Queue_ID = " & currentStudentID
    rs.Open
    
    Do Until rs.EOF
        
        Exit Sub
    Loop
ProcExit:
    Exit Sub
    
ProcError:
    MsgBox Err.Description, vbExclamation
    Resume ProcExit
End Sub

