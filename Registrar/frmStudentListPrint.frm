VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStudentListPrint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Student List"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
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
      Left            =   360
      TabIndex        =   4
      Top             =   6840
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   4440
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   6480
      TabIndex        =   3
      Top             =   6840
      Width           =   1215
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
      Left            =   7800
      TabIndex        =   2
      Top             =   6840
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid gridStudents 
      Height          =   5895
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   10398
      _Version        =   393216
      Cols            =   5
      BackColorFixed  =   16777215
      BackColorSel    =   16777215
      BackColorBkg    =   16777215
      AllowUserResizing=   1
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
   Begin VB.Label lblGrade 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Grade 3"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmStudentListPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public studentList As Collection
Public listGrade As String

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook

Const heightDifference As Integer = 1965
Const widthDifference As Integer = 780
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    Set xlObject = New Excel.Application
 
    'This Adds a new woorkbook, you could open the workbook from file also
    Set xlWB = xlObject.Workbooks.Add
                
    Clipboard.Clear 'Clear the Clipboard
    With gridStudents
        'Select Full Contents (You could also select partial content)
        .Col = 0               'From first column
        .Row = 0               'From first Row (header)
        .ColSel = .Cols - 1    'Select all columns
        .RowSel = .rows - 1    'Select all rows
        Clipboard.SetText .Clip 'Send to Clipboard
    End With
            
    With xlObject.ActiveWorkbook.ActiveSheet
        .Range("B5").Select 'Select Cell A1 (will paste from here, to different cells)
        .Paste              'Paste clipboard contents
    End With
    
    xlObject.Columns.EntireColumn.AutoFit
    ' This makes Excel visible
    xlObject.Visible = True
End Sub

Private Sub cmdPrint_Click()
    
    Dim BeginPage, EndPage, NumCopies, Orientation, i
    ' Set Cancel to True.
    cmnDlg.PrinterDefault = True
    cmnDlg.CancelError = True
    On Error GoTo ErrHandler
    ' Display the Print dialog box.
    cmnDlg.ShowPrinter
    
    ' Get user-selected values from the dialog box.
    BeginPage = cmnDlg.FromPage
    EndPage = cmnDlg.ToPage
    NumCopies = cmnDlg.Copies
    Orientation = cmnDlg.Orientation
    For i = 1 To NumCopies
        cmdPrint.Visible = False
        cmdClose.Visible = False
        PrintForm
        cmdPrint.Visible = True
        cmdClose.Visible = True
     'Printer.EndDoc
   Next
ErrHandler:
   ' User pressed Cancel button.
   Exit Sub
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub Form_Load()
    gridStudents.Cols = 5
    gridStudents.rows = studentList.Count + 1
    gridStudents.TextMatrix(0, 1) = "First Name"
    gridStudents.TextMatrix(0, 2) = "Middle Name"
    gridStudents.TextMatrix(0, 3) = "Last Name"
    gridStudents.TextMatrix(0, 4) = "Gender"
    
    Dim i As Integer
    For i = 0 To 4
        gridStudents.Row = 0
        gridStudents.Col = i
        gridStudents.CellFontBold = True
        gridStudents.ColWidth(i) = TextWidth(gridStudents.TextMatrix(0, i))
    Next
    
    For i = 1 To studentList.Count
        Dim studentInfo As Dictionary
        Set studentInfo = studentList(i)
        gridStudents.TextMatrix(i, 0) = Format(i, String(4, "0"))
        gridStudents.TextMatrix(i, 1) = studentInfo("first_name")
        gridStudents.TextMatrix(i, 2) = studentInfo("middle_name")
        gridStudents.TextMatrix(i, 3) = studentInfo("last_name")
        gridStudents.TextMatrix(i, 4) = studentInfo("gender")
        
        Dim j As Integer
        For j = 0 To 4
            If TextWidth(gridStudents.TextMatrix(i, j)) > gridStudents.ColWidth(j) Then
                gridStudents.ColWidth(j) = TextWidth(gridStudents.TextMatrix(i, j))
            End If
        Next
    Next
    
    lblGrade.Caption = listGrade
End Sub

Private Sub Form_Resize()
    gridStudents.Width = Me.Width - widthDifference
    lblGrade.Width = Me.Width - widthDifference
    cmdClose.Left = Me.Width - 2955
    cmdPrint.Left = Me.Width - 1635
    
    gridStudents.Height = Me.Height - heightDifference
    cmdExport.Top = Me.Height - 1050
    cmdClose.Top = Me.Height - 1050
    cmdPrint.Top = Me.Height - 1050
End Sub

