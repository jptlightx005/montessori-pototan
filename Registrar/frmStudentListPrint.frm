VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStudentListPrint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Student List"
   ClientHeight    =   9435
   ClientLeft      =   225
   ClientTop       =   -5265
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
   ScaleHeight     =   9435
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   4440
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
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
      Left            =   6480
      TabIndex        =   3
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
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
      Left            =   7800
      TabIndex        =   2
      Top             =   8640
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid gridStudents 
      Height          =   5895
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   10398
      _Version        =   393216
      Cols            =   5
      BackColorFixed  =   16777215
      BackColorSel    =   16777215
      BackColorBkg    =   16777215
      GridColor       =   0
      WordWrap        =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      BorderStyle     =   0
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
   Begin VB.Label lblTeacher 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Teacher"
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
      Left            =   360
      TabIndex        =   8
      Top             =   9000
      Width           =   2775
   End
   Begin VB.Label lblTeacherName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "____________________"
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
      TabIndex        =   7
      Top             =   8640
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pototan, Iloilo"
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
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   9135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "F. Parcon St."
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
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   9135
   End
   Begin VB.Label lbl_exel 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "EXEL Montessori de Pototan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   9135
   End
   Begin VB.Label lblGrade 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Grade 3"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1440
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

Const heightDifference As Integer = 4110
Const widthDifference As Integer = 780
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()

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
    gridStudents.rows = 20
    gridStudents.TextMatrix(0, 1) = "First Name"
    gridStudents.TextMatrix(0, 2) = "Middle Name"
    gridStudents.TextMatrix(0, 3) = "Last Name"
    gridStudents.TextMatrix(0, 4) = "Gender"

    Dim i As Integer
    Dim totalWidth As Integer
    Dim totalHeight As Integer
    totalWidth = 0
    totalHeight = gridStudents.RowHeight(0) * gridStudents.rows
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
        totalHeight = totalHeight + gridStudents.RowHeight(i)
    Next
    For i = 0 To 4
        totalWidth = totalWidth + gridStudents.ColWidth(i)
    Next
    Me.width = totalWidth + widthDifference + 50
    Me.Height = totalHeight + heightDifference + 50
    lblGrade.Caption = listGrade
End Sub

Private Sub Form_Resize()
    gridStudents.width = Me.width - widthDifference
    lblGrade.width = Me.width - widthDifference
    lbl_exel.width = Me.width
    Label1.width = Me.width
    Label2.width = Me.width
    cmdClose.Left = Me.width - 2955
    cmdPrint.Left = Me.width - 1635

    gridStudents.Height = Me.Height - heightDifference
    cmdClose.Top = Me.Height - 1360
    cmdPrint.Top = Me.Height - 1360
    lblTeacherName.Top = Me.Height - 1360
    lblTeacher.Top = Me.Height - 1000
End Sub

