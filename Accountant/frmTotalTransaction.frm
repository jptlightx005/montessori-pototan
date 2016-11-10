VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTotalTransaction 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Total Transaction"
   ClientHeight    =   7455
   ClientLeft      =   5685
   ClientTop       =   1800
   ClientWidth     =   9000
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
   ScaleHeight     =   7455
   ScaleWidth      =   9000
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
      Left            =   120
      TabIndex        =   6
      Top             =   6840
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   1920
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
      Height          =   5295
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9340
      _Version        =   393216
      BackColorFixed  =   16777215
      BackColorSel    =   16777215
      BackColorBkg    =   16777215
      GridColor       =   0
      GridLines       =   3
      GridLinesFixed  =   3
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
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total:"
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
      Left            =   6480
      TabIndex        =   8
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
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
      Left            =   7320
      TabIndex        =   7
      Top             =   6240
      Width           =   1455
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
Const heightDifference As Integer = 2730
Const widthDifference As Integer = 705

Dim transRecord As Collection
Dim dateNow As Date

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
    
    gridStudents.ColSel = 0
    gridStudents.RowSel = 0
End Sub

Private Sub cmdPrint_Click()
    
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
        cmdExport.Visible = False
        cmdPrint.Visible = False
        cmdClose.Visible = False
        PrintForm
        cmdPrint.Visible = True
        cmdClose.Visible = True
        cmdExport.Visible = True
     'Printer.EndDoc
   Next
errHandler:
   ' User pressed Cancel button.
   Exit Sub
End Sub

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

Private Sub RefreshTableView()
    gridStudents.Cols = 5
    gridStudents.rows = 20
    gridStudents.TextMatrix(0, 1) = "First Name"
    gridStudents.TextMatrix(0, 2) = "Last Name"
    gridStudents.TextMatrix(0, 3) = "Grade"
    gridStudents.TextMatrix(0, 4) = "Payment"
    
    Dim totalWidth As Integer
    totalWidth = 0
        
    Dim i As Integer
    For i = 0 To 4
        gridStudents.ColWidth(i) = TextWidth(gridStudents.TextMatrix(0, i))
    Next
    Dim total As Long
    total = 0
    For i = 1 To transRecord.Count
        Dim studentInfo As Dictionary
        Set studentInfo = transRecord(i)
        gridStudents.TextMatrix(i, 0) = Format(i, String(4, "0"))
        gridStudents.TextMatrix(i, 1) = studentInfo("first_name")
        gridStudents.TextMatrix(i, 2) = studentInfo("last_name")
        gridStudents.TextMatrix(i, 3) = grade(studentInfo("current_grade"))
        gridStudents.TextMatrix(i, 4) = studentInfo("payment")
        total = total + CLng(studentInfo("payment"))
        Dim j As Integer
        
        For j = 0 To 4
            If TextWidth(gridStudents.TextMatrix(i, j)) > gridStudents.ColWidth(j) Then
                gridStudents.ColWidth(j) = TextWidth(gridStudents.TextMatrix(i, j))
            End If
        Next
    Next
    For i = 0 To 4
        totalWidth = totalWidth + gridStudents.ColWidth(i)
    Next
    
    lblTotal.Caption = Format(total, "P##,##0.00")
    
    Me.Width = totalWidth + 750
    Me.Height = Me.Height + 1500
End Sub

Private Sub Form_Resize()
    gridStudents.Width = Me.Width - widthDifference
    Label1.Width = Me.Width - widthDifference
    cmdClose.Left = Me.Width - 2955
    cmdPrint.Left = Me.Width - 1635
    lblTotal.Left = Me.Width - 1920
    Label3.Left = Me.Width - 2760
    gridStudents.Height = Me.Height - heightDifference
    cmdExport.Top = Me.Height - 1185
    cmdClose.Top = Me.Height - 1185
    cmdPrint.Top = Me.Height - 1185
    lblTotal.Top = Me.Height - 1785
    Label3.Top = Me.Height - 1785
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
        Set transRecord = p.Item("message")
        
        RefreshTableView
    Else
        MsgBox p.Item("message"), vbExclamation
        cmdPrint.enabled = False
    End If
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbExclamation, "Connection Error"
    sckMain.Close
End Sub

Private Sub sckMain_Close()
    blnConnected = False
    sckMain.Close
End Sub
