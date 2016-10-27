VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReceiptPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6015
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   360
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   3240
      TabIndex        =   6
      Top             =   6720
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
      Left            =   4560
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid gridAmount 
      Height          =   3735
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6588
      _Version        =   393216
      BackColor       =   16777215
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
   Begin VB.Line Line5 
      X1              =   1320
      X2              =   5400
      Y1              =   1000
      Y2              =   1000
   End
   Begin VB.Line Line4 
      X1              =   1320
      X2              =   5400
      Y1              =   500
      Y2              =   500
   End
   Begin VB.Line Line3 
      X1              =   4080
      X2              =   5760
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label lblTotalAmount 
      BackColor       =   &H00FFFFFF&
      Caption         =   "P1,000,000.00"
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
      Left            =   4080
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total"
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
      Left            =   2640
      TabIndex        =   9
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   5760
      Y1              =   6300
      Y2              =   6300
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   5760
      Y1              =   6025
      Y2              =   6025
   End
   Begin VB.Label lblAmountWords 
      BackColor       =   &H00FFFFFF&
      Caption         =   "One Million, Two Hundred and Fifty Four Thousand, Three hundred forty five"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   5760
      Width           =   5415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Amount in Words:"
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
      TabIndex        =   7
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Brgy. St. Pototan, Iloilo"
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
      TabIndex        =   3
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label lblFullName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nicholas A. Cage"
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
      TabIndex        =   2
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name:"
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
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmReceiptPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim small_Numbers() As Variant
Dim tens_() As Variant
Dim scalar_() As Variant

Public fName As String
Public fAddress As String

Public pAmount As Long

Private Sub cmdCancel_Click()
    Unload Me
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
        cmdPrint.Visible = False
        cmdCancel.Visible = False
        PrintForm
        cmdPrint.Visible = True
        cmdCancel.Visible = True
     'Printer.EndDoc
   Next
errHandler:
   ' User pressed Cancel button.
   Exit Sub
End Sub

Private Sub Form_Load()
    small_Numbers = Array("Zero", "One", "Two", "Three", "Four", _
        "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", _
        "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", _
        "Seventeen", "Eighteen", "Nineteen")
    
    tens_ = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", _
        "Sixty", "Seventy", "Eighty", "Ninety")
        
    scalar_ = Array("", "Thousand", "Million", "Billion")

    
    displayData
End Sub

Private Sub displayData()
    gridAmount.Cols = 3
    gridAmount.rows = 4
    gridAmount.TextMatrix(0, 0) = "Nature of Collection"
    gridAmount.TextMatrix(0, 1) = "Account Code"
    gridAmount.TextMatrix(0, 2) = "Amount"
    
    gridAmount.TextMatrix(1, 0) = "Matriculation"
    gridAmount.TextMatrix(1, 1) = ""
    gridAmount.TextMatrix(1, 2) = pAmount

    Dim i As Integer
    Dim totalWidth As Integer
    totalWidth = 0
    
    For i = 0 To 2
        gridAmount.Row = 0
        gridAmount.Col = i
        gridAmount.CellFontBold = True
        gridAmount.ColWidth(i) = TextWidth(gridAmount.TextMatrix(0, i))
        totalWidth = totalWidth + gridAmount.ColWidth(i)
    Next
    Me.Width = totalWidth + 580
    gridAmount.Refresh
    lblFullName.Caption = fName
    lblAddress.Caption = fAddress
    lblTotalAmount.Caption = Format(pAmount, "P##,##0.00")
    lblAmountWords.Caption = numToWords(pAmount)
    
End Sub

Private Function numToWords(num As Long) As String
    If num = 0 Then
        numToWords = "Zero"
        Exit Function
    End If
    
    Dim digitGroups(4) As Integer
    Dim groupText(4) As String
    
    Dim positive As Long
    positive = Math.Abs(num)
    
    Dim i As Integer
    For i = 0 To 3
        digitGroups(i) = positive Mod 1000
        positive = positive \ 1000
    Next
    
    For i = 0 To 3
        groupText(i) = ThreeDigitGroupToWords(digitGroups(i))
    Next
    
    'Recombine the three-digit groups
    Dim combined As String
    combined = groupText(0)
    Dim appendAnd As Boolean
     
    'Determine whether an 'and' is needed
    appendAnd = (digitGroups(0) > 0) And (digitGroups(0) < 100)
     
    'Process the remaining groups in turn, smallest to largest
    For i = 1 To 3
        'Only add non-zero items
        If digitGroups(i) > 0 Then
            'Build the string to add as a prefix
            Dim prefix As String
            prefix = groupText(i) & " " & scalar_(i)
             
            If (Len(combined) > 0) Then
                prefix = prefix & IIf(appendAnd, " and ", ", ")
            End If
             
            'Opportunity to add 'and' is ended
            appendAnd = False
     
            'Add the three-digit group to the combined string
            combined = prefix & combined
        End If
    Next
    
    If (num < 0) Then
        combined = "Negative " & combined
    End If
    numToWords = combined
End Function

Private Function ThreeDigitGroupToWords(threeDigits As Integer) As String
    Dim groupText As String
    groupText = ""
 
    'Determine the hundreds and the remainder
    Dim hundreds As Integer
    hundreds = threeDigits \ 100
    Dim tensUnits As Integer
    tensUnits = threeDigits Mod 100
 
    'Hundreds rules
    If hundreds > 0 Then
        Debug.Print hundreds
        groupText = groupText & small_Numbers(hundreds) + " Hundred"
        If tensUnits > 0 Then
            groupText = groupText & " and "
        End If
    End If
    
    Dim tens As Integer
    tens = tensUnits \ 10
    Dim units As Integer
    units = tensUnits Mod 10
     
    'Tens rules
    If tens >= 2 Then
        groupText = groupText & tens_(tens)
        If units > 0 Then
            groupText = groupText & " " + small_Numbers(units)
        End If
    ElseIf tensUnits > 0 Then
        groupText = groupText & small_Numbers(tensUnits)
    End If
    ThreeDigitGroupToWords = groupText
End Function

Private Sub Form_Resize()
    'Width Adjustment
    Line4.X2 = Me.Width - 705
    Line5.X2 = Me.Width - 705
    
    gridAmount.Width = Me.Width - 570
    
    Label4.Left = Me.Width - 3465
    lblTotalAmount.Left = Me.Width - 2025
    Line3.X1 = Me.Width - 2025
    Line3.X2 = Line3.X1 + 1680
    
    lblAmountWords.Width = Me.Width - 690
    Line1.X2 = Me.Width - 345
    Line2.X2 = Me.Width - 345
    
    cmdPrint.Left = Me.Width - 1545
    cmdCancel.Left = Me.Width - 2865
End Sub
