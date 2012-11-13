VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmNewPriceAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price Adjustments"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   8730
   Begin VB.TextBox txtPath 
      Height          =   360
      Left            =   6960
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnGetExcel 
      Height          =   375
      Left            =   1920
      TabIndex        =   23
      Top             =   1200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Get Data from Excel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnSetExcel 
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   1200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Sent Data to Excel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtPrice 
      Height          =   360
      Left            =   5400
      TabIndex        =   19
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtFATFrom 
      Height          =   360
      Left            =   2040
      TabIndex        =   16
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtFATTo 
      Height          =   360
      Left            =   3360
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtSNFFrom 
      Height          =   360
      Left            =   2040
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtSNFTo 
      Height          =   360
      Left            =   3360
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtPCID 
      Height          =   360
      Left            =   6960
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPSID 
      Height          =   360
      Left            =   8160
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtCellText 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   7560
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSDataListLib.DataCombo cmbPaymentScheme 
      Height          =   360
      Left            =   1920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnFill 
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Fill"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo cmbPriceCycle 
      Height          =   360
      Left            =   1920
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   57147395
      CurrentDate     =   39877
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   57147395
      CurrentDate     =   39877
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   5400
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label9 
      Caption         =   "Price"
      Height          =   255
      Left            =   4920
      TabIndex        =   20
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "To"
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "From"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "FAT"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "SNF"
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Price Cycle"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Payment Scheme"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmNewPriceAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTemPrice As New ADODB.Recordset
    Dim rsPrice  As New ADODB.Recordset
    Dim temSql As String
    Dim temRow As Long
    Dim temCol As Long
    Dim temText As String
    Dim temCellText As String
    Dim temBoxText As String
    Dim FSys As New FileSystemObject
    Dim Topic As String
    Dim SubTopic As String
    
    
Private Sub btnAdd_Click()
    If IsNumeric(txtSNFFrom.Text) = False Then
        Exit Sub
    End If
    If IsNumeric(txtSNFTo.Text) = False Then
        Exit Sub
    End If
    If IsNumeric(txtFATFrom.Text) = False Then
        Exit Sub
    End If
    If Val(txtFATFrom.Text) > Val(txtFATTo.Text) Then
        Exit Sub
    End If
    If Val(txtSNFFrom.Text) > Val(txtSNFTo.Text) Then
        Exit Sub
    End If
    Dim temSNF As Double
    Dim temFAT As Double
    For temSNF = Val(txtSNFFrom.Text) To Val(txtSNFTo.Text) Step 0.1
        For temFAT = Val(txtFATFrom.Text) To Val(txtFATTo.Text) Step 0.1
            With rsTemPrice
                If .State = 1 Then .Close
                temSql = "Select * from tblPrice where PriceCycleID = " & Val(cmbPriceCycle.BoundText) & " AND PaymentSchemeID = " & Val(cmbPaymentScheme.BoundText) & " AND SNF = " & Format(temSNF, "0.0") & " AND FAT = " & Format(temFAT, "0.0")
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !Price = Val(txtPrice.Text)
                Else
                    .AddNew
                    !PriceCycleID = Val(cmbPriceCycle.BoundText)
                    !PaymentSchemeID = Val(cmbPaymentScheme.BoundText)
                    !SNF = Format(temSNF, "0.0")
                    !FAT = Format(temFAT, "0.0")
                    !FromDate = dtpFrom.Value
                    !ToDate = dtpTo.Value
                    !Price = Val(txtPrice.Text)
                End If
                .Update
            End With
        Next
    Next
Call FillGrid
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub FormatGrid()

End Sub

Private Sub btnFill_Click()
        
    If IsNumeric(cmbPaymentScheme.BoundText) = False Then
        MsgBox "Please select a Payment Scheme"
        cmbPaymentScheme.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPriceCycle.BoundText) = False Then
        MsgBox "Please select a Price Cycle"
        cmbPriceCycle.SetFocus
        Exit Sub
    End If
    
    txtPCID.Text = cmbPriceCycle.BoundText
    txtPSID.Text = cmbPaymentScheme.BoundText
    
    Screen.MousePointer = vbHourglass
    DoEvents
    Call FillGrid
    gridPrice.col = 0
    gridPrice.row = 0
    gridPrice.col = 1
    gridPrice.row = 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub btnGetExcel_Click()
    CommonDialog1.FileName = Database
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNNoChangeDir
    CommonDialog1.DefaultExt = "xls"
    CommonDialog1.Filter = "Miscrosoft Excel|*.xls"
    CommonDialog1.ShowOpen
    If CommonDialog1.CancelError = False Then
        txtDatabase.Text = CommonDialog1.FileName
        MsgBox "This change will be effective from the next time you log on"
    Else
        MsgBox "You have not selected valid database. The program may not function", vbCritical, "No database"
        bttnSelectDatabase.SetFocus
    End If

End Sub

Private Sub btnSetExcel_Click()
    CommonDialog1.InitDir = txtPath.Text
    CommonDialog1.DefaultExt = "xls"
    CommonDialog1.Filter = "Microsoft Excel|*.xls"
    CommonDialog1.FileName = "Milk Prices - " & cmbPaymentScheme.Text & " - " & cmbPriceCycle.Text
    CommonDialog1.ShowSave
    If CommonDialog1.CancelError = True Then Exit Sub
    If CommonDialog1.FileName = "" Then
        MsgBox "Please select a name"
        Exit Sub
    End If
    
    Exit Sub
    
    Dim AppExcel As Excel.Application
    Dim myworkbook As Excel.Workbook
    Dim myWorkSheet1 As Excel.Worksheet
    Dim myWorkSheet2 As Excel.Worksheet
    Dim temRow As Integer
    Dim temCol As Integer
    Dim rsTem As New ADODB.Recordset
    
    Dim row As Integer
    Dim col As Integer
    
    Dim SNF As Double
    Dim FAT As Double
    
    Topic = "Milk Prices"
    SubTopic = cmbPaymentScheme.Text & " - " & cmbPriceCycle.Text

    Screen.MousePointer = vbHourglass

    Set AppExcel = CreateObject("Excel.Application")

    Set myworkbook = AppExcel.Workbooks.Add
    Set myWorkSheet1 = AppExcel.Worksheets(1)
    myWorkSheet1.Name = Topic

    myWorkSheet1.Cells(1, 1) = Topic
    myWorkSheet1.Cells(2, 1) = SubTopic


    For col = 1 To 30
    
        For row = 1 To 30
        
        Next row
    
    Next col


'    myWorkSheet1.Cells(5, 1) = "Depot"
'    myWorkSheet1.Cells(5, 2) = "Rep Name"
'    myWorkSheet1.Cells(5, 3) = "Stock Value"
'    myWorkSheet1.Cells(5, 4) = "Cash Collection"
'    myWorkSheet1.Cells(5, 5) = "Cheque Collection"
'    myWorkSheet1.Cells(5, 6) = "Returns"
'    myWorkSheet1.Cells(5, 5 + 2) = "Damages"
'    myWorkSheet1.Cells(5, 6 + 2) = "Summery Short"
'    myWorkSheet1.Cells(5, 7 + 2) = "Summery Excess"
'    myWorkSheet1.Cells(5, 8 + 2) = "Return %"
'    myWorkSheet1.Cells(5, 9 + 2) = "Net Sale"
'    myWorkSheet1.Cells(5, 10 + 2) = "Initial Amount"
'    myWorkSheet1.Cells(5, 11 + 2) = "Incentive 1"
'    myWorkSheet1.Cells(5, 13 + 1) = "(2)50,000.00"
'    myWorkSheet1.Cells(5, 14 + 1) = "Incentive 2"
'    myWorkSheet1.Cells(5, 15 + 1) = "(3)50,000.00"
'    myWorkSheet1.Cells(5, 16 + 1) = "Incentive 3"
'    myWorkSheet1.Cells(5, 17 + 1) = "(4)50,000.00"
'    myWorkSheet1.Cells(5, 18 + 1) = "Incentive 4"
'    myWorkSheet1.Cells(5, 19 + 1) = "(5)50,000.00"
'    myWorkSheet1.Cells(5, 20 + 1) = "Incentive 5"
'    myWorkSheet1.Cells(5, 21 + 1) = "(6)50,000.00"
'    myWorkSheet1.Cells(5, 22 + 1) = "Incentive 6"
'    myWorkSheet1.Cells(5, 23 + 1) = "(7)50,000.00"
'    myWorkSheet1.Cells(5, 24 + 1) = "Incentive 7"
'    myWorkSheet1.Cells(5, 25 + 1) = "(8)50,000.00"
'    myWorkSheet1.Cells(5, 26 + 1) = "Incentive 8"
'    myWorkSheet1.Cells(5, 27 + 1) = "(9)50,000.00"
'    myWorkSheet1.Cells(5, 28 + 1) = "Incentive 9"
'    myWorkSheet1.Cells(5, 29 + 1) = "Remaining Amount"
'    myWorkSheet1.Cells(5, 30 + 1) = "Incentive 10"
'    myWorkSheet1.Cells(5, 31 + 1) = "Total Cash Incentive"
'    myWorkSheet1.Cells(5, 32 + 1) = "Cheque Incentive"
'    myWorkSheet1.Cells(5, 33 + 1) = "Total Incentive"
'    myWorkSheet1.Cells(5, 34 + 1) = "Self Handling Allowance"
'    myWorkSheet1.Cells(5, 35 + 1) = "Total Incentives And Allowances"
'
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "SELECT tblRep.RepID, tblRep.Rep, tblAgent.AgentID, tblAgent.Agent FROM tblAgent RIGHT JOIN tblRep ON tblAgent.AgentID = tblRep.AgentID WHERE (((tblRep.TW)=True)) ORDER BY tblAgent.Agent, tblRep.Rep"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        ReDim MyRep(.RecordCount - 1)
'        i = 0
'        While .EOF = False
'            MyRep(i).Agent = !Agent
'            MyRep(i).AgentID = !AgentID
'            MyRep(i).ID = !RepID
'            MyRep(i).Rep = !Rep
'            i = i + 1
'            .MoveNext
'        Wend
'        .Close
'    End With
'
'
'    Dim temAgent As String
'
'    For TemRow = 0 To UBound(MyRep)
'        RepsTWData MyRep(TemRow), DateSerial(Year(dtpMonth.Value), Month(dtpMonth.Value), 1), DateSerial(Year(dtpMonth.Value), Month(dtpMonth.Value), LastDateOfMonth)
'        If temAgent <> MyRep(TemRow).Agent Then
'            myWorkSheet1.Cells(TemRow + 6, 1) = MyRep(TemRow).Agent
'            temAgent = MyRep(TemRow).Agent
'        End If
'        myWorkSheet1.Cells(TemRow + 6, 2) = MyRep(TemRow).Rep
'        myWorkSheet1.Cells(TemRow + 6, 3) = MyRep(TemRow).StockValue
'        myWorkSheet1.Cells(TemRow + 6, 4) = MyRep(TemRow).CashCollection '  "Cash Collection"
'        myWorkSheet1.Cells(TemRow + 6, 5) = MyRep(TemRow).CashCollection ' "Cheque Collection"
'        myWorkSheet1.Cells(TemRow + 6, 6) = MyRep(TemRow).Returns '  "Returns"
'        myWorkSheet1.Cells(TemRow + 6, 5 + 2) = MyRep(TemRow).Damages ' "Damages"
'        myWorkSheet1.Cells(TemRow + 6, 6 + 2) = MyRep(TemRow).SummeryShort ' "Summery Short"
'        myWorkSheet1.Cells(TemRow + 6, 7 + 2) = MyRep(TemRow).SummeryExcess  '  "Summery Excess"
'        myWorkSheet1.Cells(TemRow + 6, 8 + 2) = MyRep(TemRow).ReturnPercent ' "Return %"
'        myWorkSheet1.Cells(TemRow + 6, 9 + 2) = MyRep(TemRow).NetSale  '   "Net Sale"
'        myWorkSheet1.Cells(TemRow + 6, 10 + 2) = MyRep(TemRow).InitialAmount ' "Initial Amount"
'        myWorkSheet1.Cells(TemRow + 6, 11 + 2) = MyRep(TemRow).Incentive1  ' "Incentive 1"
'        myWorkSheet1.Cells(TemRow + 6, 13 + 1) = MyRep(TemRow).I2ndFT  ' "(2)50,000.00"
'        myWorkSheet1.Cells(TemRow + 6, 14 + 1) = MyRep(TemRow).Incentive2 ' "Incentive 2"
'        myWorkSheet1.Cells(TemRow + 6, 15 + 1) = MyRep(TemRow).I3rdFT ' "(3)50,000.00"
'        myWorkSheet1.Cells(TemRow + 6, 16 + 1) = MyRep(TemRow).Incentive3 ' "Incentive 3"
'        myWorkSheet1.Cells(TemRow + 6, 17 + 1) = MyRep(TemRow).I4thFT ' "(4)50,000.00"
'        myWorkSheet1.Cells(TemRow + 6, 18 + 1) = MyRep(TemRow).Incentive4 ' "Incentive 4"
'        myWorkSheet1.Cells(TemRow + 6, 19 + 1) = MyRep(TemRow).I5thFT ' "(5)50,000.00"
'        myWorkSheet1.Cells(TemRow + 6, 20 + 1) = MyRep(TemRow).Incentive5 ' "Incentive 5"
'        myWorkSheet1.Cells(TemRow + 6, 21 + 1) = MyRep(TemRow).I6thFT ' "(6)50,000.00"
'        myWorkSheet1.Cells(TemRow + 6, 22 + 1) = MyRep(TemRow).Incentive6 ' "Incentive 6"
'        myWorkSheet1.Cells(TemRow + 6, 23 + 1) = MyRep(TemRow).I7thFT ' "(7)50,000.00"
'        myWorkSheet1.Cells(TemRow + 6, 24 + 1) = MyRep(TemRow).Incentive7 ' "Incentive 7"
'        myWorkSheet1.Cells(TemRow + 6, 25 + 1) = MyRep(TemRow).I8thFT ' "(8)50,000.00"
'        myWorkSheet1.Cells(TemRow + 6, 26 + 1) = MyRep(TemRow).Incentive8 ' "Incentive 8"
'        myWorkSheet1.Cells(TemRow + 6, 27 + 1) = MyRep(TemRow).I9thFT ' "(9)50,000.00"
'        myWorkSheet1.Cells(TemRow + 6, 28 + 1) = MyRep(TemRow).Incentive9 ' "Incentive 9"
'        myWorkSheet1.Cells(TemRow + 6, 29 + 1) = MyRep(TemRow).OtherAmount ' "Remaining Amount"
'        myWorkSheet1.Cells(TemRow + 6, 30 + 1) = MyRep(TemRow).Incentive10 ' "Incentive 10"
'        myWorkSheet1.Cells(TemRow + 6, 31 + 1) = MyRep(TemRow).TotalCashIncentive ' "Total Cash Incentive"
'        myWorkSheet1.Cells(TemRow + 6, 32 + 1) = MyRep(TemRow).ChequeIncentive ' "Cheque Incentive"
'        myWorkSheet1.Cells(TemRow + 6, 33 + 1) = MyRep(TemRow).TotalIncentive  ' "Total Incentive"
'        myWorkSheet1.Cells(TemRow + 6, 34 + 1) = MyRep(TemRow).SelfHandlingAllowance ' "Self Handling Allowance"
'        myWorkSheet1.Cells(TemRow + 6, 35 + 1) = MyRep(TemRow).TotalIncentiveAndAllowance ' "Total Incentives And Allowances"
'    Next TemRow
'
'
'
'
'
'
'
'
'
'    Set myWorkSheet2 = AppExcel.Worksheets(2)
'    myWorkSheet2.Name = "Van Reps"
'
'    myWorkSheet2.Cells(3, 1) = Topic & " For Van Sales Reps"
'    myWorkSheet2.Cells(4, 2) = SubTopic
'
'    myWorkSheet2.Cells(5, 1) = "Depot"
'    myWorkSheet2.Cells(5, 2) = "Rep Name"
'    myWorkSheet2.Cells(5, 3) = "Stock Value"
'    myWorkSheet2.Cells(5, 4) = "Cash Collection"
'    myWorkSheet2.Cells(5, 5) = "Cheque Collection"
'    myWorkSheet2.Cells(5, 6) = "Returns"
'    myWorkSheet2.Cells(5, 5 + 2) = "Damages"
'    myWorkSheet2.Cells(5, 6 + 2) = "Summery Short"
'    myWorkSheet2.Cells(5, 7 + 2) = "Summery Excess"
'    myWorkSheet2.Cells(5, 8 + 2) = "Return %"
'    myWorkSheet2.Cells(5, 9 + 2) = "Net Sale"
'    myWorkSheet2.Cells(5, 10 + 2) = "Initial Amount"
'    myWorkSheet2.Cells(5, 11 + 2) = "Incentive 1"
'    myWorkSheet2.Cells(5, 13 + 1) = "(2)50,000.00"
'    myWorkSheet2.Cells(5, 14 + 1) = "Incentive 2"
'    myWorkSheet2.Cells(5, 15 + 1) = "(3)50,000.00"
'    myWorkSheet2.Cells(5, 16 + 1) = "Incentive 3"
'    myWorkSheet2.Cells(5, 17 + 1) = "(4)50,000.00"
'    myWorkSheet2.Cells(5, 18 + 1) = "Incentive 4"
'    myWorkSheet2.Cells(5, 19 + 1) = "(5)50,000.00"
'    myWorkSheet2.Cells(5, 20 + 1) = "Incentive 5"
'    myWorkSheet2.Cells(5, 21 + 1) = "(6)50,000.00"
'    myWorkSheet2.Cells(5, 22 + 1) = "Incentive 6"
'    myWorkSheet2.Cells(5, 23 + 1) = "(7)50,000.00"
'    myWorkSheet2.Cells(5, 24 + 1) = "Incentive 7"
'    myWorkSheet2.Cells(5, 25 + 1) = "(8)50,000.00"
'    myWorkSheet2.Cells(5, 26 + 1) = "Incentive 8"
'    myWorkSheet2.Cells(5, 27 + 1) = "(9)50,000.00"
'    myWorkSheet2.Cells(5, 28 + 1) = "Incentive 9"
'    myWorkSheet2.Cells(5, 29 + 1) = "(10)50,000.00"
'    myWorkSheet2.Cells(5, 30 + 1) = "Incentive 10"
'    myWorkSheet2.Cells(5, 31 + 1) = "(11)50,000.00"
'    myWorkSheet2.Cells(5, 32 + 1) = "Incentive 11"
'    myWorkSheet2.Cells(5, 33 + 1) = "(12)50,000.00"
'    myWorkSheet2.Cells(5, 34 + 1) = "Incentive 12"
'    myWorkSheet2.Cells(5, 35 + 1) = "Remaining Amount"
'    myWorkSheet2.Cells(5, 36 + 1) = "Incentive 13"
'    myWorkSheet2.Cells(5, 37 + 1) = "Total Cash Incentive"
'    myWorkSheet2.Cells(5, 38 + 1) = "Cheque Incentive"
'    myWorkSheet2.Cells(5, 39 + 1) = "Total Incentive"
'
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "SELECT tblRep.RepID, tblRep.Rep, tblAgent.AgentID, tblAgent.Agent FROM tblAgent RIGHT JOIN tblRep ON tblAgent.AgentID = tblRep.AgentID WHERE (((tblRep.TW)=False)) ORDER BY tblAgent.Agent, tblRep.Rep"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        ReDim MyRep(.RecordCount - 1)
'        i = 0
'        While .EOF = False
'            MyRep(i).Agent = !Agent
'            MyRep(i).AgentID = !AgentID
'            MyRep(i).ID = !RepID
'            MyRep(i).Rep = !Rep
'            i = i + 1
'            .MoveNext
'        Wend
'        .Close
'    End With
'
'
'    For TemRow = 0 To UBound(MyRep)
'        RepsVanData MyRep(TemRow), DateSerial(Year(dtpMonth.Value), Month(dtpMonth.Value), 1), DateSerial(Year(dtpMonth.Value), Month(dtpMonth.Value), LastDateOfMonth)
'        If temAgent <> MyRep(TemRow).Agent Then
'            myWorkSheet2.Cells(TemRow + 6, 1) = MyRep(TemRow).Agent
'            temAgent = MyRep(TemRow).Agent
'        End If
'        myWorkSheet2.Cells(TemRow + 6, 2) = MyRep(TemRow).Rep
'        myWorkSheet2.Cells(TemRow + 6, 3) = MyRep(TemRow).StockValue
'        myWorkSheet2.Cells(TemRow + 6, 4) = MyRep(TemRow).CashCollection '  "Cash Collection"
'        myWorkSheet2.Cells(TemRow + 6, 5) = MyRep(TemRow).CashCollection ' "Cheque Collection"
'        myWorkSheet2.Cells(TemRow + 6, 6) = MyRep(TemRow).Returns '  "Returns"
'        myWorkSheet2.Cells(TemRow + 6, 5 + 2) = MyRep(TemRow).Damages ' "Damages"
'        myWorkSheet2.Cells(TemRow + 6, 6 + 2) = MyRep(TemRow).SummeryShort ' "Summery Short"
'        myWorkSheet2.Cells(TemRow + 6, 7 + 2) = MyRep(TemRow).SummeryExcess  '  "Summery Excess"
'        myWorkSheet2.Cells(TemRow + 6, 8 + 2) = MyRep(TemRow).ReturnPercent ' "Return %"
'        myWorkSheet2.Cells(TemRow + 6, 9 + 2) = MyRep(TemRow).NetSale  '   "Net Sale"
'        myWorkSheet2.Cells(TemRow + 6, 10 + 2) = MyRep(TemRow).InitialAmount ' "Initial Amount"
'        myWorkSheet2.Cells(TemRow + 6, 11 + 2) = MyRep(TemRow).Incentive1  ' "Incentive 1"
'        myWorkSheet2.Cells(TemRow + 6, 13 + 1) = MyRep(TemRow).I2ndFT  ' "(2)50,000.00"
'        myWorkSheet2.Cells(TemRow + 6, 14 + 1) = MyRep(TemRow).Incentive2 ' "Incentive 2"
'        myWorkSheet2.Cells(TemRow + 6, 15 + 1) = MyRep(TemRow).I3rdFT ' "(3)50,000.00"
'        myWorkSheet2.Cells(TemRow + 6, 16 + 1) = MyRep(TemRow).Incentive3 ' "Incentive 3"
'        myWorkSheet2.Cells(TemRow + 6, 17 + 1) = MyRep(TemRow).I4thFT ' "(4)50,000.00"
'        myWorkSheet2.Cells(TemRow + 6, 18 + 1) = MyRep(TemRow).Incentive4 ' "Incentive 4"
'        myWorkSheet2.Cells(TemRow + 6, 19 + 1) = MyRep(TemRow).I5thFT ' "(5)50,000.00"
'        myWorkSheet2.Cells(TemRow + 6, 20 + 1) = MyRep(TemRow).Incentive5 ' "Incentive 5"
'        myWorkSheet2.Cells(TemRow + 6, 21 + 1) = MyRep(TemRow).I6thFT ' "(6)50,000.00"
'        myWorkSheet2.Cells(TemRow + 6, 22 + 1) = MyRep(TemRow).Incentive6 ' "Incentive 6"
'        myWorkSheet2.Cells(TemRow + 6, 23 + 1) = MyRep(TemRow).I7thFT ' "(7)50,000.00"
'        myWorkSheet2.Cells(TemRow + 6, 24 + 1) = MyRep(TemRow).Incentive7 ' "Incentive 7"
'        myWorkSheet2.Cells(TemRow + 6, 25 + 1) = MyRep(TemRow).I8thFT ' "(8)50,000.00"
'        myWorkSheet2.Cells(TemRow + 6, 26 + 1) = MyRep(TemRow).Incentive8 ' "Incentive 8"
'        myWorkSheet2.Cells(TemRow + 6, 27 + 1) = MyRep(TemRow).I9thFT ' "(9)50,000.00"
'        myWorkSheet2.Cells(TemRow + 6, 28 + 1) = MyRep(TemRow).Incentive9 ' "Incentive 9"
'        myWorkSheet2.Cells(TemRow + 6, 29 + 1) = MyRep(TemRow).I10thFT ' "(9)50,000.00"
'        myWorkSheet2.Cells(TemRow + 6, 30 + 1) = MyRep(TemRow).Incentive10 ' "Incentive 9"
'        myWorkSheet2.Cells(TemRow + 6, 31 + 1) = MyRep(TemRow).I11thFT ' "(9)50,000.00"
'        myWorkSheet2.Cells(TemRow + 6, 32 + 1) = MyRep(TemRow).Incentive11 ' "Incentive 9"
'        myWorkSheet2.Cells(TemRow + 6, 33 + 1) = MyRep(TemRow).I12thFT ' "(9)50,000.00"
'        myWorkSheet2.Cells(TemRow + 6, 34 + 1) = MyRep(TemRow).Incentive12 ' "Incentive 9"
'        myWorkSheet2.Cells(TemRow + 6, 35 + 1) = MyRep(TemRow).OtherAmount ' "Remaining Amount"
'        myWorkSheet2.Cells(TemRow + 6, 36 + 1) = MyRep(TemRow).Incentive13 ' "Incentive 10"
'        myWorkSheet2.Cells(TemRow + 6, 37 + 1) = MyRep(TemRow).TotalCashIncentive ' "Total Cash Incentive"
'        myWorkSheet2.Cells(TemRow + 6, 38 + 1) = MyRep(TemRow).ChequeIncentive ' "Cheque Incentive"
'        myWorkSheet2.Cells(TemRow + 6, 39 + 1) = MyRep(TemRow).TotalIncentive  ' "Total Incentive"
'    Next TemRow
'
'    myWorkSheet1.Activate
'
'
'    myworkbook.SaveAs (App.Path & "\" & Topic & " - " & SubTopic & ".xls")
'    myworkbook.Save
'    myworkbook.Close
'
'    ShellExecute 0&, "open", App.Path & "\" & Topic & " - " & SubTopic & ".xls", "", "", vbMaximizedFocus
'
'
'    Screen.MousePointer = vbDefault
'

End Sub

Private Sub cmbPaymentScheme_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnFill.SetFocus
    End If
End Sub


Private Sub cmbPriceCycle_Change()
    Dim rsPS As New ADODB.Recordset
    With rsPS
        If .State = 1 Then .Close
        temSql = "Select * from tblPriceCycle where PriceCycleID = " & Val(cmbPriceCycle.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            dtpFrom.Value = !FromDate
            dtpTo.Value = !ToDate
        End If
        .Close
    End With
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FormatGrid
    txtCellText.Visible = False
    txtPath.Text = GetSetting(App.EXEName, Me.Name, txtPath.Name, App.Path)
End Sub

Private Sub FillCombos()
    Dim PaymentMethod As New clsFillCombos
    PaymentMethod.FillAnyCombo cmbPaymentScheme, "PaymentScheme", True
    Dim PC As New clsFillCombos
    PC.FillAnyCombo cmbPriceCycle, "PriceCycle", True
End Sub

Private Sub FillGrid()
    If IsNumeric(cmbPaymentScheme.BoundText) = False Then Exit Sub
    Dim rsPrice As New ADODB.Recordset
    gridPrice.Visible = False
    Dim row As Integer
    Dim col As Integer
    Dim temFAT As Double
    Dim temSNF As Double
    With gridPrice
        If rsPrice.State = 1 Then rsPrice.Close
        For row = 1 To .Rows - 1
            For col = 1 To .Cols - 1
                temSNF = Val(.TextMatrix(0, col))
                temFAT = Val(.TextMatrix(row, 0))
                temSql = "Select * from tblPrice where SNF = " & temSNF & " And FAT = " & temFAT & " And PaymentSchemeID = " & Val(txtPSID.Text) & " AND PriceCycleID = " & Val(txtPCID.Text)
                rsPrice.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If rsPrice.RecordCount > 0 Then
                    If IsNull(rsPrice!Price) = False Then
                        .TextMatrix(row, col) = Format(rsPrice!Price, "0.00")
                    Else
                        .TextMatrix(row, col) = Format(0, "0.00")
                    End If
                End If
                rsPrice.Close
            Next

        Next
    End With
    gridPrice.Visible = True
End Sub


Private Sub gridPrice_EnterCell()
    temRow = gridPrice.row
    temCol = gridPrice.col
    temCellText = gridPrice.TextMatrix(temRow, temCol)
    txtCellText.Top = gridPrice.Top + gridPrice.CellTop
    txtCellText.Left = gridPrice.Left + gridPrice.CellLeft
    txtCellText.Height = gridPrice.CellHeight
    txtCellText.Width = gridPrice.CellWidth
    txtCellText.Text = temCellText
    txtCellText.Visible = True
    txtCellText.SetFocus
    SendKeys "{Home}+{end}"
End Sub

Private Sub gridPrice_LeaveCell()
    temBoxText = txtCellText.Text
    Dim temFAT As Double
    Dim temSNF As Double
    Dim temPrice As Double
    temFAT = Val(gridPrice.TextMatrix(temRow, 0))
    temSNF = Val(gridPrice.TextMatrix(0, temCol))
    temPrice = Val(temBoxText)
    If temBoxText <> temCellText Then
        gridPrice.TextMatrix(temRow, temCol) = temBoxText
        With rsPrice
            If .State = 1 Then .Close
            temSql = "Select * from tblPrice where SNF = " & temSNF & " And FAT = " & temFAT & " And PaymentSchemeID = " & Val(txtPSID.Text) & " AND PriceCycleID = " & Val(txtPCID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Price = temPrice
                !FromDate = dtpFrom.Value
                !ToDate = dtpTo.Value
                .Update
            Else
                .AddNew
                !Price = temPrice
                !SNF = temSNF
                !FAT = temFAT
                !PriceCycleID = Val(txtPCID.Text)
                !FromDate = dtpFrom.Value
                !ToDate = dtpTo.Value
                .Update
            End If
            .Close
        End With
    End If
End Sub

Private Sub gridPrice_Scroll()
    txtCellText.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.Path, Me.Name, txtPath.Name, txtPath.Text
End Sub

Private Sub txtCellText_KeyDown(KeyCode As Integer, Shift As Integer)
    With gridPrice
        If KeyCode = vbKeyReturn Then
            If temCol < .Cols - 1 Then
                .col = temCol + 1
            Else
                .col = 1
                .row = temRow + 1
            End If
        ElseIf KeyCode = vbKeyEscape Then
            txtCellText.Text = temText
        ElseIf KeyCode = vbKeyTab Then
            If temCol < .Cols - 1 Then
                .col = temCol + 1
            Else
                .col = 1
                .row = temRow + 1
            End If
        ElseIf KeyCode = vbKeyUp Then
            If temRow > 1 Then
                .row = temRow - 1
            End If
        ElseIf KeyCode = vbKeyDown Then
            If temRow < .Rows - 1 Then
                .row = temRow + 1
            End If
        ElseIf KeyCode = vbKeyLeft Then
            If temCol > 1 Then
                .col = temCol - 1
            End If
        ElseIf KeyCode = vbKeyRight Then
            If temCol < .Cols - 1 Then
                .col = temCol + 1
            End If
        End If
    End With
End Sub
