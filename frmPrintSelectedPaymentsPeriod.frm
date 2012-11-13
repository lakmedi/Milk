VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPrintSelectedPaymentsPeriod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print All Payments for a Period"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10290
   ClipControls    =   0   'False
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
   ScaleHeight     =   9420
   ScaleWidth      =   10290
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   7440
      TabIndex        =   27
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Process"
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
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1320
      TabIndex        =   26
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   284295171
      CurrentDate     =   40254
   End
   Begin VB.OptionButton optAll 
      Caption         =   "&All"
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   1080
      Width           =   1575
   End
   Begin VB.OptionButton optCom 
      Caption         =   "&Commercial Bank"
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   1080
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton optSB 
      Caption         =   "&Sampath Bank"
      Height          =   255
      Left            =   720
      TabIndex        =   20
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox lstSelectedValue 
      Height          =   300
      Left            =   7560
      TabIndex        =   19
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstSelectedSuppliers 
      Height          =   300
      Left            =   7560
      TabIndex        =   18
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstSelectedSupplierIDs 
      Height          =   300
      Left            =   7560
      TabIndex        =   17
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstPaymentMethod 
      Height          =   300
      Left            =   7560
      TabIndex        =   16
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstAllSuppliers 
      Height          =   300
      Left            =   7560
      TabIndex        =   15
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstAllVaue 
      Height          =   300
      Left            =   7560
      TabIndex        =   14
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstAllSupplierIDs 
      Height          =   300
      Left            =   7560
      TabIndex        =   13
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   5280
      Width           =   2655
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4320
      Width           =   2655
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8880
      TabIndex        =   1
      Top             =   8760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
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
   Begin VB.ListBox lstBank 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7035
      Left            =   720
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   1440
      Width           =   4095
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   8760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Print"
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
   Begin VB.ListBox lstBankIDs 
      Height          =   1500
      Left            =   4080
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin btButtonEx.ButtonEx btnPrintBankC 
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   3360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Commercial Bank Print"
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
   Begin btButtonEx.ButtonEx btnPrintBankS 
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Top             =   2880
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Sampath Bank Print"
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
   Begin VB.ListBox lstBankValue 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7065
      IntegralHeight  =   0   'False
      ItemData        =   "frmPrintSelectedPaymentsPeriod.frx":0000
      Left            =   4680
      List            =   "frmPrintSelectedPaymentsPeriod.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin btButtonEx.ButtonEx btnExcelCom 
      Height          =   375
      Left            =   7440
      TabIndex        =   23
      Top             =   1920
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Commercial Bank Excel"
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
   Begin btButtonEx.ButtonEx btnExcelSam 
      Height          =   375
      Left            =   7440
      TabIndex        =   24
      Top             =   1440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Sampath Bank Excel"
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
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4680
      TabIndex        =   28
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   284295171
      CurrentDate     =   40254
   End
   Begin MSDataListLib.DataCombo cmbCC 
      Height          =   360
      Left            =   1320
      TabIndex        =   30
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label9 
      Caption         =   "Center"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   4080
      TabIndex        =   29
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Printer"
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Paper"
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label lblBankValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   9000
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Total Bank Value"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   9000
      Width           =   2175
   End
End
Attribute VB_Name = "frmPrintSelectedPaymentsPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    Dim rsCCPS As New ADODB.Recordset
    Dim Supplier() As String
    Dim SupplierID() As Long
    Dim SupplierSelected() As Boolean
    Dim SupplierPaymentsID() As Long
    Dim PaymentMethodID() As Byte
    Dim SupplierValue() As Double
    Dim FSys As New FileSystemObject
    Dim CSetPrinter As New cSetDfltPrinter
    
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean
    Dim SuppliedWord As String
    
    Dim SettingValues As Boolean
    Dim SettingNames As Boolean
    
    Dim CombankID As String
    Dim CombankBranchCode As String
    Dim ComBankAccountName As String
    Dim ComBankAccountNo As String
    
    Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long
    
Private Sub ComBankDetails()
     CombankID = "7056"
     CombankBranchCode = "104"
     ComBankAccountName = "Lucky Lanka Dairies (Pvt) Ltd"
     ComBankAccountNo = "001104013300"
End Sub

    
Private Sub btnBankAll_Click()
    Dim i As Integer
    For i = 0 To lstBankIDs.ListCount - 1
        lstBank.Selected(i) = True
        lstBankIDs.Selected(i) = True
        lstBankValue.Selected(i) = True
    Next i
    CalculateValues
End Sub

Private Sub btnBankNone_Click()
    Dim i As Integer
    For i = 0 To lstBankIDs.ListCount - 1
        lstBank.Selected(i) = False
        lstBankIDs.Selected(i) = False
        lstBankValue.Selected(i) = False
    Next i
    CalculateValues
End Sub


Private Sub btnClose_Click()
    Unload Me
End Sub

Private Function ComBankValue(Value As Double) As String
    ComBankValue = Format(Val(Format(Value, "0.00") * 100), "000000000000")
End Function

Private Sub btnExcelCom_Click()
    'On Error Resume Next
    
    Dim AppExcel As Excel.Application
    Dim myworkbook As Excel.Workbook
    Dim myworksheet As Excel.Worksheet
    Dim mychart As Excel.Chart
    Dim tempath As String
    Dim FSys As New Scripting.FileSystemObject
    Dim TemTopic As String
    Dim temSubTopic As String
    Dim RetVal As Integer
    Dim temText As String
    Dim temMoney As String
    Dim rsReport As New ADODB.Recordset
    Dim i As Integer
    Dim AccNo As String
   
    Dim AllFarmers() As Double
    
    Dim MyNumberWord As String
    Dim MyNumberWordLength As Long
    Dim RupeeWord As String
    Dim CentsWord As String
    
    Dim myNum As New clsNumbers
    
    Dim temFromDate As Date
    
    tempath = App.Path

    If FSys.FileExists(tempath & "\combank.xls") = False Then
        MsgBox "THe file named combank.xls is not found the the database folder. Please find the file and place it at " & App.Path
        Exit Sub
    End If


    i = ShellExecute(0, "open", tempath & "\combank.xls", 0, 0, 1)
    
    Set myworkbook = GetObject(tempath & "\combank.xls")
    Set myworksheet = myworkbook.WorkSheets(1)
    
    Dim TotalPayment As Double
    
                
    MyNumberWord = Format(Val(lblBankValue.Caption), "0.00")
    MyNumberWordLength = Len(MyNumberWord)
    RupeeWord = Left(MyNumberWord, MyNumberWordLength - 3)
    CentsWord = Right(MyNumberWord, 2)
    
    temMoney = RupeeWord
    
    i = 6
    
    myworksheet.Range("A6:Z5500").ClearContents

    With rsReport
        If .State = 1 Then .Close
        temSQL = "SELECT  tblSupplier.SupplierID, tblBank.BankCode, tblCity.BankCode, tblSupplier.AccountNo, tblSupplier.AccountHolder, sum(tblSupplierPayments.Value) as TotalValue , max(tblCollectingCenterPaymentSummery.FromDate) as MaxFromDate " & _
                    "FROM (((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID) LEFT JOIN tblCity ON tblSupplier.CityID = tblCity.CityId " & _
                    "Where  tblSupplier.CollectingCenterID = " & Val(cmbCC.BoundText) & " AND tblCollectingCenterPaymentSummery.FromDate Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And (tblSupplier.BankID = 2 or tblSupplier.BankID = 1 ) " & _
                    "GROUP BY tblBank.BankCode, tblCity.BankCode, tblSupplier.AccountNo, tblSupplier.SupplierID, tblSupplier.AccountHolder " & _
                    "ORDER BY tblSupplier.AccountHolder"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        .MoveLast
        ReDim AllFarmers(.RecordCount)
        .MoveFirst
        Dim FarmerCount As Integer
        While .EOF = False
            AllFarmers(FarmerCount) = !SupplierID
            FarmerCount = FarmerCount + 1
            .MoveNext
        Wend
       
        If .State = 1 Then .Close
        temSQL = "SELECT  tblSupplier.SupplierID, tblBank.BankCode, tblCity.BankCode, tblSupplier.AccountNo, tblSupplier.AccountHolder, sum(tblSupplierPayments.Value) as TotalValue , max(tblCollectingCenterPaymentSummery.FromDate) as MaxFromDate " & _
                    "FROM (((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID) LEFT JOIN tblCity ON tblSupplier.CityID = tblCity.CityId " & _
                    "Where tblSupplier.CollectingCenterID = " & Val(cmbCC.BoundText) & " AND tblCollectingCenterPaymentSummery.FromDate Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And tblSupplier.BankID = 2 " & _
                    "GROUP BY tblBank.BankCode, tblCity.BankCode, tblSupplier.AccountNo, tblSupplier.SupplierID, tblSupplier.AccountHolder " & _
                    "ORDER BY tblSupplier.AccountHolder"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        
        While .EOF = False
            
            If !totalValue > 0 Then
            
                myworksheet.Cells(i, 1) = "0000"
                myworksheet.Cells(i, 2) = ![BankCode]
                myworksheet.Cells(i, 3) = Format(![BankCode], "")
                myworksheet.Cells(i, 4) = Format(Val(!AccountNo), "000000000000")
                
                myworksheet.Cells(i, 5) = !AccountHolder
                myworksheet.Cells(i, 6) = "23"
                myworksheet.Cells(i, 7) = "00"
                myworksheet.Cells(i, 8) = "0"
                
                myworksheet.Cells(i, 9) = "000000"
                myworksheet.Cells(i, 10) = ComBankValue(!totalValue)
                myworksheet.Cells(i, 11) = "SLR"
                myworksheet.Cells(i, 12) = CombankID
                myworksheet.Cells(i, 13) = CombankBranchCode
                myworksheet.Cells(i, 14) = ComBankAccountNo
                myworksheet.Cells(i, 15) = ComBankAccountName
                myworksheet.Cells(i, 16) = 0
                myworksheet.Cells(i, 17) = "MilkPay" & Format(![MaxFromDate], "yy MM dd")
                myworksheet.Cells(i, 18) = Format(Date, "yyMMdd")
                myworksheet.Cells(i, 19) = ""
                myworksheet.Cells(i, 20) = "@"
                
                i = i + 1
            
                temFromDate = !MaxFromDate
                TotalPayment = TotalPayment + !totalValue
            End If
                
            .MoveNext
        Wend
        .Close
    End With
            Dim temFindVolumePaymentDue As Double
            temFindVolumePaymentDue = FindVolumePaymentDue(AllFarmers, dtpFrom.Value, dtpTo.Value)
            AddVolumeDeductionToCombankExcel temFindVolumePaymentDue, CInt(i), myworksheet, dtpFrom.Value
            i = i + 1
            
            TotalPayment = TotalPayment + temFindVolumePaymentDue
            
            myworksheet.Cells(i, 1) = "0000"
            myworksheet.Cells(i, 2) = CombankID
            myworksheet.Cells(i, 3) = CombankBranchCode
            myworksheet.Cells(i, 4) = ComBankAccountNo
            
            myworksheet.Cells(i, 5) = ComBankAccountName
            myworksheet.Cells(i, 6) = "23"
            myworksheet.Cells(i, 7) = "00"
            myworksheet.Cells(i, 8) = "1"
            
            myworksheet.Cells(i, 9) = "000000"
            myworksheet.Cells(i, 10) = ComBankValue(TotalPayment)
            myworksheet.Cells(i, 11) = "SLR"
            myworksheet.Cells(i, 12) = CombankID
            myworksheet.Cells(i, 13) = CombankBranchCode
            myworksheet.Cells(i, 14) = ComBankAccountNo
            myworksheet.Cells(i, 15) = ComBankAccountName
            myworksheet.Cells(i, 16) = "0"
            myworksheet.Cells(i, 17) = "MilkTotal" & Format(temFromDate, "yy MM dd")
            myworksheet.Cells(i, 18) = Format(Date, "yyMMdd")
            myworksheet.Cells(i, 19) = ""
            myworksheet.Cells(i, 20) = "@"
    
            myworksheet.Cells(6, 16) = ""
    
    
            Dim SaveAsName As String
            SaveAsName = App.Path & "\MilkPaymentComBank" & ".xls"

            myworksheet.Activate
            myworkbook.Save
            
            myworkbook.SaveAs SaveAsName

            Dim r As Long
            r = ShellExecute(0, "open", SaveAsName, 0, 0, 1)

End Sub

Private Sub btnExcelSam_Click()
    
    'On Error Resume Next
    
    Dim AppExcel As Excel.Application
    Dim myworkbook As Excel.Workbook
    Dim myworksheet As Excel.Worksheet
    Dim mychart As Excel.Chart
    Dim tempath As String
    Dim FSys As New Scripting.FileSystemObject
    Dim TemTopic As String
    Dim temSubTopic As String
    Dim RetVal As Integer
    Dim temText As String
    Dim temMoney As String
    Dim rsReport As New ADODB.Recordset
    Dim i As Integer
    Dim AccNo As String
   
    
    
    Dim MyNumberWord As String
    Dim MyNumberWordLength As Long
    Dim RupeeWord As String
    Dim CentsWord As String
    
    Dim myNum As New clsNumbers
    
    Dim temFromDate As Date
    
    tempath = App.Path

    If FSys.FileExists(tempath & "\sampathbank.xls") = False Then
        MsgBox "THe file named sampathbank.xls is not found the the database folder. Please find the file and place it at " & App.Path
        Exit Sub
    End If


    i = ShellExecute(0, "open", tempath & "\sampathbank.xls", 0, 0, 1)
    
    Set myworkbook = GetObject(tempath & "\sampathbank.xls")
    Set myworksheet = myworkbook.WorkSheets(1)
    
    Dim TotalPayment As Double
    
                
    MyNumberWord = Format(Val(lblBankValue.Caption), "0.00")
    MyNumberWordLength = Len(MyNumberWord)
    RupeeWord = Left(MyNumberWord, MyNumberWordLength - 3)
    CentsWord = Right(MyNumberWord, 2)
    
    temMoney = RupeeWord
    
    i = 2

    myworksheet.Range("A2:Z5500").ClearContents

    With rsReport
        If .State = 1 Then .Close
'        temSql = "SELECT tblBank.BankCode, tblCity.BankCode, tblSupplier.AccountNo, tblSupplier.AccountHolder, sum(tblSupplierPayments.Value) as TotalValue , Max( tblCollectingCenterPaymentSummery.FromDate) as MaxFromDate " & _
'                    "FROM (((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID) LEFT JOIN tblCity ON tblSupplier.CityID = tblCity.CityId " & _
'                    "Where tblCollectingCenterPaymentSummery.FromDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And tblSupplier.BankID = 1 " & _
'                    "GROUP BY tblBank.BankCode, tblCity.BankCode, tblSupplier.AccountNo, tblSupplier.AccountHolder " & _
'                    "ORDER BY tblSupplier.AccountHolder"
        
        temSQL = "SELECT tblSupplier.Supplier, tblSupplier.SupplierID, tblSupplier.AccountNo, max(tblSupplierPayments.FromDate) as MaxFromDate, tblBank.BankCode, tblSupplier.AccountHolder, tblSupplier.SupplierID, sum(tblSupplierPayments.Value) as TotalValue , tblSupplierPayments.GeneratedPaymentMethodID, tblBank.Bank, tblCity.City " & _
                    "FROM ((tblSupplierPayments LEFT JOIN tblSupplier ON tblSupplierPayments.SupplierID = tblSupplier.SupplierID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID) LEFT JOIN tblCity ON tblSupplier.CityID = tblCity.CityId " & _
                    "Where ((tblSupplier.CollectingCenterID)= " & Val(cmbCC.BoundText) & ") AND tblSupplierPayments.FromDate   between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'  And tblSupplier.BankID = 1 " & _
                    "GROUP BY tblSupplier.Supplier, tblSupplier.SupplierID, tblSupplier.AccountNo, tblBank.BankCode, tblSupplier.AccountHolder, tblSupplier.SupplierID, tblSupplierPayments.GeneratedPaymentMethodID, tblBank.Bank, tblCity.City " & _
                    "ORDER BY tblSupplier.AccountHolder"
        
        
        
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            
            
            If UCase(Trim(!AccountHolder)) <> UCase(Trim(!Supplier)) Then
                myworksheet.Cells(i, 1) = !AccountHolder & " For " & !Supplier
            Else
                myworksheet.Cells(i, 1) = !AccountHolder
            End If
            myworksheet.Cells(i, 2) = !AccountNo
            myworksheet.Cells(i, 3) = ![BankCode]
            myworksheet.Cells(i, 4) = Format(!City, "")
            myworksheet.Cells(i, 5) = Format(!totalValue, "0.00")
            
'            myworksheet.Cells(i, 1) = "0000"
'            myworksheet.Cells(i, 3) = Format(![tblCity.BankCode], "")
'            myworksheet.Cells(i, 4) = !AccountNo
'
'            myworksheet.Cells(i, 6) = "23"
'            myworksheet.Cells(i, 7) = "00"
'            myworksheet.Cells(i, 8) = "0"
'
'            myworksheet.Cells(i, 9) = ""
'            myworksheet.Cells(i, 11) = "SLR"
'            myworksheet.Cells(i, 12) = CombankID
'            myworksheet.Cells(i, 15) = ComBankAccountName
'            myworksheet.Cells(i, 16) = ""
'            myworksheet.Cells(i, 17) = "MilkPay" & Format(![FromDate], "yy MM dd")
'            myworksheet.Cells(i, 18) = Format(Date, "yyMMdd")
'            myworksheet.Cells(i, 19) = ""
'            myworksheet.Cells(i, 20) = "@"
            
    
            temFromDate = !MaxFromDate
            TotalPayment = TotalPayment + !totalValue
    
            .MoveNext
            i = i + 1
        Wend
        .Close
    End With
    
            
'            myworksheet.Cells(i, 1) = "0000"
'            myworksheet.Cells(i, 2) = CombankID
'            myworksheet.Cells(i, 3) = CombankBranchCode
'            myworksheet.Cells(i, 4) = ComBankAccountNo
'
'            myworksheet.Cells(i, 5) = ComBankAccountName
'            myworksheet.Cells(i, 6) = "23"
'            myworksheet.Cells(i, 7) = "00"
'            myworksheet.Cells(i, 8) = "1"
'
'            myworksheet.Cells(i, 9) = ""
'            myworksheet.Cells(i, 10) = ComBankValue(TotalPayment)
'            myworksheet.Cells(i, 11) = "SLR"
'            myworksheet.Cells(i, 12) = CombankID
'            myworksheet.Cells(i, 13) = CombankBranchCode
'            myworksheet.Cells(i, 14) = ComBankAccountNo
'            myworksheet.Cells(i, 15) = ComBankAccountName
'            myworksheet.Cells(i, 16) = ""
'            myworksheet.Cells(i, 17) = "MilkTotal" & Format(temFromDate, "yy MM dd")
'            myworksheet.Cells(i, 18) = Format(Date, "yyMMdd")
'            myworksheet.Cells(i, 19) = ""
'            myworksheet.Cells(i, 20) = "@"
    
            Dim SaveAsName As String
            SaveAsName = App.Path & "\MilkPaymentSampathBank" & Format(dtpFrom.Value, "dd MMM yyyy") & ".xls"

            myworksheet.Activate
            myworkbook.Save
            
            myworkbook.SaveAs SaveAsName

            Dim r As Long
            r = ShellExecute(0, "open", SaveAsName, 0, 0, 1)


End Sub

Private Sub btnPrint_Click()
    Dim RetVal As Integer
    
    Dim rsReport As New ADODB.Recordset
    Dim i As Integer
    With rsReport
        If .State = 1 Then .Close
        temSQL = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')' AS SupAccount , tblBank.Bank, tblSupplier.AccountNo AS Account, sum(tblSupplierPayments.Value) as TotalValue  " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where (  ((tblSupplier.CollectingCenterID)= " & Val(cmbCC.BoundText) & ") AND            ((tblCollectingCenterPaymentSummery.FromDate)  between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' ) And ((tblSupplierPayments.GeneratedPaymentMethodID) = " & 2 & ")) " & _
                    "GROUP BY tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')' AS SupAccount , tblBank.Bank, tblSupplier.AccountNo AS Account " & _
                    "ORDER BY tblSupplier.Supplier"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtrPayAdvice
        Set .DataSource = rsReport
        .Sections("Section1").Controls("txtSupplier").DataField = "SupAccount"
        .Sections("Section1").Controls("txtBank").DataField = "Bank"
        .Sections("Section1").Controls("txtAccount").DataField = "Account"
        .Sections("Section1").Controls("txtValue").DataField = "Value"
        .Sections("Section5").Controls("funValue").DataField = "Value"
        
        .Sections("Section4").Controls("lblName").Caption = InstitutionName
        .Sections("Section4").Controls("lblTopic").Caption = "Payment Advice From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
        .Sections("Section4").Controls("lblSubTopic").Caption = "CASH PAYMENTS"
        
        .Show
        i = MsgBox("Print CASH report?", vbYesNo)
        If i = vbYes Then .PrintReport True
        i = MsgBox("Save CASH report?", vbYesNo)
        If i = vbYes Then .ExportReport , "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " - Cash ", True, True
    End With
    With rsReport
        If .State = 1 Then .Close
        temSQL = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.AccountNo AS Account, sum(tblSupplierPayments.Value) as TotalValue  " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where (((tblCollectingCenterPaymentSummery.FromDate)  between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' ) And ((tblSupplierPayments.GeneratedPaymentMethodID) = " & 1 & ")) " & _
                    "GROUP By tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')' AS SupAccount , tblBank.Bank, tblSupplier.AccountNo AS Account " & _
                    "ORDER BY tblSupplier.Supplier"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtrPayAdviceBank
        Set .DataSource = rsReport
        .Sections("Section1").Controls("txtCode").DataField = "SupplierCode"
        .Sections("Section1").Controls("txtSupplier").DataField = "SupAccount"
        .Sections("Section1").Controls("txtBank").DataField = "Bank"
        .Sections("Section1").Controls("txtAccount").DataField = "Account"
        .Sections("Section1").Controls("txtValue").DataField = "Value"
        .Sections("Section5").Controls("funValue").DataField = "Value"
        
        .Sections("Section4").Controls("lbldate").Caption = Format(Date, "dd MMMM yyyy")
        .Sections("Section4").Controls("lblCenter").Caption = "On " & Format(dtpFrom.Value, "dd MMMM yyyy")
        '.Sections("Section4").Controls("lblSubTopic").Caption = "BANK PAYMENTS"
        
        .Show
        i = MsgBox("Print BANK report?", vbYesNo)
        If i = vbYes Then .PrintReport True
        i = MsgBox("Save BANK report?", vbYesNo)
        If i = vbYes Then .ExportReport , "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " - Bank ", True, True
    End With
    With rsReport
        If .State = 1 Then .Close
        temSQL = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.AccountNo AS Account, sum(tblSupplierPayments.Value) as TotalValue  " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where (((tblCollectingCenterPaymentSummery.FromDate)  between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') And ((tblSupplierPayments.GeneratedPaymentMethodID) = " & 3 & ")) " & _
                    "GROUP BY tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.AccountNo AS Account " & _
                    "ORDER BY tblSupplier.Supplier"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtrPayAdviceCheque
        Set .DataSource = rsReport
        .Sections("Section1").Controls("txtSupplier").DataField = "SupAccount"
        .Sections("Section1").Controls("txtBank").DataField = "Bank"
        .Sections("Section1").Controls("txtAccount").DataField = "Account"
        .Sections("Section1").Controls("txtValue").DataField = "Value"
        .Sections("Section5").Controls("funValue").DataField = "Value"
        .Sections("Section4").Controls("lblName").Caption = InstitutionName
        .Sections("Section4").Controls("lblTopic").Caption = "Payment Advice From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " To " & Format(dtpTo.Value, "dd MMMM yyyy")
        .Sections("Section4").Controls("lblSubTopic").Caption = "CHEQUE PAYMENTS"
        .Show
        i = MsgBox("Print CHEQUE report?", vbYesNo)
        If i = vbYes Then .PrintReport True
        i = MsgBox("Save CHEQUE report?", vbYesNo)
        If i = vbYes Then .ExportReport , "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " - Cheque ", True, True
    End With
    
    btnClose.Enabled = True

End Sub

Private Sub btnPrintBankC_Click()
    Dim RetVal As Integer
    Dim temText As String
    Dim temMoney As String
    Dim rsReport As New ADODB.Recordset
    Dim rsReport1 As New ADODB.Recordset
    Dim i As Integer
    Dim AccNo As String
    Dim AllFarmers() As Double
   
    Dim MyNumberWord As String
    Dim MyNumberWordLength As Long
    Dim RupeeWord As String
    Dim CentsWord As String
    
    Dim myNum As New clsNumbers
    
    
    
        With rsReport1
             If .State = 1 Then .Close
             temSQL = "SELECT  tblSupplier.SupplierID, tblBank.BankCode, tblCity.BankCode, tblSupplier.AccountNo, tblSupplier.AccountHolder, sum(tblSupplierPayments.Value) as TotalValue , max(tblCollectingCenterPaymentSummery.FromDate) as MaxFromDate " & _
                         "FROM (((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID) LEFT JOIN tblCity ON tblSupplier.CityID = tblCity.CityId " & _
                         "Where ((tblSupplier.CollectingCenterID)= " & Val(cmbCC.BoundText) & ") AND tblCollectingCenterPaymentSummery.FromDate Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And (tblSupplier.BankID = 2 or tblSupplier.BankID = 1 ) " & _
                         "GROUP BY tblBank.BankCode, tblCity.BankCode, tblSupplier.AccountNo, tblSupplier.SupplierID, tblSupplier.AccountHolder " & _
                         "ORDER BY tblSupplier.AccountHolder"
             .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
             .MoveLast
             ReDim AllFarmers(.RecordCount)
             .MoveFirst
             Dim FarmerCount As Integer
             While .EOF = False
                 AllFarmers(FarmerCount) = !SupplierID
                 FarmerCount = FarmerCount + 1
                 .MoveNext
             Wend
        End With
        Dim temFindVolumePaymentDue As Double
        temFindVolumePaymentDue = FindVolumePaymentDue(AllFarmers, dtpFrom.Value, dtpTo.Value)
    
   AccNo = "1104013300"
                
    MyNumberWord = Format(Val(lblBankValue.Caption), "0.00")
    MyNumberWordLength = Len(MyNumberWord)
    RupeeWord = Left(MyNumberWord, MyNumberWordLength - 3)
    CentsWord = Right(MyNumberWord, 2)
    
    temMoney = RupeeWord
    
    If Val(RupeeWord) <> 1 Then
        temMoney = myNum.NumberToWord(RupeeWord) & "Rupees"
    Else
        temMoney = myNum.NumberToWord(RupeeWord) & "Rupee"
    End If

    If Val(CentsWord) > 0 Then
        If Val(CentsWord) <> 1 Then
            temMoney = temMoney & " and " & myNum.NumberToWord(CentsWord) & " Cents"
        Else
            temMoney = temMoney & " and " & myNum.NumberToWord(CentsWord) & " Cent"
        End If
    End If
    temMoney = temMoney & " only "

    temText = "Please make necessary arrangements to transfer of " & temMoney & "(Rs. " & lblBankValue.Caption & ") from current account No. " & AccNo & " to Accounts numbers mentioned below as follows."

    With rsReport
        If .State = 1 Then .Close
'        temSql = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo AS Account, sum(tblSupplierPayments.Value) as TotalValue  " & _
'                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
'                    "Where tblCollectingCenterPaymentSummery.FromDate  between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'  And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And tblSupplier.BankID = 2 " & _
'                    "GROUP BY SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo AS Account " & _
'                    "ORDER BY tblSupplier.Supplier"
        
        
              temSQL = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo AS Account, sum(tblSupplierPayments.Value) as TotalValue  " & _
                        "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                        "Where ((tblSupplier.CollectingCenterID)= " & Val(cmbCC.BoundText) & ") AND  tblCollectingCenterPaymentSummery.FromDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'  And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And tblSupplier.BankID = 2 " & _
                        "GROUP BY tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo " & _
                        ""
    
        
        
        
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtrPayAdviceBank
        Set .DataSource = rsReport
        .Sections("Section1").Controls("txtCode").DataField = "SupplierCode"
        .Sections("Section1").Controls("txtSupplier").DataField = "SupAccount"
        .Sections("Section1").Controls("txtBank").DataField = "Bank"
        .Sections("Section1").Controls("txtAccount").DataField = "Account"
        .Sections("Section1").Controls("txtValue").DataField = "TotalValue"
        .Sections("Section5").Controls("funValue").DataField = "TotalValue"
        
        .Sections("Section4").Controls("lbldate").Caption = Format(Date, "dd MMMM yyyy")
        
        .Sections("Section5").Controls("lblWellfair").Caption = Format(temFindVolumePaymentDue, "#,##0.00")
        
        .Sections("Section4").Controls("lblCenter").Caption = "All Centers"
        .Sections("Section4").Controls("lblPeriod").Caption = "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
        
        .Sections("Section4").Controls("lblBank").Caption = "Commercial Bank"
        .Sections("Section4").Controls("lblBranch").Caption = "Kamburupitiya"
        .Sections("Section4").Controls("lblText").Caption = temText
        
        'lblWellfair
        
        .Show
        i = MsgBox("Print BANK report?", vbYesNo)
        If i = vbYes Then .PrintReport True
        i = MsgBox("Save BANK report?", vbYesNo)
        If i = vbYes Then .ExportReport , "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " - Bank ", True, True
    End With

End Sub

Private Sub btnPrintBankS_Click()
    Dim RetVal As Integer
    Dim temText As String
    Dim temMoney As String
    Dim rsReport As New ADODB.Recordset
    Dim i As Integer
    Dim AccNo As String
   
    Dim MyNumberWord As String
    Dim MyNumberWordLength As Long
    Dim RupeeWord As String
    Dim CentsWord As String
    
    Dim myNum As New clsNumbers
    
    AccNo = "001010006216"
                
    MyNumberWord = Format(Val(lblBankValue.Caption), "0.00")
    MyNumberWordLength = Len(MyNumberWord)
    RupeeWord = Left(MyNumberWord, MyNumberWordLength - 3)
    CentsWord = Right(MyNumberWord, 2)
    
    temMoney = RupeeWord
    
    If Val(RupeeWord) <> 1 Then
        temMoney = myNum.NumberToWord(RupeeWord) & "Rupees"
    Else
        temMoney = myNum.NumberToWord(RupeeWord) & "Rupee"
    End If

    If Val(CentsWord) > 0 Then
        If Val(CentsWord) <> 1 Then
            temMoney = temMoney & " and " & myNum.NumberToWord(CentsWord) & " Cents"
        Else
            temMoney = temMoney & " and " & myNum.NumberToWord(CentsWord) & " Cent"
        End If
    End If
    temMoney = temMoney & " only "

    temText = "Please make necessary arrangements to transfer of " & temMoney & "(Rs. " & lblBankValue.Caption & ") from current account No. " & AccNo & " to Accounts numbers mentioned below as follows."

    With rsReport
        If .State = 1 Then .Close
    '        temSql = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo AS Account, sum(tblSupplierPayments.Value) as TotalValue  " & _
    '                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
    '                    "Where tblCollectingCenterPaymentSummery.FromDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'  And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And tblSupplier.BankID = 1 " & _
    '                    "GROUP BY tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo " & _
    '                    "ORDER BY tblSupplier.Supplier"
    '
    '
        
            temSQL = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo AS Account, sum(tblSupplierPayments.Value) as TotalValue  " & _
                        "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                        "Where ((tblSupplier.CollectingCenterID)= " & Val(cmbCC.BoundText) & ") AND  tblCollectingCenterPaymentSummery.FromDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'  And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And tblSupplier.BankID = 1 " & _
                        "GROUP BY tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo " & _
                        ""
    
    
        
        
        
        
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtrPayAdviceBank
        Set .DataSource = rsReport
        .Sections("Section1").Controls("txtCode").DataField = "SupplierCode"
        .Sections("Section1").Controls("txtSupplier").DataField = "SupAccount"
        .Sections("Section1").Controls("txtBank").DataField = "Bank"
        .Sections("Section1").Controls("txtAccount").DataField = "Account"
        .Sections("Section1").Controls("txtValue").DataField = "TotalValue"
        .Sections("Section5").Controls("funValue").DataField = "TotalValue"
        
        .Sections("Section4").Controls("lbldate").Caption = Format(Date, "dd MMMM yyyy")
        
        .Sections("Section4").Controls("lblCenter").Caption = "All Centres"
        .Sections("Section4").Controls("lblPeriod").Caption = "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
        
        .Sections("Section4").Controls("lblBank").Caption = "Sampath Bank"
        .Sections("Section4").Controls("lblBranch").Caption = "Matara"
        .Sections("Section4").Controls("lblText").Caption = temText
        
        
        
        .Show
        i = MsgBox("Print BANK report?", vbYesNo)
        If i = vbYes Then .PrintReport True
        i = MsgBox("Save BANK report?", vbYesNo)
        If i = vbYes Then .ExportReport , "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " - Bank ", True, True
    End With

End Sub




Private Sub btnProcess_Click()
    Screen.MousePointer = vbHourglass
    FillLists
    CalculateValues
    BankEnable
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmbPrinter_Change()
    Call ListPapers
End Sub

Private Sub cmbPrinter_Click()
    Call ListPapers
End Sub

Private Sub Form_Load()
    FillCombos
    Call ListPrinters
    Call ComBankDetails
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, "Options", "Printer", "")
    Call ListPapers
    cmbPaper.Text = GetSetting(App.EXEName, "Options", "Paper", "")
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = Date
End Sub

Private Sub FillCombos()
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCC, "CollectingCenter", True

End Sub

Private Sub ListPrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub


Private Sub ListPapers()
    cmbPaper.Clear
    CSetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        With FormSize
            .cx = BillPaperHeight
            .cy = BillPaperWidth
        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                cmbPaper.AddItem PtrCtoVbString(.pName)
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub


Private Sub FillLists()
    lstAllSupplierIDs.Clear
    lstAllSuppliers.Clear
    lstAllVaue.Clear
    lstBank.Clear
    lstBankIDs.Clear
    lstBankValue.Clear
    lstPaymentMethod.Clear
    lstSelectedSupplierIDs.Clear
    lstSelectedSuppliers.Clear
    lstSelectedValue.Clear
    
    Dim rsSuppliers As New ADODB.Recordset
    Dim i As Integer
    With rsSuppliers
        If optAll.Value = True Then
            temSQL = "SELECT tblSupplier.Supplier, tblSupplier.SupplierID,  sum(tblSupplierPayments.Value) as TotalValue ,tblSupplierPayments.GeneratedPaymentMethodID " & _
                        "FROM tblSupplierPayments LEFT JOIN tblSupplier ON tblSupplierPayments.SupplierID = tblSupplier.SupplierID " & _
                        "Where (((tblSupplier.CollectingCenterID)= " & Val(cmbCC.BoundText) & ") AND ((tblSupplierPayments.FromDate)  between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' )) " & _
                        "GROUP By tblSupplier.Supplier, tblSupplier.SupplierID,  tblSupplierPayments.GeneratedPaymentMethodID " & _
                        "ORDER BY tblSupplier.Supplier"
        ElseIf optCom.Value = True Then
            temSQL = "SELECT tblSupplier.Supplier, tblSupplier.SupplierID,  sum(tblSupplierPayments.Value) as TotalValue  ,tblSupplierPayments.GeneratedPaymentMethodID " & _
                        "FROM tblSupplierPayments LEFT JOIN tblSupplier ON tblSupplierPayments.SupplierID = tblSupplier.SupplierID " & _
                        "Where ((tblSupplier.CollectingCenterID)= " & Val(cmbCC.BoundText) & ") AND tblSupplierPayments.FromDate   between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'  AND tblSupplier.BankID = 2 " & _
                        "GROUP By tblSupplier.Supplier, tblSupplier.SupplierID,  tblSupplierPayments.GeneratedPaymentMethodID " & _
                        "ORDER BY tblSupplier.Supplier"
        ElseIf optSB.Value = True Then
            temSQL = "SELECT tblSupplier.Supplier, tblSupplier.SupplierID,  sum(tblSupplierPayments.Value) as TotalValue  ,tblSupplierPayments.GeneratedPaymentMethodID " & _
                        "FROM tblSupplierPayments LEFT JOIN tblSupplier ON tblSupplierPayments.SupplierID = tblSupplier.SupplierID " & _
                        "Where ((tblSupplier.CollectingCenterID)= " & Val(cmbCC.BoundText) & ") AND tblSupplierPayments.FromDate   between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'  AND tblSupplier.BankID = 1 " & _
                        "GROUP By tblSupplier.Supplier, tblSupplier.SupplierID,  tblSupplierPayments.GeneratedPaymentMethodID " & _
                        "ORDER BY tblSupplier.Supplier"
        
        Else
            Exit Sub
        End If
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            ReDim Supplier(.RecordCount)
            ReDim SupplierID(.RecordCount)
            ReDim SupplierSelected(.RecordCount)
            ReDim PaymentMethodID(.RecordCount)
            ReDim SupplierPaymentsID(.RecordCount)
            ReDim SupplierValue(.RecordCount)
            
            i = 0
            While .EOF = False
                Supplier(i) = !Supplier
                SupplierID(i) = !SupplierID
'                SupplierPaymentsID(i) = !SupplierPaymentsID
'                lstAllSupplierIDs.AddItem !SupplierPaymentsID
                lstAllSuppliers.AddItem !Supplier
                lstAllVaue.AddItem !totalValue
                SupplierValue(i) = !totalValue
                
                If IsNull(!GeneratedPaymentMethodID) = False Then
                    Select Case !GeneratedPaymentMethodID
                        Case 1:
                                lstBank.AddItem !Supplier
                                lstBankValue.AddItem Right(Space(20) & Format(!totalValue, "0.00"), 20)
                                lstBankIDs.AddItem !SupplierID
                    End Select
                                    
                End If
                i = i + 1
                .MoveNext
            Wend
        End If
        .Close
    End With
End Sub

Private Sub CalculateValues()
    Dim i As Integer
    Dim temValue As Double
    
    temValue = 0
    For i = 0 To lstBankValue.ListCount - 1
        temValue = temValue + Val(lstBankValue.List(i))
    Next i
    lblBankValue.Caption = Format(temValue, "0.00")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, "Options", "Printer", cmbPrinter.Text
    SaveSetting App.EXEName, "Options", "Paper", cmbPaper.Text
End Sub

Private Sub lstBank_Click()
    Dim i As Integer
    For i = 0 To lstBank.ListCount - 1
        lstBankValue.Selected(i) = lstBank.Selected(i)
    Next
End Sub

Private Sub lstBank_LostFocus()
    Dim i As Integer
    For i = 0 To lstBank.ListCount - 1
        lstBankValue.Selected(i) = lstBank.Selected(i)
    Next
End Sub

Private Sub lstBank_Scroll()
    lstBankValue.TopIndex = lstBank.TopIndex
End Sub

Private Sub lstBankValue_Click()
    Dim i As Integer
    For i = 0 To lstBank.ListCount - 1
        lstBankValue.Selected(i) = lstBank.Selected(i)
    Next
End Sub

Private Sub lstBankValue_LostFocus()
    Dim i As Integer
    For i = 0 To lstBank.ListCount - 1
        lstBankValue.Selected(i) = lstBank.Selected(i)
    Next
End Sub

Private Sub lstBankValue_Scroll()
    lstBank.TopIndex = lstBankValue.TopIndex
End Sub

Private Sub lstSelectedSuppliers_Click()
    If SettingNames = True Then Exit Sub
    SettingValues = True
    Dim i As Integer
    For i = 0 To lstSelectedSuppliers.ListCount - 1
        lstSelectedValue.Selected(i) = lstSelectedSuppliers.Selected(i)
    Next
    SettingValues = False
End Sub

Private Sub lstSelectedSuppliers_ItemCheck(Item As Integer)
    Dim i As Integer
    For i = 0 To lstSelectedSuppliers.ListCount - 1
        lstSelectedValue.Selected(i) = lstSelectedSuppliers.Selected(i)
    Next
End Sub

Private Sub lstSelectedSuppliers_LostFocus()
    Dim i As Integer
    For i = 0 To lstSelectedSuppliers.ListCount - 1
        lstSelectedValue.Selected(i) = lstSelectedSuppliers.Selected(i)
    Next
End Sub

Private Sub lstSelectedSuppliers_Scroll()
    lstSelectedValue.TopIndex = lstSelectedSuppliers.TopIndex
End Sub

Private Sub lstSelectedValue_Click()
    If SettingValues = True Then Exit Sub
    SettingNames = True
    Dim i As Integer
    For i = 0 To lstSelectedSuppliers.ListCount - 1
        lstSelectedSuppliers.Selected(i) = lstSelectedValue.Selected(i)
    Next
    SettingNames = False
End Sub

Private Sub lstSelectedValue_LostFocus()
    Dim i As Integer
    For i = 0 To lstSelectedSuppliers.ListCount - 1
        lstSelectedSuppliers.Selected(i) = lstSelectedValue.Selected(i)
    Next
End Sub

Private Sub lstSelectedValue_Scroll()
    lstSelectedSuppliers.TopIndex = lstSelectedValue.TopIndex
End Sub

Private Sub BankEnable()
        If optAll.Value = True Then
            btnPrintBankC.Enabled = False
            btnPrintBankS.Enabled = False
        ElseIf optCom.Value = True Then
            btnPrintBankC.Enabled = True
            btnPrintBankS.Enabled = False
        ElseIf optSB.Value = True Then
            btnPrintBankC.Enabled = False
            btnPrintBankS.Enabled = True
        Else
            btnPrintBankC.Enabled = False
            btnPrintBankS.Enabled = False
        End If

End Sub
