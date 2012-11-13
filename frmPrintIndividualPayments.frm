VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPrintIndividualPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Payments"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
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
   ScaleWidth      =   10620
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   7560
      TabIndex        =   34
      Top             =   6720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   31719427
      CurrentDate     =   40254
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   7560
      TabIndex        =   35
      Top             =   7200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   31850499
      CurrentDate     =   40254
   End
   Begin btButtonEx.ButtonEx btnSampathDetails 
      Height          =   375
      Left            =   7560
      TabIndex        =   33
      Top             =   5760
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Sampath Bank Details"
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
   Begin btButtonEx.ButtonEx btnCommercialDetails 
      Height          =   375
      Left            =   7560
      TabIndex        =   32
      Top             =   6240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Commercial Bank Details"
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
   Begin VB.OptionButton optAll 
      Caption         =   "&All"
      Height          =   255
      Left            =   4440
      TabIndex        =   26
      Top             =   1320
      Width           =   1575
   End
   Begin VB.OptionButton optCom 
      Caption         =   "&Commercial Bank"
      Height          =   255
      Left            =   2400
      TabIndex        =   25
      Top             =   1320
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton optSB 
      Caption         =   "&Sampath Bank"
      Height          =   255
      Left            =   720
      TabIndex        =   24
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ListBox lstSelectedValue 
      Height          =   300
      Left            =   7560
      TabIndex        =   22
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstSelectedSuppliers 
      Height          =   300
      Left            =   7560
      TabIndex        =   21
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstSelectedSupplierIDs 
      Height          =   300
      Left            =   7560
      TabIndex        =   20
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstPaymentMethod 
      Height          =   300
      Left            =   7560
      TabIndex        =   19
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstAllSuppliers 
      Height          =   300
      Left            =   7560
      TabIndex        =   18
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstAllVaue 
      Height          =   300
      Left            =   7560
      TabIndex        =   17
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstAllSupplierIDs 
      Height          =   300
      Left            =   7560
      TabIndex        =   16
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   5280
      Width           =   2655
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4320
      Width           =   2655
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8880
      TabIndex        =   2
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
      Height          =   6810
      ItemData        =   "frmPrintIndividualPayments.frx":0000
      Left            =   720
      List            =   "frmPrintIndividualPayments.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   1680
      Width           =   4455
   End
   Begin MSDataListLib.DataCombo cmbCCPS 
      Height          =   360
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   7440
      TabIndex        =   3
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
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin btButtonEx.ButtonEx btnPrintBankC 
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Top             =   3000
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
      TabIndex        =   13
      Top             =   2040
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
   Begin MSDataListLib.DataCombo cmbCC 
      Height          =   360
      Left            =   720
      TabIndex        =   23
      Top             =   240
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
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
      Height          =   6825
      IntegralHeight  =   0   'False
      ItemData        =   "frmPrintIndividualPayments.frx":0004
      Left            =   4440
      List            =   "frmPrintIndividualPayments.frx":0006
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   1680
      Width           =   2895
   End
   Begin btButtonEx.ButtonEx btnExcelCom 
      Height          =   375
      Left            =   7440
      TabIndex        =   27
      Top             =   3480
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
      TabIndex        =   28
      Top             =   2520
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
   Begin btButtonEx.ButtonEx btnAll 
      Height          =   375
      Left            =   7320
      TabIndex        =   29
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&All Payments"
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
   Begin btButtonEx.ButtonEx btnAllPeriod 
      Height          =   375
      Left            =   7320
      TabIndex        =   30
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&All Payments for Period"
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
   Begin btButtonEx.ButtonEx btnSelectedPeriod 
      Height          =   375
      Left            =   7320
      TabIndex        =   31
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Centre Payments for Period"
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
   Begin VB.Label Label11 
      Caption         =   "Payment Summery"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Printer"
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Paper"
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label lblBankValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   9000
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Total Bank Value"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   9000
      Width           =   2175
   End
End
Attribute VB_Name = "frmPrintIndividualPayments"
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
    
    
    
Private Sub btnSampathDetails_Click()
    
    Screen.MousePointer = vbHourglass
    
    Dim temCol As Integer
    Dim temRow As Integer
    
    Dim temLiters As Double
    Dim temCommission As Double
    Dim temDeductions As Double
    
        
    Dim totalLiters As Double
    Dim totelCommission As Double
    Dim totalDeduction As Double
    Dim totalValue As Double
    
    
    Dim SupplierID As Long
    
    
    Dim rsTem As New ADODB.Recordset
    
    
    Dim myExcel As New clsExcel
    
    With myExcel
        .Topic = "Sampath Bank Milk Payment Details"
        .SubTopic = cmbCCPS.Text
        .Path = App.Path
        .FileName = .Topic & " " & .SubTopic
        
        temRow = 4
        .setValue temRow, 1, "Code"
        .setValue temRow, 2, "Farmer"
        .setValue temRow, 3, "Liters"
        .setValue temRow, 4, "Commission"
        .setValue temRow, 5, "Deductions"
        .setValue temRow, 6, "Account"
        .setValue temRow, 7, "Amount"
    
    
    End With
    
    temRow = 5
    
    With rsTem
        If .State = 1 Then .Close
'        temSQL = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , " & _
'                    "tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo AS Account, " & _
'                    " tblSupplierPayments.Value, tblSupplier.Supplierid " & _
'                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
'                    "Where tblCollectingCenterPaymentSummery.FromDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And tblSupplier.BankID = 1 " & _
'                    "ORDER BY tblSupplier.Supplier"
'
'
'
        temSQL = "SELECT tblSupplier.SupplierID,  tblCollectingCenterPaymentSummery.FromDate, tblCollectingCenterPaymentSummery.ToDate, tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo AS Account, tblSupplierPayments.Value " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText) & " And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And tblSupplier.BankID = 1 " & _
                    "ORDER BY tblSupplier.Supplier"
        
        
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            SupplierID = !SupplierID
            
            
            dtpFrom.Value = !FromDate
            dtpTo.Value = !ToDate
            
            temLiters = PeriodMilkSupply(dtpFrom.Value, dtpTo.Value, SupplierID, 1).Liters + PeriodMilkSupply(dtpFrom.Value, dtpTo.Value, SupplierID, 2).Liters
            temCommission = OwnCommision(SupplierID, dtpFrom.Value, dtpTo.Value, 0) + OthersCommision(SupplierID, dtpFrom.Value, dtpTo.Value, 0) + AdditionalCommision(SupplierID, dtpFrom.Value, dtpTo.Value)
            temDeductions = PeriodDeductions(SupplierID, dtpFrom.Value, dtpTo.Value)
            
            myExcel.setValue temRow, 1, !SupplierCode
            myExcel.setValue temRow, 2, !SupAccount
            myExcel.setValue temRow, 3, CDbl(temLiters)
            myExcel.setValue temRow, 4, CDbl(temCommission)
            myExcel.setValue temRow, 5, CDbl(temDeductions)
            myExcel.setValue temRow, 6, "A/C No. " & !Account
            myExcel.setValue temRow, 7, !Value
            
            totalLiters = temLiters + totalLiters
            totelCommission = totelCommission + temCommission
            totalDeduction = totalDeduction + temDeductions
            totalValue = totalValue + !Value
            
            .MoveNext
            temRow = temRow + 1
            
        Wend
        .Close
    End With
    
    
    myExcel.setValue temRow, 3, CDbl(totalLiters)
    myExcel.setValue temRow, 4, CDbl(totelCommission)
    myExcel.setValue temRow, 5, CDbl(totalDeduction)
    myExcel.setValue temRow, 7, CDbl(totalValue)
    myExcel.autofitAll
    myExcel.finalizeExcel
    myExcel.saveExcel
    myExcel.openExcel
    
    
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Sub btnCommercialDetails_Click()
    Screen.MousePointer = vbHourglass
    
    Dim temCol As Integer
    Dim temRow As Integer
    
    Dim temLiters As Double
    Dim temCommission As Double
    Dim temDeductions As Double
    
        
    Dim totalLiters As Double
    Dim totelCommission As Double
    Dim totalDeduction As Double
    Dim totalValue As Double
    
    
    Dim SupplierID As Long
    
    
    Dim rsTem As New ADODB.Recordset
    
    
    Dim myExcel As New clsExcel
    
    With myExcel
        .Topic = "Commercial Bank Milk Payment Details"
        .SubTopic = cmbCCPS.Text
        .Path = App.Path
        .FileName = .Topic & " " & .SubTopic
        
        temRow = 4
        .setValue temRow, 1, "Code"
        .setValue temRow, 2, "Farmer"
        .setValue temRow, 3, "Liters"
        .setValue temRow, 4, "Commission"
        .setValue temRow, 5, "Deductions"
        .setValue temRow, 6, "Account"
        .setValue temRow, 7, "Amount"
    
    
    End With
    
    temRow = 5
    
    With rsTem
        If .State = 1 Then .Close
'        temSQL = "SELECT tblSupplier.SupplierID , tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo AS Account, tblSupplierPayments.Value " & _
'                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
'                    "Where tblCollectingCenterPaymentSummery.FromDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And tblSupplier.BankID = 2 " & _
'                    "ORDER BY tblSupplier.Supplier"
        
        temSQL = "SELECT  tblCollectingCenterPaymentSummery.FromDate, tblCollectingCenterPaymentSummery.ToDate, tblSupplier.SupplierID ,  tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo AS Account, tblSupplierPayments.Value " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText) & " And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And tblSupplier.BankID = 2 " & _
                    "ORDER BY tblSupplier.Supplier"
        
        
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            SupplierID = !SupplierID
            
            dtpFrom.Value = !FromDate
            dtpTo.Value = !ToDate
            
            temLiters = PeriodMilkSupply(dtpFrom.Value, dtpTo.Value, SupplierID, 1).Liters + PeriodMilkSupply(dtpFrom.Value, dtpTo.Value, SupplierID, 2).Liters
            temCommission = OwnCommision(SupplierID, dtpFrom.Value, dtpTo.Value, 0) + OthersCommision(SupplierID, dtpFrom.Value, dtpTo.Value, 0) + AdditionalCommision(SupplierID, dtpFrom.Value, dtpTo.Value)
            temDeductions = PeriodDeductions(SupplierID, dtpFrom.Value, dtpTo.Value)
            
            myExcel.setValue temRow, 1, !SupplierCode
            myExcel.setValue temRow, 2, !SupAccount
            myExcel.setValue temRow, 3, CDbl(temLiters)
            myExcel.setValue temRow, 4, CDbl(temCommission)
            myExcel.setValue temRow, 5, CDbl(temDeductions)
            myExcel.setValue temRow, 6, "A/C No. " & !Account
            myExcel.setValue temRow, 7, !Value
            
            totalLiters = temLiters + totalLiters
            totelCommission = totelCommission + temCommission
            totalDeduction = totalDeduction + temDeductions
            totalValue = totalValue + !Value
            
            .MoveNext
            temRow = temRow + 1
            
        Wend
        .Close
    End With
    
    
    myExcel.setValue temRow, 3, CDbl(totalLiters)
    myExcel.setValue temRow, 4, CDbl(totelCommission)
    myExcel.setValue temRow, 5, CDbl(totalDeduction)
    myExcel.setValue temRow, 7, CDbl(totalValue)
    myExcel.autofitAll
    myExcel.finalizeExcel
    myExcel.saveExcel
    myExcel.openExcel
    
    
    Screen.MousePointer = vbDefault
    
    

End Sub
    
    
    
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


Private Sub btnAll_Click()
    frmPrintAllPayments.Show
    frmPrintAllPayments.ZOrder 0
End Sub

Private Sub btnAllPeriod_Click()
    frmPrintAllPaymentsPeriod.Show
    frmPrintAllPaymentsPeriod.ZOrder 0
End Sub

Private Sub btnAllCentre_Click()

End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Function ComBankValue(Value As Double) As String
    ComBankValue = Format(Val(Format(Value, "0.00") * 100), "000000000000")
End Function

Private Sub btnExcelCom_Click()
    On Error Resume Next
    
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
    
    myworksheet.Range("A7:Z5500").ClearContents

    With rsReport
        If .State = 1 Then .Close
        temSQL = "SELECT tblBank.BankCode, tblCity.BankCode, tblSupplier.AccountNo, tblSupplier.AccountHolder, tblSupplierPayments.Value, tblCollectingCenterPaymentSummery.FromDate " & _
                    "FROM (((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID) LEFT JOIN tblCity ON tblSupplier.CityID = tblCity.CityId " & _
                    "Where tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText) & " And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And tblSupplier.BankID = 2 " & _
                    "ORDER BY tblSupplier.AccountHolder"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            
            myworksheet.Cells(i, 1) = "0000"
            myworksheet.Cells(i, 2) = ![BankCode]
            myworksheet.Cells(i, 3) = Format(![BankCode], "")
            myworksheet.Cells(i, 4) = Format(Val(!AccountNo), "000000000000")
            
            myworksheet.Cells(i, 5) = !AccountHolder
            myworksheet.Cells(i, 6) = "23"
            myworksheet.Cells(i, 7) = "00"
            myworksheet.Cells(i, 8) = "0"
            
            myworksheet.Cells(i, 9) = "000000"
            myworksheet.Cells(i, 10) = ComBankValue(!Value)
            myworksheet.Cells(i, 11) = "SLR"
            myworksheet.Cells(i, 12) = CombankID
            myworksheet.Cells(i, 13) = CombankBranchCode
            myworksheet.Cells(i, 14) = ComBankAccountNo
            myworksheet.Cells(i, 15) = ComBankAccountName
            myworksheet.Cells(i, 16) = ""
            myworksheet.Cells(i, 17) = "MilkPay" & Format(![FromDate], "yy MM dd")
            myworksheet.Cells(i, 18) = Format(Date, "yyMMdd")
            myworksheet.Cells(i, 19) = ""
            myworksheet.Cells(i, 20) = "@"
            
    
            temFromDate = !FromDate
            TotalPayment = TotalPayment + !Value
            i = i + 1
    
            .MoveNext
        Wend
        .Close
    End With
    
            
            myworksheet.Cells(i, 1) = "0000"
            myworksheet.Cells(i, 2) = CombankID
            myworksheet.Cells(i, 3) = CombankBranchCode
            myworksheet.Cells(i, 4) = ComBankAccountNo
            
            myworksheet.Cells(i, 5) = ComBankAccountName
            myworksheet.Cells(i, 6) = "23"
            myworksheet.Cells(i, 7) = "00"
            myworksheet.Cells(i, 8) = "1"
            
            myworksheet.Cells(i, 9) = ""
            myworksheet.Cells(i, 10) = ComBankValue(TotalPayment)
            myworksheet.Cells(i, 11) = "SLR"
            myworksheet.Cells(i, 12) = CombankID
            myworksheet.Cells(i, 13) = CombankBranchCode
            myworksheet.Cells(i, 14) = ComBankAccountNo
            myworksheet.Cells(i, 15) = ComBankAccountName
            myworksheet.Cells(i, 16) = ""
            myworksheet.Cells(i, 17) = "MilkTotal" & Format(temFromDate, "yy MM dd")
            myworksheet.Cells(i, 18) = Format(Date, "yyMMdd")
            myworksheet.Cells(i, 19) = ""
            myworksheet.Cells(i, 20) = "@"
            
            myworksheet.Cells(6, 16) = ""
            
            myworksheet.Cells(1, 1) = ""
            
            Dim SaveAsName As String
            SaveAsName = App.Path & "\MilkPaymentComBank" & cmbCC.Text & cmbCCPS.Text & ".xls"

            myworksheet.Activate
            myworkbook.Save
            
            myworkbook.SaveAs SaveAsName

            Dim r As Long
            r = ShellExecute(0, "open", SaveAsName, 0, 0, 1)

End Sub

Private Sub btnExcelSam_Click()
    
    On Error Resume Next
    
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
        temSQL = "SELECT tblBank.BankCode, tblCity.BankCode, tblSupplier.AccountNo, tblSupplier.AccountHolder, tblSupplierPayments.Value, tblCollectingCenterPaymentSummery.FromDate " & _
                    "FROM (((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID) LEFT JOIN tblCity ON tblSupplier.CityID = tblCity.CityId " & _
                    "Where tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText) & " And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And tblSupplier.BankID = 1 " & _
                    "ORDER BY tblSupplier.AccountHolder"
        
        temSQL = "SELECT tblSupplier.Supplier, tblSupplier.AccountNo, tblSupplierPayments.FromDate, tblBank.BankCode, tblSupplier.AccountHolder, tblSupplier.SupplierID, tblSupplierPayments.SupplierPaymentsID, tblSupplierPayments.Value, tblSupplierPayments.GeneratedPaymentMethodID, tblBank.Bank, tblCity.City " & _
                    "FROM ((tblSupplierPayments LEFT JOIN tblSupplier ON tblSupplierPayments.SupplierID = tblSupplier.SupplierID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID) LEFT JOIN tblCity ON tblSupplier.CityID = tblCity.CityId " & _
                    "Where tblSupplierPayments.CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText) & " And tblSupplier.BankID = 1 " & _
                    "ORDER BY tblSupplier.AccountHolder"
        
        
        
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            
            myworksheet.Cells(i, 1) = !AccountHolder
            myworksheet.Cells(i, 2) = !AccountNo
            myworksheet.Cells(i, 3) = ![BankCode]
            myworksheet.Cells(i, 4) = Format(!City, "")
            myworksheet.Cells(i, 5) = Format(!Value, "0.00")
            
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
            
    
            temFromDate = !FromDate
            TotalPayment = TotalPayment + !Value
    
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
            SaveAsName = App.Path & "\MilkPaymentSampathBank" & cmbCC.Text & cmbCCPS.Text & ".xls"

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
        temSQL = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')' AS SupAccount , tblBank.Bank, tblSupplier.AccountNo AS Account, tblSupplierPayments.Value " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where (((tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) = " & Val(cmbCCPS.BoundText) & ") And ((tblSupplierPayments.GeneratedPaymentMethodID) = " & 2 & ")) " & _
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
        .Sections("Section4").Controls("lblTopic").Caption = "Payment Advice For " & cmbCCPS.Text
        .Sections("Section4").Controls("lblSubTopic").Caption = "CASH PAYMENTS"
        
        .Show
        i = MsgBox("Print CASH report?", vbYesNo)
        If i = vbYes Then .PrintReport True
        i = MsgBox("Save CASH report?", vbYesNo)
        If i = vbYes Then .ExportReport , cmbCCPS.Text & " - Cash ", True, True
    End With
    With rsReport
        If .State = 1 Then .Close
        temSQL = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.AccountNo AS Account, tblSupplierPayments.Value " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where (((tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) = " & Val(cmbCCPS.BoundText) & ") And ((tblSupplierPayments.GeneratedPaymentMethodID) = " & 1 & ")) " & _
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
        .Sections("Section4").Controls("lblCenter").Caption = cmbCCPS.Text
        '.Sections("Section4").Controls("lblSubTopic").Caption = "BANK PAYMENTS"
        
        .Show
        i = MsgBox("Print BANK report?", vbYesNo)
        If i = vbYes Then .PrintReport True
        i = MsgBox("Save BANK report?", vbYesNo)
        If i = vbYes Then .ExportReport , cmbCCPS.Text & " - Bank ", True, True
    End With
    With rsReport
        If .State = 1 Then .Close
        temSQL = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.AccountNo AS Account, tblSupplierPayments.Value " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where (((tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) = " & Val(cmbCCPS.BoundText) & ") And ((tblSupplierPayments.GeneratedPaymentMethodID) = " & 3 & ")) " & _
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
        .Sections("Section4").Controls("lblTopic").Caption = "Payment Advice For " & cmbCCPS.Text
        .Sections("Section4").Controls("lblSubTopic").Caption = "CHEQUE PAYMENTS"
        .Show
        i = MsgBox("Print CHEQUE report?", vbYesNo)
        If i = vbYes Then .PrintReport True
        i = MsgBox("Save CHEQUE report?", vbYesNo)
        If i = vbYes Then .ExportReport , cmbCCPS.Text & " - Cheque ", True, True
    End With
    
'    CSetPrinter.SetPrinterAsDefault cmbPrinter.Text
'
'    If RetVal = SelectForm(cmbPaper.Text, Me.hdc) Then
'        i = MsgBox("Print Bank Slips", vbYesNo)
'        If i = vbYes Then
'            With rsReport
'                If .State = 1 Then .Close
'                temSql = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')' AS SupAccount , tblBank.Bank, tblSupplier.AccountNo AS Account, tblSupplierPayments.Value " & _
'                            "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
'                            "Where (((tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) = " & Val(cmbCCPS.BoundText) & ") And ((tblSupplierPayments.GeneratedPaymentMethodID) = " & 1 & ")) " & _
'                            "ORDER BY tblSupplier.Supplier"
'                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                If .RecordCount > 0 Then
'                    While .EOF = False
'                        Printer.CurrentX = 100
'                        Printer.CurrentY = 100
'                        Printer.Print !SupAccount
'                        Printer.CurrentX = 200
'                        Printer.CurrentY = 200
'                        Printer.Print !Bank
'                        Printer.CurrentX = 200
'                        Printer.CurrentY = 200
'                        Printer.Print !Account
'                        Printer.CurrentX = 200
'                        Printer.CurrentY = 200
'                        Printer.Print !Bank
'
'
'
'                        Printer.EndDoc
'
'                        .MoveNext
'                    Wend
'                End If
'                .Close
'            End With
'        End If
'    Else
'        MsgBox "Printer error"
'    End If
    
    
    btnClose.Enabled = True
    cmbCCPS.Enabled = True
End Sub

Private Sub btnPrintBank_Click()
    Dim RetVal As Integer
    
    Dim rsReport As New ADODB.Recordset
    Dim i As Integer

    With rsReport
        If .State = 1 Then .Close
        temSQL = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo AS Account, tblSupplierPayments.Value " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where (((tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) = " & Val(cmbCCPS.BoundText) & ") And ((tblSupplierPayments.GeneratedPaymentMethodID) = " & 1 & ")) " & _
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
        .Sections("Section4").Controls("lblCenter").Caption = cmbCCPS.Text
        '.Sections("Section4").Controls("lblSubTopic").Caption = "BANK PAYMENTS"
        
        .Show
        i = MsgBox("Print BANK report?", vbYesNo)
        If i = vbYes Then .PrintReport True
        i = MsgBox("Save BANK report?", vbYesNo)
        If i = vbYes Then .ExportReport , cmbCCPS.Text & " - Bank ", True, True
    End With

End Sub


Private Sub btnPrintCheque_Click()
    Dim RetVal As Integer
    
    Dim rsReport As New ADODB.Recordset
    Dim i As Integer
    With rsReport
        If .State = 1 Then .Close
        temSQL = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo AS Account, tblSupplierPayments.Value " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where (((tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) = " & Val(cmbCCPS.BoundText) & ") And ((tblSupplierPayments.GeneratedPaymentMethodID) = " & 3 & ")) " & _
                    "ORDER BY tblSupplier.Supplier"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtrPayAdviceCheque
        Set .DataSource = rsReport
        .Sections("Section1").Controls("txtCode").DataField = "SupplierCode"
        .Sections("Section1").Controls("txtSupplier").DataField = "SupAccount"
'        .Sections("Section1").Controls("txtAccount").DataField = "Account"
        .Sections("Section1").Controls("txtValue").DataField = "Value"
        .Sections("Section5").Controls("funValue").DataField = "Value"
        .Sections("Section4").Controls("lblName").Caption = InstitutionName
        .Sections("Section4").Controls("lblTopic").Caption = "Payment Advice For " & cmbCCPS.Text
        .Sections("Section4").Controls("lblSubTopic").Caption = "CHEQUE PAYMENTS"
        .Show
        i = MsgBox("Print CHEQUE report?", vbYesNo)
        If i = vbYes Then .PrintReport True
        i = MsgBox("Save CHEQUE report?", vbYesNo)
        If i = vbYes Then .ExportReport , cmbCCPS.Text & " - Cheque ", True, True
    End With

End Sub


Private Sub btnPrintSlips2_Click()
    Dim i As Integer
    Dim MySupplier As New clsSupplier
    
    Dim ValueX As Long
    Dim ValueY As Long
    
    Dim ValueWordX As Long
    Dim ValueWordY As Long
    
    Dim NameX As Long
    Dim NameY As Long
    
    Dim ACX1 As Long
    Dim ACX2 As Long
    Dim ACX3 As Long
    Dim ACX4 As Long
    
    Dim ACY As Long
    
    Dim TotalX As Long
    Dim TotalY As Long
    
    
    Dim DateX As Long
    Dim DateY As Long
    
    Dim ValueNoX As Long
    Dim ValueNoY As Long
    
    NameX = 2400
    NameY = 600
    
    ACX1 = 4025
    ACX2 = 7632
    ACX3 = 8640
    ACX4 = 10476
    
'    ACY = 864 - 375
    ACY = 200
    
    DateX = 8600
    DateY = 200
    
    ValueX = 2160
    ValueY = 2736 - 432
    
    ValueNoX = 9500
    ValueNoY = 3300
    
    Dim myNum As New clsNumbers
    Dim AddSpacesToNo As New clsAddSpacesInbetween
    
    
    
    CSetPrinter.SetPrinterAsDefault cmbPrinter.Text
    If SelectForm(cmbPaper.Text, Me.hdc) = 1 Then
        For i = 0 To lstBank.ListCount - 1
            
            If lstBank.Selected(i) = True Then
            
                MySupplier.ID = Val(lstBankIDs.List(i))
                
                Printer.CurrentX = NameX
                Printer.CurrentY = NameY
                If MySupplier.AccountHolder <> "" Then
                    Printer.Print MySupplier.Name & " (" & MySupplier.AccountHolder & ")"
                Else
                    Printer.Print MySupplier.Name
                End If
'                Printer.CurrentX = ValueX
'                Printer.CurrentY = ValueY - 250
'                Printer.Print myNum.NumberToWord(Val(lstBankValue.List(i)))
                
'                Printer.CurrentX = ValueX
'                Printer.CurrentY = ValueY
'                Printer.Print "And " & myNum.NumberToWord(Right(Val(lstBankValue.List(i)), 2)) & "Cents only"
                
                Printer.CurrentX = ValueNoX - Printer.TextHeight(AddSpacesToNo.AddSPace((Format(Val(lstBankValue.List(i)), "0.00"))))
                Printer.CurrentY = ValueNoY
                Printer.Print AddSpacesToNo.AddSPace(Format(Val(lstBankValue.List(i)), "0.00"))
                
                
                Printer.CurrentX = ACX1
                Printer.CurrentY = ACY
                Printer.Print AddSpacesToNo.AddSPace(MySupplier.AccountNo)
                
                
'                Printer.CurrentX = ACX2
'                Printer.CurrentY = ACY
'                Printer.Print Mid(MySupplier.AccountNo, 5, 3)
'
'                Printer.CurrentX = ACX3
'                Printer.CurrentY = ACY
'
'                Printer.Print Mid(MySupplier.AccountNo, 8, 6)
'
'                Printer.CurrentX = ACX4
'                Printer.CurrentY = ACY
'                Printer.Print Mid(MySupplier.AccountNo, 14, 1)
                
                Printer.CurrentX = DateX
                Printer.CurrentY = DateY
                Printer.Print AddSpacesToNo.AddSPace(Format(Date, "dd MM yyyy"))
                Printer.NewPage
            End If
            Printer.EndDoc
        Next i
        If i < lstBank.ListCount - 1 Then
            Printer.NewPage
        End If
        Printer.EndDoc
    Else
        MsgBox "Printer Error"
    End If
    

End Sub


Private Sub btnPrintSlips_Click()
    Dim i As Integer
    Dim MySupplier As New clsSupplier
    
    Dim MyNumberWord As String
    Dim MyNumberWordLength As Long
    Dim RupeeWord As String
    Dim CentsWord As String
    
    Dim ValueX As Long
    Dim ValueY As Long
    
    Dim ValueWordX As Long
    Dim ValueWordY As Long
    
    Dim NameX As Long
    Dim NameY As Long
    
    Dim ACX1 As Long
    Dim ACX2 As Long
    Dim ACX3 As Long
    Dim ACX4 As Long
    
    Dim ACY As Long
    
    Dim TotalX As Long
    Dim TotalY As Long
    
    
    Dim DateX As Long
    Dim DateY As Long
    
    Dim ValueNoX As Long
    Dim ValueNoY As Long
    
    NameX = 700
    NameY = 2016 - 432
    
    ACX1 = 6375
    ACX2 = 7632
    ACX3 = 8640
    ACX4 = 10476
    
    ACY = 864 - 375
    
    DateX = 1008
    DateY = 1296 - 325
    
    ValueX = 2160
    ValueY = 2736 - 432
    
    ValueNoX = 9800 - (1440 * 0.2)
    ValueNoY = 3744 - 432
    
    Dim myNum As New clsNumbers
    Dim AddSpacesToNo As New clsAddSpacesInbetween
    
    CSetPrinter.SetPrinterAsDefault cmbPrinter.Text
    If SelectForm(cmbPaper.Text, Me.hdc) = 1 Then
        For i = 0 To lstBank.ListCount - 1
            
            If lstBank.Selected(i) = True Then
            
                MySupplier.ID = Val(lstBankIDs.List(i))
                
                Printer.Font.Name = "Verdana"
                Printer.Font.Size = 11
                
                
                Printer.CurrentX = NameX
                Printer.CurrentY = NameY
                If MySupplier.AccountHolder <> "" Then
                    Printer.Print MySupplier.Name & " (" & MySupplier.AccountHolder & ")"
                Else
                    Printer.Print MySupplier.Name
                End If
                
                MyNumberWord = Format(Val(lstBankValue.List(i)), "0.00")
                MyNumberWordLength = Len(MyNumberWord)
                RupeeWord = Left(MyNumberWord, MyNumberWordLength - 3)
                CentsWord = Right(MyNumberWord, 2)
                
                Printer.CurrentX = ValueX
                Printer.CurrentY = ValueY - 250
                
                If Val(RupeeWord) <> 1 Then
                    Printer.Print myNum.NumberToWord(RupeeWord) & " Rupees"
                Else
                    Printer.Print myNum.NumberToWord(RupeeWord) & "Rupee"
                End If
                
                
                Printer.CurrentX = ValueX
                Printer.CurrentY = ValueY
                If Val(RupeeWord) <> 1 Then
                    Printer.Print "and " & myNum.NumberToWord(CentsWord) & " Cents ONLY"
                Else
                    Printer.Print "and " & myNum.NumberToWord(CentsWord) & " Cent ONLY"
                End If
                
                Printer.CurrentX = ValueNoX - Printer.TextHeight(AddSpacesToNo.AddSPace((Format(Val(lstBankValue.List(i)), "0.00"))))
                Printer.CurrentY = ValueNoY
                Printer.Print AddSpacesToNo.AddSPace(Format(Val(lstBankValue.List(i)), "0.00"))
                
                
                Printer.CurrentX = ACX1
                Printer.CurrentY = ACY
                Printer.Print AddSpacesToNo.AddSPace(MySupplier.AccountNo)
                
                Printer.CurrentX = DateX
                Printer.CurrentY = DateY
                Printer.Print AddSpacesToNo.AddSPace(Format(Date, "dd MM yyyy"))
                Printer.NewPage
            End If
            Printer.EndDoc
        Next i
        If i < lstBank.ListCount - 1 Then
            Printer.NewPage
        End If
        Printer.EndDoc
    Else
        MsgBox "Printer Error"
    End If
    
End Sub

Private Sub bttnPrintSlips2_Click()
    Dim i As Integer
    Dim MySupplier As New clsSupplier
    
    Dim ValueX As Long
    Dim ValueY As Long
    
    Dim ValueWordX As Long
    Dim ValueWordY As Long
    
    Dim NameX As Long
    Dim NameY As Long
    
    Dim ACX1 As Long
    Dim ACX2 As Long
    Dim ACX3 As Long
    Dim ACX4 As Long
    
    Dim ACY As Long
    
    Dim TotalX As Long
    Dim TotalY As Long
    
    
    Dim DateX As Long
    Dim DateY As Long
    
    Dim ValueNoX As Long
    Dim ValueNoY As Long
    
    NameX = 2400
    NameY = 600
    
    ACX1 = 4025
    ACX2 = 7632
    ACX3 = 8640
    ACX4 = 10476
    
    ACY = 200
    
    DateX = 8600
    DateY = 200
    
    ValueX = 2160 - (1440 * 0.1) - (1440 / 8)
    ValueY = 2736 - 432
    
    ValueNoX = 9500
    ValueNoY = 3300
    
    Dim myNum As New clsNumbers
    Dim AddSpacesToNo As New clsAddSpacesInbetween
    
    
    
    CSetPrinter.SetPrinterAsDefault cmbPrinter.Text
    If SelectForm(cmbPaper.Text, Me.hdc) = 1 Then
        For i = 0 To lstBank.ListCount - 1
            
            If lstBank.Selected(i) = True Then
            
                MySupplier.ID = Val(lstBankIDs.List(i))
                
                Printer.Font.Name = "Verdana"
                Printer.Font.Size = 11
                
                
                Printer.CurrentX = NameX
                Printer.CurrentY = NameY
                If MySupplier.AccountHolder <> "" Then
                    Printer.Print MySupplier.Name & " (" & MySupplier.AccountHolder & ")"
                Else
                    Printer.Print MySupplier.Name
                End If
                
                Printer.CurrentX = ValueNoX - Printer.TextHeight(AddSpacesToNo.AddSPace((Format(Val(lstBankValue.List(i)), "0.00"))))
                Printer.CurrentY = ValueNoY
                Printer.Print AddSpacesToNo.AddSPace(Format(Val(lstBankValue.List(i)), "0.00"))
                
                
                Printer.CurrentX = ACX1
                Printer.CurrentY = ACY
                Printer.Print AddSpacesToNo.AddSPace(MySupplier.AccountNo)
                
                Printer.CurrentX = DateX
                Printer.CurrentY = DateY
                Printer.Print AddSpacesToNo.AddSPace(Format(Date, "dd MM yyyy"))
                Printer.NewPage
            End If
            Printer.EndDoc
        Next i
        If i < lstBank.ListCount - 1 Then
            Printer.NewPage
        End If
        Printer.EndDoc
    Else
        MsgBox "Printer Error"
    End If
    


End Sub

Private Sub btnPrintBankC_Click()
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
        temSQL = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo AS Account, tblSupplierPayments.Value " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText) & " And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And tblSupplier.BankID = 2 " & _
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
        
        .Sections("Section4").Controls("lblCenter").Caption = cmbCC.Text
        .Sections("Section4").Controls("lblPeriod").Caption = cmbCCPS.Text
        
        .Sections("Section4").Controls("lblBank").Caption = "Commercial Bank"
        .Sections("Section4").Controls("lblBranch").Caption = "Kamburupitiya"
        .Sections("Section4").Controls("lblText").Caption = temText
        
        
        
        .Show
        i = MsgBox("Print BANK report?", vbYesNo)
        If i = vbYes Then .PrintReport True
        i = MsgBox("Save BANK report?", vbYesNo)
        If i = vbYes Then .ExportReport , cmbCCPS.Text & " - Bank ", True, True
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
        temSQL = "SELECT tblSupplier.Supplier + ' (' +  tblSupplier.AccountHolder + ')'  AS SupAccount , tblBank.Bank, tblSupplier.SupplierCode, tblSupplier.AccountNo AS Account, tblSupplierPayments.Value " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText) & " And tblSupplierPayments.GeneratedPaymentMethodID = " & 1 & " And tblSupplier.BankID = 1 " & _
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
        
        .Sections("Section4").Controls("lblCenter").Caption = cmbCC.Text
        .Sections("Section4").Controls("lblPeriod").Caption = cmbCCPS.Text
        
        .Sections("Section4").Controls("lblBank").Caption = "Sampath Bank"
        .Sections("Section4").Controls("lblBranch").Caption = "Matara"
        .Sections("Section4").Controls("lblText").Caption = temText
        
        
        
        .Show
        i = MsgBox("Print BANK report?", vbYesNo)
        If i = vbYes Then .PrintReport True
        i = MsgBox("Save BANK report?", vbYesNo)
        If i = vbYes Then .ExportReport , cmbCCPS.Text & " - Bank ", True, True
    End With

End Sub



Private Sub btnSelectedPeriod_Click()
    frmPrintSelectedPaymentsPeriod.Show
    frmPrintSelectedPaymentsPeriod.ZOrder 0
End Sub

Private Sub cmbCC_Change()
    With rsCCPS
        temSQL = "SELECT ( convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102)) AS Display, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID " & _
                    "FROM tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID " & _
                    "Where (((tblCollectingCenterPaymentSummery.PaymentsProcessStarted) = 1) AND ((tblCollectingCenterPaymentSummery.CollectingcenterID) = " & Val(cmbCC.BoundText) & ") AND ((tblCollectingCenterPaymentSummery.PaymentsPrinted) = 0 ) ) " & _
                    "ORDER BY convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) "
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCCPS
        Set .RowSource = rsCCPS
        .ListField = "Display"
        .BoundColumn = "CollectingCenterPaymentSummeryID"
    End With
End Sub

Private Sub cmbCCPS_Change()
    If IsNumeric(cmbCCPS.BoundText) = False Then Exit Sub
    FillLists
    CalculateValues
    BankEnable
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
End Sub

Private Sub FillCombos()
'    With rsCCPS
'        temSql = "SELECT (tblCollectingCenter.CollectingCenter + ' - From ' +  convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) + ' To ' +  convert(varchar, tblCollectingCenterPaymentSummery.ToDate, 102)) AS Display, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID " & _
'                    "FROM tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID " & _
'                    "Where (((tblCollectingCenterPaymentSummery.PaymentsProcessStarted) = True) AND ((tblCollectingCenterPaymentSummery.PaymentsPrinted) = 0 ) ) " & _
'                    "ORDER BY (tblCollectingCenter.CollectingCenter + ' - From ' +  convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) + ' To ' +  convert(varchar, tblCollectingCenterPaymentSummery.ToDate, 102))"
'        If .State = 1 Then .Close
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'    With cmbCCPS
'        Set .RowSource = rsCCPS
'        .ListField = "Display"
'        .BoundColumn = "CollectingCenterPaymentSummeryID"
'    End With
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
            temSQL = "SELECT tblSupplier.Supplier, tblSupplier.SupplierID, tblSupplierPayments.SupplierPaymentsID , tblSupplierPayments.Value ,tblSupplierPayments.GeneratedPaymentMethodID " & _
                        "FROM tblSupplierPayments LEFT JOIN tblSupplier ON tblSupplierPayments.SupplierID = tblSupplier.SupplierID " & _
                        "Where (((tblSupplierPayments.CollectingCenterPaymentSummeryID) = " & Val(cmbCCPS.BoundText) & ")) " & _
                        "ORDER BY tblSupplier.Supplier"
        ElseIf optCom.Value = True Then
            temSQL = "SELECT tblSupplier.Supplier, tblSupplier.SupplierID, tblSupplierPayments.SupplierPaymentsID , tblSupplierPayments.Value ,tblSupplierPayments.GeneratedPaymentMethodID " & _
                        "FROM tblSupplierPayments LEFT JOIN tblSupplier ON tblSupplierPayments.SupplierID = tblSupplier.SupplierID " & _
                        "Where tblSupplierPayments.CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText) & " AND tblSupplier.BankID = 2 " & _
                        "ORDER BY tblSupplier.Supplier"
        ElseIf optSB.Value = True Then
            temSQL = "SELECT tblSupplier.Supplier, tblSupplier.SupplierID, tblSupplierPayments.SupplierPaymentsID , tblSupplierPayments.Value ,tblSupplierPayments.GeneratedPaymentMethodID " & _
                        "FROM tblSupplierPayments LEFT JOIN tblSupplier ON tblSupplierPayments.SupplierID = tblSupplier.SupplierID " & _
                        "Where tblSupplierPayments.CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText) & " AND tblSupplier.BankID = 1 " & _
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
                SupplierPaymentsID(i) = !SupplierPaymentsID
                lstAllSupplierIDs.AddItem !SupplierPaymentsID
                lstAllSuppliers.AddItem !Supplier
                lstAllVaue.AddItem !Value
                SupplierValue(i) = !Value
                
                If IsNull(!GeneratedPaymentMethodID) = False Then
                    Select Case !GeneratedPaymentMethodID
                        Case 1:
                                lstBank.AddItem !Supplier
                                lstBankValue.AddItem Right(Space(20) & Format(!Value, "0.00"), 20)
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

Private Sub optAll_Click()
    If IsNumeric(cmbCCPS.BoundText) = False Then Exit Sub
    FillLists
    CalculateValues
    BankEnable
End Sub

Private Sub optCom_Click()
    If IsNumeric(cmbCCPS.BoundText) = False Then Exit Sub
    FillLists
    CalculateValues
    BankEnable
End Sub

Private Sub optSB_Click()
    If IsNumeric(cmbCCPS.BoundText) = False Then Exit Sub
    FillLists
    CalculateValues
    BankEnable
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
