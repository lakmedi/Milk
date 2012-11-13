VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmGenerateIndividualPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Approve Payments"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
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
   ScaleWidth      =   10905
   Begin VB.TextBox txtValues 
      Height          =   1455
      Left            =   9120
      MultiLine       =   -1  'True
      TabIndex        =   34
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton btnToText 
      Caption         =   "Command1"
      Height          =   375
      Left            =   7440
      TabIndex        =   33
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Show All"
      Height          =   240
      Left            =   5880
      TabIndex        =   32
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   8520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   8880
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ListBox lstBankValue 
      Height          =   5100
      Left            =   9600
      MultiSelect     =   2  'Extended
      TabIndex        =   18
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ListBox lstSelectedValue 
      Height          =   5100
      Left            =   4320
      MultiSelect     =   2  'Extended
      TabIndex        =   17
      Top             =   1800
      Width           =   1335
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9480
      TabIndex        =   6
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
      Height          =   5100
      Left            =   6600
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   1800
      Width           =   4215
   End
   Begin MSDataListLib.DataCombo cmbCCPS 
      Height          =   360
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnBankAdd 
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   3600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   ">"
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
   Begin btButtonEx.ButtonEx btnBankRemove 
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   4080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "<"
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Top             =   8760
      Visible         =   0   'False
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
   Begin btButtonEx.ButtonEx btnAll 
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   6960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Select All"
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
   Begin btButtonEx.ButtonEx btnNone 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   15
      Top             =   6960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Select None"
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
   Begin VB.ListBox lstPaymentMethod 
      Height          =   5100
      Left            =   2880
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstAllSuppliers 
      Height          =   5100
      Left            =   1560
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lstAllSupplierIDs 
      Height          =   4920
      Left            =   4920
      Style           =   1  'Checkbox
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstSelectedSupplierIDs 
      Height          =   4920
      Left            =   5280
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstBankIDs 
      Height          =   1740
      Left            =   10440
      MultiSelect     =   2  'Extended
      TabIndex        =   12
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   8040
      TabIndex        =   16
      Top             =   8760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Save"
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
   Begin VB.ListBox lstAllVaue 
      Height          =   5100
      Left            =   4800
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstSelectedSuppliers 
      Height          =   5100
      Left            =   720
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   1800
      Width           =   3495
   End
   Begin MSDataListLib.DataCombo cmbCC 
      Height          =   360
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label5 
      Caption         =   "Paper"
      Height          =   255
      Left            =   600
      TabIndex        =   31
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Printer"
      Height          =   255
      Left            =   600
      TabIndex        =   30
      Top             =   8640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Payment Summery"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblSelectedValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   3360
      TabIndex        =   25
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label Label12 
      Caption         =   "Total Selected Value"
      Height          =   255
      Left            =   600
      TabIndex        =   24
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label lblAllValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   3360
      TabIndex        =   23
      Top             =   7440
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "Total Value"
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   7440
      Width           =   2175
   End
   Begin VB.Label lblBankValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   8640
      TabIndex        =   21
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Total Bank Value"
      Height          =   255
      Left            =   6600
      TabIndex        =   20
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "To Bank"
      Height          =   255
      Left            =   6600
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "frmGenerateIndividualPayments"
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
    
Private Sub btnAll_Click()
    Dim i As Integer
    For i = 0 To lstSelectedSupplierIDs.ListCount - 1
        lstSelectedSupplierIDs.Selected(i) = True
        lstSelectedSuppliers.Selected(i) = True
        lstSelectedValue.Selected(i) = True
    Next
    CalculateValues
End Sub

Private Sub btnBankAdd_Click()
    Dim i As Integer
    Dim ii As Integer
    Dim AlreadySelected As Boolean
    For i = 0 To lstSelectedSupplierIDs.ListCount - 1
        If lstSelectedSuppliers.Selected(i) = True Then
            AlreadySelected = False

            For ii = 0 To lstBankIDs.ListCount - 1
                If Val(lstBankIDs.List(ii)) = Val(lstSelectedSupplierIDs.List(i)) Then
                    AlreadySelected = True
                    MsgBox "Supplier " & lstSelectedSuppliers.List(i) & " is already listed under Bank"
                End If
            Next
            If AlreadySelected = False Then
                lstBank.AddItem lstSelectedSuppliers.List(i)
                lstBankIDs.AddItem lstSelectedSupplierIDs.List(i)
                lstBankValue.AddItem lstSelectedValue.List(i)
            End If
        End If
    Next
    CalculateValues
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

Private Sub btnBankRemove_Click()
    Dim i As Integer
    For i = lstBank.ListCount - 1 To 0 Step -1
        If lstBank.Selected(i) = True Then
            lstBank.RemoveItem (i)
            lstBankIDs.RemoveItem (i)
            lstBankValue.RemoveItem (i)
        End If
    Next
    CalculateValues
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
 
Private Sub btnNone_Click()
    Dim i As Integer
    For i = 0 To lstSelectedSupplierIDs.ListCount - 1
        lstSelectedSupplierIDs.Selected(i) = False
        lstSelectedSuppliers.Selected(i) = False
        lstSelectedValue.Selected(i) = False
    Next
End Sub

Private Sub btnPrint_Click()
    Dim RetVal As Integer
    
    Dim rsReport As New ADODB.Recordset
    Dim i As Integer
    With rsReport
        If .State = 1 Then .Close
        temSQL = "SELECT tblSupplier.Supplier, tblBank.Bank, tblSupplier.AccountNo + ' ' +  tblSupplier.AccountHolder AS Account, tblSupplierPayments.Value " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where (((tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) = " & Val(cmbCCPS.BoundText) & ") And ((tblSupplierPayments.PrintedPaymentMethodID) = " & 2 & ")) " & _
                    "ORDER BY tblSupplier.Supplier"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtrPayAdvice
        Set .DataSource = rsReport
        .Sections("Section1").Controls("txtSupplier").DataField = "Supplier"
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
        temSQL = "SELECT tblSupplier.Supplier, tblBank.Bank, tblSupplier.AccountNo + ' ' +  tblSupplier.AccountHolder AS Account, tblSupplierPayments.Value " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where (((tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) = " & Val(cmbCCPS.BoundText) & ") And ((tblSupplierPayments.PrintedPaymentMethodID) = " & 1 & ")) " & _
                    "ORDER BY tblSupplier.Supplier"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtrPayAdviceBank
        Set .DataSource = rsReport
        .Sections("Section1").Controls("txtCode").DataField = "SupplierCode"
        .Sections("Section1").Controls("txtSupplier").DataField = "Supplier"
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
        temSQL = "SELECT tblSupplier.Supplier, tblBank.Bank, tblSupplier.AccountNo + ' ' +  tblSupplier.AccountHolder AS Account, tblSupplierPayments.Value " & _
                    "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                    "Where (((tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) = " & Val(cmbCCPS.BoundText) & ") And ((tblSupplierPayments.PrintedPaymentMethodID) = " & 3 & ")) " & _
                    "ORDER BY tblSupplier.Supplier"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtrPayAdvice
        Set .DataSource = rsReport
        .Sections("Section1").Controls("txtSupplier").DataField = "Supplier"
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
    
    CSetPrinter.SetPrinterAsDefault cmbPrinter.Text
    If RetVal = SelectForm(cmbPaper.Text, Me.hdc) Then
        i = MsgBox("Print Bank Slips", vbYesNo)
        If i = vbYes Then
            With rsReport
                If .State = 1 Then .Close
                temSQL = "SELECT tblSupplier.Supplier, tblBank.Bank, tblSupplier.AccountNo + ' ' +  tblSupplier.AccountHolder AS Account, tblSupplierPayments.Value " & _
                            "FROM ((tblSupplier RIGHT JOIN tblSupplierPayments ON tblSupplier.SupplierID = tblSupplierPayments.SupplierID) RIGHT JOIN (tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID) ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblBank ON tblSupplier.BankID = tblBank.BankID " & _
                            "Where (((tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) = " & Val(cmbCCPS.BoundText) & ") And ((tblSupplierPayments.PrintedPaymentMethodID) = " & 1 & ")) " & _
                            "ORDER BY tblSupplier.Supplier"
                .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    While .EOF = False
                        Printer.CurrentX = 100
                        Printer.CurrentY = 100
                        Printer.Print !Supplier
                        Printer.CurrentX = 200
                        Printer.CurrentY = 200
                        Printer.Print !Bank
                        Printer.CurrentX = 200
                        Printer.CurrentY = 200
                        Printer.Print !Account
                        Printer.CurrentX = 200
                        Printer.CurrentY = 200
                        Printer.Print !Bank
                        
                        
                        
                        Printer.EndDoc
                        
                        .MoveNext
                    Wend
                End If
                .Close
            End With
        End If
    Else
        MsgBox "Printer error"
    End If
    
    
    btnClose.Enabled = True
    cmbCCPS.Enabled = True
End Sub

Private Sub btnSave_Click()
    Dim tr As Long
    If lstBank.ListCount <> lstAllSupplierIDs.ListCount Then
        tr = MsgBox("Still some suppliers are to be added to a payment method. Is this intentional?", vbYesNo)
        If vbYes = True Then
        
        Else
            AddRemaining
            Exit Sub
        End If
    End If
    Dim i As Integer
    Dim rsSupplierPayments As New ADODB.Recordset
    With rsSupplierPayments
        If .State = 1 Then .Close
        temSQL = "Select * from tblCollectingCenterPaymentSummery where CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !PaymentsProcessStarted = True
            !PaymentsProcessStartedUserID = UserID
            !PaymentsProcessStartedDate = Date
            !PaymentsProcessStartedTime = Time
            .Update
        End If
        .Close
    End With
    For i = 0 To lstBankIDs.ListCount
        With rsSupplierPayments
            If .State = 1 Then .Close
            temSQL = "Select * from tblSupplierPayments where SupplierPaymentsID = " & Val(lstBankIDs.List(i)) & " And CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText)
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !GeneratedPaymentMethodID = 1
                !Generated = True
                !GeneratedUserID = UserID
                !GeneratedDate = Date
                !GeneratedTime = Time
                .Update
            End If
            .Close
        End With
    Next i
    btnSave.Enabled = False
    cmbCCPS.Enabled = False
End Sub


Private Sub btnToText_Click()
    Dim i As Integer
    For i = 0 To lstAllVaue.ListCount - 1
        txtValues.Text = txtValues.Text & vbNewLine & lstAllVaue.List(i)
    Next i
End Sub

Private Sub chkAll_Click()
    If chkAll.Value = 1 Then
    
        With rsCCPS
            If IsNumeric(cmbCC.BoundText) = False Then
                temSQL = "SELECT tblCollectingCenter.CollectingCenter + ' - On ' + convert(varchar, tblCollectingCenterPaymentSummery.FromDate,102) AS Display, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID " & _
                            "FROM tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID " & _
                            " " & _
                            "ORDER BY tblCollectingCenter.CollectingCenter + ' - On ' + convert(varchar, tblCollectingCenterPaymentSummery.FromDate 102) "
            Else
                temSQL = "SELECT tblCollectingCenter.CollectingCenter + ' - On ' + convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) AS Display, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID " & _
                            "FROM tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID " & _
                            "Where (((tblCollectingCenterPaymentSummery.CollectingCenterID)=" & Val(cmbCC.BoundText) & " )) " & _
                            "ORDER BY tblCollectingCenter.CollectingCenter + ' - From ' + convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) "
            
            End If
            If .State = 1 Then .Close
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        End With
        With cmbCCPS
            Set .RowSource = rsCCPS
            .ListField = "Display"
            .BoundColumn = "CollectingCenterPaymentSummeryID"
        End With
    
    
    Else
        
        With rsCCPS
            If IsNumeric(cmbCC.BoundText) = False Then
                temSQL = "SELECT tblCollectingCenter.CollectingCenter + ' - On ' + convert(varchar, tblCollectingCenterPaymentSummery.FromDate,102) AS Display, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID " & _
                            "FROM tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID " & _
                            " " & _
                            "ORDER BY tblCollectingCenter.CollectingCenter + ' - On ' + convert(varchar, tblCollectingCenterPaymentSummery.FromDate 102) "
            Else
                temSQL = "SELECT tblCollectingCenter.CollectingCenter + ' - On ' + convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) AS Display, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID " & _
                            "FROM tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID " & _
                            "Where (((tblCollectingCenterPaymentSummery.CollectingCenterID)=" & Val(cmbCC.BoundText) & " )) " & _
                            "ORDER BY tblCollectingCenter.CollectingCenter + ' - From ' + convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) "
            
            End If
            If .State = 1 Then .Close
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        End With
        With cmbCCPS
            Set .RowSource = rsCCPS
            .ListField = "Display"
            .BoundColumn = "CollectingCenterPaymentSummeryID"
        End With
        
    
    End If
End Sub

Private Sub cmbCC_Change()
    With rsCCPS
        If IsNumeric(cmbCC.BoundText) = False Then
            temSQL = "SELECT tblCollectingCenter.CollectingCenter + ' - On ' + convert(varchar, tblCollectingCenterPaymentSummery.FromDate,102) AS Display, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID " & _
                        "FROM tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID " & _
                        "Where (((tblCollectingCenterPaymentSummery.PaymentsProcessStarted) = 0 )) " & _
                        "ORDER BY tblCollectingCenter.CollectingCenter + ' - On ' + convert(varchar, tblCollectingCenterPaymentSummery.FromDate 102) "
        Else
            temSQL = "SELECT tblCollectingCenter.CollectingCenter + ' - On ' + convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) AS Display, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID " & _
                        "FROM tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID " & _
                        "Where (((tblCollectingCenterPaymentSummery.PaymentsProcessStarted) = 0 ) AND ((tblCollectingCenterPaymentSummery.CollectingCenterID)=" & Val(cmbCC.BoundText) & " )) " & _
                        "ORDER BY tblCollectingCenter.CollectingCenter + ' - From ' + convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) "
        
        End If
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCCPS
        Set .RowSource = rsCCPS
        .ListField = "Display"
        .BoundColumn = "CollectingCenterPaymentSummeryID"
    End With
End Sub

Private Sub cmbCC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbCCPS.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbCC.Text = Empty
    End If
End Sub

Private Sub cmbCCPS_Change()
    If IsNumeric(cmbCCPS.BoundText) = False Then Exit Sub
    FillLists
    FillSelected
    btnPrint.Enabled = False
    btnSave.Enabled = True
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
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, "Options", "Printer", "")
    Call ListPapers
    cmbPaper.Text = GetSetting(App.EXEName, "Options", "Paper", "")
    
    Select Case UserAuthorityLevel
    
    
    Case Authority.PowerUser '4
    btnSave.Visible = False

    Case Authority.SuperUser '5
    btnSave.Visible = True
    
    Case Authority.Administrator '6
    btnSave.Visible = True

    Case Else
    
    End Select

End Sub

Private Sub FillCombos()
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCC, "CollectingCenter", True
    
    With rsCCPS
        temSQL = "SELECT (tblCollectingCenter.CollectingCenter  + ' - From ' + convert(varchar , tblCollectingCenterPaymentSummery.FromDate,102) + ' To ' + convert(varchar, tblCollectingCenterPaymentSummery.ToDate,102 )) AS Display, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID " & _
                    "FROM tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID " & _
                    "Where (((tblCollectingCenterPaymentSummery.PaymentsProcessStarted) = 0 )) " & _
                    "ORDER BY (tblCollectingCenter.CollectingCenter + ' - From ' + convert(varchar, tblCollectingCenterPaymentSummery.FromDate,102) + ' To ' + convert(varchar, tblCollectingCenterPaymentSummery.ToDate, 102 ))"
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCCPS
        Set .RowSource = rsCCPS
        .ListField = "Display"
        .BoundColumn = "CollectingCenterPaymentSummeryID"
    End With
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


Private Sub AddRemaining()
    
    lstSelectedSupplierIDs.Clear
    lstSelectedSuppliers.Clear
    lstSelectedValue.Clear
    
    Dim i As Integer
    Dim ii As Integer
    Dim AlreadySelected As Boolean
    For i = 0 To lstAllSupplierIDs.ListCount - 1
        AlreadySelected = False
        For ii = 0 To lstBankIDs.ListCount - 1
            If Val(lstBankIDs.List(ii)) = Val(lstAllSupplierIDs.List(i)) Then
                AlreadySelected = True
            End If
        Next
        If AlreadySelected = False Then
            lstSelectedSuppliers.AddItem lstAllSuppliers.List(i)
            lstSelectedSupplierIDs.AddItem lstAllSupplierIDs.List(i)
            lstSelectedValue.AddItem lstAllVaue.List(i)
        End If
    Next
    CalculateValues
    
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
        temSQL = "SELECT tblSupplier.Supplier, tblSupplier.SupplierID, tblSupplierPayments.SupplierPaymentsID , tblSupplierPayments.Value ,tblSupplierPayments.GeneratedPaymentMethodID " & _
                    "FROM tblSupplierPayments LEFT JOIN tblSupplier ON tblSupplierPayments.SupplierID = tblSupplier.SupplierID " & _
                    "Where (((tblSupplierPayments.CollectingCenterPaymentSummeryID) = " & Val(cmbCCPS.BoundText) & ")) " & _
                    "ORDER BY tblSupplier.Supplier"
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
                
                
                lstAllVaue.AddItem Format(!Value, "0.00")
                SupplierValue(i) = Format(!Value, "0.00")
                
'                lstAllVaue.AddItem !Value
'                SupplierValue(i) = !Value
                
                
                If IsNull(!GeneratedPaymentMethodID) = False Then
                    lstPaymentMethod.AddItem !GeneratedPaymentMethodID
                    PaymentMethodID(i) = !GeneratedPaymentMethodID
                Else
                    lstPaymentMethod.AddItem 0
                    PaymentMethodID(i) = 0
                End If
                i = i + 1
                .MoveNext
            Wend
        End If
        .Close
    End With
End Sub

Private Sub FillSelected()
    Dim i As Integer
    
    lstSelectedSupplierIDs.Clear
    lstSelectedSuppliers.Clear
    lstSelectedValue.Clear
    
    For i = 0 To lstAllSupplierIDs.ListCount - 1
        lstAllSupplierIDs.Selected(i) = False
        lstAllSuppliers.Selected(i) = False
        lstAllVaue.Selected(i) = False
        SupplierSelected(i) = False
    Next i
    
    For i = 0 To lstAllSupplierIDs.ListCount - 1
        lstAllSupplierIDs.Selected(i) = True
        lstAllSuppliers.Selected(i) = True
        lstAllVaue.Selected(i) = True
        SupplierSelected(i) = True
    Next
    
    For i = 0 To lstAllSupplierIDs.ListCount - 1
        If lstAllSupplierIDs.Selected(i) = True Then
            lstSelectedSupplierIDs.AddItem lstAllSupplierIDs.List(i)
            lstSelectedSuppliers.AddItem lstAllSuppliers.List(i)
            lstSelectedValue.AddItem lstAllVaue.List(i)
        End If
    Next i
    CalculateValues
End Sub

Private Sub CalculateValues()
    Dim i As Integer
    Dim temValue As Double
    For i = 0 To lstAllVaue.ListCount - 1
        temValue = temValue + Val(lstAllVaue.List(i))
    Next i
    lblAllValue.Caption = Format(temValue, "#,##0.00")
    
    temValue = 0
    For i = 0 To lstSelectedValue.ListCount - 1
        temValue = temValue + Val(lstSelectedValue.List(i))
    Next i
    lblSelectedValue.Caption = Format(temValue, "#,##0.00")
    
'    TemValue = 0
'    For i = 0 To lstSelectedValue.ListCount - 1
'        TemValue = TemValue + Val(lstCashValue.List(i))
'    Next i
'    lblCashValue.Caption = Format(TemValue, "#,##0.00")
'
'    TemValue = 0
'    For i = 0 To lstChequeValue.ListCount - 1
'        TemValue = TemValue + Val(lstChequeValue.List(i))
'    Next i
'    lblChequeValue.Caption = Format(TemValue, "#,##0.00")
    
    temValue = 0
    For i = 0 To lstBankValue.ListCount - 1
        temValue = temValue + Val(lstBankValue.List(i))
    Next i
    lblBankValue.Caption = Format(temValue, "#,##0.00")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    SaveSetting App.EXEName, "Options", "Printer", cmbPrinter.Text
'    SaveSetting App.EXEName, "Options", "Paper", cmbPaper.Text
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
