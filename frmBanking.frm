VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmBanking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Approve Collecting Center Payments"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   8880
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   8520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ListBox lstBankValue 
      Height          =   1740
      Left            =   9600
      MultiSelect     =   2  'Extended
      TabIndex        =   39
      Top             =   6240
      Width           =   1215
   End
   Begin VB.ListBox lstChequeValue 
      Height          =   1740
      Left            =   9600
      MultiSelect     =   2  'Extended
      TabIndex        =   38
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ListBox lstCashValue 
      Height          =   1740
      Left            =   9600
      MultiSelect     =   2  'Extended
      TabIndex        =   37
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox lstSelectedValue 
      Height          =   5820
      Left            =   4320
      MultiSelect     =   2  'Extended
      TabIndex        =   36
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox chkOther 
      Caption         =   "Other"
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   840
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkCheque 
      Caption         =   "Cheque"
      Height          =   255
      Left            =   3120
      TabIndex        =   18
      Top             =   840
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkBank 
      Caption         =   "Bank"
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   840
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkCash 
      Caption         =   "Cash"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   840
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9480
      TabIndex        =   11
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
   Begin btButtonEx.ButtonEx btnCashAdd 
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   960
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
   Begin VB.ListBox lstBank 
      Height          =   1740
      Left            =   6600
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   6240
      Width           =   4215
   End
   Begin VB.ListBox lstCheque 
      Height          =   1740
      Left            =   6600
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   3360
      Width           =   4215
   End
   Begin VB.ListBox lstCash 
      Height          =   1740
      Left            =   6600
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   480
      Width           =   4215
   End
   Begin MSDataListLib.DataCombo cmbCCPS 
      Height          =   360
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnCashRemove 
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   1440
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
   Begin btButtonEx.ButtonEx btnChequeAdd 
      Height          =   375
      Left            =   6000
      TabIndex        =   7
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
   Begin btButtonEx.ButtonEx btnChequeRemove 
      Height          =   375
      Left            =   6000
      TabIndex        =   8
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
   Begin btButtonEx.ButtonEx btnBankAdd 
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   6480
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
      Left            =   6000
      TabIndex        =   10
      Top             =   6960
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
      TabIndex        =   12
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
      TabIndex        =   27
      Top             =   7080
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
      TabIndex        =   28
      Top             =   7080
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
   Begin btButtonEx.ButtonEx btnCashAll 
      Height          =   255
      Left            =   9480
      TabIndex        =   29
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "All"
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
   Begin btButtonEx.ButtonEx btnCashNone 
      Height          =   255
      Left            =   10200
      TabIndex        =   30
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "None"
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
   Begin btButtonEx.ButtonEx btnBankAll 
      Height          =   255
      Left            =   9120
      TabIndex        =   31
      Top             =   5880
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "All"
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
   Begin btButtonEx.ButtonEx btnBankNone 
      Height          =   255
      Left            =   9840
      TabIndex        =   32
      Top             =   5880
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "None"
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
   Begin btButtonEx.ButtonEx btnChequeAll 
      Height          =   255
      Left            =   9480
      TabIndex        =   33
      Top             =   3000
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "All"
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
   Begin btButtonEx.ButtonEx btnChequeNone 
      Height          =   255
      Left            =   10200
      TabIndex        =   34
      Top             =   3000
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "None"
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
      Height          =   6060
      Left            =   2880
      TabIndex        =   22
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstAllSuppliers 
      Height          =   6060
      Left            =   1560
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lstAllSupplierIDs 
      Height          =   6270
      Left            =   4920
      Style           =   1  'Checkbox
      TabIndex        =   21
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstSelectedSupplierIDs 
      Height          =   6270
      Left            =   5280
      Style           =   1  'Checkbox
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstCashIDs 
      Height          =   1740
      Left            =   9960
      MultiSelect     =   2  'Extended
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstChequeIDs 
      Height          =   1740
      Left            =   9960
      MultiSelect     =   2  'Extended
      TabIndex        =   25
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstBankIDs 
      Height          =   1740
      Left            =   9960
      MultiSelect     =   2  'Extended
      TabIndex        =   24
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   8040
      TabIndex        =   35
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
      Height          =   6060
      Left            =   4800
      TabIndex        =   40
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstSelectedSuppliers 
      Height          =   5820
      Left            =   720
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   1200
      Width           =   3495
   End
   Begin btButtonEx.ButtonEx btnPrintSlips 
      Height          =   255
      Left            =   9120
      TabIndex        =   55
      Top             =   8040
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Print Slips"
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
   Begin VB.Label Label7 
      Caption         =   "Bank Slip Printer"
      Height          =   255
      Left            =   600
      TabIndex        =   52
      Top             =   8640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Bank Slip Paper"
      Height          =   255
      Left            =   600
      TabIndex        =   51
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblSelectedValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   3360
      TabIndex        =   50
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label Label12 
      Caption         =   "Total Selected Value"
      Height          =   255
      Left            =   600
      TabIndex        =   49
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label lblAllValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   3360
      TabIndex        =   48
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label Label10 
      Caption         =   "Total Value"
      Height          =   255
      Left            =   600
      TabIndex        =   47
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label lblBankValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   8640
      TabIndex        =   46
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Total Bank Value"
      Height          =   255
      Left            =   6600
      TabIndex        =   45
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Label lblChequeValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   8640
      TabIndex        =   44
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Total Cheque Value"
      Height          =   255
      Left            =   6600
      TabIndex        =   43
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label lblCashValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   8640
      TabIndex        =   42
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Total Cash Value"
      Height          =   255
      Left            =   6600
      TabIndex        =   41
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Bank"
      Height          =   255
      Left            =   6600
      TabIndex        =   15
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Cheque"
      Height          =   255
      Left            =   6600
      TabIndex        =   14
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Cash"
      Height          =   255
      Left            =   6600
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmBanking"
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
            For ii = 0 To lstCashIDs.ListCount - 1
                If Val(lstCashIDs.List(ii)) = Val(lstSelectedSupplierIDs.List(i)) Then
                    AlreadySelected = True
                    MsgBox "Supplier " & lstSelectedSuppliers.List(i) & " is already listed under Cash"
                End If
            Next
            For ii = 0 To lstChequeIDs.ListCount - 1
                If Val(lstChequeIDs.List(ii)) = Val(lstSelectedSupplierIDs.List(i)) Then
                    AlreadySelected = True
                    MsgBox "Supplier " & lstSelectedSuppliers.List(i) & " is already listed under Cheque"
                End If
            Next
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

Private Sub btnCashAdd_Click()
    Dim i As Integer
    Dim ii As Integer
    Dim AlreadySelected As Boolean
    For i = 0 To lstSelectedSupplierIDs.ListCount - 1
        If lstSelectedSuppliers.Selected(i) = True Then
            AlreadySelected = False
            For ii = 0 To lstCashIDs.ListCount - 1
                If Val(lstCashIDs.List(ii)) = Val(lstSelectedSupplierIDs.List(i)) Then
                    AlreadySelected = True
                    MsgBox "Supplier " & lstSelectedSuppliers.List(i) & " is already listed under Cash"
                End If
            Next
            For ii = 0 To lstChequeIDs.ListCount - 1
                If Val(lstChequeIDs.List(ii)) = Val(lstSelectedSupplierIDs.List(i)) Then
                    AlreadySelected = True
                    MsgBox "Supplier " & lstSelectedSuppliers.List(i) & " is already listed under Cheque"
                End If
            Next
            For ii = 0 To lstBankIDs.ListCount - 1
                If Val(lstBankIDs.List(ii)) = Val(lstSelectedSupplierIDs.List(i)) Then
                    AlreadySelected = True
                    MsgBox "Supplier " & lstSelectedSuppliers.List(i) & " is already listed under Bank"
                End If
            Next
            If AlreadySelected = False Then
                lstCash.AddItem lstSelectedSuppliers.List(i)
                lstCashIDs.AddItem lstSelectedSupplierIDs.List(i)
                lstCashValue.AddItem lstSelectedValue.List(i)
            End If
        End If
    Next
    CalculateValues
End Sub

Private Sub btnCashAll_Click()
    Dim i As Integer
    For i = 0 To lstCashIDs.ListCount - 1
        lstCash.Selected(i) = True
        lstCashIDs.Selected(i) = True
        lstCashValue.Selected(i) = True
    Next i
    CalculateValues
End Sub

Private Sub btnCashNone_Click()
    Dim i As Integer
    For i = 0 To lstCashIDs.ListCount - 1
        lstCash.Selected(i) = False
        lstCashIDs.Selected(i) = False
        lstCashValue.Selected(i) = False
    Next i
    CalculateValues
End Sub

Private Sub btnCashRemove_Click()
    Dim i As Integer
    For i = lstCash.ListCount - 1 To 0 Step -1
        If lstCash.Selected(i) = True Then
            lstCash.RemoveItem (i)
            lstCashIDs.RemoveItem (i)
            lstCashValue.RemoveItem (i)
        End If
    Next
    CalculateValues
End Sub

Private Sub btnChequeAdd_Click()
    Dim i As Integer
    Dim ii As Integer
    Dim AlreadySelected As Boolean
    For i = 0 To lstSelectedSupplierIDs.ListCount - 1
        If lstSelectedSuppliers.Selected(i) = True Then
            AlreadySelected = False
            For ii = 0 To lstCashIDs.ListCount - 1
                If Val(lstCashIDs.List(ii)) = Val(lstSelectedSupplierIDs.List(i)) Then
                    AlreadySelected = True
                    MsgBox "Supplier " & lstSelectedSuppliers.List(i) & " is already listed under Cash"
                End If
            Next
            For ii = 0 To lstChequeIDs.ListCount - 1
                If Val(lstChequeIDs.List(ii)) = Val(lstSelectedSupplierIDs.List(i)) Then
                    AlreadySelected = True
                    MsgBox "Supplier " & lstSelectedSuppliers.List(i) & " is already listed under Cheque"
                End If
            Next
            For ii = 0 To lstBankIDs.ListCount - 1
                If Val(lstBankIDs.List(ii)) = Val(lstSelectedSupplierIDs.List(i)) Then
                    AlreadySelected = True
                    MsgBox "Supplier " & lstSelectedSuppliers.List(i) & " is already listed under Bank"
                End If
            Next
            If AlreadySelected = False Then
                lstCheque.AddItem lstSelectedSuppliers.List(i)
                lstChequeIDs.AddItem lstSelectedSupplierIDs.List(i)
                lstChequeValue.AddItem lstSelectedValue.List(i)
            End If
        End If
    Next
    CalculateValues
End Sub

Private Sub btnChequeAll_Click()
    Dim i As Integer
    For i = 0 To lstChequeIDs.ListCount - 1
        lstCheque.Selected(i) = True
        lstChequeIDs.Selected(i) = True
        lstAllVaue.Selected(i) = True
    Next i
End Sub

Private Sub btnChequeNone_Click()
    Dim i As Integer
    For i = 0 To lstChequeIDs.ListCount - 1
        lstCheque.Selected(i) = False
        lstChequeIDs.Selected(i) = False
        lstChequeValue.Selected(i) = False
    Next i
End Sub

Private Sub btnChequeRemove_Click()
    Dim i As Integer
    For i = lstCheque.ListCount - 1 To 0 Step -1
        If lstCheque.Selected(i) = True Then
            lstCheque.RemoveItem (i)
            lstChequeIDs.RemoveItem (i)
            lstChequeValue.RemoveItem (i)
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

Private Sub btnPrintSlips_Click()
    Dim i As Integer
    Dim MySupplier As New clsSupplier
    
    Dim ValueX As Long
    Dim ValueY As Long
    
    Dim ValueWordX As Long
    Dim ValueWordY As Long
    
    Dim NameX As Long
    Dim NameY As Long
    
    Dim ACX As Long
    Dim ACY As Long
    
    Dim TotalX As Long
    Dim TotalY As Long
    
    
    Dim DateX As Long
    Dim DateY As Long
    
    
    NameX = 1500
    NameY = 2300
    
    ACX = 1100
    ACY = 800
    
    DateX = 11000
    DateY = 800
    
    ValueX = 1700
    ValueY = 5500
    
    CSetPrinter.SetPrinterAsDefault cmbPrinter.Text
    If SelectForm(cmbPaper.Text, Me.hdc) = 1 Then
        For i = 0 To lstBank.ListCount - 1
            MySupplier.ID = Val(lstBankIDs.List(i))
            
            Printer.CurrentX = NameX
            Printer.CurrentY = NameY
            Printer.Print MySupplier.Name
            
            Printer.CurrentX = ValueX
            Printer.CurrentY = ValueY
            Printer.Print lstBankValue.List(i)
            
            Printer.CurrentX = ACX
            Printer.CurrentY = ACY
            Printer.Print MySupplier.AccountNo
            
            Printer.CurrentX = DateX
            Printer.CurrentY = DateY
            Printer.Print Format(Date, "dd MM yyyy")
            
        Next i
        If i < lstBank.ListCount - 1 Then
            Printer.NewPage
        End If
        Printer.EndDoc
    Else
        MsgBox "Printer Error"
    End If
    
End Sub

Private Sub btnSave_Click()
    If lstCash.ListCount + lstCheque.ListCount + lstBank.ListCount <> lstAllSupplierIDs.ListCount Then
        MsgBox "Still some suppliers are to be added to a payment method"
        AddRemaining
        Exit Sub
    End If
    Dim i As Integer
    Dim rsSupplierPayments As New ADODB.Recordset
    With rsSupplierPayments
        If .State = 1 Then .Close
        temSQL = "Select * from tblCollectingCenterPaymentSummery where CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !PaymentsProcessStarted = True
            .Update
        End If
        .Close
    End With
    For i = 0 To lstCashIDs.ListCount
        With rsSupplierPayments
            If .State = 1 Then .Close
            temSQL = "Select * from tblSupplierPayments where SupplierPaymentsID = " & Val(lstCashIDs.List(i)) & " And CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText)
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !GeneratedPaymentMethodID = 2
                !Generated = True
                !GeneratedUserID = UserID
                !GeneratedDate = Date
                !GeneratedTime = Time
                .Update
            End If
            .Close
        End With
    Next i
    For i = 0 To lstChequeIDs.ListCount
        With rsSupplierPayments
            If .State = 1 Then .Close
            temSQL = "Select * from tblSupplierPayments where SupplierPaymentsID = " & Val(lstChequeIDs.List(i)) & " And CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText)
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !GeneratedPaymentMethodID = 3
                !Generated = True
                !GeneratedUserID = UserID
                !GeneratedDate = Date
                !GeneratedTime = Time
                .Update
            End If
            .Close
        End With
    Next i
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
    btnPrint.Enabled = True
    cmbCCPS.Enabled = False
    btnClose.Enabled = False
End Sub

Private Sub chkBank_Click()
    FillSelected
End Sub

Private Sub chkCash_Click()
    FillSelected
End Sub

Private Sub chkCheque_Click()
    FillSelected
End Sub

Private Sub chkOther_Click()
    FillSelected
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
End Sub

Private Sub FillCombos()
    With rsCCPS
        temSQL = "SELECT (tblCollectingCenter.CollectingCenter + ' - From ' +  convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) + ' To ' +  convert(varchar, tblCollectingCenterPaymentSummery.ToDate, 102)) AS Display, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID " & _
                    "FROM tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID " & _
                    "Where (((tblCollectingCenterPaymentSummery.PaymentsProcessStarted) = 0 )) " & _
                    "ORDER BY (tblCollectingCenter.CollectingCenter + ' - From ' +  convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) + ' To ' +  convert(varchar, tblCollectingCenterPaymentSummery.ToDate, 102))"
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
    chkBank.Value = 0
    chkCash.Value = 0
    chkCheque.Value = 0
    chkOther.Value = 0
    
    lstSelectedSupplierIDs.Clear
    lstSelectedSuppliers.Clear
    lstSelectedValue.Clear
    
    Dim i As Integer
    Dim ii As Integer
    Dim AlreadySelected As Boolean
    For i = 0 To lstAllSupplierIDs.ListCount - 1
        AlreadySelected = False
        For ii = 0 To lstCashIDs.ListCount - 1
            If Val(lstCashIDs.List(ii)) = Val(lstAllSupplierIDs.List(i)) Then
                AlreadySelected = True
            End If
        Next
        For ii = 0 To lstChequeIDs.ListCount - 1
            If Val(lstChequeIDs.List(ii)) = Val(lstAllSupplierIDs.List(i)) Then
                AlreadySelected = True
            End If
        Next
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
    lstCash.Clear
    lstCashIDs.Clear
    lstCashValue.Clear
    lstCheque.Clear
    lstChequeIDs.Clear
    lstChequeValue.Clear
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
                lstAllVaue.AddItem !Value
                SupplierValue(i) = !Value
                
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
        
        If chkBank.Value = 1 And lstPaymentMethod.List(i) = "1" Then
            lstAllSupplierIDs.Selected(i) = True
            lstAllSuppliers.Selected(i) = True
            lstAllVaue.Selected(i) = True
            SupplierSelected(i) = True
        ElseIf chkCash.Value = 1 And lstPaymentMethod.List(i) = "2" Then
            lstAllSupplierIDs.Selected(i) = True
            lstAllSuppliers.Selected(i) = True
            lstAllVaue.Selected(i) = True
            SupplierSelected(i) = True
        ElseIf chkCheque.Value = 1 And lstPaymentMethod.List(i) = "3" Then
            lstAllSupplierIDs.Selected(i) = True
            lstAllSuppliers.Selected(i) = True
            lstAllVaue.Selected(i) = True
            SupplierSelected(i) = True
        ElseIf chkOther.Value = 1 And lstPaymentMethod.List(i) = "0" Then
            lstAllSupplierIDs.Selected(i) = True
            lstAllSuppliers.Selected(i) = True
            lstAllVaue.Selected(i) = True
            SupplierSelected(i) = True
        End If
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
    
    temValue = 0
    For i = 0 To lstSelectedValue.ListCount - 1
        temValue = temValue + Val(lstSelectedValue.List(i))
    Next i
    lblCashValue.Caption = Format(temValue, "#,##0.00")
    
    temValue = 0
    For i = 0 To lstCashValue.ListCount - 1
        temValue = temValue + Val(lstSelectedValue.List(i))
    Next i
    lblCashValue.Caption = Format(temValue, "#,##0.00")
    
    temValue = 0
    For i = 0 To lstChequeValue.ListCount - 1
        temValue = temValue + Val(lstChequeValue.List(i))
    Next i
    lblChequeValue.Caption = Format(temValue, "#,##0.00")
    
    temValue = 0
    For i = 0 To lstBankValue.ListCount - 1
        temValue = temValue + Val(lstBankValue.List(i))
    Next i
    lblBankValue.Caption = Format(temValue, "#,##0.00")
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

Private Sub lstCash_Click()
    Dim i As Integer
    For i = 0 To lstCash.ListCount - 1
        lstCashValue.Selected(i) = lstCash.Selected(i)
    Next
End Sub

Private Sub lstCash_LostFocus()
    Dim i As Integer
    For i = 0 To lstCash.ListCount - 1
        lstCashValue.Selected(i) = lstCash.Selected(i)
    Next
End Sub

Private Sub lstCash_Scroll()
    lstCashValue.TopIndex = lstCash.TopIndex
End Sub

Private Sub lstCashValue_Click()
    Dim i As Integer
    For i = 0 To lstCash.ListCount - 1
        lstCash.Selected(i) = lstCashValue.Selected(i)
    Next

End Sub

Private Sub lstCashValue_LostFocus()
    Dim i As Integer
    For i = 0 To lstCash.ListCount - 1
        lstCash.Selected(i) = lstCashValue.Selected(i)
    Next

End Sub

Private Sub lstCashValue_Scroll()
    lstCash.TopIndex = lstCashValue.TopIndex
End Sub

Private Sub lstCheque_Click()
    Dim i As Integer
    For i = 0 To lstCheque.ListCount - 1
        lstCashValue.Selected(i) = lstCash.Selected(i)
    Next
End Sub

Private Sub lstCheque_LostFocus()
    Dim i As Integer
    For i = 0 To lstCheque.ListCount - 1
        lstCashValue.Selected(i) = lstCash.Selected(i)
    Next
End Sub

Private Sub lstCheque_Scroll()
    lstChequeValue.TopIndex = lstCheque.TopIndex
End Sub

Private Sub lstChequeValue_Click()
    lstCheque.TopIndex = lstChequeValue.TopIndex
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
