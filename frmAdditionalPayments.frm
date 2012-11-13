VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAdditionalPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Additional Payments"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
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
   ScaleHeight     =   4005
   ScaleWidth      =   7275
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   3360
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
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   2880
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   24707075
      CurrentDate     =   39749
   End
   Begin VB.TextBox txtComments 
      Height          =   1215
      Left            =   2040
      TabIndex        =   7
      Top             =   1560
      Width           =   4935
   End
   Begin VB.TextBox txtAmount 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1080
      Width           =   4935
   End
   Begin MSDataListLib.DataCombo cmbCC 
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSupplier 
      Height          =   360
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label5 
      Caption         =   "&Date"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Co&mments"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "&Amount"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "&Supplier"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmAdditionalPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim i As Integer
    Dim rsSup As New ADODB.Recordset
    Dim rsAD As New ADODB.Recordset
    
Private Sub btnAdd_Click()
    SaveDetails
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cmbCC_Change()
    FillSuppliers
End Sub


Private Sub Form_Load()
    dtpDate.Value = Date
    FillCombos
End Sub

Private Sub FillCombos()
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCC, "CollectingCenter", True
End Sub

Private Sub FillSuppliers()
    With rsSup
        If .State = 1 Then .Close
        temSql = "Select * from tblSupplier where CollectingCenterID = " & Val(cmbCC.BoundText) & " And Deleted = False "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbSupplier
        Set .RowSource = rsSup
        .ListField = "Supplier"
        .BoundColumn = "SupplierID"
    End With
End Sub

Private Sub ClearDetails()
    txtAmount.Text = Empty
    txtComments.Text = Empty
    cmbCC.Text = Empty
    cmbSupplier.Text = Empty
End Sub

Private Sub SaveDetails()
    If IsNumeric(cmbCC.BoundText) = False Then
        MsgBox "Select a Collecting Center"
        cmbCC.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbSupplier.BoundText) = False Then
        MsgBox "Select a Supplier"
        cmbSupplier.SetFocus
        Exit Sub
    End If
    If Val(txtAmount.Text) <= 0 Then
        MsgBox "Please enter a value"
        txtAmount.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If Trim(txtComments.Text) = "" Then
        MsgBox "Please enter some comments"
        txtComments.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    With rsAD
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblAdditionalCommision"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !SupplierID = Val(cmbSupplier.BoundText)
        !CommisionDate = dtpDate.Value
        !AddedDate = Date
        !AddedTime = Time
        !AddedUserID = UserID
        !Value = Val(txtAmount.Text)
        !Comments = txtComments.Text
        .Update
        .Close
    End With
    MsgBox "Successfully Saved"
    ClearDetails
    cmbCC.SetFocus
End Sub
