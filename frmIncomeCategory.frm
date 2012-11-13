VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmIncomeCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Income Category"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
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
   ScaleHeight     =   8160
   ScaleWidth      =   10185
   Begin VB.TextBox txtCategory 
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   1080
      Width           =   4935
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6600
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   5940
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   10478
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   6600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Edit"
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   6600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Delete"
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
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
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
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   495
      Left            =   7920
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Cancel"
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8640
      TabIndex        =   9
      Top             =   7560
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label Label2 
      Caption         =   "&Income Category"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "&Income Category"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmIncomeCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String

Private Sub btnAdd_Click()
    Call ClearDetails
    Call EditMode
    cmbCategory.Text = Empty
    txtCategory.SetFocus
End Sub

Private Sub btnCancel_Click()
    Call ClearDetails
    Call SelectMode
    cmbCategory.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()

    Dim i As Integer
    i = MsgBox("Are you sure?", vbYesNo)
    If i = vbNo Then Exit Sub

    If IsNumeric(cmbCategory.BoundText) = False Then
        MsgBox "Please select"
        cmbCategory.SetFocus
        Exit Sub
    End If

    i = MsgBox("Are You sure", vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsDelete As New ADODB.Recordset
    With rsDelete
        If .State = 1 Then .Close
        temSQL = "Select * from tblIncomeCategory where IncomeCategoryID = " & Val(cmbCategory.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            .Update
        End If
        .Close
    End With
    Set rsDelete = Nothing
    Call FillCombos
    cmbCategory.Text = Empty
    cmbCategory.SetFocus
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbCategory.BoundText) = False Then
        MsgBox "Please select"
        cmbCategory.SetFocus
        Exit Sub
    End If
    Call EditMode
    txtCategory.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtCategory.Text) = "" Then
        MsgBox "Nothing to Save"
        txtCategory.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbCategory.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call FillCombos
    Call SelectMode
    cmbCategory.Text = Empty
    cmbCategory.SetFocus
End Sub

Private Sub cmbCategory_Change()
    Call ClearDetails
    Call DisplayDetails
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call SelectMode
End Sub

Private Sub FillCombos()
    Dim Category As New clsFillCombos
    Category.FillAnyCombo cmbCategory, "IncomeCategory", True
End Sub

Private Sub ClearDetails()
    txtCategory.Text = Empty
End Sub

Private Sub DisplayDetails()
    Dim i As Integer
    Dim rsDisplay As New ADODB.Recordset
    With rsDisplay
        If .State = 1 Then .Close
        temSQL = "Select * from tblIncomeCategory where IncomeCategoryID = " & Val(cmbCategory.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Not IsNull(!IncomeCategory) Then
                txtCategory.Text = !IncomeCategory
            End If
        End If
        .Close
    End With
    Set rsDisplay = Nothing
End Sub

Private Sub SaveNew()
    Dim rsEdit As New ADODB.Recordset
    With rsEdit
        If .State = 1 Then .Close
        temSQL = "Select * from tblIncomeCategory where IncomeCategoryID =0"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !IncomeCategory = txtCategory.Text
        .Update
        .Close
    End With
    Set rsEdit = Nothing
End Sub

Private Sub SaveOld()
    Dim rsEdit As New ADODB.Recordset
    With rsEdit
        If .State = 1 Then .Close
        temSQL = "Select * from tblIncomeCategory where IncomeCategoryID = " & Val(cmbCategory.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !IncomeCategory = txtCategory.Text
            .Update
        End If
        .Close
    End With
    Set rsEdit = Nothing
End Sub

Private Sub SelectMode()
    cmbCategory.Enabled = True
    btnAdd.Enabled = True
    btnDelete.Enabled = True
    btnEdit.Enabled = True
    
    txtCategory.Enabled = False
    btnSave.Enabled = False
    btnCancel.Enabled = False
End Sub

Private Sub EditMode()
    cmbCategory.Enabled = False
    btnAdd.Enabled = False
    btnDelete.Enabled = False
    btnEdit.Enabled = False
    
    txtCategory.Enabled = True
    btnSave.Enabled = True
    btnCancel.Enabled = True
End Sub
