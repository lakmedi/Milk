VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Items"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
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
   ScaleHeight     =   4905
   ScaleWidth      =   10335
   Begin VB.TextBox txtComments 
      Height          =   1455
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   4095
   End
   Begin VB.TextBox txtValue 
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   360
      Width           =   4095
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9000
      TabIndex        =   12
      Top             =   4320
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
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   6600
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
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
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
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
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4320
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
   Begin MSDataListLib.DataCombo cmbItem 
      Height          =   3780
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6668
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
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
      Left            =   2760
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Label Label1 
      Caption         =   "Comments"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Item"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Unit Value"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsItem As New ADODB.Recordset
    Dim temSql As String
    
    
Private Sub btnAdd_Click()
    Dim temString As String
    temString = cmbItem.Text
    ClearValues
    txtName.Text = Empty
    EditMode
    txtName.SetFocus
    txtName.Text = temString
    SendKeys "{Home}+{end}"
End Sub

Private Sub btnAdd_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtName.SetFocus
    End If
End Sub

Private Sub btnCancel_Click()
    ClearValues
    SelectMode
    cmbItem.SetFocus
    cmbItem.SetFocus
    cmbItem.Text = Empty
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim ItemID As Long
    
    Dim i As Integer
    i = MsgBox("Are you sure?", vbYesNo)
    If i = vbNo Then Exit Sub
    
    If IsNumeric(cmbItem.BoundText) = True Then
        ItemID = Val(cmbItem.BoundText)
    Else
        MsgBox "Nothing to delete"
    End If
    With rsItem
        If .State = 1 Then .Close
        temSql = "Select * from tblItem where ItemID = " & ItemID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            .Update
        Else
            .Close
        End If
    End With
    MsgBox "Deleted"
    FillCombos
    SelectMode
    cmbItem.SetFocus
End Sub

Private Sub btnEdit_Click()
    EditMode
    txtName.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnEdit_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
       txtName.SetFocus
    End If
End Sub

Private Sub btnSave_Click()
    Dim ItemID As Long
    If IsNumeric(cmbItem.BoundText) = True Then
        SaveOld (Val(cmbItem.BoundText))
    Else
        SaveNew
    End If
    SelectMode
    FillCombos
    cmbItem.SetFocus
    cmbItem.Text = Empty
End Sub

Private Sub cmbItem_Change()
    ClearValues
    If IsNumeric(cmbItem.BoundText) = True Then
        DisplayDetails (Val(cmbItem.BoundText))
    End If
End Sub

Private Sub Form_Load()
    
    FillCombos
    SelectMode
    ClearValues
End Sub

Private Sub FillCombos()
    Dim Item As New clsFillCombos
    Item.FillAnyCombo cmbItem, "Item", True
End Sub

Private Sub EditMode()
    btnAdd.Enabled = False
    btnDelete.Enabled = False
    btnEdit.Enabled = False
    cmbItem.Enabled = False
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    txtName.Enabled = True
    txtValue.Enabled = True
    txtComments.Enabled = True
End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnDelete.Enabled = True
    btnEdit.Enabled = True
    cmbItem.Enabled = True
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
    txtName.Enabled = False
    txtValue.Enabled = False
    txtComments.Enabled = False
End Sub

Private Sub ClearValues()
    txtName.Text = Empty
    txtValue.Text = Empty
    txtComments.Text = Empty
End Sub

Private Sub DisplayDetails(ItemID As Long)
    With rsItem
        If .State = 1 Then .Close
        temSql = "Select * from tblItem where ItemID = " & ItemID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!Item) = False Then txtName.Text = !Item
            If IsNull(!Value) = False Then txtValue.Text = Format(!Value, "#,##0.00")
            If IsNull(!Comments) = False Then txtComments.Text = !Comments
        Else
            MsgBox "Error"
        End If
        .Close
    End With
End Sub

Private Sub SaveNew()
    With rsItem
        If .State = 1 Then .Close
        temSql = "Select * from tblItem"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Item = Trim(txtName.Text)
        !Value = Val(txtValue.Text)
        !Comments = txtComments.Text
        .Update
        .Close
    End With
End Sub

Private Sub SaveOld(ItemID As Long)
    With rsItem
        If .State = 1 Then .Close
        temSql = "Select * from tblItem where ItemID = " & ItemID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Item = Trim(txtName.Text)
            !Value = Val(txtValue.Text)
            !Comments = txtComments.Text
            .Update
        Else
            .Close
        End If
    End With
End Sub

Private Sub txtComments_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnSave.SetFocus
    End If
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtValue.SetFocus
    End If
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtComments.SetFocus
    End If
End Sub
