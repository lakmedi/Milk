VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmUnit1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Collecting Centers"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
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
   ScaleHeight     =   7410
   ScaleWidth      =   10080
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   360
      Width           =   3855
   End
   Begin VB.TextBox txtCode 
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8760
      TabIndex        =   3
      Top             =   6840
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
      Left            =   6000
      TabIndex        =   4
      Top             =   6120
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
      Left            =   7320
      TabIndex        =   5
      Top             =   6120
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
      TabIndex        =   6
      Top             =   6120
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
   Begin MSDataListLib.DataCombo cmbCollectingCenter 
      Height          =   5700
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   10054
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   6120
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
      TabIndex        =   8
      Top             =   6120
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
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Code"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Unit"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmUnit1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsCollectingCenter As New ADODB.Recordset
    Dim temSql As String

Private Sub btnAdd_Click()
    Dim temString As String
    temString = cmbCollectingCenter.Text
    cmbCollectingCenter.Text = Empty
    ClearValues
    txtName.Text = temString
    EditMode
    txtName.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    ClearValues
    SelectMode
    cmbCollectingCenter.Text = Empty
    cmbCollectingCenter.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
On Error GoTo eh

    Dim i As Integer
    i = MsgBox("Are you sure?", vbYesNo)
    If i = vbNo Then Exit Sub

    If IsNumeric(cmbCollectingCenter.BoundText) = False Then
        MsgBox "Nothing to Delete"
        cmbCollectingCenter.SetFocus
        Exit Sub
    End If
    With rsCollectingCenter
        If .State = 1 Then .Close
        temSql = "Select * from tblCollectingCenter where CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            .Update
        Else
            MsgBox "Error"
        End If
        .Close
        
        ClearValues
        FillCombos
        cmbCollectingCenter.Text = Empty
        cmbCollectingCenter.SetFocus
        
        
        Exit Sub
eh:
        MsgBox Err.Description
        If .State = 1 Then .CancelUpdate
        If .State = 1 Then .Close
    End With
    ClearValues
    FillCombos
    cmbCollectingCenter.Text = Empty
    cmbCollectingCenter.SetFocus
    
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then
        MsgBox "Nothing to Edit"
        cmbCollectingCenter.SetFocus
        Exit Sub
    End If
    EditMode
    txtName.SetFocus
    SendKeys "{home}+{end}"
    
End Sub

Private Sub btnSave_Click()
    If Trim(txtName.Text) = "" Then
        MsgBox "Nothing to Save"
        txtName.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then
        SaveNew
    Else
        SaveOld
    End If
    ClearValues
    SelectMode
    FillCombos
    cmbCollectingCenter.Text = Empty
    cmbCollectingCenter.SetFocus
    
End Sub

Private Sub cmbCollectingCenter_Click(Area As Integer)
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then
        ClearValues
    Else
        DisplayDetails
    End If
End Sub

Private Sub cmbPaymentScheme_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnSave_Click
    End If
End Sub

Private Sub Form_Load()
    ClearValues
    SelectMode
    FillCombos
    

End Sub

Private Sub FillCombos()
    Dim CollectingCenter    As New clsFillCombos
    CollectingCenter.FillAnyCombo cmbCollectingCenter, "CollectingCenter", True
End Sub

Private Sub EditMode()

    cmbCollectingCenter.Enabled = False
    btnDelete.Enabled = False
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    
    txtName.Enabled = True
    txtCode.Enabled = True
    btnSave.Enabled = True
    btnCancel.Enabled = True

End Sub

Private Sub SelectMode()

    
    cmbCollectingCenter.Enabled = True
    btnDelete.Enabled = True
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    
    txtName.Enabled = False
    txtCode.Enabled = False
    btnSave.Enabled = False
    btnCancel.Enabled = False
    

End Sub

Private Sub SaveNew()
    With rsCollectingCenter
        If .State = 1 Then .Close
        temSql = "Select * from tblCollectingCenter"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !CollectingCenter = Trim(txtName.Text)
        !CollectingCenterCode = txtCode.Text
        .Update
        .Close
    End With
End Sub

Private Sub SaveOld()
    With rsCollectingCenter
        If .State = 1 Then .Close
        temSql = "Select * from tblCollectingCenter where CollectingCenterID =" & Val(cmbCollectingCenter.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !CollectingCenter = Trim(txtName.Text)
            !CollectingCenterCode = txtCode.Text
            .Update
        Else
            MsgBox "Error"
        End If
        .Close
    End With
End Sub

Private Sub ClearValues()
    txtName.Text = Empty
    cmbCollectingCenter.Text = Empty
    txtCode.Text = Empty
End Sub


Private Sub DisplayDetails()
    With rsCollectingCenter
        If .State = 1 Then .Close
        temSql = "Select * from tblCollectingCenter where CollectingCenterID =" & Val(cmbCollectingCenter.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!CollectingCenter) = False Then txtName.Text = !CollectingCenter
            If IsNull(!CollectingCenterCode) = False Then txtCode.Text = !CollectingCenterCode
        Else
            MsgBox "Error"
        End If
        .Close
    End With
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtCode.SetFocus
    End If
End Sub
