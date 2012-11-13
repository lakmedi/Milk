VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAuthority 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Authority Details"
   ClientHeight    =   4665
   ClientLeft      =   2130
   ClientTop       =   1635
   ClientWidth     =   10920
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
   ScaleHeight     =   4665
   ScaleWidth      =   10920
   Begin VB.Frame Frame2 
      Caption         =   "Authority "
      ForeColor       =   &H00FF0000&
      Height          =   3975
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3615
      Begin MSDataListLib.DataCombo dtcAuthority 
         Height          =   2820
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   4974
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
         Text            =   ""
      End
      Begin btButtonEx.ButtonEx bttnAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
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
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Authority Details"
      ForeColor       =   &H00FF0000&
      Height          =   3975
      Left            =   3960
      TabIndex        =   9
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtName 
         Height          =   360
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtcomment 
         Height          =   2175
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   4455
      End
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   3360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   375
         Left            =   4560
         TabIndex        =   7
         Top             =   3360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   3360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
      Begin VB.Label Label1 
         Caption         =   "A&uthority"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "C&omments"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2055
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   4200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
End
Attribute VB_Name = "frmAuthority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsAuthority As New ADODB.Recordset
    Dim rsViewAuthority As New ADODB.Recordset
    Dim temSQL As String

Private Sub bttnCancel_Click()
    Call BeforeAddEdit
    Call ClearValues
    dtcAuthority.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub dtcAuthority_Change()
    If IsNumeric(dtcAuthority.BoundText) = True Then
        Call DisplaySelected
        bttnEdit.Enabled = True
        bttnAdd.Enabled = False
    Else
        ClearValues
        bttnAdd.Enabled = True
        bttnEdit.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    FillCombos
    BeforeAddEdit
    ClearValues
End Sub

Private Sub bttnAdd_Click()
    ClearValues
    AfterAdd
    txtName.SetFocus
    txtName.Text = dtcAuthority.Text
    SendKeys "{home}+{end}"
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    AfterEdit
    txtName.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub bttnChange_Click()
    Dim TemResponce As Integer
    If txtName.Text = "" Then NoName: Exit Sub
    With rsAuthority
    On Error GoTo ErrorHandler
        If .State = 1 Then .Close
        .Open "Select * From tblAuthority Where AuthorityID = " & Val(dtcAuthority.BoundText), cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount = 0 Then Exit Sub
        !Authority = Trim(txtName.Text)
        !Comments = txtcomment.Text
        .Update
        If .State = 1 Then .Close
        FillCombos
        BeforeAddEdit
        ClearValues
        dtcAuthority.Text = Empty
        dtcAuthority.SetFocus
        Exit Sub
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Save Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        dtcAuthority.Text = Empty
        dtcAuthority.SetFocus
        If .State = 1 Then .Close
    End With
End Sub

Private Sub bttnSave_Click()
    Dim TemResponce As Integer
    If Trim(txtName.Text) = "" Then NoName: Exit Sub
    With rsAuthority
    On Error GoTo ErrorHandler
        If .State = 1 Then .Close
        .Open "Select * From tblAuthority", cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Authority = Trim(txtName.Text)
        !Comments = txtcomment.Text
        .Update
        If .State = 1 Then .Close
        FillCombos
        BeforeAddEdit
        ClearValues
        dtcAuthority.Text = Empty
        dtcAuthority.SetFocus
        Exit Sub
    
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Save Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        dtcAuthority.Text = Empty
        dtcAuthority.SetFocus
        If .State = 1 Then .Close
    End With
    
End Sub

Private Sub NoName()
    Dim TemResponce As Integer
    TemResponce = MsgBox("You have not entered an Authority to save", vbCritical, "No Name")
    txtName.SetFocus
End Sub

Private Sub FillCombos()
    With rsViewAuthority
        If .State = 1 Then .Close
        temSQL = "Select * From tblAuthority Order By Authority"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcAuthority
        Set .RowSource = rsViewAuthority
        .ListField = "Authority"
        .BoundColumn = "AuthorityID"
    End With
End Sub

Private Sub AfterAdd()
    
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    dtcAuthority.Enabled = False
    
    bttnSave.Enabled = True
    bttnChange.Enabled = False
    bttnCancel.Enabled = True
    txtcomment.Enabled = True
    txtName.Enabled = True
    
    bttnSave.Visible = True
    bttnChange.Visible = False
    
End Sub

Private Sub AfterEdit()
    
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    dtcAuthority.Enabled = False
    
    bttnSave.Enabled = False
    bttnChange.Enabled = True
    bttnCancel.Enabled = True
    txtcomment.Enabled = True
    txtName.Enabled = True
    
    bttnSave.Visible = False
    bttnChange.Visible = True
    
End Sub

Private Sub BeforeAddEdit()
    
    bttnAdd.Enabled = True
    bttnEdit.Enabled = False
    dtcAuthority.Enabled = True
    
    bttnSave.Enabled = False
    bttnChange.Enabled = False
    bttnCancel.Enabled = False
    txtcomment.Enabled = False
    txtName.Enabled = False
    
    bttnSave.Visible = True
    bttnChange.Visible = True
    
    On Error Resume Next
    dtcAuthority.SetFocus
    
End Sub

Private Sub ClearValues()
    txtName.Text = Empty
    txtcomment.Text = Empty
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsAuthority.State = 1 Then rsAuthority.Close: Set rsAuthority = Nothing
    If rsViewAuthority.State = 1 Then rsViewAuthority.Close: Set rsViewAuthority = Nothing
End Sub

Private Sub DisplaySelected()
    If Not IsNumeric(dtcAuthority.BoundText) Then Exit Sub
    With rsAuthority
        If .State = 1 Then .Close
        .Open "Select * From tblAuthority Where (AuthorityID = " & dtcAuthority.BoundText & ")", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Call ClearValues
        txtName.Text = !Authority
        If Not IsNull(!Comments) Then txtcomment.Text = !Comments
        If .State = 1 Then .Close
    End With
End Sub

