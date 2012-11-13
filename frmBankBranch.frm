VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmBankBranch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Branch"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11490
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
   ScaleHeight     =   7125
   ScaleWidth      =   11490
   Begin VB.TextBox txtCode 
      Height          =   360
      Left            =   5760
      TabIndex        =   15
      Tag             =   "E"
      Top             =   3000
      Width           =   5295
   End
   Begin VB.TextBox txtBranch 
      Height          =   360
      Left            =   5760
      TabIndex        =   13
      Tag             =   "E"
      Top             =   2055
      Width           =   5295
   End
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Tag             =   "S"
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
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
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Tag             =   "S"
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Tag             =   "S"
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
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
   Begin MSDataListLib.DataCombo cmbBankS 
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Tag             =   "S"
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbBranch 
      Height          =   5220
      Left            =   120
      TabIndex        =   0
      Tag             =   "S"
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   9208
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   10080
      TabIndex        =   10
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
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
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   7200
      TabIndex        =   11
      Tag             =   "E"
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
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
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   495
      Left            =   8520
      TabIndex        =   12
      Tag             =   "E"
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
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
   Begin MSDataListLib.DataCombo cmbBankE 
      Height          =   360
      Left            =   5760
      TabIndex        =   14
      Tag             =   "E"
      Top             =   2520
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Code"
      Height          =   240
      Left            =   4920
      TabIndex        =   8
      Tag             =   "E"
      Top             =   3000
      Width           =   435
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Branch"
      Height          =   240
      Left            =   4920
      TabIndex        =   7
      Tag             =   "E"
      Top             =   2040
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Bank"
      Height          =   240
      Left            =   4920
      TabIndex        =   6
      Tag             =   "E"
      Top             =   2520
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Branch"
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Tag             =   "S"
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bank"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Tag             =   "S"
      Top             =   120
      Width           =   405
   End
End
Attribute VB_Name = "frmBankBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Branch As New clsFillCombos
    Dim temSql As String
    
Private Sub btnAdd_Click()
    Dim temText As String
    EditMode Me
    temText = cmbBranch.Text
    cmbBranch.Text = Empty
    txtBranch.Text = temText
    cmbBankE.BoundText = cmbBankS.BoundText
    txtBranch.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
        ClearEditDetails Me
        SelectMode Me
        cmbBranch.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnEdit_Click()
    EditMode Me
    txtBranch.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
        If CanSave.CanSave = False Then
            MsgBox CanSave.Message
            CanSave.Control.SetFocus
            Exit Sub
        End If
        If IsNumeric(cmbBranch.BoundText) = True Then
            Call SaveOld
        Else
            Call SaveNew
        End If
        ClearEditDetails Me
        SelectMode Me
        cmbBranch.SetFocus
End Sub

Private Sub SaveNew()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblCity"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !City = txtBranch.Text
        !BankID = Val(cmbBankE.BoundText)
        !BankCode = txtCode.Text
        .Update
        .Close
    End With
End Sub

Private Sub SaveOld()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblCity where CityID = " & Val(cmbBranch.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !City = txtBranch.Text
            !BankID = Val(cmbBankE.BoundText)
            !BankCode = txtCode.Text
            .Update
        End If
        .Close
    End With
End Sub


Private Function CanSave() As SaveCheck
    If Trim(txtBranch.Text) = Empty Then
        CanSave.CanSave = False
        CanSave.Control = txtBranch
        CanSave.Message = "Please enter a Branch Name"
        Exit Function
    End If
    If Trim(txtCode.Text) = Empty Then
        CanSave.CanSave = False
        CanSave.Control = txtCode
        CanSave.Message = "Please enter a Branch Code"
        Exit Function
    End If
    If IsNumeric(cmbBankE.BoundText) = False Then
        CanSave.CanSave = False
        CanSave.Control = cmbBankE
        CanSave.Message = "Please select a Bank Name"
        Exit Function
    End If
    CanSave.CanSave = True
End Function

Private Sub cmbBankS_Click(Area As Integer)
    Branch.FillLongCombo cmbBranch, "City", "City", "BankID", Val(cmbBankS.BoundText), True
End Sub

Private Sub cmbBranch_Change()
    Call DisplayDetails
End Sub

Private Sub cmbBranch_Click(Area As Integer)
    Call DisplayDetails
End Sub

Private Sub DisplayDetails()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblCity where CityID = " & Val(cmbBranch.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtBranch.Text = Format(!City, "")
            cmbBankE.BoundText = !BankID
            txtCode.Text = Format(!BankCode, "")
        End If
        .Close
    End With
End Sub

Private Sub Form_Load()
    SelectMode Me
    FillCombos
End Sub

Private Sub FillCombos()
    Dim B1 As New clsFillCombos
    Dim B2 As New clsFillCombos
    B1.FillAnyCombo cmbBankE, "Bank", True
    B2.FillAnyCombo cmbBankS, "Bank", True
End Sub

