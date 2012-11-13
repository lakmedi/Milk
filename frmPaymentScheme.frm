VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPaymentScheme 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Scheme"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
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
   ScaleHeight     =   3900
   ScaleWidth      =   10215
   Begin VB.TextBox txtPS 
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   480
      Width           =   4095
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   6960
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
      Left            =   6240
      TabIndex        =   1
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
      Left            =   7560
      TabIndex        =   2
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
      Left            =   240
      TabIndex        =   3
      Top             =   3120
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
   Begin MSDataListLib.DataCombo cmbPS 
      Height          =   2580
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4551
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   3120
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
      Left            =   2880
      TabIndex        =   6
      Top             =   3120
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
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   8880
      TabIndex        =   9
      Top             =   3120
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
   Begin VB.Label Label4 
      Caption         =   "Payment Center"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Payment Scheme"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frmPaymentScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    
Private Sub btnAdd_Click()
    Dim temText As String
    If IsNumeric(cmbPS.BoundText) = False Then
        temText = cmbPS.Text
    Else
        temText = Empty
    End If
    cmbPS.Text = Empty
    Call EditMode
    txtPS.Text = temText
    txtPS.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbPS.Text = Empty
    cmbPS.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete " & cmbPS.Text, vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblPaymentScheme where PaymentSchemeID = " & Val(cmbPS.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedTime = Now
            !DeletedUserID = UserID
            .Update
            MsgBox "Deleted"
        Else
            MsgBox "Nothing to Delete"
        End If
        .Close
    End With
    Set rsTem = Nothing
    Call FillCombos
    cmbPS.SetFocus
    cmbPS.Text = Empty
    
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbPS.BoundText) = False Then Exit Sub
    Call EditMode
    txtPS.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtPS.Text) = Empty Then
        MsgBox "You have not entered an Investigation"
        txtPS.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPS.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbPS.Text = Empty
    cmbPS.SetFocus
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub cmbPS_Change()
    Call ClearValues
    If IsNumeric(cmbPS.BoundText) = True Then Call DisplayDetails
End Sub

Private Sub Form_Load()
    Call SelectMode
    Call FillCombos
End Sub

Private Sub EditMode()
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    cmbPS.Enabled = False
    
    txtPS.Enabled = True
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    
End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    cmbPS.Enabled = True
    
    txtPS.Enabled = False
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
End Sub

Private Sub ClearValues()
    txtPS.Text = Empty
End Sub

Private Sub SaveNew()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblPaymentScheme where PaymentSchemeID =0 "
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !PaymentScheme = txtPS.Text
        .Update
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub SaveOld()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblPaymentScheme where PaymentSchemeID = " & Val(cmbPS.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !PaymentScheme = txtPS.Text
        .Update
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim PaymentScheme As New clsFillCombos
    PaymentScheme.FillAnyCombo cmbPS, "PaymentScheme", True
End Sub

Private Sub DisplayDetails()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblPaymentScheme where PaymentSchemeID = " & Val(cmbPS.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtPS.Text = !PaymentScheme
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub


