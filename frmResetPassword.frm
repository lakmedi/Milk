VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmResetPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reset Password"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
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
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6990
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   4200
      TabIndex        =   9
      Top             =   2280
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Exit"
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
   Begin VB.TextBox txtReenterNewPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtNewPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txtOldPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtUserName 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Re-enter New Password"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "New Password"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Old Password"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmResetPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
    End
End Sub

Private Sub btnSave_Click()
    Dim rsStaff As New ADODB.Recordset
    Dim temSql As String
    
    If txtNewPassword.Text <> txtReenterNewPassword.Text Then
        MsgBox "The New Password and the Reenter Password are NOT matching. Please reckeck"
        txtNewPassword.SetFocus
        Exit Sub
    End If
    
    If txtNewPassword.Text = txtOldPassword.Text Then
        MsgBox "This password has been used.Please enter the new password"
        txtNewPassword.SetFocus
        Exit Sub
    End If
    
    If txtNewPassword.Text = "" Then
        MsgBox "Please enter a password"
        txtNewPassword.SetFocus
        Exit Sub
    End If

    
    Dim TemResponce As Byte
    Dim UserNameFound As Boolean
    UserNameFound = False
    
    TemUserPassward = txtNewPassword.Text
    With rsStaff
        If .State = 1 Then .Close
        temSql = "Select tblstaff.* from tblstaff where StaffID = " & UserID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If DecreptedWord(!Password) <> txtOldPassword.Text Then
            MsgBox "The Old Password is wrong"
            txtOldPassword.SetFocus
            .Close
            Exit Sub
        End If
        !Password = EncreptedWord(txtNewPassword.Text)
        !NeedPasswordReset = False
        !ResetDate = Date
        .Update
        .Close
    End With
    MsgBox "The Change of Password is done successfully"
    Unload Me
End Sub

Private Sub Form_Load()
    txtUserName.Text = UserName
End Sub
