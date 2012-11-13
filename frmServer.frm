VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Details"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6285
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
   ScaleHeight     =   3150
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin btButtonEx.ButtonEx btnConnect 
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Connect"
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
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2160
      Width           =   4215
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox txtDatabase 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox txtSQLServer 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   4215
   End
   Begin VB.TextBox txtServer 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Database"
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "SQL 2005 Instance"
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server"
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   570
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
    End
End Sub

Private Sub btnConnect_Click()
    
    
    If connectToDatabase Then
        MsgBox "Successfully Connected"
    Else
        MsgBox "Connection Failure"
    End If
End Sub

Private Function connectToDatabase() As Boolean
    On Error GoTo eh:
    Dim constr As String
    Dim cnnTest As New ADODB.Connection
    connectToDatabase = False
    constr = "Provider=MSDataShape.1;Persist Security Info=True;Data Source=" & txtServer.Text & _
        "\" & txtSQLServer.Text & _
        ";User ID=" & txtUserName.Text & _
        ";Password=" & txtPassword.Text & _
        ";Initial Catalog=" & txtDatabase.Text & _
        ";Data Provider=SQLOLEDB.1"
    If cnnTest.State = 1 Then cnnTest.Close
    cnnTest.Open constr
    connectToDatabase = True
    Exit Function
eh:
    connectToDatabase = False
    
End Function


Private Sub btnSave_Click()
    saveToMemory
    saveToComputer
    Unload Me
End Sub

Private Sub Form_Load()
    getDetails
    
End Sub

Private Sub getDetails()
    txtDatabase.Text = ServerDatabase
    txtPassword.Text = ServerPassword
    txtServer.Text = Server
    txtSQLServer.Text = SQLServer
    txtUserName.Text = ServerUserName
End Sub

Private Sub saveToMemory()
    ServerDatabase = txtDatabase.Text
    ServerPassword = txtPassword.Text
    Server = txtServer.Text
     SQLServer = txtSQLServer.Text
    ServerUserName = txtUserName.Text

End Sub

Private Sub saveToComputer()
    SaveSetting App.EXEName, frmLogin.Name, "ServerDatabase", EncreptedWord(txtDatabase.Text)
    SaveSetting App.EXEName, frmLogin.Name, "ServerPassword", EncreptedWord(txtPassword.Text)
    SaveSetting App.EXEName, frmLogin.Name, "Server", EncreptedWord(txtServer.Text)
    SaveSetting App.EXEName, frmLogin.Name, "SQLServer", EncreptedWord(txtSQLServer.Text)
    SaveSetting App.EXEName, frmLogin.Name, "ServerUserName", EncreptedWord(txtUserName.Text)
End Sub
