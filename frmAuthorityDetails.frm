VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAuthorityDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Authority Details"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
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
   ScaleHeight     =   1785
   ScaleWidth      =   3600
   Begin VB.Frame Frame1 
      Caption         =   "Select Authority"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3375
      Begin MSDataListLib.DataCombo dtcAuthority 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3375
      Begin btButtonEx.ButtonEx bttnPrint 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "&Print"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnClose 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "&Close"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAuthorityDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsAuthority As New ADODB.Recordset
    
Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnPrint_Click()
'    If Not IsNumeric(dtcAuthority.BoundText) Then Exit Sub
'    With DataEnvironment1.rscmmdAuthority
'        If .State = 1 Then .Close
'        temSQL = "SELECT tblStaff.* " & _
'                    "From tblStaff " & _
'                    "Where (((tblStaff.AuthorityID) = " & dtcAuthority.BoundText & "))" & _
'                    "ORDER BY tblStaff.Name"
'       .Open temSQL
'    End With
'    With dtrAuUsers
'        .Sections.Item("Section4").Controls.Item("LblTopic").Caption = "Authority Details"
'        .Show
'    End With
End Sub

Private Sub Form_Load()
    Call FillCombos
End Sub


Private Sub FillCombos()
    With rsAuthority
        If .State = 1 Then .Close
        temSql = "SELECT * from tblAuthority order by Authority"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcAuthority
        Set .RowSource = rsAuthority
        .ListField = "Authority"
        .BoundColumn = "AuthorityID"
    End With

End Sub

