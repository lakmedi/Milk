VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmStaffComment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Staff Comments"
   ClientHeight    =   4755
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
   ScaleHeight     =   4755
   ScaleWidth      =   10920
   Begin VB.Frame Frame2 
      ForeColor       =   &H00FF0000&
      Height          =   3975
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3615
      Begin MSDataListLib.DataCombo dtcStaff 
         Height          =   3540
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   6244
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
         Text            =   ""
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00FF0000&
      Height          =   3975
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtcomment 
         Height          =   2895
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   3240
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
         TabIndex        =   3
         Top             =   3240
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
      Begin VB.Label Label2 
         Caption         =   "C&omments"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   9000
      TabIndex        =   4
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
Attribute VB_Name = "frmStaffComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsStaff As New ADODB.Recordset
    Dim rsAgent As New ADODB.Recordset
    Dim rsViewStaff As New ADODB.Recordset
    Dim temSql As String

Private Sub bttnCancel_Click()
    Call ClearValues
    dtcStaff.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub Form_Load()
    FillCombos
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnSave_Click()
    Dim TemResponce As Integer
    If Trim(txtcomment.Text) = "" Then
        TemResponce = MsgBox("You have NOT entered any comment to add", vbCritical, "No Comments")
        txtcomment.SetFocus
        Exit Sub
    End If
    If IsNumeric(dtcStaff.BoundText) = False Then
        TemResponce = MsgBox("Please select the Staff Member", vbCritical, "Staff member?")
        dtcStaff.SetFocus
        Exit Sub
    End If
    With rsStaff
'    On Error GoTo ErrorHandler
        If .State = 1 Then .Close
        temSql = "Select * From tblStaffComment where StaffID = " & dtcStaff.BoundText
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !StaffID = Val(dtcStaff.BoundText)
        !Comment = txtcomment.Text
        !Date = Format(Date, "dd MMMM yyyy")
        !Time = Time
        !BySTaffID = UserID
        .Update
        If .State = 1 Then .Close
        ClearValues
        dtcStaff.Text = Empty
        dtcStaff.SetFocus
        Exit Sub
    
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Save Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        dtcStaff.Text = Empty
        dtcStaff.SetFocus
        If .State = 1 Then .Close
    End With
    
End Sub

Private Sub FillCombos()
    With rsViewStaff
        If .State = 1 Then .Close
        temSql = "Select * From tblStaff Order By Name"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcStaff
        Set .RowSource = rsViewStaff
        .ListField = "Name"
        .BoundColumn = "StaffID"
    End With
End Sub

Private Sub ClearValues()
    txtcomment.Text = Empty
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsStaff.State = 1 Then rsStaff.Close: Set rsStaff = Nothing
    If rsViewStaff.State = 1 Then rsViewStaff.Close: Set rsViewStaff = Nothing
End Sub
