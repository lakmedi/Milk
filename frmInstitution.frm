VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmInstitution 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Institution's Profile"
   ClientHeight    =   5400
   ClientLeft      =   3390
   ClientTop       =   3390
   ClientWidth     =   8865
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
   Icon            =   "frmInstitution.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8865
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   24
      Top             =   4320
      Width           =   8655
      Begin btButtonEx.ButtonEx bttnClose 
         Height          =   375
         Left            =   7680
         TabIndex        =   14
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
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
      Begin btButtonEx.ButtonEx bttnOK 
         Height          =   375
         Left            =   5160
         TabIndex        =   12
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&OK"
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
         Left            =   2640
         TabIndex        =   13
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
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
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "E&dit"
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
   Begin VB.Frame framInstutions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   8655
      Begin VB.TextBox txtInsname 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   6375
      End
      Begin VB.TextBox txtDiscription 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   720
         Width           =   6375
      End
      Begin VB.TextBox txtRegistration 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   1200
         Width           =   6375
      End
      Begin VB.TextBox txtAddress01 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   1680
         Width           =   6375
      End
      Begin VB.TextBox txtTelephone01 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtTelephone02 
         Height          =   375
         Left            =   5280
         TabIndex        =   6
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtEmail01 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   3120
         Width           =   3135
      End
      Begin VB.TextBox txtEmail02 
         Height          =   375
         Left            =   5280
         TabIndex        =   9
         Top             =   3120
         Width           =   3135
      End
      Begin VB.TextBox txtwbsite01 
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   3600
         Width           =   3135
      End
      Begin VB.TextBox txtWebsite02 
         Height          =   375
         Left            =   5280
         TabIndex        =   11
         Top             =   3600
         Width           =   3135
      End
      Begin VB.TextBox txtFax 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   2640
         Width           =   6375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Institution &Name"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Line &1"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Line &2"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Line &3"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Tele&phone No"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "&Email "
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "&Website"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fax"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmInstitution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim SuppliedWord As String
    Dim TemResponce  As Integer
    Dim temSQL As String
    Dim rsInstitution As New ADODB.Recordset
    Dim A As String

Private Sub bttnClose_Click()
        Unload Me
End Sub

Private Sub bttnEdit_Click()
Call AfterEdit
End Sub


Private Sub bttnOK_Click()
    If Date > #12/18/2008# Then
        TemResponce = MsgBox("Please contact Lakmedipro for Assistant", vbCritical, "Expired")
        End
    End If
    If Trim(txtInsname.Text) = "" Then
        TemResponce = MsgBox("You have not entered the name of the institution", vbCritical, "Institution Name?")
        txtInsname.SetFocus
        Exit Sub
    End If
    Call SaveInstitution
    Call BeforeAddEdit
End Sub

Private Sub Form_Load()
    Call Display
    Call BeforeAddEdit
End Sub

Private Sub bttnCancel_Click()
    Call BeforeAddEdit
End Sub

Private Sub Display()
    With rsInstitution
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblInstitutionDetail"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then
            TemResponce = MsgBox("Someone had altered the database outside the program. Please contact Lakmedipro for assistant" & vbNewLine & Me.Caption & vbNewLine & Err.Description, vbCritical, "Altered Database")
            Exit Sub
        End If
        .MoveFirst
        txtInsname.Text = DecreptedWord(!InstitutionName)
        txtDiscription.Text = DecreptedWord(!InstitutionDescription)
        txtRegistration.Text = DecreptedWord(!InstitutionRegistation)
        txtAddress01.Text = DecreptedWord(!InstitutionAddress)
        txtTelephone01.Text = DecreptedWord(!institutiontelephone1)
        txtTelephone02.Text = DecreptedWord(!InstitutionTelephone2)
        txtFax.Text = DecreptedWord(!InstitutionFax)
        txtEmail01.Text = DecreptedWord(!InstitutionEmail)
        txtEmail02.Text = DecreptedWord(!InstitutionEmail2)
        txtwbsite01.Text = DecreptedWord(!InstitutionWebSite1)
        txtWebsite02.Text = DecreptedWord(!InstitutionWebSite2)
        .Close
    End With
End Sub


Private Sub SaveInstitution()
    With rsInstitution
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblInstitutionDetail"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount = 0 Then
            .AddNew
        End If
        .MoveFirst
        If txtInsname.Text = "" Then
            TemResponce = MsgBox("You have not entered an Institution Name", vbCritical, "? Institution Name")
            txtInsname.SetFocus
            Exit Sub
        End If
        !InstitutionName = EncreptedWord(txtInsname.Text)
        !InstitutionDescription = EncreptedWord(txtDiscription.Text)
        !InstitutionRegistation = EncreptedWord(txtRegistration.Text)
        !InstitutionAddress = EncreptedWord(txtAddress01.Text)
        !institutiontelephone1 = EncreptedWord(txtTelephone01.Text)
        !InstitutionTelephone2 = EncreptedWord(txtTelephone02.Text)
        !InstitutionFax = EncreptedWord(txtFax.Text)
        !InstitutionEmail = EncreptedWord(txtEmail01.Text)
        !InstitutionEmail2 = EncreptedWord(txtEmail02.Text)
        !InstitutionWebSite1 = EncreptedWord(txtwbsite01.Text)
        !InstitutionWebSite2 = EncreptedWord(txtWebsite02.Text)
        .Update
        Exit Sub
ErrorHandler:
        If Err.Number = -2147217887 Then
            TemResponce = MsgBox("The Doctor name, " & txtInsname.Text & " is already there in the database. If you want to make changes, click the Edit button", , "Alredy in the database")
        Else
            MsgBox ("An Error Occured during Updating" & vbNewLine & Me.Caption & vbNewLine & Err.Number & vbNewLine & Err.Description)
        End If
        If .State = 1 Then .CancelUpdate
        If .State = 1 Then .Close
    End With
End Sub

Private Sub BeforeAddEdit()
    bttnEdit.Enabled = True
    bttnCancel.Enabled = False
    bttnOK.Enabled = False
    txtInsname.Enabled = False
    txtDiscription.Enabled = False
    txtRegistration.Enabled = False
    txtAddress01.Enabled = False
    txtTelephone01.Enabled = False
    txtTelephone02.Enabled = False
    txtFax.Enabled = False
    txtEmail01.Enabled = False
    txtEmail02.Enabled = False
    txtwbsite01.Enabled = False
    txtWebsite02.Enabled = False
    
End Sub

Private Sub AfterEdit()
    bttnEdit.Enabled = True
    bttnCancel.Enabled = True
    bttnOK.Enabled = True
    txtInsname.Enabled = True
    txtDiscription.Enabled = True
    txtRegistration.Enabled = True
    txtAddress01.Enabled = True
    txtTelephone01.Enabled = True
    txtTelephone02.Enabled = True
    txtFax.Enabled = True
    txtEmail01.Enabled = True
    txtEmail02.Enabled = True
    txtwbsite01.Enabled = True
    txtWebsite02.Enabled = True
End Sub
Private Sub ClearValues()
    txtInsname.Text = Empty
    txtRegistration.Text = Empty
    txtAddress01.Text = Empty
    txtTelephone01.Text = Empty
    txtTelephone02.Text = Empty
    txtFax.Text = Empty
    txtEmail01.Text = Empty
    txtEmail02.Text = Empty
    txtwbsite01.Text = Empty
    txtWebsite02.Text = Empty
End Sub

Private Sub NameEmpty()
    A = MsgBox("Enter Correct Name", vbCritical + vbExclamation, "Name Empty")
End Sub


