VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPriceCycle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price Cycles"
   ClientHeight    =   8145
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
   ScaleHeight     =   8145
   ScaleWidth      =   10080
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   103088131
      CurrentDate     =   39877
   End
   Begin VB.TextBox txtPriceCycle 
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   1920
      Width           =   4935
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6600
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSDataListLib.DataCombo cmbPriceCycle 
      Height          =   5940
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   10478
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   6600
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   3000
      TabIndex        =   4
      Top             =   6600
      Width           =   1335
      _ExtentX        =   2355
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
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   5640
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
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
      TabIndex        =   12
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8640
      TabIndex        =   13
      Top             =   7560
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   103088131
      CurrentDate     =   39877
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "From"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Price Cycle"
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Price Cycle"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmPriceCycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String

Private Sub btnAdd_Click()
    Call ClearDetails
    Call EditMode
    cmbPriceCycle.Text = Empty
    txtPriceCycle.SetFocus
End Sub

Private Sub btnCancel_Click()
    Call ClearDetails
    Call SelectMode
    cmbPriceCycle.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()

    Dim i As Integer
    i = MsgBox("Are you sure?", vbYesNo)
    If i = vbNo Then Exit Sub
    
    If IsNumeric(cmbPriceCycle.BoundText) = False Then
        MsgBox "Please select"
        cmbPriceCycle.SetFocus
        Exit Sub
    End If

    i = MsgBox("Are You sure", vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsDelete As New ADODB.Recordset
    With rsDelete
        If .State = 1 Then .Close
        temSQL = "Select * from tblPriceCycle where PriceCycleID = " & Val(cmbPriceCycle.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedDate = Date
            !DeletedTime = Time
            !DeletedUserID = UserID
            .Update
        End If
        .Close
    End With
    Set rsDelete = Nothing
    Call FillCombos
    cmbPriceCycle.Text = Empty
    cmbPriceCycle.SetFocus
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbPriceCycle.BoundText) = False Then
        MsgBox "Please select"
        cmbPriceCycle.SetFocus
        Exit Sub
    End If
    Call EditMode
    txtPriceCycle.SetFocus
    On Error Resume Next
    SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtPriceCycle.Text) = "" Then
        MsgBox "Nothing to Save"
        txtPriceCycle.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPriceCycle.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call FillCombos
    Call SelectMode
    cmbPriceCycle.Text = Empty
    cmbPriceCycle.SetFocus
End Sub

Private Sub cmbPriceCycle_Change()
    Call ClearDetails
    Call DisplayDetails
End Sub

Private Sub dtpFrom_Change()
    If txtPriceCycle.Enabled = True Then
        txtPriceCycle.Text = "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    End If
End Sub

Private Sub dtpTo_Change()
    If txtPriceCycle.Enabled = True Then
        txtPriceCycle.Text = "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    End If
End Sub

Private Sub Form_Load()
    dtpFrom.Value = Date
    dtpTo.Value = Date
    Call FillCombos
    Call SelectMode
End Sub

Private Sub FillCombos()
    Dim Category As New clsFillCombos
    Category.FillAnyCombo cmbPriceCycle, "PriceCycle", True
End Sub

Private Sub ClearDetails()
    txtPriceCycle.Text = Empty
End Sub

Private Sub DisplayDetails()
    Dim i As Integer
    Dim rsDisplay As New ADODB.Recordset
    With rsDisplay
        If .State = 1 Then .Close
        temSQL = "Select * from tblPriceCycle where PriceCycleID = " & Val(cmbPriceCycle.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            dtpFrom.Value = !FromDate
            dtpTo.Value = !ToDate
            If Not IsNull(!PriceCycle) Then
                txtPriceCycle.Text = !PriceCycle
            End If
        End If
        .Close
    End With
    Set rsDisplay = Nothing
End Sub

Private Sub SaveNew()
    Dim rsEdit As New ADODB.Recordset
    With rsEdit
        If .State = 1 Then .Close
        temSQL = "Select * from tblPriceCycle where PriceCycleID =0"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !PriceCycle = txtPriceCycle.Text
        !FromDate = dtpFrom.Value
        !ToDate = dtpTo.Value
        .Update
        .Close
    End With
    Set rsEdit = Nothing
End Sub

Private Sub SaveOld()
    Dim rsEdit As New ADODB.Recordset
    With rsEdit
        If .State = 1 Then .Close
        temSQL = "Select * from tblPriceCycle where PriceCycleID = " & Val(cmbPriceCycle.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !PriceCycle = txtPriceCycle.Text
            !FromDate = dtpFrom.Value
            !ToDate = dtpTo.Value
            .Update
        End If
        .Close
        temSQL = "Select * from tblPrice where PriceCycleID = " & Val(cmbPriceCycle.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            !FromDate = dtpFrom.Value
            !ToDate = dtpTo.Value
            .Update
            .MoveNext
        Wend
        .Close
    End With
    Set rsEdit = Nothing
End Sub

Private Sub SelectMode()
    cmbPriceCycle.Enabled = True
    btnAdd.Enabled = True
    btnDelete.Enabled = True
    btnEdit.Enabled = True
        
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    txtPriceCycle.Enabled = False
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
End Sub

Private Sub EditMode()
    cmbPriceCycle.Enabled = False
    btnAdd.Enabled = False
    btnDelete.Enabled = False
    btnEdit.Enabled = False
    
    txtPriceCycle.Enabled = True
    dtpFrom.Enabled = True
    dtpTo.Enabled = True
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
End Sub

