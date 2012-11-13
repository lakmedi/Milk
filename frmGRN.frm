VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmGRN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Good Receive Note"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
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
   ScaleHeight     =   2730
   ScaleWidth      =   6930
   Begin VB.TextBox txtTLMR 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtTFAT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   5640
      TabIndex        =   10
      Top             =   2040
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
   Begin VB.TextBox txtLiters 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   160759811
      CurrentDate     =   39785
   End
   Begin MSDataListLib.DataCombo cmbCC 
      Height          =   360
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpDate1 
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   160759811
      CurrentDate     =   39785
   End
   Begin MSDataListLib.DataCombo cmbCC1 
      Height          =   360
      Left            =   2760
      TabIndex        =   13
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label17 
      Caption         =   "Tested Average of &LMR"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label18 
      Caption         =   "Tested Average of &FAT%"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "liters"
      Height          =   255
      Left            =   5400
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "&Volume"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "&Date"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmGRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cmbCC_Change()
    Call WriteDailyCollection
    Call DisplayDailyCollection
    cmbCC1.BoundText = Val(cmbCC.BoundText)
End Sub

Private Sub cmbCC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpDate.SetFocus
    End If
End Sub

Private Sub dtpDate_Change()
    Call WriteDailyCollection
    Call DisplayDailyCollection
    dtpDate1.Value = dtpDate.Value
End Sub

Private Sub Form_Load()
    Call FillCombos
    dtpDate.Value = Date
    dtpDate1.Value = Date
End Sub

Private Sub FillCombos()
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCC, "CollectingCenter", True
    Dim cc1 As New clsFillCombos
    cc1.FillAnyCombo cmbCC1, "CollectingCenter", True
End Sub

Private Sub DisplayDailyCollection()
    Dim rsDailyCollection As New ADODB.Recordset
    With rsDailyCollection
        If .State = 1 Then .Close
        temSql = "SELECT tblDailyCollection.* FROM tblDailyCollection Where ProgramDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' And CollectingCenterID = " & Val(cmbCC.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 1 Then
            txtLiters.Text = Format(!ActualVolume, "0.00")
            txtTLMR.Text = Format(!TestedLMR, "0.00")
            txtTFAT.Text = Format(!TestedFAT, "0.00")
        Else
            txtLiters.Text = Format(0, "0.00")
            txtTLMR.Text = Format(0, "0.00")
            txtTFAT.Text = Format(0, "0.00")
        End If
        .Close
    End With
End Sub

Private Sub WriteDailyCollection()
    Dim rsDailyCollection As New ADODB.Recordset
    Dim rsSecessionCollection As New ADODB.Recordset
    
    Dim MyPrice As Double
    Dim MySNF As Double
    
    
    If IsNumeric(cmbCC1.BoundText) = False Then Exit Sub
    With rsDailyCollection
        If .State = 1 Then .Close
        temSql = "SELECT tblDailyCollection.* FROM tblDailyCollection Where ProgramDate = '" & Format(dtpDate1.Value, "dd MMMM yyyy") & "' And CollectingCenterID = " & Val(cmbCC1.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount < 1 Then
            .AddNew
            !Date = Format(dtpDate1.Value, "dd MMMM yyyy")
            !CollectionDate = dtpDate1.Value
            !DeliveryDate = dtpDate1.Value + 1
            !ProgramDate = dtpDate1.Value
            !CollectingCenterID = cmbCC1.BoundText
        End If
        !ActualVolume = Val(txtLiters.Text)
        !TestedLMR = Val(txtTLMR.Text)
        !TestedFAT = Val(txtTFAT.Text)
        MySNF = SNF(Val(txtTLMR.Text), Val(txtTFAT.Text))
        MyPrice = Price(Val(txtTFAT.Text), MySNF, Val(cmbCC.BoundText), 0, dtpDate.Value)
        !actualValue = Val(txtLiters.Text) * MyPrice
        !ValueDifference = !TotalValue - !actualValue
        .Update
        .Close
    End With
    With rsDailyCollection
        If .State = 1 Then .Close
        temSql = "SELECT tblSecessionCollection.* FROM tblSecessionCollection Where ProgramDate = '" & Format(dtpDate1.Value, "dd MMMM yyyy") & "' And CollectingCenterID = " & Val(cmbCC1.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 1 Then
            While .EOF = False
'                !ActualVolume = Val(txtLiters.Text)
                !TestedLMR = Val(txtTLMR.Text)
                !TestedFAT = Val(txtTFAT.Text)
                .Update
                .MoveNext
            Wend
        End If
        .Close
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call WriteDailyCollection
End Sub
