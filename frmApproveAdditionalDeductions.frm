VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmApproveAdditionalDeductions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Approve Additional Deductions"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11175
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
   ScaleHeight     =   7200
   ScaleWidth      =   11175
   Begin btButtonEx.ButtonEx btnAll 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "A&ll"
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
   Begin MSDataListLib.DataCombo cmbOver 
      Height          =   360
      Left            =   6960
      TabIndex        =   12
      Top             =   1680
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Frame frameApprove 
      Caption         =   "Approval"
      Height          =   2535
      Left            =   5520
      TabIndex        =   19
      Top             =   3720
      Width           =   5535
      Begin VB.TextBox txtAAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1440
         TabIndex        =   23
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox txtAComments 
         Height          =   975
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   3975
      End
      Begin btButtonEx.ButtonEx btnApprove 
         Height          =   495
         Left            =   3480
         TabIndex        =   25
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "&App&rove"
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
      Begin btButtonEx.ButtonEx btnCancelApproval 
         Height          =   495
         Left            =   1440
         TabIndex        =   24
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "&Can&cel Approval"
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
      Begin VB.Label Label9 
         Caption         =   "Comme&nts"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Am&ount"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   2175
      End
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6960
      TabIndex        =   18
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox txtComments 
      Height          =   615
      Left            =   6960
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   2640
      Width           =   3975
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   495
      Left            =   9720
      TabIndex        =   26
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
   Begin VB.ListBox lstIDs 
      Height          =   4620
      Left            =   4200
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox lstAdditionalDeductions 
      Height          =   4380
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   1680
      Width           =   5295
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   95879171
      CurrentDate     =   39817
   End
   Begin MSDataListLib.DataCombo cmbCC 
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1200
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   95879171
      CurrentDate     =   39817
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   6960
      TabIndex        =   27
      Top             =   1680
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   95879171
      CurrentDate     =   39817
   End
   Begin MSDataListLib.DataCombo cmbSupplier 
      Height          =   360
      Left            =   6960
      TabIndex        =   14
      Top             =   2160
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnNone 
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   6120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&None"
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
   Begin btButtonEx.ButtonEx btnApproveSelected 
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   6120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "A&pprove Selected"
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
   Begin VB.Label Label7 
      Caption         =   "&Farmer"
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "&Date"
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Co&mments"
      Height          =   255
      Left            =   5520
      TabIndex        =   15
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Am&ount"
      Height          =   255
      Left            =   5520
      TabIndex        =   17
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "&To"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "&From"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "frmApproveAdditionalDeductions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'0777 374588
'mr.Galalgamuwa

'GeneticEnginaring

'Total Cost for the course 4,500 - 5,700
'TOtal Cost for other expences 43,000

'

'Degree Details


Option Explicit
    Dim rsAP As New ADODB.Recordset
    Dim temSql As String
    Dim rsViewSup As New ADODB.Recordset
    
Private Sub btnAll_Click()
    Dim i As Integer
    With lstAdditionalDeductions
        For i = 0 To .ListCount - 1
            .Selected(i) = True
        Next
    End With
End Sub

Private Sub btnApprove_Click()
    With rsAP
        If .State = 1 Then .Close
        temSql = "Select * from tblAdditionalDeduction where AdditionalDeductionID = " & Val(lstIDs.List(lstAdditionalDeductions.ListIndex))
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Value = Val(txtAAmount.Text)
            !ApprovedComments = txtAComments.Text
            !ApprovedUserID = UserID
            !ApprovedDate = Date
            !ApprovedTime = Time
            !Approved = True
            .Update
        End If
        .Close
    End With
    Call ClearDetails
End Sub

Private Sub btnApproveSelected_Click()
    Dim i As Integer
    For i = 0 To lstAdditionalDeductions.ListCount - 1
        If lstAdditionalDeductions.Selected(i) = True Then
        With rsAP
            If .State = 1 Then .Close
            temSql = "Select * from tblAdditionalDeduction where AdditionalDeductionID = " & Val(lstIDs.List(i))
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Value = !AddedValue
                !ApprovedComments = "Approved"
                !ApprovedUserID = UserID
                !ApprovedDate = Date
                !ApprovedTime = Time
                !Approved = True
                .Update
            End If
            .Close
        End With
        End If
    Next
    MsgBox "All Requesrs are approved"
    btnNone_Click
End Sub

Private Sub btnCancelApproval_Click()
    With rsAP
        If .State = 1 Then .Close
        temSql = "Select * from tblAdditionalDeduction where AdditionalDeductionID = " & Val(lstIDs.List(lstAdditionalDeductions.ListIndex))
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Value = 0
            !ApprovedComments = txtAComments.Text
            !ApprovedUserID = UserID
            !ApprovedDate = Date
            !ApprovedTime = Time
            !Approved = False
            .Update
        End If
        .Close
    End With
    Call ClearDetails
    
End Sub

Private Sub btnNone_Click()
    Dim ii As Integer
    With lstAdditionalDeductions
        For ii = 0 To .ListCount - 1
            .Selected(ii) = False
        Next
    End With

End Sub

Private Sub ButtonEx1_Click()
    Unload Me
End Sub

Private Sub cmbCC_Change()
    With rsViewSup
        If .State = 1 Then .Close
        If IsNumeric(cmbCC.BoundText) = True Then
            temSql = "Select * from tblSupplier where Deleted = 0  and CollectingCenterID = " & Val(cmbCC.BoundText) & " ORder by Supplier"
        Else
            temSql = "Select * from tblSupplier where Deleted = 0  ORder by Supplier"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbSupplier
        Set .RowSource = rsViewSup
        .ListField = "Supplier"
        .BoundColumn = "SupplierID"
    End With
    Call ClearDetails
    Call FillLists
End Sub

Private Sub FillLists()
    lstAdditionalDeductions.Clear
    lstIDs.Clear
    With rsAP
        If .State = 1 Then .Close
        temSql = "SELECT tblSupplier.Supplier, tblAdditionalDeduction.*, tblAdditionalDeduction.Deleted " & _
                    "FROM tblAdditionalDeduction LEFT JOIN tblSupplier ON tblAdditionalDeduction.SupplierID = tblSupplier.SupplierID " & _
                    "WHERE (((tblAdditionalDeduction.Deleted) = 0) AND ((tblSupplier.CollectingCenterID)=" & Val(cmbCC.BoundText) & ") AND ((tblAdditionalDeduction.DeductionDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "DD MMMM yyyy") & "'))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                lstIDs.AddItem !AdditionalDeductionID
                lstAdditionalDeductions.AddItem !DeductionDate & " " & !Supplier
                .MoveNext
            Wend
        End If
        .Close
        
    End With
End Sub

Private Sub dtpFrom_Change()
    Call ClearDetails
    Call FillLists

End Sub

Private Sub dtpTo_Change()
    Call ClearDetails
    Call FillLists

End Sub

Private Sub Form_Load()
    Call FillCombos
    dtpFrom.Value = Date
    dtpTo.Value = Date
End Sub

Private Sub FillCombos()
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCC, "CollectingCenter", True
End Sub

Private Sub ClearDetails()
    txtAAmount.Text = Empty
    txtAComments.Text = Empty
    txtComments.Text = Empty
    txtAComments.Text = Empty
    cmbOver.Visible = True
    cmbSupplier.Text = Empty
    txtAmount.Text = Empty
End Sub


Private Sub lstAdditionalDeductions_Click()
    Dim i As Integer
    i = lstAdditionalDeductions.ListIndex
    Call ClearDetails
    With rsAP
        If .State = 1 Then .Close
        temSql = "Select * from tblAdditionalDeduction where AdditionalDeductionID = " & Val(lstIDs.List(i))
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbSupplier.BoundText = !SupplierID
            txtComments.Text = !Comments
            txtAmount.Text = Format(!AddedValue, "0.00")
            cmbOver.Visible = False
            If IsNull(!ApprovedComments) = False Then
                txtAComments.Text = !ApprovedComments
            Else
                txtAComments.Text = Empty
            End If
            dtpDate.Value = !DeductionDate
            If !Approved = True Then
                txtAAmount.Text = Format(!Value, "0.00")
            Else
                txtAAmount.Text = Empty
            End If
        End If
        .Close
    
    End With
End Sub
