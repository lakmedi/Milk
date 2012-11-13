VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAddAdditionalCommisions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Additional Payments"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
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
   ScaleHeight     =   7830
   ScaleWidth      =   7560
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Add"
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
   Begin MSFlexGridLib.MSFlexGrid gridCommision 
      Height          =   3135
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5530
      _Version        =   393216
      AllowBigSelection=   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.TextBox txtComments 
      Height          =   1215
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   3240
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   55967747
      CurrentDate     =   39817
   End
   Begin MSDataListLib.DataCombo cmbCC 
      Height          =   360
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSupplier 
      Height          =   360
      Left            =   2400
      TabIndex        =   5
      Top             =   1440
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Delete"
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Print"
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
      Left            =   6480
      TabIndex        =   14
      Top             =   7320
      Width           =   975
      _ExtentX        =   1720
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
   Begin VB.Label Label5 
      Caption         =   "Amount"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmAddAdditionalCommisions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsCom As New ADODB.Recordset
    Dim rsViewSup As New ADODB.Recordset
    Dim temSql As String
    Dim i As Integer

Private Sub FormatGrid()
    With gridDeduction
        .Rows = 1
        
        .Cols = 8
        
        .row = 0
        
        .col = 0
        .CellAlignment = 4
        .Text = "No"
        
        .col = 1
        .CellAlignment = 4
        .Text = "Supplier"
        
        .col = 2
        .CellAlignment = 4
        .Text = "Comments"
        
        .col = 3
        .CellAlignment = 4
        .Text = "Amount"
    
        .col = 4
        .CellAlignment = 4
        .Text = "Approved on"
        
        .col = 5
        .Text = "Approval Comments"
        .CellAlignment = 4
    
        .col = 6
        .CellAlignment = 4
        .Text = "Approved Amount"
        
        .col = 7
        .Text = "ID"
        
        
    End With
End Sub

Private Sub FillGrid()
    Dim temRows As Integer
    With rsCom
        If .State = 1 Then .Close
        temSql = "SELECT tblSupplier.Supplier, tblAdditionalDeduction.*, tblAdditionalDeduction.Deleted " & _
                    "FROM tblAdditionalDeduction LEFT JOIN tblSupplier ON tblAdditionalDeduction.SupplierID = tblSupplier.SupplierID " & _
                    "WHERE (((tblSupplier.CollectingCenterID)=" & Val(cmbCC.BoundText) & ") AND ((tblAdditionalDeduction.DeductionDate)=#" & Format(dtpDate.Value, "dd MMMM yyyy") & "#) AND ((tblAdditionalDeduction.Deleted)=False))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            temRows = .RecordCount
            .MoveFirst
            gridDeduction.Rows = temRows + 1
            For i = 1 To temRows
                gridDeduction.TextMatrix(i, 0) = i
                gridDeduction.TextMatrix(i, 1) = !Supplier
                gridDeduction.TextMatrix(i, 2) = !Comments
                gridDeduction.TextMatrix(i, 3) = Format(!AddedValue, "0.00")
                If !Approved = True Then
                    gridDeduction.TextMatrix(i, 4) = Format(!ApprovedDate, "DD MMMM yyyy")
                    gridDeduction.TextMatrix(i, 5) = !ApprovedComments
                    gridDeduction.TextMatrix(i, 6) = !Value
                End If
                gridDeduction.TextMatrix(i, 7) = !AdditionalDeductionID
                .MoveNext
            Next
        End If
        .Close
    End With
    
End Sub

Private Sub ClearDetails()
    cmbSupplier.Text = Empty
    txtAmount.Text = Empty
    txtComments.Text = Empty
End Sub

Private Sub btnAdd_Click()
    If IsNumeric(cmbCC.BoundText) = False Then
        MsgBox "Please select a Collecting Center"
        cmbCC.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbSupplier.BoundText) = False Then
        MsgBox "Please select a supplier"
        cmbSupplier.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtAmount.Text) = False Then
        MsgBox "Please enter a value"
        txtAmount.SetFocus
        Exit Sub
    End If
    If Trim(txtComments.Text) = "" Then
        MsgBox "Please enter a comment"
        txtComments.SetFocus
        Exit Sub
    End If
    With rsCom
        If .State = 1 Then .Close
        temSql = "Select * from tblAdditionalDeduction"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !AddedDate = Date
        !AddedTime = Time
        !AddedValue = Val(txtAmount.Text)
        !Comments = Trim(txtComments.Text)
        !SupplierID = Val(cmbSupplier.BoundText)
        !DeductionDate = Format(dtpDate.Value, "dd MMMM yyyy")
        .Update
    End With
    Call ClearDetails
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim temRow As Integer
    If gridDeduction.Rows < 2 Then Exit Sub
    If gridDeduction.row < 1 Then Exit Sub
    temRow = gridDeduction.row
    If IsNumeric(gridDeduction.TextMatrix(temRow, 7)) = False Then Exit Sub
    With rsCom
        If .State = 1 Then .Close
        temSql = "Select * from tblAdditionalDeduction where AdditionalDeductionID = " & Val(gridDeduction.TextMatrix(temRow, 7))
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If !Approved = True Then
                MsgBox "This payment is alreasy approved. You can't delete it"
            Else
                !Deleted = True
                !DeletedUserID = UserID
                !DeletedDate = Date
                !DeletedTime = Time
                .Update
            End If
        End If
        .Close
    End With
    Call ClearDetails
    Call FormatGrid
    Call FillGrid
    
End Sub

Private Sub cmbCC_Change()
    With rsViewSup
        If .State = 1 Then .Close
        If IsNumeric(cmbCC.BoundText) = True Then
            temSql = "Select * from tblSupplier where Deleted = false and CollectingCenterID = " & Val(cmbCC.BoundText) & " ORder by Supplier"
        Else
            temSql = "Select * from tblSupplier where Deleted = false ORder by Supplier"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbSupplier
        Set .RowSource = rsViewSup
        .ListField = "Supplier"
        .BoundColumn = "SupplierID"
    End With
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbSupplier_Click(Area As Integer)

End Sub

Private Sub dtpDate_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call FillCombos
    dtpDate.Value = Date
    
End Sub

Private Sub FillCombos()
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCC, "CollectingCenter", True
End Sub
