VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmFarmerVolumeDeduction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Farmer Volume Deductions"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8580
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
   ScaleHeight     =   7440
   ScaleWidth      =   8580
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
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
   Begin VB.TextBox txtDeduction 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   1800
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid gridDeductions 
      Height          =   4695
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8281
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSDataListLib.DataCombo cmbSupplierName 
      Height          =   360
      Left            =   2160
      TabIndex        =   0
      Top             =   720
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbCollectingCenter 
      Height          =   360
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbDeduction 
      Height          =   360
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   6960
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label Label3 
      Caption         =   "Deduction per Liter"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Deduction Type"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Farmer &Name"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmFarmerVolumeDeduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsItemIssue As New ADODB.Recordset
    Dim rsViewSuppliers As New ADODB.Recordset
    Dim temSql As String

Private Sub btnAdd_Click()
    If IsNumeric(cmbSupplierName.BoundText) = False Then
        MsgBox "Supplier?"
        cmbSupplierName.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbDeduction.BoundText) = False Then
        MsgBox "Deduction?"
        cmbDeduction.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtDeduction.Text) = False Then
        MsgBox "Deduction per liter?"
        txtDeduction.SetFocus
        Exit Sub
    End If
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Insert into tblFarmerVolumeDeduction(FarmerID, VolumeDeductionID, DeductionPerLiter) Values(" & Val(cmbSupplierName.BoundText) & ", " & Val(cmbDeduction.BoundText) & ", " & Val(txtDeduction.Text) & ")"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
    End With
    Call FormatGrid
    Call FillGrid
    
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure?", vbYesNo)
    If i = vbNo Then Exit Sub
    Dim temID As Long
    temID = Val(gridDeductions.TextMatrix(gridDeductions.row, 0))
    temSql = "SELECT * FROM  tblFarmerVolumeDeduction where FarmerVolumeDeductionID = " & temID
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedUserID = UserID
            !DeletedDate = Date
            !DeletedTime = Time
            .Update
        End If
        .Close
    End With
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call FillCombos
    
End Sub


Private Sub FillCombos()
    
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCollectingCenter, "CollectingCenter", True

    Dim PD As New clsFillCombos
    PD.FillAnyCombo cmbDeduction, "VolumeDeduction", True
End Sub

Private Sub cmbCollectingCenter_Change()
    With rsViewSuppliers
        If .State = 1 Then .Close
        temSql = "Select * from tblSupplier where farmer = 1 and Deleted = 0  and CollectingCenterID = " & Val(cmbCollectingCenter.BoundText) & " ORDER BY SUPPLIER"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbSupplierName
        Set .RowSource = rsViewSuppliers
        .ListField = "Supplier"
        .BoundColumn = "SupplierID"
    End With
End Sub

Private Sub cmbSupplierName_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub FillGrid()
    Dim rsDeductions As New ADODB.Recordset
    temSql = "SELECT  tblFarmerVolumeDeduction.FarmerVolumeDeductionID,  tblVolumeDeduction.VolumeDeduction, tblFarmerVolumeDeduction.DeductionPerLiter from " & _
                "tblFarmerVolumeDeduction INNER JOIN tblVolumeDeduction ON tblFarmerVolumeDeduction.VolumeDeductionID = tblVolumeDeduction.VolumeDeductionID " & _
                "WHERE tblFarmerVOlumeDeduction.FarmerID = " & Val(cmbSupplierName.BoundText) & " AND tblFarmerVOlumeDeduction.Deleted = 0"
                
    With rsDeductions
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                gridDeductions.Rows = gridDeductions.Rows + 1
                gridDeductions.row = gridDeductions.Rows - 1
                gridDeductions.col = 0
                gridDeductions.Text = !FarmerVolumeDeductionID
                gridDeductions.col = 1
                gridDeductions.Text = !VolumeDeduction
                gridDeductions.col = 2
                gridDeductions.Text = Format(!DeductionPerLiter, "0.00")
                .MoveNext
            Wend
        End If
        .Close
    End With
End Sub

Private Sub FormatGrid()
    With gridDeductions
        .Clear
        .Rows = 1
        .Cols = 3
        
        .row = 0
        
        .col = 0
        .Text = "ID"
        
        .col = 1
        .Text = "Deduction"
        
        .col = 2
        .Text = "Deduction per Liter"
        
        
        .ColWidth(0) = 0
        .ColWidth(2) = 3500 '
        .ColWidth(1) = .Width - 100 - 3500 '
        
    End With
    
End Sub

