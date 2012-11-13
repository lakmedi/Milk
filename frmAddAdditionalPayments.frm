VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAddAdditionalPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Additional Payments"
   ClientHeight    =   8610
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
   ScaleHeight     =   8610
   ScaleWidth      =   7560
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   6480
      TabIndex        =   14
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
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
   Begin MSFlexGridLib.MSFlexGrid gridCommision 
      Height          =   3135
      Left            =   120
      TabIndex        =   15
      Top             =   4560
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
      TabIndex        =   11
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   4080
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
      Format          =   181010435
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
      TabIndex        =   9
      Top             =   2280
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
      TabIndex        =   16
      Top             =   4560
      Width           =   975
      _ExtentX        =   1720
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   7800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Print"
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
      TabIndex        =   18
      Top             =   8160
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
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMM yyyy"
      Format          =   181010435
      CurrentDate     =   39682
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMM yyyy"
      Format          =   181010435
      CurrentDate     =   39682
   End
   Begin VB.Label Label6 
      Caption         =   "&From"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "&To"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Am&ount"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Co&mments"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "D&ate"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Fa&rmer"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmAddAdditionalPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsCom As New ADODB.Recordset
    Dim rsViewSup As New ADODB.Recordset
    Dim temSQL As String
    Dim i As Integer

Private Sub FormatGrid()
    With gridCommision
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
        temSQL = "SELECT tblSupplier.Supplier, tblAdditionalCommision.*, tblAdditionalCommision.Deleted " & _
                    "FROM tblAdditionalCommision LEFT JOIN tblSupplier ON tblAdditionalCommision.SupplierID = tblSupplier.SupplierID " & _
                    "WHERE (((tblSupplier.CollectingCenterID)=" & Val(cmbCC.BoundText) & ") AND ((tblAdditionalCommision.CommisionDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND ((tblAdditionalCommision.Deleted) = 0))"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            temRows = .RecordCount
            .MoveFirst
            gridCommision.Rows = temRows + 1
            For i = 1 To temRows
                gridCommision.TextMatrix(i, 0) = i
                gridCommision.TextMatrix(i, 1) = !Supplier
                gridCommision.TextMatrix(i, 2) = !Comments
                gridCommision.TextMatrix(i, 3) = Format(!AddedValue, "0.00")
                If !Approved = True Then
                    gridCommision.TextMatrix(i, 4) = Format(!ApprovedDate, "DD MMMM yyyy")
                    gridCommision.TextMatrix(i, 5) = !ApprovedComments
                    gridCommision.TextMatrix(i, 6) = !Value
                End If
                gridCommision.TextMatrix(i, 7) = !AdditionalCommisionID
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
    If dtpTo.Value <= dtpFrom.Value Then
        MsgBox "From date has to be earlier then To date"
        dtpTo.SetFocus
        Exit Sub
    End If
    If dtpDate.Value < dtpFrom.Value Or dtpDate.Value > dtpTo.Value Then
        MsgBox "You have to select a date between From date and To date"
        dtpDate.SetFocus
        Exit Sub
    End If
    With rsCom
        If .State = 1 Then .Close
        temSQL = "Select * from tblAdditionalCommision where AdditionalCommisionID =0"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !AddedDate = Date
        !AddedTime = Time
        !AddedValue = Val(txtAmount.Text)
        !Comments = Trim(txtComments.Text)
        !SupplierID = Val(cmbSupplier.BoundText)
        !CommisionDate = Format(dtpDate.Value, "dd MMMM yyyy")
        !AddedUserID = UserID
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
    Dim i As Integer
    i = MsgBox("Are you sure?", vbYesNo)
    If i = vbNo Then Exit Sub
    
    Dim temRow As Integer
    If gridCommision.Rows < 2 Then Exit Sub
    If gridCommision.row < 1 Then Exit Sub
    temRow = gridCommision.row
    If IsNumeric(gridCommision.TextMatrix(temRow, 7)) = False Then Exit Sub
    With rsCom
        If .State = 1 Then .Close
        temSQL = "Select * from tblAdditionalCommision where AdditionalCommisionID = " & Val(gridCommision.TextMatrix(temRow, 7))
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
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
            temSQL = "Select * from tblSupplier where Deleted = 0  and CollectingCenterID = " & Val(cmbCC.BoundText) & " ORder by Supplier"
        Else
            temSQL = "Select * from tblSupplier where Deleted = 0  ORder by Supplier"
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbSupplier
        Set .RowSource = rsViewSup
        .ListField = "Supplier"
        .BoundColumn = "SupplierID"
    End With
    Call FormatGrid
    Call FillGrid
    If LastPaymentGeneratedDate(Val(cmbCC.BoundText)) < dtpDate.Value Then
        btnAdd.Enabled = True
        btnDelete.Enabled = True
    Else
        btnAdd.Enabled = False
        btnDelete.Enabled = False
    End If
End Sub

Private Sub cmbCC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpDate.SetFocus
    End If
End Sub

Private Sub cmbSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtComments.SetFocus
    End If
End Sub

Private Sub dtpDate_Change()
    Call FormatGrid
    Call FillGrid
    If LastPaymentGeneratedDate(Val(cmbCC.BoundText)) < dtpDate.Value Then
        btnAdd.Enabled = True
        btnDelete.Enabled = True
    Else
        btnAdd.Enabled = False
        btnDelete.Enabled = False
    End If
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpFrom.SetFocus
    End If
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpTo.SetFocus
    End If
End Sub

Private Sub dtpTo_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    If KeyCode = vbKeyReturn Then
        cmbSupplier.SetFocus
    End If
End Sub

Private Sub Form_Load()
    If UserAuthorityLevel = OrdinaryUser Then
        btnDelete.Enabled = False
    End If
    
    Call FillCombos
    Call FormatGrid
    
    dtpDate.Value = Date
    dtpFrom.Value = Date
    dtpTo.Value = Date
    
    Select Case UserAuthorityLevel
    
    
    Case Authority.OrdinaryUser '3
    btnDelete.Visible = False
    
    Case Authority.PowerUser '4
    btnDelete.Visible = True

    Case Authority.SuperUser '5
    btnDelete.Visible = True
    
    Case Authority.Administrator '6
    btnDelete.Visible = True

    Case Else
    
    End Select
    If ItemSuppiersEditAllowed = True Then
       btnDelete.Visible = True
    Else
        btnDelete.Visible = False
    End If

End Sub

Private Sub FillCombos()
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCC, "CollectingCenter", True
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnAdd_Click
    End If
End Sub

Private Sub txtComments_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtAmount.SetFocus
    End If
End Sub
