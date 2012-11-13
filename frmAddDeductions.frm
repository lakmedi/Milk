VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAddDeductions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Deductions"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8235
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
   ScaleHeight     =   8220
   ScaleWidth      =   8235
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   360
      Left            =   5520
      TabIndex        =   17
      Top             =   2520
      Width           =   1335
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   6960
      TabIndex        =   29
      Top             =   7560
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
   Begin VB.TextBox txtNetPayments 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   4920
      TabIndex        =   28
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox txtTotalPayments 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   4920
      TabIndex        =   26
      Top             =   6600
      Width           =   2175
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   7200
      TabIndex        =   22
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   7200
      TabIndex        =   20
      Top             =   3000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
   Begin VB.TextBox txtTotalDeductions 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   4920
      TabIndex        =   24
      Top             =   6120
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid GridDeductions 
      Height          =   2535
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4471
      _Version        =   393216
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   360
      Left            =   5520
      TabIndex        =   19
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2040
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtToPay 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   360
      Left            =   5520
      TabIndex        =   15
      Top             =   2040
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo cmbSupplierName 
      Height          =   360
      Left            =   2040
      TabIndex        =   3
      Top             =   600
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
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbItem 
      Height          =   360
      Left            =   2040
      TabIndex        =   9
      Top             =   1560
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMM yyyy"
      Format          =   181534723
      CurrentDate     =   39682
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMM yyyy"
      Format          =   181534723
      CurrentDate     =   39682
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   2520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMM yyyy"
      Format          =   181534723
      CurrentDate     =   39682
   End
   Begin VB.Label Label13 
      Caption         =   "Ded&ucted Date"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "Q&ty. to Deduct"
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "&Rate"
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "&Net Payment"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Total Pa&yments"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Total D&eductions"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "&From"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "&To"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "&Value"
      Height          =   375
      Left            =   4200
      TabIndex        =   18
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "&Quenty"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "&Item"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Farmer &Name"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmAddDeductions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsItemIssue As New ADODB.Recordset
    Dim rsViewSuppliers As New ADODB.Recordset
    Dim temSQL As String
    

Private Sub btnAdd_Click()
    If IsNumeric(cmbSupplierName.BoundText) = False Then
        MsgBox "Supplier?"
        cmbSupplierName.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbItem.BoundText) = False Then
        MsgBox "Item?"
        cmbItem.SetFocus
        Exit Sub
    End If
    If Val(txtQty.Text) <= 0 Then
        MsgBox "Quentity"
        txtQty.SetFocus
        SendKeys "{Home}+{end}"
        Exit Sub
    End If
    If Val(txtQty.Text) > Val(txtToPay.Text) Then
        MsgBox "Can't pay more than due"
        txtQty.SetFocus
        SendKeys "{Home}+{end}"
        Exit Sub
    End If
    If Val(txtValue.Text) > Val(txtNetPayments.Text) Then
        MsgBox "Can't pay more than the net Payment"
        txtQty.SetFocus
        SendKeys "{home}+{end}"
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
    Dim rsDeduction As New ADODB.Recordset
    With rsDeduction
        temSQL = "Select * from tblDeduction where DeductionID =0"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !DeductedDate = dtpDate.Value
        !ItemID = Val(cmbItem.BoundText)
        !SupplierID = Val(cmbSupplierName.BoundText)
        !Quentity = Val(txtQty.Text)
        !Rate = Val(txtRate.Text)
        !Value = Val(txtValue.Text)
        !AddedDate = Date
        !AddedTime = Time
        !AddedUserID = UserID
        .Update
        .Close
    End With
    Dim rsItemIssue As New ADODB.Recordset
    Dim remainingQty As Double
    Dim SettlingQty As Double
    remainingQty = Val(txtQty.Text)
    With rsItemIssue
        If .State = 1 Then .Close
        temSQL = "SELECT tblItemIssue.Quentity, tblItemIssue.Paid, tblItemIssue.ToPay, tblItemIssue.Quentity, tblItemIssue.IssueDate " & _
                    "From tblItemIssue " & _
                    "WHERE (((tblItemIssue.Quentity)>[tblItemIssue].[Paid]) AND ((tblItemIssue.ItemID)=" & Val(cmbItem.BoundText) & ") AND ((tblItemIssue.SupplierID)=" & Val(cmbSupplierName.BoundText) & ")) " & _
                    "ORDER BY tblItemIssue.IssueDate DESC"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            Do While .EOF = False
                If !ToPay >= remainingQty Then
                    !ToPay = !ToPay - remainingQty
                    !Paid = !Paid + remainingQty
                    .Update
                    remainingQty = 0
                    Exit Do
                Else
                    SettlingQty = !ToPay
                    !ToPay = 0
                    !Paid = !Paid + SettlingQty
                    .Update
                    remainingQty = remainingQty - SettlingQty
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    Call ClearAddValues
    Call FormatGrid
    Call FillGrid
    Call FindPaymentsAndDeductions
    cmbItem.SetFocus
End Sub

Private Sub ClearAddValues()
    cmbItem.Text = Empty
    txtQty.Text = Empty
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete?", vbYesNo)
    If i = vbNo Then Exit Sub
    
    Dim MyRow As Integer
    Dim DeductionID As Long
    Dim ItemID As Long
    With GridDeductions
        If .Rows < 2 Then Exit Sub
        If .row < 1 Then Exit Sub
        MyRow = .row
        If IsNumeric(.TextMatrix(MyRow, 5)) = False Then Exit Sub
        DeductionID = Val(.TextMatrix(MyRow, 5))
        ItemID = Val(.TextMatrix(MyRow, 6))
    End With
    Dim rsDeduction As New ADODB.Recordset
    With rsDeduction
        If .State = 1 Then .Close
        temSQL = "Select * from tblDeduction where DeductionID = " & DeductionID
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            .Update
        End If
        .Close
    End With
    
    
    
    Dim rsItemIssue As New ADODB.Recordset
    Dim remainingQty As Double
    Dim SettlingQty As Double
    remainingQty = Val(GridDeductions.TextMatrix(MyRow, 2))
    With rsItemIssue
        If .State = 1 Then .Close
        temSQL = "SELECT tblItemIssue.Quentity, tblItemIssue.Paid, tblItemIssue.ToPay, tblItemIssue.Quentity, tblItemIssue.IssueDate " & _
                    "From tblItemIssue " & _
                    "WHERE (((tblItemIssue.ItemID)=" & ItemID & ") AND ((tblItemIssue.SupplierID)=" & Val(cmbSupplierName.BoundText) & ")) " & _
                    "ORDER BY tblItemIssue.IssueDate DESC"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            Do While .EOF = False
                If !Paid >= remainingQty Then
                    !ToPay = !ToPay + remainingQty
                    !Paid = !Paid - remainingQty
                    .Update
                    remainingQty = 0
                    Exit Do
                Else
                    SettlingQty = !Paid
                    !ToPay = !ToPay + SettlingQty
                    !Paid = !Paid - SettlingQty
                    .Update
                    remainingQty = remainingQty - SettlingQty
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    
    
    
    Call ClearAddValues
    Call FormatGrid
    Call FillGrid
    cmbItem.SetFocus
End Sub

Private Sub cmbCollectingCenter_Change()
    With rsViewSuppliers
        If .State = 1 Then .Close
        temSQL = "Select * from tblSupplier where farmer  = 1  and Deleted = 0  and CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbSupplierName
        Set .RowSource = rsViewSuppliers
        .ListField = "Supplier"
        .BoundColumn = "SupplierID"
    End With
    If LastPaymentGeneratedDate(Val(cmbCollectingCenter.BoundText)) < dtpDate.Value Then
        btnAdd.Enabled = True
        btnDelete.Enabled = True
    Else
        btnAdd.Enabled = False
        btnDelete.Enabled = False
    End If
End Sub

Private Sub FillCombos()
    Dim Items As New clsFillCombos
    Items.FillAnyCombo cmbItem, "Item", True
    
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCollectingCenter, "CollectingCenter", True
End Sub

Private Sub cmbCollectingCenter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbSupplierName.SetFocus
    End If
End Sub

Private Sub cmbItem_Change()
    txtRate.Text = Empty
    With rsItemIssue
        If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblItem Where ItemID = " & Val(cmbItem.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!Value) = False Then txtRate.Text = Format(!Value, "0.00")
        End If
        .Close
    End With
    CalculateToPay
End Sub

Private Sub cmbItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtQty.SetFocus
    End If
End Sub

Private Sub cmbSupplierName_Change()
    CalculateToPay
    Call FormatGrid
    Call FillGrid
    Call FindPaymentsAndDeductions
End Sub


Private Sub cmbSupplierName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpFrom.SetFocus
    End If
End Sub

Private Sub dtpDate_Change()
    If LastPaymentGeneratedDate(Val(cmbCollectingCenter.BoundText)) < dtpDate.Value Then
        btnAdd.Enabled = True
        btnDelete.Enabled = True
    Else
        btnAdd.Enabled = False
        btnDelete.Enabled = False
    End If
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnAdd_Click
    End If
End Sub

Private Sub dtpFrom_Change()
    Call FormatGrid
    Call FillGrid
    Call FindPaymentsAndDeductions
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpTo.SetFocus
    End If
End Sub

Private Sub dtpTo_Change()
    Call FormatGrid
    Call FillGrid
    Call FindPaymentsAndDeductions
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbItem.SetFocus
    End If
End Sub

Private Sub Form_Load()
    If UserAuthorityLevel = OrdinaryUser Then
        btnDelete.Enabled = False
    End If
    FillCombos

    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = Date
    dtpDate.Value = Date
        
    Call FormatGrid
        
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

Private Sub CalculateValue()
    txtValue.Text = Val(txtRate.Text) * Val(txtQty.Text)
End Sub

Private Sub FindPaymentsAndDeductions()
    If IsNumeric(cmbSupplierName.BoundText) = False Then Exit Sub
    
    Dim Supplier As String
    Dim SupplierID As Long
    Dim MyMilkCollectionM As MilkCollection
    Dim MyMilkCollectionE As MilkCollection
    Dim temOwnCommision As Double
    Dim temOthersCommision As Double
    Dim temAdditionalCommision As Double
    Dim SupplierDeduction As Double
    Dim MinusValue As Boolean
    MinusValue = False

    Supplier = cmbSupplierName.Text
    SupplierID = Val(cmbSupplierName.BoundText)

    MyMilkCollectionM = PeriodMilkSupply(dtpFrom.Value, dtpTo.Value, SupplierID, 1)
    MyMilkCollectionE = PeriodMilkSupply(dtpFrom.Value, dtpTo.Value, SupplierID, 2)
    
    temOwnCommision = OwnCommision(SupplierID, dtpFrom.Value, dtpTo.Value, 0)
    temOthersCommision = OthersCommision(SupplierID, dtpFrom.Value, dtpTo.Value, 0)
    temAdditionalCommision = AdditionalCommision(SupplierID, dtpFrom.Value, dtpTo.Value)
    
    SupplierDeduction = PeriodDeductions(SupplierID, dtpFrom.Value, dtpTo.Value)
    
    If MyMilkCollectionM.Value + MyMilkCollectionE.Value - SupplierDeduction + temOwnCommision + temOthersCommision + temAdditionalCommision < 0 Then
        MinusValue = True
    End If
    
    txtTotalPayments.Text = Format(MyMilkCollectionM.Value + MyMilkCollectionE.Value + temOwnCommision + temOthersCommision + temAdditionalCommision, "0.00")
    txtTotalDeductions.Text = Format(SupplierDeduction, "0.00")
    txtNetPayments.Text = Format(MyMilkCollectionM.Value + MyMilkCollectionE.Value - SupplierDeduction + temOwnCommision + temOthersCommision + temAdditionalCommision, "0.00")
    
    If MinusValue = True Then
        btnAdd.Enabled = False
    Else
        btnAdd.Enabled = True
        If LastPaymentGeneratedDate(Val(cmbCollectingCenter.BoundText)) < dtpDate.Value Then
            btnAdd.Enabled = True
            btnDelete.Enabled = True
        Else
            btnAdd.Enabled = False
            btnDelete.Enabled = False
        End If
    End If
    
    
    
    
End Sub


Private Sub txtQty_Change()
    Call CalculateValue
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpDate.SetFocus
    End If
End Sub

Private Sub txtRate_Change()
    Call CalculateValue
End Sub

Private Sub CalculateToPay()
    Dim rsToPay As New ADODB.Recordset
    With rsToPay
        If .State = 1 Then .Close
        temSQL = "SELECT Sum(tblItemIssue.ToPay) AS SumOfToPay FROM tblItemIssue WHERE (((tblItemIssue.Deleted) = 0) AND ((tblItemIssue.ItemID)=" & Val(cmbItem.BoundText) & ") AND ((tblItemIssue.SupplierID)=" & Val(cmbSupplierName.BoundText) & "))"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If IsNull(!SumOfToPay) = False Then
            txtToPay.Text = !SumOfToPay
        Else
            txtToPay.Text = 0
        End If
        .Close
    End With
End Sub

Private Sub FillGrid()
    Dim rsDeductions As New ADODB.Recordset
    temSQL = "SELECT tblDeduction.DeductedDate, tblItem.Item, tblItem.ItemID, tblDeduction.Quentity, tblDeduction.Rate, tblDeduction.Value, tblDeduction.DeductionID " & _
                "FROM tblDeduction LEFT JOIN tblItem ON tblDeduction.ItemID = tblItem.ItemID " & _
                "WHERE (((tblDeduction.DeductedDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND ((tblDeduction.SupplierID)=" & Val(cmbSupplierName.BoundText) & ") AND ((tblDeduction.Deleted) = 0))"
    With rsDeductions
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                GridDeductions.Rows = GridDeductions.Rows + 1
                GridDeductions.row = GridDeductions.Rows - 1
                GridDeductions.col = 0
                GridDeductions.Text = Format(!DeductedDate, ShortDateFormat)
                GridDeductions.col = 1
                GridDeductions.Text = !Item
                GridDeductions.col = 2
                GridDeductions.Text = Format(!Quentity, "0.00")
                GridDeductions.col = 3
                GridDeductions.Text = Format(!Rate, "0.00")
                GridDeductions.col = 4
                GridDeductions.Text = Format(!Value, "0.00")
                GridDeductions.col = 5
                GridDeductions.Text = !DeductionID
                GridDeductions.col = 6
                GridDeductions.Text = !ItemID
               
                .MoveNext
            Wend
        End If
        .Close
    End With
End Sub

Private Sub FormatGrid()
    With GridDeductions
        .Clear
        .Rows = 1
        .Cols = 7
        
        .row = 0
        
        .col = 0
        .Text = "Date"
        
        .col = 1
        .Text = "Item"
        
        .col = 2
        .Text = "Quentity"
        
        .col = 3
        .Text = "Rate"
        
        .col = 4
        .Text = "Value"
        
        .col = 5
        .Text = "ID"
        
        .col = 5
        .Text = "Item ID"
        
        .ColWidth(5) = 1
        .ColWidth(6) = 1
        
    End With
    
End Sub
