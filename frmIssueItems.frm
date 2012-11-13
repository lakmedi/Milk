VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmIssueItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Item Issues"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13245
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
   ScaleHeight     =   8445
   ScaleWidth      =   13245
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Paid"
      TabPicture(0)   =   "frmIssueItems.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btnDelete"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btnAdd"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmbItem"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "gridItem"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbSupplierName"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtRate"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtValue"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtQty"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "To Pay"
      TabPicture(1)   =   "frmIssueItems.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "gridToPay"
      Tab(1).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid gridToPay 
         Height          =   5895
         Left            =   -74760
         TabIndex        =   21
         Top             =   480
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   10398
         _Version        =   393216
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   8400
         TabIndex        =   11
         Top             =   780
         Width           =   855
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   780
         Width           =   975
      End
      Begin MSDataListLib.DataCombo cmbSupplierName 
         Height          =   360
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSFlexGridLib.MSFlexGrid gridItem 
         Height          =   4815
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8493
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowBigSelection=   -1  'True
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSDataListLib.DataCombo cmbItem 
         Height          =   360
         Left            =   5160
         TabIndex        =   9
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin btButtonEx.ButtonEx btnAdd 
         Height          =   375
         Left            =   11760
         TabIndex        =   16
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   12583104
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
      Begin btButtonEx.ButtonEx btnDelete 
         Height          =   375
         Left            =   11760
         TabIndex        =   18
         Top             =   1440
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   12583104
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
      Begin VB.Label Label6 
         Caption         =   "Farmer &Name"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Item"
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Qty"
         Height          =   375
         Left            =   8400
         TabIndex        =   10
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Rate"
         Height          =   375
         Left            =   9360
         TabIndex        =   12
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Value"
         Height          =   375
         Left            =   10440
         TabIndex        =   14
         Top             =   420
         Width           =   975
      End
   End
   Begin VB.CheckBox chkByCode 
      Caption         =   "By Code"
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   11760
      TabIndex        =   20
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   12583104
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   12583104
      CalendarTitleForeColor=   12583104
      CustomFormat    =   "dd MM yyyy"
      Format          =   293470211
      CurrentDate     =   39682
   End
   Begin MSDataListLib.DataCombo cmbCollectingCenter 
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   10440
      TabIndex        =   19
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   12583104
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
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Da&te"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "frmIssueItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsItemIssue As New ADODB.Recordset
    Dim rsViewSuppliers As New ADODB.Recordset
    Dim rsViewSupplier As New ADODB.Recordset
    Dim temSQL As String
    
Private Sub btnAdd_Click()
    If IsNumeric(cmbSupplierName.BoundText) = False Then
        MsgBox "Farmer?"
        cmbSupplierName.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbItem.BoundText) = False Then
        MsgBox "Item"
        cmbItem.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtQty.Text) = False Then
        MsgBox "Quentity"
        txtQty.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    With rsItemIssue
        If .State = 1 Then .Close
        temSQL = "Select * from tblItemIssue where ItemIssueID =0"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ItemID = Val(cmbItem.BoundText)
        !SupplierID = Val(cmbSupplierName.BoundText)
        !Quentity = Val(txtQty.Text)
        !ToPay = Val(txtQty.Text)
        !Value = Val(txtValue.Text)
        !Rate = Val(txtRate.Text)
        !IssuedUserID = UserID
        !IssueDate = Format(dtpDate.Value, "dd MMMM yyyy")
        .Update
        .Close
    End With
    ClearAddValues
    FillGrid
    cmbSupplierName.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    With gridItem
        i = .row
        If i < 1 Then
            MsgBox "Nothing To Delete"
            Exit Sub
        End If
        If IsNumeric(.TextMatrix(i, 7)) = False Then
            MsgBox "nothing to delete"
            Exit Sub
        End If
        If Val(.TextMatrix(i, 6)) > 0 Then
            MsgBox "The Items issued is already deducted from the farmer. You can't delete this"
            Exit Sub
        End If
        Dim ItemIssueID As Long
    End With
    With rsItemIssue
        If .State = 1 Then .Close
        temSQL = "Select * from tblItemIssue where ItemIssueID = " & Val(gridItem.TextMatrix(i, 7))
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedUserID = UserID
'            !DeletedDate = Date
'            !DeletedTime = Time
            
            .Update
        Else
            MsgBox "Error deleting"
        End If
        .Close
    End With
    Call ClearAddValues
    Call FillGrid
    cmbSupplierName.SetFocus
End Sub

Private Sub btnPrint_Click()
    Dim tabReport As Long
    Dim tab1 As Long
    Dim tab2 As Long

    tabReport = 70
    tab1 = 5
    tab2 = 40
    
    Printer.Print
    Printer.Font.Bold = True
    Printer.Print Tab(tabReport); "Items Issue Report"
    Printer.Font.Bold = False
    Printer.Print
    Printer.Print Tab(tab1); "Collecting Center :";
    Printer.Print Tab(tab2); cmbCollectingCenter.Text
    Printer.Print Tab(tab1); "Date  :";
    Printer.Print Tab(tab2); dtpDate.Value;
    Printer.Print

    Dim i As Integer
    Dim tabNo As Long
    Dim tabSupplier As Long
    Dim tabItem As Long
    Dim tabQuentity As Long
    Dim tabRate As Long
    Dim tabValue As Long
    Dim tabDeducted As Long
    
    tabNo = 10
    tabSupplier = 20
    tabItem = 60
    tabQuentity = 115
    tabRate = 130
    tabValue = 145
    tabDeducted = 160

    With gridItem
        For i = 0 To .Rows - 1
            Printer.Print
            Printer.Print Tab(tabNo - Len(.TextMatrix(i, 0))); .TextMatrix(i, 0);
            Printer.Print Tab(tabSupplier); .TextMatrix(i, 1);
            Printer.Print Tab(tabItem); .TextMatrix(i, 2);
            Printer.Print Tab(tabQuentity - Len(.TextMatrix(i, 3))); .TextMatrix(i, 3);
            Printer.Print Tab(tabRate - Len(.TextMatrix(i, 4))); .TextMatrix(i, 4);
            Printer.Print Tab(tabValue - Len(.TextMatrix(i, 5))); .TextMatrix(i, 5);
            Printer.Print Tab(tabDeducted - Len(.TextMatrix(i, 6))); .TextMatrix(i, 6);
            Printer.Print
        Next
    End With

   Printer.EndDoc
End Sub

Private Sub FillToPay()
    cmbSupplierName.Text = Empty
    cmbItem.Text = Empty
    txtRate.Text = Empty
    txtValue.Text = Empty
    txtQty.Text = Empty
End Sub

Private Sub FillToPayGrid()
    With rsItemIssue
        If .State = 1 Then .Close
        temSQL = "SELECT tblSupplier.Supplier, tblSupplier.SupplierCode, tblItem.Item, tblItemIssue.Rate , sum(tblItemIssue.Quentity) as MQuentity, sum(tblItemIssue.ToPay) as MToPay, sum(tblItemIssue.Value) as MValue, sum(tblItemIssue.Paid) as MPaid " & _
                    "FROM tblCollectingCenter LEFT JOIN ((tblItem RIGHT JOIN tblItemIssue ON tblItem.ItemID = tblItemIssue.ItemID) RIGHT JOIN tblSupplier ON tblItemIssue.SupplierID = tblSupplier.SupplierID) ON tblCollectingCenter.CollectingCenterID = tblSupplier.CollectingCenterID " & _
                    "Where (((tblCollectingCenter.CollectingCenterID) = " & Val(cmbCollectingCenter.BoundText) & ") And ((tblItemIssue.Deleted) = 0 )   And ((tblItemIssue.ToPay) > 0 )) " & _
                    "GROUP BY tblSupplier.Supplier, tblSupplier.SupplierCode, tblItem.Item, tblItemIssue.Rate"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        With gridToPay
            .Rows = 1
            .Cols = 10
            .Clear
            
            .row = 0
            
            .ColWidth(1) = 2600
            .ColWidth(2) = 2200
            .ColWidth(7) = 0
            .ColWidth(4) = 0
            
            .col = 0
            .Text = "No."
            
            .col = 1
            .Text = "Supplier"
            
            .col = 2
            .Text = "Item"
            
            .col = 3
            .Text = "Quentity"
            
            
            .col = 5
            .Text = "Value"
            
            .col = 6
            .Text = "Deducted"
            
            .col = 7
            .Text = "ItemIssueID"
            
            .col = 8
            .Text = "To Pay QTY"
            
            .col = 9
            .Text = "To Pay Value"
            
            
        End With
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Dim i As Integer
            gridToPay.Rows = .RecordCount + 1
            For i = 1 To .RecordCount
                gridToPay.TextMatrix(i, 0) = i
                If IsNull(!Supplier) = False Then gridToPay.TextMatrix(i, 1) = !Supplier & " " & !SupplierCode
                If IsNull(!Item) = False Then gridToPay.TextMatrix(i, 2) = !Item
                If IsNull(!mQuentity) = False Then gridToPay.TextMatrix(i, 3) = Format(!mQuentity, "0.00")
                If IsNull(!mValue) = False Then gridToPay.TextMatrix(i, 5) = Format(!mValue, "0.00")
                If IsNull(!mPaid) = False Then gridToPay.TextMatrix(i, 6) = Format(!mPaid, "0.00")
                If IsNull(!mToPay) = False Then gridToPay.TextMatrix(i, 8) = Format(!mToPay, "0.00")
                If IsNull(!mToPay) = False And IsNull(!Rate) = False Then gridToPay.TextMatrix(i, 9) = Format(!mToPay * !Rate, "0.00")
                .MoveNext
            Next
        End If
        
    End With
End Sub


Private Sub chkByCode_Click()
    cmbCollectingCenter_Change
End Sub

Private Sub cmbCollectingCenter_Change()

    If chkByCode.Value = 0 Then
        With rsViewSupplier
            If .State = 1 Then .Close
            temSQL = "Select ( Supplier + ' (' + SupplierCode + ')') as Display, SupplierID from tblSupplier where CollectingCenterID = " & Val(cmbCollectingCenter.BoundText) & " And Deleted = 0  Order by Supplier"
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        End With
        With cmbSupplierName
            Set .RowSource = rsViewSupplier
            .ListField = "Display"
            .BoundColumn = "SupplierID"
            .Text = Empty
        End With
    Else
        With rsViewSupplier
            If .State = 1 Then .Close
            temSQL = "Select (SupplierCode + ' (' + Supplier + ')' ) as Display, * from tblSupplier where CollectingCenterID = " & Val(cmbCollectingCenter.BoundText) & " And Deleted = 0  Order by SupplierCode"
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        End With
        With cmbSupplierName
            Set .RowSource = rsViewSupplier
            .ListField = "Display"
            .BoundColumn = "SupplierID"
            .Text = Empty
        End With
    End If
    
    
    Call FillGrid
    
    Call FillToPayGrid
    
    EnableEdit
    
End Sub

Private Sub EnableEdit()
    Dim rsCCSummery As New ADODB.Recordset
    With rsCCSummery
        If .State = 1 Then .Close
        temSQL = "SELECT Max(tblCollectingCenterPaymentSummery.ToDate) AS MaxOfToDate " & _
                    "From tblCollectingCenterPaymentSummery " & _
                    "WHERE tblCollectingCenterPaymentSummery.CollectingCenterID =" & Val(cmbCollectingCenter.BoundText)
               .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If IsNull(!MaxOfToDate) = False Then
            btnAdd.Enabled = True
            btnDelete.Enabled = True
        Else
            If !MaxOfToDate < dtpDate.Value Then
                btnAdd.Enabled = False
                btnDelete.Enabled = False
            Else
                btnAdd.Enabled = True
                btnDelete.Enabled = True
            End If
        End If
    End With
End Sub

Private Sub ClearAddValues()
    cmbSupplierName.Text = Empty
    cmbItem.Text = Empty
    txtRate.Text = Empty
    txtValue.Text = Empty
    txtQty.Text = Empty
End Sub

Private Sub FillGrid()
    With rsItemIssue
        If .State = 1 Then .Close
        If chkByCode.Value = 0 Then
            temSQL = "SELECT tblSupplier.Supplier + '(' + tblSupplier.SupplierCode + ')' as MySupplier  , tblItem.Item, tblItemIssue.Quentity, tblItemIssue.ToPay, tblItemIssue.Rate, tblItemIssue.Value, tblItemIssue.Paid , tblItemIssue.ItemIssueID " & _
                        "FROM tblCollectingCenter LEFT JOIN ((tblItem RIGHT JOIN tblItemIssue ON tblItem.ItemID = tblItemIssue.ItemID) RIGHT JOIN tblSupplier ON tblItemIssue.SupplierID = tblSupplier.SupplierID) ON tblCollectingCenter.CollectingCenterID = tblSupplier.CollectingCenterID " & _
                        "Where (((tblCollectingCenter.CollectingCenterID) = " & Val(cmbCollectingCenter.BoundText) & ") And ((tblItemIssue.Deleted) = 0 ) AND ((tblItemIssue.IssueDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') )"
            
        Else
            temSQL = "SELECT tblSupplier.SupplierCode + '(' + tblSupplier.Supplier + ')' as MySupplier ,  tblItem.Item, tblItemIssue.Quentity, tblItemIssue.ToPay, tblItemIssue.Rate, tblItemIssue.Value, tblItemIssue.Paid , tblItemIssue.ItemIssueID " & _
                        "FROM tblCollectingCenter LEFT JOIN ((tblItem RIGHT JOIN tblItemIssue ON tblItem.ItemID = tblItemIssue.ItemID) RIGHT JOIN tblSupplier ON tblItemIssue.SupplierID = tblSupplier.SupplierID) ON tblCollectingCenter.CollectingCenterID = tblSupplier.CollectingCenterID " & _
                        "Where (((tblCollectingCenter.CollectingCenterID) = " & Val(cmbCollectingCenter.BoundText) & ") And ((tblItemIssue.Deleted) = 0 ) AND ((tblItemIssue.IssueDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') )"
                    
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        With gridItem
            .Rows = 1
            .Cols = 10
            .Clear
            
            .row = 0
            
            .ColWidth(1) = 2600
            .ColWidth(2) = 2200
            .ColWidth(7) = 1
            
            .col = 0
            .Text = "No."
            
            .col = 1
            .Text = "Supplier"
            
            .col = 2
            .Text = "Item"
            
            .col = 3
            .Text = "Quentity"
            
            .col = 4
            .Text = "Rate"
            
            .col = 5
            .Text = "Value"
            
            .col = 6
            .Text = "Deducted"
            
            .col = 7
            .Text = "ItemIssueID"
            
            .col = 8
            .Text = "To Pay QTY"
            
            .col = 9
            .Text = "To Pay Value"
            
            
        End With
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Dim i As Integer
            gridItem.Rows = .RecordCount + 1
            For i = 1 To .RecordCount
                gridItem.TextMatrix(i, 0) = i
                If IsNull(!MySupplier) = False Then gridItem.TextMatrix(i, 1) = !MySupplier
                If IsNull(!Item) = False Then gridItem.TextMatrix(i, 2) = !Item
                If IsNull(!Quentity) = False Then gridItem.TextMatrix(i, 3) = !Quentity
                If IsNull(!Rate) = False Then gridItem.TextMatrix(i, 4) = Format(!Rate, "0.00")
                If IsNull(!Value) = False Then gridItem.TextMatrix(i, 5) = Format(!Value, "0.00")
                If IsNull(!Paid) = False Then gridItem.TextMatrix(i, 6) = Format(!Paid, "0.00")
                If IsNull(!ToPay) = False Then gridItem.TextMatrix(i, 8) = !ToPay
                If IsNull(!ToPay) = False And IsNull(!Rate) = False Then gridItem.TextMatrix(i, 9) = Format(!ToPay * !Rate, "0.00")
                If IsNull(!ItemIssueID) = False Then gridItem.TextMatrix(i, 7) = !ItemIssueID
                .MoveNext
            Next
        End If
        
    End With
End Sub

Private Sub FillCombos()
    Dim Items As New clsFillCombos
    Items.FillAnyCombo cmbItem, "Item", True
    
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCollectingCenter, "CollectingCenter", True
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
End Sub

Private Sub cmbItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtQty.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbItem.Text = Empty
    End If
End Sub

Private Sub cmbSupplierName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbItem.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbSupplierName.Text = Empty
    End If
End Sub

Private Sub dtpDate_Change()
    Call FillGrid
    Call EnableEdit
End Sub

Private Sub Form_Load()
    If UserAuthorityLevel = OrdinaryUser Then
        btnDelete.Enabled = False
    End If
    dtpDate.Value = Date
    Call FillCombos
    
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
    
    If IssuePaymentDeleteAllowed = True Then
       btnDelete.Visible = True
    Else
        btnDelete.Visible = False
    End If
    dtpDate.Value = Date
End Sub

Private Sub CalculateValue()
    txtValue.Text = Val(txtRate.Text) * Val(txtQty.Text)
End Sub

Private Sub txtQty_Change()
    Call CalculateValue
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnAdd_Click
    End If
End Sub

Private Sub txtRate_Change()
    Call CalculateValue
End Sub
