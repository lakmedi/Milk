VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmItemPurchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Purchase"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   10680
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   8040
      Width           =   735
   End
   Begin VB.TextBox txtNet 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   7440
      TabIndex        =   26
      Text            =   "0.00"
      Top             =   7080
      Width           =   2055
   End
   Begin VB.TextBox txtDiscount 
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
      Left            =   7440
      TabIndex        =   7
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox txtGross 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   7440
      TabIndex        =   25
      Text            =   "0.00"
      Top             =   6120
      Width           =   2055
   End
   Begin VB.TextBox txtFQty 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   9720
      TabIndex        =   5
      Top             =   600
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   9720
      TabIndex        =   6
      Top             =   1080
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   9480
      TabIndex        =   13
      Top             =   8040
      Width           =   1095
      _ExtentX        =   1931
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
   Begin MSDataListLib.DataCombo cmbItemSupplierName 
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   6480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
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
   Begin MSFlexGridLib.MSFlexGrid gridItem 
      Height          =   4815
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8493
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowBigSelection=   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
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
      Left            =   120
      TabIndex        =   9
      Top             =   7320
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MM yyyy"
      Format          =   90112003
      CurrentDate     =   39682
   End
   Begin btButtonEx.ButtonEx btnUpdate 
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   8040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Update"
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
   Begin MSDataListLib.DataCombo cmbItem 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
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
   Begin MSDataListLib.DataCombo cmbCC 
      Height          =   360
      Left            =   120
      TabIndex        =   10
      Top             =   8160
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   ""
      Text            =   ""
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
   Begin VB.Label Label11 
      Caption         =   "Store"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Net"
      Height          =   255
      Left            =   5280
      TabIndex        =   24
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "Discount"
      Height          =   255
      Left            =   5280
      TabIndex        =   23
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Total"
      Height          =   255
      Left            =   5280
      TabIndex        =   22
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Free.Qty"
      Height          =   255
      Left            =   5880
      TabIndex        =   21
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Supplier &Name"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Da&te"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Item"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Qty"
      Height          =   255
      Left            =   4680
      TabIndex        =   17
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Rate"
      Height          =   255
      Left            =   7080
      TabIndex        =   16
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Value"
      Height          =   255
      Left            =   8400
      TabIndex        =   15
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmItemPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsItemPurchase As New ADODB.Recordset
    Dim rsViewItemSuppliers As New ADODB.Recordset
    Dim temSQL As String
        
Private Sub FormatGrid()
    With gridItem
        .Rows = 1
        .row = 0
        
        .Cols = 7
        
        .col = 0
        .ColWidth(0) = 600
        .Text = "No"
        
        .col = 1
        .ColWidth(1) = 3900
        .Text = "Item"
        
        .col = 2
        .ColWidth(2) = 1000
        .Text = "Quentity"
        
        .col = 3
        .ColWidth(3) = 1000
        .Text = "F.Quentity"
        
        .col = 4
        .ColWidth(4) = 1300
        .Text = "Rate"
        
         .col = 5
        .ColWidth(5) = 1500
        .Text = "Value"
        
        .col = 6
        .ColWidth(6) = 0
        .Text = "Item ID"
        
    End With
End Sub
    
Private Sub btnAdd_Click()
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

    With gridItem
        .Rows = .Rows + 1
        .row = .Rows - 1
        .col = 0
        .Text = .row
        .col = 1
        .Text = cmbItem.Text
        .col = 2
        .Text = Val(txtQty.Text)
        .col = 3
        .Text = Val(txtFQty.Text)
        .col = 4
        .Text = Val(txtRate.Text)
        .col = 5
        .Text = Format((Val(txtValue.Text)), "0.00")
        .col = 6
        .Text = Val(cmbItem.BoundText)
    End With
    
    Call CalculateTotal
    
    ClearAddValues
    cmbItem.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    Dim temRow As Integer
    With gridItem
        i = .row
        If i < 1 Then
            MsgBox "Nothing To Delete"
            Exit Sub
        End If
        .RemoveItem i
        For temRow = 0 To .Rows - 1
            .TextMatrix(temRow, 0) = temRow
        Next
    End With
        
    Call ClearAddValues
    Call CalculateTotal
    cmbItem.SetFocus
    
End Sub

Private Sub ClearAddValues()
    cmbItemSupplierName.Text = Empty
    cmbCC.Text = Empty
    cmbItem.Text = Empty
    txtRate.Text = Empty
    txtValue.Text = Empty
    txtQty.Text = Empty
    txtFQty.Text = Empty
End Sub

Private Sub FillCombos()
    Dim Items As New clsFillCombos
    Items.FillAnyCombo cmbItem, "Item", True
    
    Dim sup As New clsFillCombos
    sup.FillAnyCombo cmbItemSupplierName, "ItemSupplier", True
    
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCC, "CollectingCenter", True
    
End Sub

Private Sub btnDelete_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbItemSupplierName.SetFocus
        SendKeys "{down}"
    End If
End Sub

Private Sub btnUpdate_Click()
    Dim i As Integer
    Dim temPurchaseBillID As Long
    Dim rsTemItemPurchase As New ADODB.Recordset
    Dim rsTemItemPurchaseBill As New ADODB.Recordset
    Dim rsStock As New ADODB.Recordset
    
    
'    If IsNumeric(cmbItemSupplierName.BoundText) = False Then
'        MsgBox "Item Supplier?"
'        cmbItemSupplierName.SetFocus
'        Exit Sub
'    End If
    
'    If IsNumeric(dtpDate.Value) = False Then
'        MsgBox "Purchase Date?"
'        dtpDate.SetFocus
'        Exit Sub
'    End If
    If IsNumeric(cmbCC.BoundText) = False Then
        MsgBox "Collecting Center?"
        cmbCC.SetFocus
        Exit Sub
    End If
    
    With rsTemItemPurchaseBill
        If .State = 1 Then .Close
        temSQL = "Select * from tblItemPurchaseBIll where ItemPurchaseBIllID =0"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !GrossTotal = txtGross.Text
        If IsNumeric(!Discount) Then txtDiscount.Text = !Discount
        If IsNumeric(!NetTotal) Then txtNet.Text = !NetTotal
        .Update
        temPurchaseBillID = !ItemPurchaseBillID
    End With
    
    With gridItem
        For i = 1 To .Rows - 1
            If rsItemPurchase.State = 1 Then rsItemPurchase.Close
                temSQL = "Select * from tblItemPurchase where ItemPurchaseID =0"
                rsItemPurchase.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                rsItemPurchase.AddNew
                rsItemPurchase!ItemPurchaseBillID = temPurchaseBillID
                rsItemPurchase!ItemID = Val(.TextMatrix(i, 6))
'                rsItemPurchase!Item = Val(.TextMatrix(i, 1))
                rsItemPurchase!Quentity = Val(.TextMatrix(i, 2))
                rsItemPurchase!FreeQuentity = Val(.TextMatrix(i, 3))
                rsItemPurchase!Rate = Val(.TextMatrix(i, 4))
                rsItemPurchase!Value = Val(.TextMatrix(i, 5))
                rsItemPurchase.Update
            
            If rsStock.State = 1 Then rsStock.Close
                temSQL = "Select * from tblItemStock where CollectingCenterID = " & Val(cmbCC.BoundText)
                rsStock.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If rsStock.RecordCount < 1 Then
                rsStock.AddNew
                rsStock!CollectingCenterID = Val(cmbCC.BoundText)
                rsStock!ItemID = Val(.TextMatrix(i, 6))
                rsStock!Stock = Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 3))
                rsStock.Update
            Else
                rsStock!Stock = rsStock!Stock + (Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 3)))
                rsStock.Update
            End If
            
        Next
    End With
    
    Call ClearAddValues
    Call ClearUpdateValues
    
    cmbItem.SetFocus
    
End Sub

Private Sub ClearUpdateValues()

End Sub

Private Sub cmbItem_Change()
    txtRate.Text = Empty
    With rsItemPurchase
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
        SendKeys "{down}"
    End If
End Sub


Private Sub cmbItemSupplierName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtpDate.SetFocus
        SendKeys "{down}"
    End If
End Sub

'Private Sub dtpDate_Change()
'    Call FillGrid
'End Sub

Private Sub Form_Load()
    dtpDate.Value = Date
    Call FillCombos
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

Private Sub txtDiscount_Change()
    Call CalculateNetTotal
End Sub

Private Sub txtFQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnAdd.SetFocus
        SendKeys "{down}"
    End If
End Sub

Private Sub txtGross_Change()
    Call CalculateNetTotal
End Sub

Private Sub txtQty_Change()
    Call CalculateValue
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtFQty.SetFocus
        SendKeys "{down}"
    End If
End Sub

Private Sub txtRate_Change()
    Call CalculateValue
End Sub

Private Sub CalculateTotal()
    Dim i As Integer
    Dim TemTotal As Double
    With gridItem
        For i = 0 To .Rows - 1
            TemTotal = TemTotal + Val(.TextMatrix(i, 5))
        Next
    End With
    txtGross.Text = Format(TemTotal, "0.00")
End Sub

Private Sub CalculateNetTotal()
        txtNet.Text = Format((Val(txtGross.Text) - Val(txtDiscount.Text)), "0.00")
End Sub

Private Sub txtValue_Change()
    txtValue.Text = Format(txtValue, "0.00")
End Sub
