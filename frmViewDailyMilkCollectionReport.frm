VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmViewDailyMilkCollectionReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Milk Collection Report"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
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
   ScaleHeight     =   8370
   ScaleWidth      =   11880
   Begin VB.TextBox txtTotalValueDisplay 
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
      Height          =   375
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CheckBox chkByCode 
      Caption         =   "By Code"
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtValueDifference 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtTSNF 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtCValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtTotalLMR 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtTotalFAT 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtTFAT 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6480
      TabIndex        =   36
      Top             =   7080
      Width           =   1335
   End
   Begin VB.TextBox txtTLMR 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6480
      TabIndex        =   34
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox txtCFAT 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   7080
      Width           =   1335
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtSNF 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtAverageValue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox txtCLMR 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox txtTotalValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox txtTotalLiters 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   6120
      Width           =   1335
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   9600
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
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
      Left            =   9600
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   10560
      TabIndex        =   42
      Top             =   7800
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
   Begin VB.TextBox txtFatXLiters 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtLMRXLiters 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtLiters 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7800
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtFat 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6360
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtLMR 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo cmbSupplierName 
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid gridMilk 
      Height          =   4815
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8493
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MM yyyy"
      Format          =   180813827
      CurrentDate     =   39682
   End
   Begin MSDataListLib.DataCombo cmbCollectingCenter 
      Height          =   360
      Left            =   1920
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
   Begin MSDataListLib.DataCombo cmbSecession 
      Height          =   360
      Left            =   4560
      TabIndex        =   5
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   9240
      TabIndex        =   41
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
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
   Begin VB.TextBox txtTFAT1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtTLMR1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label19 
      Caption         =   "Average Per Leter"
      Height          =   375
      Left            =   8040
      TabIndex        =   39
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label Label18 
      Caption         =   "Tested Average of FAT%"
      Height          =   375
      Left            =   4200
      TabIndex        =   35
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label Label17 
      Caption         =   "Tested Average of LMR"
      Height          =   375
      Left            =   4200
      TabIndex        =   33
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label16 
      Caption         =   "Calculated Average of FAT%"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Value"
      Height          =   375
      Left            =   7800
      TabIndex        =   24
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "SNF"
      Height          =   375
      Left            =   4920
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Price"
      Height          =   375
      Left            =   6360
      TabIndex        =   22
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Total Value"
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
      Left            =   8040
      TabIndex        =   37
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label Label11 
      Caption         =   "Calculated Average of LMR"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label10 
      Caption         =   "Total Leters"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Fat % X Liters"
      Height          =   375
      Left            =   3480
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "LMR  X Liters"
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "&Liters"
      Height          =   375
      Left            =   7800
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Supplier &Name"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "L&MR"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "&Fat %"
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Secess&ion"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Da&te"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmViewDailyMilkCollectionReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    Dim i As Integer
    
    Dim rsCollection As New ADODB.Recordset
    Dim rsDailyCollection As New ADODB.Recordset
    
    Dim rsViewCollectingCenter As New ADODB.Recordset
    Dim rsViewSupplier As New ADODB.Recordset
    Dim rsViewSecession As New ADODB.Recordset
    
    Dim NewRecord As Boolean
    
    
Private Sub DisplayDailyCollectionDetails()
    txtTLMR.Text = Empty
    txtTFAT.Text = Empty
    txtTLMR1.Text = Empty
    txtTFAT1.Text = Empty
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then Exit Sub
    If IsNumeric(cmbSecession.BoundText) = False Then Exit Sub
    With rsDailyCollection
        If .State = 1 Then .Close
        If cmbSecession.BoundText = 1 Then
            temSQL = "SELECT tblSecessionCollection.* FROM tblSecessionCollection Where ProgramDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' And SecessionID = " & Val(cmbSecession.BoundText) & " And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
        Else
            temSQL = "SELECT tblSecessionCollection.* FROM tblSecessionCollection Where ProgramDate = '" & Format(dtpDate.Value + 1, "dd MMMM yyyy") & "' And SecessionID = " & Val(cmbSecession.BoundText) & " And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            txtTLMR.Text = !TestedLMR
            txtTFAT.Text = !TestedFAT
            txtTLMR1.Text = !TestedLMR
            txtTFAT1.Text = !TestedFAT
            '.Update
        Else
            txtTLMR.Text = Empty
            txtTFAT.Text = Empty
            txtTLMR1.Text = Empty
            txtTFAT1.Text = Empty
        End If
        .Close
    End With
End Sub


Private Sub btnAdd_Click()
    Dim temCr As Double
    Dim temC As Double
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then
        MsgBox "No Collecting Center"
        cmbCollectingCenter.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbSecession.BoundText) = False Then
        MsgBox "No Secession"
        cmbSecession.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbSupplierName.BoundText) = False Then
        MsgBox "No Supplier"
        cmbSupplierName.SetFocus
        Exit Sub
    End If
    If Val(txtLiters.Text) <= 0 Then
        MsgBox "You have not entered the volume"
        txtLiters.SetFocus
        Exit Sub
    End If
    If Val(txtFat.Text) < 2.5 Or Val(txtFat.Text) > 8 Then
        MsgBox "The FAT value you entered is not correct. Please reckeck"
        txtFat.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If Val(txtLMR.Text) < 23 Or Val(txtLMR.Text) > 32 Then
        MsgBox "The LMR value you entered is not correct. Please reckeck"
        txtLMR.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If Val(txtPrice.Text) <= 0 Then
        MsgBox "The values you entered is not correct. Please reckeck"
        txtFat.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    Dim i As Integer
    With gridMilk
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 11)) = Val(cmbSupplierName.BoundText) Then
                MsgBox "Details of this supplier is already entered"
                cmbSupplierName.SetFocus
                Exit Sub
            End If
        Next
    End With
    With rsCollection
        If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblCollection where CollectionID =0"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Date = Format(dtpDate.Value, "dd MMMM yyyy")
        !SecessionID = cmbSecession.BoundText
        If cmbSecession.BoundText = 1 Then ' Morning
            !CollectionDate = dtpDate.Value
            !DeliveryDate = dtpDate.Value + 1
            !ProgramDate = dtpDate.Value
        Else
            !CollectionDate = dtpDate.Value
            !DeliveryDate = dtpDate.Value + 2
            !ProgramDate = dtpDate.Value + 1
        End If
        !CollectingCenterID = cmbCollectingCenter.BoundText
        !SupplierID = cmbSupplierName.BoundText
        !LMR = Val(txtLMR.Text)
        !FAT = Val(txtFat.Text)
        !Liters = Val(txtLiters.Text)
        !LMRXLiters = Val(txtLMRXLiters.Text)
        !FATXLiters = Val(txtFatXLiters.Text)
        !SNF = Val(txtSNF.Text)
        !Price = Val(txtPrice.Text)
        !Value = Val(txtValue.Text)
        !AddedUserID = UserID
        !AddedDate = Format(Date, "dd MMMM yyyy")
        !AddedTime = Time
        !Deleted = False
        temCr = OwnCommisionRate(cmbSupplierName.BoundText, Val(txtLiters.Text))
        !CommisionRate = temCr
        !Commision = temCr * Val(txtLiters)
        .Update
        .Close
    End With
    Call FormatGrid
    Call FillGrid
    
    
    Call CalculateTotals
    Call WriteDailyCollection
    
    
    'Call ClearValues
    
    
    ' *********************
    
    txtLMR.Text = Empty
    txtFat.Text = Empty
    txtLiters.Text = Empty
    txtLMRXLiters.Text = Empty
    txtFatXLiters.Text = Empty
    txtSNF.Text = Empty
    txtPrice.Text = Empty
    txtValue.Text = Empty
    
    If gridMilk.Rows > 9 Then
        gridMilk.TopRow = gridMilk.Rows - 9
    End If
    
    
    cmbSupplierName.SetFocus
'    SendKeys "{down}"
    
    
    
    ' *********************
    

End Sub

Private Sub WriteDailyCollection()
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then Exit Sub
    If IsNumeric(cmbSecession.BoundText) = False Then Exit Sub
    
    Dim TotalVolume As Double
    Dim TotalLMR As Double
    Dim TotalFAT As Double
    Dim TestedLMR As Double
    Dim TestedFAT As Double
    Dim totalValue As Double
    Dim TotalActualValue As Double
    Dim TotalValueDifference As Double
    
    With rsDailyCollection
        If .State = 1 Then .Close
        If cmbSecession.BoundText = 1 Then
            temSQL = "SELECT tblSecessionCollection.* FROM tblSecessionCollection Where ProgramDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' And SecessionID = " & Val(cmbSecession.BoundText) & " And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
        Else
            temSQL = "SELECT tblSecessionCollection.* FROM tblSecessionCollection Where ProgramDate = '" & Format(dtpDate.Value + 1, "dd MMMM yyyy") & "' And SecessionID = " & Val(cmbSecession.BoundText) & " And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 1 Then
            If .State = 1 Then .Close
            If cmbSecession.BoundText = 1 Then
                temSQL = "Delete tblSecessionCollection.* FROM tblSecessionCollection Where ProgramDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' And SecessionID = " & Val(cmbSecession.BoundText) & " And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
            Else
                temSQL = "Delete tblSecessionCollection.* FROM tblSecessionCollection Where ProgramDate = '" & Format(dtpDate.Value + 1, "dd MMMM yyyy") & "' And SecessionID = " & Val(cmbSecession.BoundText) & " And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
            End If
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If .State = 1 Then .Close
            If cmbSecession.BoundText = 1 Then
                temSQL = "SELECT tblSecessionCollection.* FROM tblSecessionCollection Where ProgramDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' And SecessionID = " & Val(cmbSecession.BoundText) & " And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
            Else
                temSQL = "SELECT tblSecessionCollection.* FROM tblSecessionCollection Where ProgramDate = '" & Format(dtpDate.Value + 1, "dd MMMM yyyy") & "' And SecessionID = " & Val(cmbSecession.BoundText) & " And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
            End If
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !Date = Format(dtpDate.Value, "dd MMMM yyyy")
            !SecessionID = cmbSecession.BoundText
            If cmbSecession.BoundText = 1 Then ' Morning
                !CollectionDate = dtpDate.Value
                !DeliveryDate = dtpDate.Value + 1
                !ProgramDate = dtpDate.Value
            Else
                !CollectionDate = dtpDate.Value
                !DeliveryDate = dtpDate.Value + 2
                !ProgramDate = dtpDate.Value + 1
            End If
            !CollectingCenterID = cmbCollectingCenter.BoundText
        ElseIf .RecordCount < 1 Then
            .AddNew
            !Date = Format(dtpDate.Value, "dd MMMM yyyy")
            !SecessionID = cmbSecession.BoundText
            If cmbSecession.BoundText = 1 Then ' Morning
                !CollectionDate = dtpDate.Value
                !DeliveryDate = dtpDate.Value + 1
                !ProgramDate = dtpDate.Value
            Else
                !CollectionDate = dtpDate.Value
                !DeliveryDate = dtpDate.Value + 2
                !ProgramDate = dtpDate.Value + 1
            End If
            !CollectingCenterID = cmbCollectingCenter.BoundText
        End If
        !TotalVolume = Val(txtTotalLiters.Text)
        !TotalLMR = Val(txtTotalLMR.Text)
        !TotalFAT = Val(txtTotalFAT.Text)
        !totalValue = Val(txtTotalValue.Text)
        !CalculatedLMR = Val(txtCLMR.Text)
        !CalculatedFat = Val(txtCFAT.Text)
        !TestedLMR = Val(txtTLMR.Text)
        !TestedFAT = Val(txtTFAT.Text)
        !AverageValue = Val(txtAverageValue.Text)
        !actualValue = Val(txtCValue.Text)
        !ValueDifference = Val(txtValueDifference.Text)
        .Update
        .Close
    End With
    
    With rsDailyCollection
        If .State = 1 Then .Close
        If cmbSecession.BoundText = 1 Then
            temSQL = "SELECT Sum(tblSecessionCollection.TotalVolume) AS SumOfTotalVolume, Sum(tblSecessionCollection.TotalLMR) AS SumOfTotalLMR, Sum(tblSecessionCollection.TotalFAT) AS SumOfTotalFAT, Sum(tblSecessionCollection.ActualValue) AS SumOfActualValue, Sum(tblSecessionCollection.ValueDifference) AS SumOfValueDifference,  Sum(tblSecessionCollection.TotalValue) AS SumOfTotalValue, Avg(tblSecessionCollection.TestedLMR) AS AvgOfTestedLMR, Avg(tblSecessionCollection.TestedFAT) AS AvgOfTestedFAT " & _
                        "FROM tblSecessionCollection Where ProgramDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
        Else
            temSQL = "SELECT Sum(tblSecessionCollection.TotalVolume) AS SumOfTotalVolume, Sum(tblSecessionCollection.TotalLMR) AS SumOfTotalLMR, Sum(tblSecessionCollection.TotalFAT) AS SumOfTotalFAT, Sum(tblSecessionCollection.ActualValue) AS SumOfActualValue, Sum(tblSecessionCollection.ValueDifference) AS SumOfValueDifference,  Sum(tblSecessionCollection.TotalValue) AS SumOfTotalValue, Avg(tblSecessionCollection.TestedLMR) AS AvgOfTestedLMR, Avg(tblSecessionCollection.TestedFAT) AS AvgOfTestedFAT " & _
                        "FROM tblSecessionCollection Where ProgramDate = '" & Format(dtpDate.Value + 1, "dd MMMM yyyy") & "' And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            TotalVolume = !SumOfTotalVolume
            TotalLMR = !SumOfTotalLMR
            TotalFAT = !SumOfTotalFAT
            TestedLMR = !AvgOfTestedLMR
            TestedFAT = !AvgOfTestedFAT
            totalValue = !SumOfTotalValue
            TotalActualValue = !SumOfActualValue
            TotalValueDifference = !SumOfValueDifference
        End If
        .Close
    End With
    
    With rsDailyCollection
        If .State = 1 Then .Close
        If cmbSecession.BoundText = 1 Then
            temSQL = "SELECT tblDailyCollection.* FROM tblDailyCollection Where ProgramDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
        Else
            temSQL = "SELECT tblDailyCollection.* FROM tblDailyCollection Where ProgramDate = '" & Format(dtpDate.Value + 1, "dd MMMM yyyy") & "' And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 1 Then
            If .State = 1 Then .Close
            If cmbSecession.BoundText = 1 Then
                temSQL = "Delete tblDailyCollection.* FROM tblDailyCollection Where ProgramDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
            Else
                temSQL = "Delete tblDailyCollection.* FROM tblDailyCollection Where ProgramDate = '" & Format(dtpDate.Value + 1, "dd MMMM yyyy") & "' And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
            End If
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If .State = 1 Then .Close
            If cmbSecession.BoundText = 1 Then
                temSQL = "SELECT tblDailyCollection.* FROM tblDailyCollection Where ProgramDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
            Else
                temSQL = "SELECT tblDailyCollection.* FROM tblDailyCollection Where ProgramDate = '" & Format(dtpDate.Value + 1, "dd MMMM yyyy") & "' And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
            End If
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !Date = Format(dtpDate.Value, "dd MMMM yyyy")
            If cmbSecession.BoundText = 1 Then ' Morning
                !CollectionDate = dtpDate.Value
                !DeliveryDate = dtpDate.Value + 1
                !ProgramDate = dtpDate.Value
            Else
                !CollectionDate = dtpDate.Value
                !DeliveryDate = dtpDate.Value + 2
                !ProgramDate = dtpDate.Value + 1
            End If
            !CollectingCenterID = cmbCollectingCenter.BoundText
        ElseIf .RecordCount < 1 Then
            .AddNew
            !Date = Format(dtpDate.Value, "dd MMMM yyyy")
            If cmbSecession.BoundText = 1 Then ' Morning
                !CollectionDate = dtpDate.Value
                !DeliveryDate = dtpDate.Value + 1
                !ProgramDate = dtpDate.Value
            Else
                !CollectionDate = dtpDate.Value
                !DeliveryDate = dtpDate.Value + 2
                !ProgramDate = dtpDate.Value + 1
            End If
            !CollectingCenterID = cmbCollectingCenter.BoundText
        End If
        !TotalVolume = TotalVolume
        !TotalLMR = TotalLMR
        !TotalFAT = TotalFAT
        !TestedLMR = TestedLMR
        !TestedFAT = TestedFAT
        !actualValue = TotalActualValue
        !ValueDifference = TotalValueDifference
        If TotalVolume <> 0 Then
            !totalValue = totalValue
            !CalculatedLMR = TotalLMR / TotalVolume
            !CalculatedFat = TotalFAT / TotalVolume
            !AverageValue = totalValue / TotalVolume
        Else
            !totalValue = 0
            !CalculatedLMR = 0
            !CalculatedFat = 0
            !AverageValue = 0
        End If
        .Update
        .Close
    End With
    
End Sub

Private Sub CalculateTotals()
    Dim TotalLeters As Double
    Dim TotalLMR As Double
    Dim TotalFAT As Double
    Dim totalValue As Double
    txtTotalLiters.Text = Empty
    txtCLMR.Text = Empty
    txtCFAT.Text = Empty
    txtTotalValue.Text = Empty
    txtAverageValue.Text = Empty
    txtTotalLMR.Text = Empty
    txtTotalFAT.Text = Empty
    With gridMilk
        For i = 1 To .Rows - 1
            TotalLeters = TotalLeters + Val(.TextMatrix(i, 4))
            TotalLMR = TotalLMR + Val(.TextMatrix(i, 5))
            TotalFAT = TotalFAT + Val(.TextMatrix(i, 6))
            totalValue = totalValue + Val(.TextMatrix(i, 9))
        Next
    End With
    If TotalLeters <> 0 Then
        txtTotalLiters.Text = TotalLeters
        txtTotalLMR.Text = TotalLMR
        txtTotalFAT.Text = TotalFAT
        txtCLMR.Text = Format((TotalLMR / TotalLeters), "0.00")
        txtCFAT.Text = Format((TotalFAT / TotalLeters), "0.00")
        txtTotalValue.Text = Format(totalValue, "0.00")
        txtAverageValue.Text = Format((totalValue / TotalLeters), "0.00")
        
        txtTSNF.Text = SNF(Val(txtTLMR.Text), Val(txtTFAT.Text))
        txtCValue.Text = (Price(Val(txtTFAT.Text), Val(txtTSNF.Text), Val(cmbCollectingCenter.BoundText), 0, dtpDate.Value)) * Val(txtTotalLiters.Text)
        
        txtValueDifference.Text = Val(txtTotalValue.Text) - Val(txtCValue.Text)
        
    End If
'   0   "No."
'   1   "Supplier"
'   2   "LMR"
'   3   "FAT %"
'   4   "Liters"
'   5   "LMR X Liters"
'   6   "FAT% X Liters"
'   7   "SNF"
'   8   "Price"
'   9   "Value"
'   10  "ID"
'   11  "Supplier ID"
End Sub

Private Sub ClearValues()
    cmbSupplierName.Text = Empty
    txtLMR.Text = Empty
    txtFat.Text = Empty
    txtLiters.Text = Empty
    txtLMRXLiters.Text = Empty
    txtFatXLiters.Text = Empty
    txtSNF.Text = Empty
    txtPrice.Text = Empty
    txtValue.Text = Empty
End Sub

Private Sub FillGrid()
    With rsCollection
        If .State = 1 Then .Close
        temSQL = "SELECT tblCollection.*, tblSupplier.* FROM tblCollection LEFT JOIN tblSupplier ON tblCollection.SupplierID = tblSupplier.SupplierID Where tblCollection.Deleted = 0 And tblCollection.Date = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' And tblCollection.SecessionID = " & Val(cmbSecession.BoundText) & " And tblCollection.CollectingCenterID = " & Val(cmbCollectingCenter.BoundText) & " And tblCollection.Deleted = 0 "
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            gridMilk.Rows = .RecordCount + 1
            i = 0
            While .EOF = False
                i = i + 1
                gridMilk.TextMatrix(i, 0) = i
                If IsNull(!SupplierCode) = False Then
                    If !SupplierCode <> "" Then
                        gridMilk.TextMatrix(i, 1) = !Supplier & " (" & !SupplierCode & ")"
                    Else
                        gridMilk.TextMatrix(i, 1) = !Supplier
                    End If
                Else
                    gridMilk.TextMatrix(i, 1) = !Supplier
                End If
                gridMilk.TextMatrix(i, 2) = Format(!LMR, "0.00")
                gridMilk.TextMatrix(i, 3) = Format(!FAT, "0.00")
                gridMilk.TextMatrix(i, 4) = Format(!Liters, "0.00")
                gridMilk.TextMatrix(i, 5) = Format(!LMRXLiters, "0.00")
                gridMilk.TextMatrix(i, 6) = Format(!FATXLiters, "0.00")
                gridMilk.TextMatrix(i, 7) = Format(!SNF, "0.00")
                gridMilk.TextMatrix(i, 8) = Format(!Price, "0.00")
                gridMilk.TextMatrix(i, 9) = Format(!Value, "0.00")
                gridMilk.TextMatrix(i, 10) = !CollectionID
                gridMilk.TextMatrix(i, 11) = ![tblSupplier.SupplierID]
'   0   "No."
'   1   "Supplier"
'   2   "LMR"
'   3   "FAT %"
'   4   "Liters"
'   5   "LMR X Liters"
'   6   "FAT% X Liters"
'   7   "SNF"
'   8   "Price"
'   9   "Value"
'   10  "ID"
'   11  "Supplier ID"
                .MoveNext
            Wend
        End If
        .Close
    End With
    gridMilk.row = 0
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Call WriteDailyCollection
    Dim temID As Long
    If gridMilk.Rows < 2 Then
        MsgBox "Nothing to Delete"
        Exit Sub
    End If
    If gridMilk.row < 1 Then
        MsgBox "Nothing to Delete"
        Exit Sub
    End If
    If IsNumeric(gridMilk.TextMatrix(gridMilk.row, 10)) = False Then
        MsgBox "Nothing to Delete"
        Exit Sub
    Else
        temID = Val(gridMilk.TextMatrix(gridMilk.row, 10))
    End If
    With rsCollection
        If .State = 1 Then .Close
        temSQL = "Select * from tblCollection Where CollectionID = " & temID
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedUserID = UserID
            !DeletedDate = Format(Date, "dd MMMM yyyy")
            !DeletedTime = Time
            .Update
        End If
        If .State = 1 Then .Close
    End With
    Call FormatGrid
    Call FillGrid
    Call ClearValues
    Call CalculateTotals
    Call WriteDailyCollection
    cmbSupplierName.SetFocus
End Sub

Private Sub chkByCode_Click()
    cmbCollectingCenter_Change
    cmbSupplierName.SetFocus
End Sub

Private Sub cmbCollectingCenter_Change()
    If chkByCode.Value = 0 Then
        With rsViewSupplier
            If .State = 1 Then .Close
            temSQL = "Select * from tblSupplier where CollectingCenterID = " & Val(cmbCollectingCenter.BoundText) & " And Deleted = 0  Order by Supplier"
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        End With
        With cmbSupplierName
            Set .RowSource = rsViewSupplier
            .ListField = "Supplier"
            .BoundColumn = "SupplierID"
        End With
    Else
        With rsViewSupplier
            If .State = 1 Then .Close
            temSQL = "Select * from tblSupplier where CollectingCenterID = " & Val(cmbCollectingCenter.BoundText) & " And Deleted = 0  Order by SupplierCode"
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        End With
        With cmbSupplierName
            Set .RowSource = rsViewSupplier
            .ListField = "SupplierCode"
            .BoundColumn = "SupplierID"
        End With
    End If
    Call FormatGrid
    Call FillGrid
    Call CalculateTotals
    Call DisplayDailyCollectionDetails
End Sub

Private Sub cmbCollectingCenter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsNumeric(cmbCollectingCenter.BoundText) = True Then
            dtpDate.SetFocus
        End If
    End If
End Sub


Private Sub cmbSecession_Change()
    Call FormatGrid
    Call FillGrid
    Call CalculateTotals
    Call DisplayDailyCollectionDetails
End Sub


Private Sub cmbSecession_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbSupplierName.SetFocus
        SendKeys "{down}"
    End If
End Sub


Private Sub cmbSupplierName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtLMR.SetFocus
        SendKeys "{Home}+{End}"
    ElseIf KeyCode = vbKeyEscape Then
        cmbSupplierName.Text = Empty
    End If
End Sub

Private Sub dtpDate_Change()
    Call FormatGrid
    Call FillGrid
    Call CalculateTotals
    Call DisplayDailyCollectionDetails
    Call WriteDailyCollection
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbSecession.SetFocus
    End If
End Sub

Private Sub Form_Load()
    NewRecord = True
    dtpDate.Value = Date
    Call FillCombos
    Call FormatGrid
    cmbSecession.Text = "Morning"
    
Select Case UserAuthorityLevel
    
    Case Authority.Analyzer '2
        btnAdd.Visible = False
        btnDelete.Visible = False
        btnPrint.Visible = False
    Case Else
End Select
    
End Sub

Private Sub FormatGrid()
    With gridMilk
        
        .Rows = 1
        .Cols = 12
        
        .row = 0
        
        For i = 0 To .Cols - 1
            Select Case i
                Case 0:
                    .ColWidth(i) = 400
                    .col = i
                    .CellAlignment = 4
                    .Text = "No."
                Case 1:
                    .ColWidth(i) = 3400
                    .col = i
                    .CellAlignment = 4
                    .Text = "Supplier"
                Case 2:
                    .ColWidth(i) = 900
                    .col = i
                    .CellAlignment = 4
                    .Text = "LMR"
                Case 3:
                    .ColWidth(i) = 700
                    .col = i
                    .CellAlignment = 4
                    .Text = "FAT %"
                Case 4:
                    .ColWidth(i) = 800
                    .col = i
                    .CellAlignment = 4
                    .Text = "Liters"
                Case 5:
                    .ColWidth(i) = 800
                    .col = i
                    .CellAlignment = 4
                    .Text = "LMR X l"
                Case 6:
                    .ColWidth(i) = 1300
                    .col = i
                    .CellAlignment = 4
                    .Text = "FAT% X Liters"
                Case 7:
                    .ColWidth(i) = 800
                    .col = i
                    .CellAlignment = 4
                    .Text = "SNF"
                Case 8:
                    .ColWidth(i) = 1000
                    .col = i
                    .CellAlignment = 4
                    .Text = "Price"
                Case 9:
                    .ColWidth(i) = 1000
                    .col = i
                    .CellAlignment = 4
                    .Text = "Value"
                Case 10:
                    .ColWidth(i) = 1
                    .col = i
                    .CellAlignment = 4
                    .Text = "ID"
                Case 11:
                    .ColWidth(i) = 1
                    .col = i
                    .CellAlignment = 4
                    .Text = "Supplier ID"
            End Select
        Next
    End With
    
'   0   "No."
'   1   "Supplier"
'   2   "LMR"
'   3   "FAT %"
'   4   "Liters"
'   5   "LMR X Liters"
'   6   "FAT% X Liters"
'   7   "SNF"
'   8   "Price"
'   9   "Value"
'   10  "ID"
'   11  "Supplier ID"

End Sub

Private Sub FillCombos()
    Dim Supplier As New clsFind
    Dim Center As New clsFind
    Dim Secession As New clsFind
    Center.FillCombo cmbCollectingCenter, "tblCollectingCenter", "CollectingCenter", "CollectingCenterID", True
    Secession.FillCombo cmbSecession, "tblSecession", "Secession", "SecessionID", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call WriteDailyCollection
End Sub

Private Sub gridMilk_DblClick()
Exit Sub
    Dim temRow As Integer
    Dim temID As Long
    Call ClearValues
    If gridMilk.Rows < 2 Then
        Exit Sub
    End If
    If gridMilk.row < 1 Then
        Exit Sub
    End If
    If IsNumeric(gridMilk.TextMatrix(gridMilk.row, 10)) = False Then
        Exit Sub
    Else
        temID = Val(gridMilk.TextMatrix(gridMilk.row, 10))
    End If
    With rsCollection
        If .State = 1 Then .Close
        temSQL = "Select * from tblCollection Where CollectionID = " & temID
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedUserID = UserID
            !DeletedDate = Format(Date, "dd MMMM yyyy")
            !DeletedTime = Time
            .Update
        End If
        If .State = 1 Then .Close
    End With
    
    With gridMilk
        temRow = .row
        cmbSupplierName.BoundText = .TextMatrix(temRow, 11)
        txtLMR.Text = .TextMatrix(temRow, 2)
        txtFat.Text = .TextMatrix(temRow, 3)
        txtLiters.Text = .TextMatrix(temRow, 4)
    End With
    
    Call FormatGrid
    Call FillGrid
    Call CalculateTotals
    Call WriteDailyCollection
    cmbSupplierName.SetFocus
'   0   "No."
'   1   "Supplier"
'   2   "LMR"
'   3   "FAT %"
'   4   "Liters"
'   5   "LMR X Liters"
'   6   "FAT% X Liters"
'   7   "SNF"
'   8   "Price"
'   9   "Value"
'   10  "ID"
'   11  "Supplier ID"

End Sub

Private Sub txtFat_Change()
    txtFatXLiters.Text = FATXLiters(Val(txtFat.Text), Val(txtLiters.Text))
    txtSNF.Text = SNF(Val(txtLMR.Text), Val(txtFat.Text))
End Sub

Private Sub txtFat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtLiters.SetFocus
        SendKeys "{Home}+{end}"
    End If
End Sub

Private Sub txtLiters_Change()
    txtFatXLiters.Text = FATXLiters(Val(txtFat.Text), Val(txtLiters.Text))
    txtLMRXLiters.Text = LMRXLiters(Val(txtLMR.Text), Val(txtLiters.Text))
    txtValue.Text = Format(Val(txtPrice.Text) * Val(txtLiters.Text), "0.00")
End Sub

Private Sub txtLiters_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnAdd_Click
    End If
End Sub

Private Sub txtLMR_Change()
    txtLMRXLiters.Text = LMRXLiters(Val(txtLMR.Text), Val(txtLiters.Text))
    txtSNF.Text = SNF(Val(txtLMR.Text), Val(txtFat.Text))
End Sub


Private Sub txtLMR_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtFat.SetFocus
        SendKeys "{Home}+{end}"
    End If
End Sub

Private Sub txtPrice_Change()
    txtValue.Text = Format(Val(txtPrice.Text) * Val(txtLiters.Text), "0.00")
End Sub

Private Sub txtSNF_Change()
    txtPrice.Text = Price(Val(txtFat.Text), Val(txtSNF.Text), Val(cmbCollectingCenter.BoundText), Val(cmbSupplierName.BoundText), dtpDate.Value)
End Sub

Private Sub txtTFAT_Change()
        txtTSNF.Text = SNF(Val(txtTLMR.Text), Val(txtTFAT.Text))
        txtCValue.Text = (Price(Val(txtTFAT.Text), Val(txtTSNF.Text), Val(cmbCollectingCenter.BoundText), Val(cmbSupplierName.BoundText), dtpDate.Value)) * Val(txtTotalLiters.Text)
        txtValueDifference.Text = Val(txtTotalValue.Text) - Val(txtCValue.Text)
End Sub

Private Sub txtTFAT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnClose.SetFocus
    End If
End Sub

Private Sub txtTFAT_LostFocus()
    If txtTFAT.Text <> txtTFAT1.Text Then WriteDailyCollection
End Sub

Private Sub txtTLMR_Change()
        txtTSNF.Text = SNF(Val(txtTLMR.Text), Val(txtTFAT.Text))
        txtCValue.Text = (Price(Val(txtTFAT.Text), Val(txtTSNF.Text), Val(cmbCollectingCenter.BoundText), Val(cmbSupplierName.BoundText), dtpDate.Value)) * Val(txtTotalLiters.Text)
        
        txtValueDifference.Text = Val(txtTotalValue.Text) - Val(txtCValue.Text)
End Sub

Private Sub txtTLMR_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtTFAT.SetFocus
    End If
End Sub

Private Sub txtTLMR_LostFocus()
    If txtTLMR.Text <> txtTLMR1.Text Then WriteDailyCollection
End Sub

Private Sub txtTotalValue_Change()
    txtTotalValueDisplay.Text = Format(Val(txtTotalValue.Text), "0.00")
End Sub
