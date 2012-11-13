VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmReportVitaminAndCattleFeedDeduction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report - Vitamin And Cattle Feed Deduction"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12480
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
   ScaleWidth      =   12480
   Begin MSFlexGridLib.MSFlexGrid gridDetails 
      Height          =   4695
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8281
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin btButtonEx.ButtonEx btnFill 
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Fill"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   158990339
      CurrentDate     =   39864
   End
   Begin MSDataListLib.DataCombo cmbCC 
      Height          =   360
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1200
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   158990339
      CurrentDate     =   39864
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   11040
      TabIndex        =   8
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   375
      Left            =   11040
      TabIndex        =   9
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Collecting Center"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmReportVitaminAndCattleFeedDeduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub FormatGrid()
    With gridDetails
        .Clear
        
        .Cols = 5
        .Rows = 1
        
        
        .ColWidth(0) = 2000
        .ColWidth(1) = 4000
        .ColWidth(2) = 3000
        .ColWidth(3) = 1250
        .ColWidth(4) = 1500
'        .ColWidth(5) = 3000
'        .ColWidth(6) = 3000
'        .ColWidth(7) = 3000
'        .ColWidth(8) = 3000
'        .ColWidth(9) = 3000
       
        
        .row = 0
        
        .col = 0
        .Text = "Deducted Date"
        
        .col = 1
        .Text = "Supplier"
        
        .col = 2
        .Text = "Item"
        
        .col = 3
        .Text = "Quentity"
       
        
        .col = 4
        .Text = "Value"
        
'        .col = 5
'        .Text = "Approved"
'
'        .col = 6
'        .Text = "Approved User"
'
'        .col = 7
'        .Text = "Approved Comments"
'
'        .col = 8
'        .Text = "Approved Date"
'
'        .col = 9
'        .Text = "Approved Value"
    
    End With
End Sub

Private Sub FillGrid()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblDeduction.DeductedDate, tblSupplier.Supplier, tblItem.Item, tblDeduction.Quentity, tblDeduction.Value, tblDeduction.DeductedDate " & _
                    "FROM (tblDeduction LEFT JOIN tblItem ON tblDeduction.ItemID = tblItem.ItemID) LEFT JOIN tblSupplier ON tblDeduction.SupplierID = tblSupplier.SupplierID " & _
                    "WHERE (((tblSupplier.CollectingCenterID)= " & Val(cmbCC.BoundText) & ") AND ((tblDeduction.DeductedDate) Between '" & Format(dtpFrom.Value, "DD MMMM YYYY") & "' And '" & Format(dtpTo.Value, "DD MMMM YYYY") & "'))"

        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridDetails.Rows = gridDetails.Rows + 1
            gridDetails.row = gridDetails.Rows - 1
            gridDetails.col = 0
            gridDetails.Text = Format(!DeductedDate, "dd MMMM yyyy")
            gridDetails.col = 1
            gridDetails.Text = Format(![Supplier], "")
            gridDetails.col = 2
            gridDetails.Text = !Item
            gridDetails.col = 3
            gridDetails.Text = !Quentity
            gridDetails.col = 4
            gridDetails.Text = Format(!Value, "0.00")
'            If !Approved = True Then
'                gridDetails.col = 5
'                gridDetails.Text = "Approved"
'            Else
'                gridDetails.col = 5
'                gridDetails.Text = "Not Approved"
'            End If
'            gridDetails.col = 6
'            gridDetails.Text = Format(![tblStaffAutherized.Staff], "")
'            gridDetails.col = 7
'            gridDetails.Text = Format(!ApprovedDate, "dd MMMM yyyy")
'            gridDetails.col = 8
'            gridDetails.Text = Format(!ApprovedComments, "")
'            gridDetails.col = 9
'            gridDetails.Text = Format(!Value, "0.00")
            .MoveNext
        Wend
    End With

End Sub

Private Sub btnFill_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub btnFill_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnClose.SetFocus
    End If
End Sub

Private Sub btnPrint_Click()
    
        Dim tabReport As Long
        Dim tab1 As Long
        Dim tab2 As Long
        Dim tab3 As Long
        Dim tab4 As Long
    
        tabReport = 60
        tab1 = 5
        tab2 = 40
        tab3 = 60
        tab4 = 95
        
        Printer.Font = "Arial"
        Printer.Font.Size = 11
        Printer.Font.Bold = True
    
        Printer.Print
        Printer.Print Tab(tabReport); "Vitami & Cattle Feed Deduction Report"
        
        Printer.Font = "Arial"
        Printer.Font.Size = 10
        Printer.Font.Bold = True
        
        Printer.Print Tab(tab1); "Collecting Center :";
        Printer.Print Tab(tab2); cmbCC.Text
        Printer.Print Tab(tab1); "From  :";
        Printer.Print Tab(tab2); dtpFrom.Value;
        Printer.Print Tab(tab3); "To  :";
        Printer.Print Tab(tab4); dtpTo.Value;
        Printer.Print
    
        Printer.Font = "Arial Narrow"
        Printer.Font.Size = 10
        Printer.Font.Bold = True
    
        Dim i As Integer
        Dim tabNo As Long
        Dim tabDate As Long
        Dim tabTotalLiters As Long
        Dim tabCLMR As Long
        Dim tabCFAT As Long
            
        tabNo = 10
        tabDate = 30
        tabTotalLiters = 70
        tabCLMR = 100
        tabCFAT = 130
        
        With gridDetails
            For i = 0 To .Rows - 1
                Printer.Print
                Printer.Print Tab(tabNo); .TextMatrix(i, 0);
                Printer.Print Tab(tabDate); .TextMatrix(i, 1);
                Printer.Print Tab(tabTotalLiters - Len(.TextMatrix(i, 2))); .TextMatrix(i, 2);
                Printer.Print Tab(tabCLMR - Len(.TextMatrix(i, 3))); .TextMatrix(i, 3);
                Printer.Print Tab(tabCFAT - Len(.TextMatrix(i, 4))); .TextMatrix(i, 4);
                Printer.Print
            Next
        End With
        Printer.EndDoc
End Sub

Private Sub cmbCC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpFrom.SetFocus
    End If
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpTo.SetFocus
    End If
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnFill.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCC, "CollectingCenter", True
    
Select Case UserAuthorityLevel
    
    Case Authority.Analyzer '2
        btnPrint.Visible = False
    Case Else
End Select
End Sub
