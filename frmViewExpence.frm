VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmViewExpence 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expence"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
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
   ScaleHeight     =   8070
   ScaleWidth      =   10200
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6480
      TabIndex        =   12
      Top             =   6720
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid gridExpence 
      Height          =   5055
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8916
      _Version        =   393216
      ScrollTrack     =   -1  'True
      AllowUserResizing=   2
   End
   Begin MSDataListLib.DataCombo cmbCollectingCenter 
      Height          =   360
      Left            =   2160
      TabIndex        =   5
      Top             =   600
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   103481347
      CurrentDate     =   39776
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   2160
      TabIndex        =   7
      Top             =   1080
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8760
      TabIndex        =   10
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
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
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   7320
      TabIndex        =   9
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   103481347
      CurrentDate     =   39776
   End
   Begin VB.Label Label5 
      Caption         =   "Total"
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Category"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmViewExpence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub bttnPrint_Click()
    Dim tabReport As Long
    Dim tab1 As Long
    Dim tab2 As Long

    tabReport = 70
    tab1 = 5
    tab2 = 40
    
    Printer.Print
    Printer.Font.Bold = True
    Printer.Print Tab(tabReport); "Expence Report"
    Printer.Font.Bold = False
    Printer.Print
    Printer.Print Tab(tab1); "Collecting Center :";
    Printer.Print Tab(tab2); cmbCollectingCenter.Text
    
    Printer.Print Tab(tab1); "From  :";
    
    Printer.Print Tab(tab2); Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    
    
    Printer.Print

    Dim i As Integer
    Dim tabNo As Long
    Dim tabCategory As Long
    Dim tabComments As Long
    Dim tabValue As Long
    
    tabNo = 5
    tabCategory = 30
    tabComments = 70
    tabValue = 115

    With gridExpence
        For i = 0 To .Rows - 1
            Printer.Print
            Printer.Print Tab(tabNo); .TextMatrix(i, 0);
            Printer.Print Tab(tabCategory); .TextMatrix(i, 1);
            Printer.Print Tab(tabComments); .TextMatrix(i, 2);
            Printer.Print Tab(tabValue - Len(.TextMatrix(i, 3))); .TextMatrix(i, 3);
            Printer.Print
        Next
    End With
    Printer.Print
    Printer.Print Tab(tab1); "Total : ";
    Printer.Print Tab(tab2); txtTotal.Text
   Printer.EndDoc
End Sub

Private Sub cmbCategory_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        bttnPrint.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbCategory.Text = Empty
    End If
End Sub

Private Sub cmbCollectingCenter_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpDate_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbCollectingCenter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbCategory.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbCollectingCenter.Text = Empty
    End If
End Sub

Private Sub dtpFrom_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpTo.SetFocus
    End If
End Sub

Private Sub dtpTo_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbCollectingCenter.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = Date
    Call FormatGrid
End Sub

Private Sub FillCombos()
    Dim CC As New clsFillCombos
    Dim Category As New clsFillCombos
    CC.FillAnyCombo cmbCollectingCenter, "CollectingCenter", True
    Category.FillAnyCombo cmbCategory, "ExpenceCategory", True
End Sub

Private Sub FormatGrid()
    With gridExpence
        .Clear
        .Cols = 5
        .Rows = 1
        .row = 0
        .col = 0
        .CellAlignment = 4
        .Text = "Date"
        .col = 1
        .CellAlignment = 4
        .Text = "Category"
        .col = 2
        .CellAlignment = 4
        .Text = "Comments"
        .col = 3
        .CellAlignment = 4
        .Text = "Value"
        .col = 4
        .Text = "ID"
        .ColWidth(0) = 1000
        .ColWidth(1) = 2500
        .ColWidth(2) = 3500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1
    End With
End Sub

Private Sub FillGrid()
    Dim TemTotal As Double
    Dim rsExpence As New ADODB.Recordset
    With rsExpence
        If .State = 1 Then .Close
        temSql = "SELECT tblExpence.*, tblExpenceCategory.ExpenceCategory FROM tblExpence LEFT JOIN tblExpenceCategory ON tblExpence.ExpenceCategoryID = tblExpenceCategory.ExpenceCategoryID where tblExpence.ExpenceDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
        If IsNumeric(cmbCategory.BoundText) = True Then
            temSql = temSql & " AND tblExpence.ExpenceCategoryID = " & Val(cmbCategory.BoundText) & " "
        End If
        
        If IsNumeric(cmbCollectingCenter.BoundText) = True Then
            temSql = temSql & "AND tblExpence.CollectingCenterID = " & Val(cmbCollectingCenter.BoundText) & " "
        End If
        temSql = temSql & " And tblExpence.Deleted = 0 "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Dim i As Integer
            gridExpence.Rows = .RecordCount + 1
            For i = 1 To .RecordCount
                gridExpence.TextMatrix(i, 0) = Format(!ExpenceDate, "dd MMM yyyy")
                gridExpence.TextMatrix(i, 1) = !expenceCategory
                gridExpence.TextMatrix(i, 2) = !ExpenceComments
                gridExpence.TextMatrix(i, 3) = Format(!ExpenceValue, "#,##0.00")
                TemTotal = TemTotal + !ExpenceValue
                gridExpence.TextMatrix(i, 4) = !ExpenceID
                .MoveNext
            Next
        End If
    End With
    txtTotal.Text = Format(TemTotal, "#,##0.00")
End Sub
