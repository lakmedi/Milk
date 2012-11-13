VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmExpence 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expences"
   ClientHeight    =   7320
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
   ScaleHeight     =   7320
   ScaleWidth      =   10200
   Begin VB.TextBox txtValue 
      Height          =   360
      Left            =   2160
      TabIndex        =   7
      Top             =   1560
      Width           =   4695
   End
   Begin MSFlexGridLib.MSFlexGrid gridExpence 
      Height          =   3495
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6165
      _Version        =   393216
      ScrollTrack     =   -1  'True
      AllowUserResizing=   2
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   495
      Left            =   8760
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
      Height          =   495
      Left            =   8760
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
   Begin VB.TextBox txtComments 
      Height          =   975
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2040
      Width           =   4695
   End
   Begin MSDataListLib.DataCombo cmbCollectingCenter 
      Height          =   360
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpDate 
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
      Format          =   90046467
      CurrentDate     =   39776
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8760
      TabIndex        =   13
      Top             =   6720
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
      TabIndex        =   14
      Top             =   6720
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
   Begin VB.Label Label5 
      Caption         =   "Value"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Category"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   2
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
Attribute VB_Name = "frmExpence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String

Private Sub btnAdd_Click()
    Dim i As Integer
    If IsNumeric(txtValue.Text) = False Then
        MsgBox "Please enter a value"
        txtValue.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then
        MsgBox "Please select a collecting center"
        cmbCollectingCenter.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbCategory.BoundText) = False Then
        MsgBox "Please select a category"
        cmbCategory.SetFocus
        Exit Sub
    End If
    
    Dim rsExpence As New ADODB.Recordset
    With rsExpence
        If .State = 1 Then .Close
        temSQL = "Select * from tblExpence where ExpenceID =0"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ExpenceValue = Val(txtValue.Text)
        !ExpenceCategoryID = Val(cmbCategory.BoundText)
        !CollectingCenterID = Val(cmbCollectingCenter.BoundText)
        !ExpenceComments = txtComments.Text
        !ExpenceDate = Format(dtpDate.Value, "dd MMMM yyyy")
        !AddedDate = Now
        !AddedUserID = UserID
        .Update
        .Close
    End With
    Call ClearAddValues
    Call FormatGrid
    Call FillGrid
    cmbCategory.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    With gridExpence
        Dim temRow As Integer
        temRow = .row
        If .Rows <= 1 Then Exit Sub
        If .row < 1 Then Exit Sub
        If IsNumeric(.TextMatrix(temRow, 4)) = False Then Exit Sub
    End With
    Dim rsExpence As New ADODB.Recordset
    With rsExpence
        If .State = 1 Then .Close
        temSQL = "Select * from tblExpence where ExpenceID = " & Val(gridExpence.TextMatrix(temRow, 4))
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedDate = Now
            !DeletedUserID = UserID
            .Update
        End If
        .Close
    End With
    Call FormatGrid
    Call FillGrid
    cmbCategory.SetFocus
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
    Printer.Print Tab(tab1); "Date  :";
    Printer.Print Tab(tab2); dtpDate.Value;
    Printer.Print

    Dim i As Integer
    Dim tabNo As Long
    Dim tabCategory As Long
    Dim tabComments As Long
    Dim tabValue As Long
    
    tabNo = 10
    tabCategory = 20
    tabComments = 60
    tabValue = 115

    With gridExpence
        For i = 0 To .Rows - 1
            Printer.Print
            Printer.Print Tab(tabNo - Len(.TextMatrix(i, 0))); .TextMatrix(i, 0);
            Printer.Print Tab(tabCategory); .TextMatrix(i, 1);
            Printer.Print Tab(tabComments); .TextMatrix(i, 2);
            Printer.Print Tab(tabValue - Len(.TextMatrix(i, 3))); .TextMatrix(i, 3);
            Printer.Print
        Next
    End With

   Printer.EndDoc
End Sub

Private Sub cmbCollectingCenter_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpDate_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call FillCombos
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
        .Text = "No."
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
        .ColWidth(0) = 600
        .ColWidth(1) = 2500
        .ColWidth(2) = 3500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1
    End With
End Sub

Private Sub FillGrid()
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then Exit Sub
    Dim rsExpence As New ADODB.Recordset
    With rsExpence
        If .State = 1 Then .Close
        temSQL = "SELECT tblExpence.*, tblExpenceCategory.ExpenceCategory FROM tblExpence LEFT JOIN tblExpenceCategory ON tblExpence.ExpenceCategoryID = tblExpenceCategory.ExpenceCategoryID where tblExpence.ExpenceDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' And tblExpence.CollectingCenterID = " & Val(cmbCollectingCenter.BoundText) & " And tblExpence.Deleted = 0 "
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Dim i As Integer
            gridExpence.Rows = .RecordCount + 1
            For i = 1 To .RecordCount
                gridExpence.TextMatrix(i, 0) = i
                gridExpence.TextMatrix(i, 1) = !expenceCategory
                gridExpence.TextMatrix(i, 2) = !ExpenceComments
                gridExpence.TextMatrix(i, 3) = Format(!ExpenceValue, "#,##0.00")
                gridExpence.TextMatrix(i, 4) = !ExpenceID
                .MoveNext
            Next
        End If
    End With
End Sub

Private Sub ClearAddValues()
    txtComments.Text = Empty
    txtValue.Text = Empty
    cmbCategory.Text = Empty
End Sub

Private Sub ClearUpdateValues()
    cmbCollectingCenter.Text = Empty
End Sub
