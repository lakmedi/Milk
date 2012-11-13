VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDisplayCommision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commision"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
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
   ScaleWidth      =   6945
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   7440
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
   Begin MSDataListLib.DataCombo cmbSupplierName 
      Height          =   360
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   4695
      _ExtentX        =   8281
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
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMM yyyy"
      Format          =   56492035
      CurrentDate     =   39682
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMM yyyy"
      Format          =   56492035
      CurrentDate     =   39682
   End
   Begin MSFlexGridLib.MSFlexGrid gridCom 
      Height          =   5775
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   10186
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   4200
      TabIndex        =   10
      Top             =   7440
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
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "&Farmer"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmDisplayCommision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Centre As New clsFind
    Dim Supplier As New clsFind
    Dim rsMil As New ADODB.Recordset
    Dim temSql As String
    Dim CSetPrinter As New cSetDfltPrinter

Private Sub FillCombos()
    Centre.FillCombo cmbCollectingCenter, "tblCollectingCenter", "CollectingCenter", "CollectingCenterID", True
    Supplier.FillCombo cmbSupplierName, "tblSupplier", "Supplier", "SupplierID", True
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    
    Dim i As Integer
    Dim SerialX As Integer
    Dim TextX As Integer
    Dim ValueX As Integer
    
    SerialX = 4
    TextX = 7
    ValueX = 50
    
    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
    If SelectForm(ReportPaperName, Me.hwnd) = 1 Then
        Printer.Print Tab(SerialX); "MILK PAY ADVICE - LUCKY LANKA DIARY PRODUCTS"
        Printer.Print Tab(SerialX); "COMMISION"
        Printer.Print
        Printer.Print Tab(SerialX); cmbSupplierName.Text
        Printer.Print
        With gridCom
            For i = 1 To .Rows - 1
                Printer.Print Tab(SerialX); .TextMatrix(i, 0);
                Printer.Print Tab(TextX); .TextMatrix(i, 1);
                Printer.Print Tab(ValueX - Len(.TextMatrix(i, 2))); .TextMatrix(i, 2);
            Next
            Printer.EndDoc
        End With
        
    Else
        MsgBox "Printer Error"
    End If

End Sub

Private Sub btnPrint_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnClose.SetFocus
    End If
End Sub

Private Sub cmbCollectingCenter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpFrom.SetFocus
    End If
End Sub

Private Sub cmbSupplierName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnPrint.SetFocus
    End If
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpTo.SetFocus
    End If
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbSupplierName.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
End Sub

Public Sub FillGrid()

    Call FormatGrid
    
    Dim i As Integer
    Dim rsS As New ADODB.Recordset
    Dim MyMilkCollection As MilkCollection
    Dim TotalMilk As Double
    Dim temDays As Integer
    Dim AvgMilk As Double
    Dim temCommsisionRate As Double
    Dim temCommsision As Double
    
    temDays = DateDiff("d", dtpFrom.Value, dtpTo.Value) + 1

    With rsS
        If .State = 1 Then .Close
        temSql = "SELECT tblSupplier.SupplierID, tblSupplier.Supplier, tblSupplier.SupplierCode " & _
                    "From tblSupplier " & _
                    "Where (((tblSupplier.CommisionCollectorID) = " & Val(cmbSupplierName.BoundText) & ")) " & _
                    "ORDER BY tblSupplier.Supplier"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            TotalMilk = 0
            gridCom.Rows = .RecordCount + 1
            For i = 1 To .RecordCount
                gridCom.TextMatrix(i, 0) = i
                gridCom.TextMatrix(i, 1) = !Supplier & " " & !SupplierCode
                MyMilkCollection = PeriodMilkSupply(dtpFrom.Value, dtpTo.Value, !SupplierID, 0)
                TotalMilk = TotalMilk + MyMilkCollection.Liters
                gridCom.TextMatrix(i, 2) = Format(MyMilkCollection.Liters, "0.00")
                .MoveNext
            Next
        End If
        .Close
    End With
    With gridCom
        .Rows = .Rows + 5
        
        .TextMatrix(i + 1, 1) = "Milk Collection for " & temDays & " days"
        .TextMatrix(i + 1, 2) = Format(TotalMilk, "0.00")
        
        AvgMilk = TotalMilk / temDays
        .TextMatrix(i + 2, 1) = "Average Day Collection"
        .TextMatrix(i + 2, 2) = Format(AvgMilk, "0.00")
        
        temCommsisionRate = OthersCommisionRate(Val(cmbSupplierName.BoundText), AvgMilk)
        .TextMatrix(i + 3, 1) = "Commision Rate"
        .TextMatrix(i + 3, 2) = temCommsisionRate
        
        temCommsision = temCommsisionRate * TotalMilk
        .TextMatrix(i + 4, 1) = "Commision"
        .TextMatrix(i + 4, 2) = Format(temCommsision, "0.00")
        
    End With
End Sub

Private Sub FormatGrid()
    With gridCom
        .Cols = 3
        .Rows = 1
        .row = 0
        .col = 0
        .CellAlignment = 4
        .Text = "No"
        .col = 1
        .CellAlignment = 4
        .Text = "Farmer"
        .col = 2
        .CellAlignment = 4
        .Text = "Liters"
        .ColWidth(0) = 800
        .ColWidth(2) = 1400
        .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + 100)
    End With
End Sub

