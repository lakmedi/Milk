VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmOutstandingBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outstanding Balance"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10545
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
   ScaleHeight     =   7665
   ScaleWidth      =   10545
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   6720
      Width           =   3855
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   6240
      Width           =   3855
   End
   Begin VB.TextBox txtOutstanding 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7560
      TabIndex        =   7
      Top             =   6240
      Width           =   2535
   End
   Begin btButtonEx.ButtonEx btnCalculate 
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Calculate"
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
   Begin MSFlexGridLib.MSFlexGrid gridOutstanding 
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9128
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbCC 
      Height          =   360
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   7200
      Width           =   1695
      _ExtentX        =   2990
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
      Left            =   8760
      TabIndex        =   5
      Top             =   7200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Close"
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
   Begin MSComctlLib.ProgressBar pgb1 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label7 
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Paper"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Total Outstanding"
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmOutstandingBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String

    Dim CSetPrinter As New cSetDfltPrinter
    
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean


Private Sub btnCalculate_Click()
    Screen.MousePointer = vbHourglass
    If IsNumeric(cmbCC.BoundText) = False Then
        Call FormatCCGrid
        Call FillCCGrid
    Else
        Call FormatSGrid
        Call FillSGrid
    End If
    pgb1.Value = 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    CSetPrinter.SetPrinterAsDefault cmbPrinter.Text
    If SelectForm(cmbPaper.Text, Me.hdc) <> 1 Then
        MsgBox "Printer Error"
        Exit Sub
    End If

    Dim tabReport As Long
    Dim tab1 As Long
    Dim tab2 As Long

    tabReport = 70
    tab1 = 5
    tab2 = 40
    
    Printer.Print
    Printer.Font.Bold = True

    If IsNumeric(cmbCC.BoundText) = True Then
        Printer.Print Tab(tabReport); "Outstanding Value of " & cmbCC.Text
    Else
        Printer.Print Tab(tabReport); "Total Outstanding"
    End If
    Printer.Font.Bold = False
    Printer.Print
    Printer.Print Tab(tab1); "As at :";
    Printer.Print Tab(tab2); Format(Date, "dd MMMM yyyy");
    Printer.Print

    Dim i As Integer
    Dim tabNo As Long
    Dim tabCategory As Long
    Dim tabValue1 As Long
    Dim tabValue2 As Long
    Dim tabValue3 As Long
    
    tabNo = 10
    tabCategory = 20
    
    tabValue1 = 90
    tabValue2 = 115
    tabValue3 = 140
    
    With gridOutstanding
        For i = 0 To .Rows - 1
            Printer.Print
            Printer.Print Tab(tabNo - Len(.TextMatrix(i, 0))); .TextMatrix(i, 0);
            Printer.Print Tab(tabCategory); .TextMatrix(i, 1);
            Printer.Print Tab(tabValue1 - Len(.TextMatrix(i, 2))); .TextMatrix(i, 2);
            Printer.Print Tab(tabValue2 - Len(.TextMatrix(i, 3))); .TextMatrix(i, 3);
            Printer.Print Tab(tabValue3 - Len(.TextMatrix(i, 4))); .TextMatrix(i, 4);
            Printer.Print
        Next
    End With
    
    Printer.Print
    Printer.Print
    Printer.Print Tab(tab1); "Outstanding Value ";
    Printer.Print Tab(tab2); txtOutstanding.Text;
    Printer.EndDoc


End Sub

Private Sub cmbCC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnCalculate.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbCC.Text = Empty
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call ListPrinters
    Call GetSettings
    pgb1.Value = 0
End Sub


Private Sub GetSettings()
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, "Printer", "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, "Paper", "")
End Sub


Private Sub ListPrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub ListPapers()
    cmbPaper.Clear
    CSetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        With FormSize
            .cx = BillPaperHeight
            .cy = BillPaperWidth
        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                cmbPaper.AddItem PtrCtoVbString(.pName)
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, Me.Name, "Printer", cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, "Paper", cmbPaper.Text
End Sub

Private Sub cmbPrinter_Change()
    Call ListPapers
End Sub

Private Sub cmbPrinter_Click()
    Call ListPapers
End Sub


Private Sub FillCombos()
    Dim CC As New clsFillCombos
    Call FormatGrid
    CC.FillAnyCombo cmbCC, "CollectingCenter", True
End Sub


Private Sub FormatGrid()
    With gridOutstanding
        .Cols = 1
        .Rows = 1
        .ColWidth(0) = .Width
    End With
End Sub

Private Sub FormatCCGrid()
    With gridOutstanding
        .Cols = 5
        .Rows = 1
        
        .row = 0
        
        .ColWidth(0) = 660
        .ColWidth(1) = 3800
        .ColWidth(2) = 1800
        .ColWidth(3) = 1800
        .ColWidth(4) = 1800
        
        .col = 0
        .CellAlignment = 4
        .Text = "No"
        
        .col = 1
        .CellAlignment = 4
        .Text = "Collecting Center"
        
        .col = 2
        .CellAlignment = 4
        .Text = "Due Payments"
        
        .col = 3
        .CellAlignment = 4
        .Text = "New Payments"
        
        .col = 4
        .CellAlignment = 4
        .Text = "Total"
    End With
End Sub

Private Sub FormatSGrid()
    With gridOutstanding
        .Cols = 5
        .Rows = 1
        
        .row = 0
        
        .ColWidth(0) = 660
        .ColWidth(1) = 3800
        .ColWidth(2) = 1800
        .ColWidth(3) = 1800
        .ColWidth(4) = 1800
        
        .col = 0
        .CellAlignment = 4
        .Text = "No"
        
        .col = 1
        .CellAlignment = 4
        .Text = "Supplier"
        
        .col = 2
        .CellAlignment = 4
        .Text = "Due Payments"
        
        .col = 3
        .CellAlignment = 4
        .Text = "New Payments"
        
        .col = 4
        .CellAlignment = 4
        .Text = "Total"
        
        
    End With
End Sub

Private Sub FillCCGrid()
    Dim TotalOutstanding As Double
    Dim DuePayments As Double
    Dim NewMilkPayments As Double
    Dim NewCommisions As Double
    Dim rsCC As New ADODB.Recordset
    Dim CCCount As Long
    Dim i As Integer
    Dim LastPaymentGeneratedDate As Date
    With rsCC
        If .State = 1 Then .Close
        temSQL = "Select * from tblCollectingCenter where Deleted = 0  Order by CollectingCenter"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            CCCount = .RecordCount
            .MoveFirst
            pgb1.Value = 0
            pgb1.Min = 0
            pgb1.Max = CCCount + 1
            gridOutstanding.Rows = CCCount + 1
            For i = 1 To CCCount
                pgb1.Value = pgb1.Value + 1
                DoEvents
                DuePayments = CCPaymentsDue(!CollectingCenterID)
                LastPaymentGeneratedDate = FindLastPaymentGenerateDate(!CollectingCenterID)
                NewMilkPayments = CCPeriodMilkSupply(LastPaymentGeneratedDate + 1, Date, !CollectingCenterID)
                NewCommisions = CCOthersCommision(LastPaymentGeneratedDate + 1, Date, !CollectingCenterID)
                gridOutstanding.TextMatrix(i, 0) = i
                gridOutstanding.TextMatrix(i, 1) = !CollectingCenter
                gridOutstanding.TextMatrix(i, 2) = Format(DuePayments, "0.00")
                gridOutstanding.TextMatrix(i, 3) = Format(NewMilkPayments + NewCommisions, "0.00")
                gridOutstanding.TextMatrix(i, 4) = Format(DuePayments + NewMilkPayments + NewCommisions, "0.00")
                TotalOutstanding = TotalOutstanding + DuePayments + NewMilkPayments + NewCommisions
                DoEvents
                .MoveNext
            Next
        End If
        .Close
    End With
    txtOutstanding.Text = Format(TotalOutstanding, "#,##0.00")
End Sub

Private Function FindLastPaymentGenerateDate(CollectingCenterID As Long) As Date
    Dim rsDate As New ADODB.Recordset
    With rsDate
        If .State = 1 Then .Close
        temSQL = "SELECT Max(tblCollectingCenterPaymentSummery.ToDate) AS MaxOfToDate FROM tblCollectingCenterPaymentSummery WHERE (((tblCollectingCenterPaymentSummery.CollectingCenterID)=" & CollectingCenterID & "))"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!MaxOfToDate) = False Then
                FindLastPaymentGenerateDate = !MaxOfToDate
            Else
                FindLastPaymentGenerateDate = #1/1/2007#
            End If
        Else
            FindLastPaymentGenerateDate = #1/1/2007#
        End If
    End With
End Function

Private Sub FillSGrid()
    Dim TotalOutstanding As Double
    Dim DuePayments As Double
    Dim NewMilkPayments As Double
    Dim NewCommisions As Double
    Dim rsCC As New ADODB.Recordset
    Dim CCCount As Long
    Dim i As Integer
    Dim LastPaymentGeneratedDate As Date
    With rsCC
        If .State = 1 Then .Close
        temSQL = "Select * from tblSupplier where Deleted = 0  AND COllectingCenterID = " & Val(cmbCC.BoundText) & " Order by Supplier"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            CCCount = .RecordCount
            pgb1.Value = 0
            pgb1.Min = 0
            pgb1.Max = CCCount + 1
            .MoveFirst
            gridOutstanding.Rows = CCCount + 1
            For i = 1 To CCCount
                pgb1.Value = pgb1.Value + 1
                DoEvents
                DuePayments = SPaymentsDue(!SupplierID)
                LastPaymentGeneratedDate = FindLastPaymentGenerateDate(Val(cmbCC.BoundText))
                NewMilkPayments = SPeriodMilkSupply(LastPaymentGeneratedDate + 1, Date, !SupplierID, Val(cmbCC.BoundText))
                NewCommisions = SOthersCommision(LastPaymentGeneratedDate + 1, Date, !SupplierID)
                gridOutstanding.TextMatrix(i, 0) = i
                gridOutstanding.TextMatrix(i, 1) = !Supplier
                gridOutstanding.TextMatrix(i, 2) = Format(DuePayments, "0.00")
                gridOutstanding.TextMatrix(i, 3) = Format(NewMilkPayments + NewCommisions, "0.00")
                gridOutstanding.TextMatrix(i, 4) = Format(DuePayments + NewMilkPayments + NewCommisions, "0.00")
                TotalOutstanding = TotalOutstanding + DuePayments + NewMilkPayments + NewCommisions
                DoEvents
                .MoveNext
            Next
        End If
        .Close
    End With
    txtOutstanding.Text = Format(TotalOutstanding, "#,##0.00")
End Sub

Private Function CCPaymentsDue(CCID As Long) As Double
    Dim rsPayments As New ADODB.Recordset
    With rsPayments
        If .State = 1 Then .Close
        temSQL = "SELECT Sum(tblSupplierPayments.Value) AS SumOfValue FROM tblSupplierPayments LEFT JOIN tblSupplier ON tblSupplierPayments.SupplierID = tblSupplier.SupplierID WHERE (((tblSupplierPayments.Deleted) = 0) AND ((tblSupplierPayments.Generated)=True) AND ((tblSupplierPayments.Completed) = 0) AND ((tblSupplier.CollectingCenterID)=" & CCID & "))"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then
                CCPaymentsDue = !SumOfValue
            Else
                CCPaymentsDue = 0
            End If
        Else
            CCPaymentsDue = 0
        End If
    End With
End Function

Private Function SPaymentsDue(SID As Long) As Double
    Dim rsPayments As New ADODB.Recordset
    With rsPayments
        If .State = 1 Then .Close
        temSQL = "SELECT Sum(tblSupplierPayments.Value) AS SumOfValue FROM tblSupplierPayments WHERE (((tblSupplierPayments.Deleted) = 0) AND ((tblSupplierPayments.Generated)=True) AND ((tblSupplierPayments.Completed) = 0) AND ((tblSupplierPayments.SupplierID)=" & SID & "))"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then
                SPaymentsDue = !SumOfValue
            Else
                SPaymentsDue = 0
            End If
        Else
            SPaymentsDue = 0
        End If
    End With
End Function



Public Function SPeriodMilkSupply(FromDate As Date, ToDate As Date, SupplierID As Long, CollectingCenterID As Long) As Double
    Dim rsMilk As New ADODB.Recordset
    Dim totalValue As Double
    Dim TotalOwnCommision As Double
    Dim rsCC As New ADODB.Recordset
    With rsMilk
        If .State = 1 Then .Close
        temSQL = "SELECT  sum(tblCollection.Value) as SumOfValue  , sum(tblCollection.Commision) as SumOfCommision  " & _
                    "From tblCollection " & _
                    "WHERE (((tblCollection.ProgramDate) Between '" & Format(FromDate, "dd MMMM yyyy") & "' And '" & Format(ToDate, "dd MMMM yyyy") & "') AND ((tblCollection.SupplierID)=" & SupplierID & ")  And ((tblCollection.Deleted) = 0 ))"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then
                totalValue = totalValue + !SumOfValue
            End If
            If IsNull(!SumOfCommision) = False Then
                TotalOwnCommision = TotalOwnCommision + !SumOfCommision
            End If
        Else
            totalValue = 0
            TotalOwnCommision = 0
        End If
        .Close
    End With
    SPeriodMilkSupply = totalValue + TotalOwnCommision
    Set rsMilk = Nothing
End Function

Public Function CCPeriodMilkSupply(FromDate As Date, ToDate As Date, CollectingCenterID As Long) As Double
    Dim rsMilk As New ADODB.Recordset
    Dim totalValue As Double
    Dim TotalOwnCommision As Double
    Dim rsCC As New ADODB.Recordset
    With rsMilk
        If .State = 1 Then .Close
        temSQL = "SELECT  sum(tblCollection.Value) as SumOfValue  , sum(tblCollection.Commision) as SumOfCommision  " & _
                    "FROM tblCollection LEFT JOIN tblSupplier ON tblCollection.SupplierID = tblSupplier.SupplierID " & _
                    "WHERE (((tblCollection.ProgramDate) Between '" & Format(FromDate, "dd MMMM yyyy") & "' And '" & Format(ToDate, "dd MMMM yyyy") & "') And ((tblCollection.Deleted) = 0 ) AND ((tblSupplier.CollectingCenterID)=" & CollectingCenterID & "))"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then
                totalValue = totalValue + !SumOfValue
            End If
            If IsNull(!SumOfCommision) = False Then
                TotalOwnCommision = TotalOwnCommision + !SumOfCommision
            End If
        Else
            totalValue = 0
            TotalOwnCommision = 0
        End If
        .Close
    End With
    CCPeriodMilkSupply = totalValue + TotalOwnCommision
    Set rsMilk = Nothing
End Function

Public Function CCOthersCommision(FromDate As Date, ToDate As Date, CollectingCenterID As Long) As Double
    Dim rsCS As New ADODB.Recordset
    With rsCS
        If .State = 1 Then .Close
        temSQL = "SELECT tblSupplierGettingCommision.SupplierID, tblSupplierGettingCommision.Supplier " & _
                    "FROM tblSupplier AS tblSupplierGivingCommision LEFT JOIN tblSupplier AS tblSupplierGettingCommision ON tblSupplierGivingCommision.CommisionCollectorID = tblSupplierGettingCommision.SupplierID " & _
                    "Where (((tblSupplierGivingCommision.CommisionCollectorID) <> 0) And ((tblSupplierGivingCommision.CollectingCenterID) = " & CollectingCenterID & ")) " & _
                     "GROUP BY tblSupplierGettingCommision.SupplierID, tblSupplierGettingCommision.Supplier"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                CCOthersCommision = CCOthersCommision + OthersCommision(!SupplierID, FromDate, ToDate, 0)
                .MoveNext
            Wend
        End If
        .Close
    End With
End Function

Public Function SOthersCommision(FromDate As Date, ToDate As Date, SupplierID As Long) As Double
    SOthersCommision = OthersCommision(SupplierID, FromDate, ToDate, 0)
End Function
