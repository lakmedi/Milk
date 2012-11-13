VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmCumulativeDailyMilkCollection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cumulative Report For Daily Milk Collection"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
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
   ScaleHeight     =   8745
   ScaleWidth      =   15270
   Begin VB.TextBox txtVolumeDifference 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7200
      TabIndex        =   25
      Top             =   7800
      Width           =   2415
   End
   Begin VB.TextBox txtDCMRLiters 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7200
      TabIndex        =   23
      Top             =   7320
      Width           =   2415
   End
   Begin VB.TextBox txtValueDifference 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   12600
      TabIndex        =   21
      Top             =   7800
      Width           =   2415
   End
   Begin VB.TextBox txtDCMRValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   12600
      TabIndex        =   19
      Top             =   7320
      Width           =   2415
   End
   Begin VB.TextBox txtGRNLiters 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7200
      TabIndex        =   17
      Top             =   6840
      Width           =   2415
   End
   Begin VB.TextBox txtGRNValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   12600
      TabIndex        =   15
      Top             =   6840
      Width           =   2415
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   6840
      Width           =   3135
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   7320
      Width           =   3135
   End
   Begin MSDataListLib.DataCombo cmbCollectingCenter 
      Height          =   360
      Left            =   1920
      TabIndex        =   0
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
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MM yyyy"
      Format          =   285016067
      CurrentDate     =   39682
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MM yyyy"
      Format          =   285016067
      CurrentDate     =   39682
   End
   Begin MSFlexGridLib.MSFlexGrid gridMilk 
      Height          =   5655
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   9975
      _Version        =   393216
      WordWrap        =   -1  'True
      SelectionMode   =   1
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   13680
      TabIndex        =   7
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
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
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   375
      Left            =   10920
      TabIndex        =   8
      Top             =   8280
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Print Cumulative Report"
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
   Begin btButtonEx.ButtonEx btnStoreReport 
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   8280
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "Print Stores Report"
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
      Caption         =   "Volume Difference"
      Height          =   255
      Left            =   5520
      TabIndex        =   24
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Label Label10 
      Caption         =   "DCMR Total Liters"
      Height          =   255
      Left            =   5520
      TabIndex        =   22
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "Value Difference"
      Height          =   255
      Left            =   10440
      TabIndex        =   20
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "DCMR Total Value"
      Height          =   255
      Left            =   10440
      TabIndex        =   18
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "GRN Total Liters"
      Height          =   255
      Left            =   5520
      TabIndex        =   16
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "GRN Total Value"
      Height          =   255
      Left            =   10440
      TabIndex        =   14
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Paper"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Print"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "&To"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "&From"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmCumulativeDailyMilkCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    Dim rsDailyCollection As New ADODB.Recordset
    Dim i As Integer
    
    Dim NumForms As Long

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


Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnStoreReport_Click()
    Dim CSetPrinter As New cSetDfltPrinter
    CSetPrinter.SetPrinterAsDefault cmbPrinter.Text
    If SelectForm(cmbPaper.Text, Me.hdc) = 1 Then
    
        
        Dim tabReport As Long
        Dim tab1 As Long
        Dim tab2 As Long
        Dim tab3 As Long
        Dim tab4 As Long
    
        tabReport = 35
        tab1 = 5
        tab2 = 40
        tab3 = 60
        tab4 = 65
        
        Printer.Print
        Printer.Font = "Arial"
        Printer.Font.Size = 11
        Printer.Font.Bold = True
        Printer.Print Tab(tabReport); "Stores Report"
        Printer.Font = "Arial"
        Printer.Font.Size = 9
        Printer.Font.Bold = True
        Printer.Print
        Printer.Print Tab(tab1); "Collecting Center :";
        Printer.Print Tab(tab2); cmbCollectingCenter.Text
        Printer.Print Tab(tab1); "From  :";
        Printer.Print Tab(tab2); dtpFrom.Value;
        Printer.Print Tab(tab3); "To  :";
        Printer.Print Tab(tab4); dtpTo.Value;
        Printer.Print
        Printer.Font = "Arial"
        Printer.Font.Size = 8
        Printer.Font.Bold = True
        Dim i As Integer
        Dim tabNo As Long
        Dim tabDate As Long
        Dim tabTotalLiters As Long
        Dim tabCLMR As Long
        Dim tabCFAT As Long
        Dim tabDCMRVALUE As Long
        Dim tabTLMR As Long
        Dim tabTFAT As Long
        Dim tabDiferences
        Dim tabTVALUE As Long
        Dim tabID As Long
            
        tabNo = 10
        tabDate = 16
        tabTotalLiters = 45
        tabCLMR = 65
        tabCFAT = 80
        tabDCMRVALUE = 95
        tabTLMR = 110
        tabTFAT = 120
        tabTVALUE = 135
        tabDiferences = 150
        tabID = 165
    
        Printer.Font.Size = 10
    
        With gridMilk
            For i = 0 To .Rows - 1
                Printer.Print
                Printer.Print Tab(tabNo - Len(.TextMatrix(i, 0))); .TextMatrix(i, 0);
                Printer.Print Tab(tabDate); .TextMatrix(i, 1);
                Printer.Print Tab(tabTotalLiters - Len(.TextMatrix(i, 2))); .TextMatrix(i, 2);
                Printer.Print Tab(tabCLMR - Len(.TextMatrix(i, 5))); .TextMatrix(i, 5);
                Printer.Print Tab(tabCFAT - Len(.TextMatrix(i, 7))); .TextMatrix(i, 7);
                Printer.Print Tab(tabDCMRVALUE - Len(.TextMatrix(i, 6))); .TextMatrix(i, 6);
    '            Printer.Print Tab(tabTLMR - Len(.TextMatrix(i, 8))); .TextMatrix(i, 8);
    '            Printer.Print Tab(tabTFAT - Len(.TextMatrix(i, 9))); .TextMatrix(i, 9);
    '            Printer.Print Tab(tabTVALUE - Len(.TextMatrix(i, 10))); .TextMatrix(i, 10);
    '            Printer.Print Tab(tabDiferences - Len(.TextMatrix(i, 12))); .TextMatrix(i, 12);
    '            Printer.Print Tab(tabID - Len(.TextMatrix(i, 13))); .TextMatrix(i, 13);
                Printer.Print
            Next
        End With
        
        
        'Grn Total Liters
        'DMCR Total Lietrs
        'Grn Total Value
        'DMCR Total values
        'Grn Total difference
        'DMCR Total differences
        
    
       Printer.EndDoc

    End If
    
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
    Dim CSetPrinter As New cSetDfltPrinter
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


Private Sub btnStoreReport_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        bttnPrint.SetFocus
    End If
End Sub

Private Sub bttnPrint_Click()
    Dim CSetPrinter As New cSetDfltPrinter
    CSetPrinter.SetPrinterAsDefault cmbPrinter.Text
    If SelectForm(cmbPaper.Text, Me.hdc) = 1 Then
    
        Dim tabReport As Long
        Dim tab1 As Long
        Dim tab2 As Long
        Dim tab3 As Long
        Dim tab4 As Long
    
        tabReport = 50
        tab1 = 5
        tab2 = 40
        tab3 = 60
        tab4 = 95
        
        On Error Resume Next
        
        Printer.Orientation = cdlLandscape
        
        Printer.Font = "Arial"
        Printer.Font.Size = 11
        Printer.Font.Bold = True
    
        Printer.Print
        Printer.Print Tab(tabReport); "Cumalative Reports For Daily Milk Collection"
        
        Printer.Font = "Arial"
        Printer.Font.Size = 10
        Printer.Font.Bold = True
        
        Printer.Print Tab(tab1); "Collecting Center :";
        Printer.Print Tab(tab2); cmbCollectingCenter.Text
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
        Dim tabDCMRVALUE As Long
        
        Dim tabDCMRAVG As Long
        Dim tabGRNLiters As Long
        
        Dim tabTLMR As Long
        Dim tabTFAT As Long
        Dim tabDiferences
        Dim tabTVALUE As Long
        Dim tabTAVG As Long
        Dim tabID As Long
            
        tabNo = 10
        tabDate = 15
        tabTotalLiters = 43
        tabCLMR = 55
        tabCFAT = 70
        tabDCMRVALUE = 85
        
        tabDCMRAVG = 100
        tabGRNLiters = 115
        
        tabTLMR = 130
        tabTFAT = 145
        tabTVALUE = 160
        tabTAVG = 175
        
        tabDiferences = 190
        
        

        
        tabID = 165

'     0: "No."
'     1: "Date"

'     2: "DMCR Leters"
'     3: "DMCR LMR"
'     4: "DMCR FAT%"
'     5: "DCMR Value"
'     6: "DMCR Average"

'     7: "GRN Leters"
'     8: "Tested LMR"
'     9: "Tested FAT%"
'     10: "Tested Value"
'     11: "GRN Average"

'     12: "Difference"

'     13:"ID"
        
        With gridMilk
            For i = 0 To .Rows - 1
                Printer.Print
                Printer.Print Tab(tabNo - Len(.TextMatrix(i, 0))); .TextMatrix(i, 0);
                Printer.Print Tab(tabDate); .TextMatrix(i, 1);
                Printer.Print Tab(tabTotalLiters - Len(.TextMatrix(i, 2))); .TextMatrix(i, 2);
                Printer.Print Tab(tabCLMR - Len(.TextMatrix(i, 3))); .TextMatrix(i, 3);
                Printer.Print Tab(tabCFAT - Len(.TextMatrix(i, 4))); .TextMatrix(i, 4);
                Printer.Print Tab(tabDCMRVALUE - Len(.TextMatrix(i, 5))); .TextMatrix(i, 5);
                
                Printer.Print Tab(tabDCMRAVG - Len(.TextMatrix(i, 6))); .TextMatrix(i, 6);
                Printer.Print Tab(tabGRNLiters - Len(.TextMatrix(i, 7))); .TextMatrix(i, 7);
                
                Printer.Print Tab(tabTLMR - Len(.TextMatrix(i, 8))); .TextMatrix(i, 8);
                Printer.Print Tab(tabTFAT - Len(.TextMatrix(i, 9))); .TextMatrix(i, 9);
                Printer.Print Tab(tabTVALUE - Len(.TextMatrix(i, 10))); .TextMatrix(i, 10);
                
                Printer.Print Tab(tabTAVG - Len(.TextMatrix(i, 11))); .TextMatrix(i, 11);
                
                Printer.Print Tab(tabDiferences - Len(.TextMatrix(i, 12))); .TextMatrix(i, 12);
                
'                Printer.Print Tab(tabID - Len(.TextMatrix(i, 13))); .TextMatrix(i, 13);
                Printer.Print
            Next
        End With
    
        Printer.Print
        
        Dim tabLabel As Long
        Dim tabValue As Long
        
        tabLabel = 10
        tabValue = 40
        Printer.Font = "Tahoma"
        Printer.Font.Size = 10
        
        Printer.Print Tab(tabLabel); "DCMR Value :"; Tab(tabValue - Len(txtDCMRLiters.Text)); txtDCMRLiters.Text
        Printer.Print Tab(tabLabel); "DCMR Volume :"; Tab(tabValue - Len(txtDCMRValue.Text)); txtDCMRValue.Text
        Printer.Print
        Printer.Print Tab(tabLabel); "GRN Value :"; Tab(tabValue - Len(txtGRNLiters.Text)); txtGRNLiters.Text
        Printer.Print Tab(tabLabel); "GRN Volume :"; Tab(tabValue - Len(txtGRNValue.Text)); txtGRNValue.Text
        Printer.Print
        Printer.Print Tab(tabLabel); "Volume Difference :"; Tab(tabValue - Len(txtVolumeDifference.Text)); txtVolumeDifference.Text
        Printer.Print Tab(tabLabel); "Value Difference :"; Tab(tabValue - Len(txtValueDifference.Text)); txtValueDifference.Text
        Printer.EndDoc
       
    End If
End Sub

Private Sub bttnPrint_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnClose.SetFocus
    End If
End Sub

Private Sub cmbCollectingCenter_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbCollectingCenter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpFrom.SetFocus
    End If
End Sub

Private Sub cmbPrinter_Change()
    Call ListPapers
End Sub

Private Sub cmbPrinter_Click()
    Call ListPapers
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

Private Sub dtpFrom_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnStoreReport.SetFocus
    End If
End Sub

Private Sub Form_Load()
    dtpFrom.Value = Date
    dtpTo.Value = Date
    Call FillCombos
    Call FormatGrid
    Call FillGrid
    Call ListPrinters
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, "Printer", "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, "Paper", "")
    
    Select Case UserAuthorityLevel
    
    
    Case Authority.Viewer '1
    btnStoreReport.Visible = False
    bttnPrint.Visible = False
    
    Case Authority.Analyzer '2
    btnStoreReport.Visible = False
    bttnPrint.Visible = False
    cmbPaper.Visible = False
    cmbPrinter.Visible = False
    
    Case Authority.OrdinaryUser '3
    btnStoreReport.Visible = True
    bttnPrint.Visible = True
   
    Case Authority.PowerUser '4
    btnStoreReport.Visible = True
    bttnPrint.Visible = True
    
    Case Authority.SuperUser '5
    btnStoreReport.Visible = True
    bttnPrint.Visible = True
    
    Case Authority.Administrator '6
    btnStoreReport.Visible = True
    bttnPrint.Visible = True
        
    Case Else
    
    End Select
    
    If CumulativeReportPrintAllowed = True Then
        btnStoreReport.Visible = True
        bttnPrint.Visible = True
    Else
        btnStoreReport.Visible = False
        bttnPrint.Visible = False
    End If
    


End Sub

Private Sub FormatGrid()
    With gridMilk
        
        .Rows = 1
        .Cols = 15
        
        .row = 0
        
        .RowHeight(0) = 700
        
        For i = 0 To .Cols - 1
            Select Case i
                Case 0:
                    .ColWidth(i) = 400
                    .col = i
                    .CellAlignment = 4
                    .Text = "No."
                Case 1:
                    .ColWidth(i) = 1300
                    .col = i
                    .CellAlignment = 4
                    .Text = "Date"
                Case 2:
                    .ColWidth(i) = 1100
                    .col = i
                    .CellAlignment = 4
                    .Text = "DCMR Vol"
                Case 3:
                    .ColWidth(i) = 1100
                    .col = i
                    .CellAlignment = 4
                    .Text = "DCMR LMR"
                Case 4:
                    .ColWidth(i) = 1100
                    .col = i
                    .CellAlignment = 4
                    .Text = "DCMR FAT%"
                Case 5:
                    .ColWidth(i) = 1400
                    .col = i
                    .CellAlignment = 4
                    .Text = "DCMR Value"
                Case 6:
                    .ColWidth(i) = 1200
                    .col = i
                    .CellAlignment = 4
                    .Text = "DCMR Average"
                    
                Case 7:
                    .ColWidth(i) = 1100
                    .col = i
                    .CellAlignment = 4
                    .Text = "GRN Vol"
                Case 8:
                    .ColWidth(i) = 1100
                    .col = i
                    .CellAlignment = 4
                    .Text = "GRN LMR"
                Case 9:
                    .ColWidth(i) = 1100
                    .col = i
                    .CellAlignment = 4
                    .Text = "GRN FAT%"
                Case 10:
                    .ColWidth(i) = 1400
                    .col = i
                    .CellAlignment = 4
                    .Text = "GRN Value"
                Case 11:
                    .ColWidth(i) = 1200
                    .col = i
                    .CellAlignment = 4
                    .Text = "GRN Average"
                    
                    
                Case 12:
                    .ColWidth(i) = 1400
                    .col = i
                    .CellAlignment = 4
                    .Text = "Difference"
                Case 13:
                    .ColWidth(i) = 0
                    .col = i
                    .CellAlignment = 4
                    .Text = "ID"
                Case 14:
                    .ColWidth(i) = 0
                    .col = i
                    .CellAlignment = 4
                    .Text = "Supplier ID"
            End Select
        Next
    End With
    
    
    
'     0: "No."
'     1: "Date"
'     2: "Total Leters"
'     3:"C. LMR"
'     4: "C. FAT%"
'     5:"DCMR Value"
'     6:"T. LMR"
'     7: "T. FAT%"
'     8:"T. Value"
'     9: "Difference"
'     10:"ID"
    
    
    
'     0: "No."
'     1: "Date"

'     2: "DMCR Leters"
'     3: "DMCR LMR"
'     4: "DMCR FAT%"
'     5: "DCMR Value"
'     6: "DMCR Average"

'     7: "GRN Leters"
'     8: "Tested LMR"
'     9: "Tested FAT%"
'     10: "Tested Value"
'     11: "GRN Average"

'     12: "Difference"

'     13:"ID"
    

End Sub

Private Sub FillGrid()
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then Exit Sub
   
    Dim NoOfDays As Long
    Dim MaxDate As Date
    Dim MinDate As Date
    Dim ThisDate As Date
    
    Dim temDCMRVolume As Double
    Dim temDCMRValue As Double
    Dim temGRNValue As Double
    Dim temGRNVolume As Double
    
    Dim temDifference As Double
    
    
    
    If dtpFrom.Value < dtpTo.Value Then
        MaxDate = dtpTo.Value
        MinDate = dtpFrom.Value
    Else
        MaxDate = dtpFrom.Value
        MinDate = dtpTo.Value
    End If
    
    dtpFrom.Value = MinDate
    dtpTo.Value = MaxDate
    
    NoOfDays = DateDiff("d", MinDate, MaxDate) + 1
    
    ThisDate = MinDate
    
    For i = 1 To NoOfDays
        
        With rsDailyCollection
            If .State = 1 Then .Close
            temSQL = "SELECT tblDailyCollection.* FROM tblDailyCollection Where ProgramDate = '" & Format(ThisDate, "dd MMMM yyyy") & "'  And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText) & " order by tblDailyCollection.DailyCollectionID DESC"
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                gridMilk.Rows = gridMilk.Rows + 1
                gridMilk.row = gridMilk.Rows - 1
                gridMilk.TextMatrix(gridMilk.row, 0) = gridMilk.row
                gridMilk.TextMatrix(gridMilk.row, 1) = Format(ThisDate, "dd MMM yyyy")
                gridMilk.TextMatrix(gridMilk.row, 2) = Format(!TotalVolume, "0.00")
                temDCMRVolume = temDCMRVolume + !TotalVolume
                gridMilk.TextMatrix(gridMilk.row, 3) = Format(Round(!CalculatedLMR, 2), "0.00")
                
                gridMilk.TextMatrix(gridMilk.row, 4) = Format(Round(!CalculatedFat, 2), "0.00")
                gridMilk.TextMatrix(gridMilk.row, 5) = Format(!totalValue, "0.00")
                
                temDCMRValue = temDCMRValue + !totalValue
                
                If IsNull(!TotalVolume) = False Then
                    If Val(!TotalVolume) <> 0 Then
                        gridMilk.TextMatrix(gridMilk.row, 6) = Format(!totalValue / !TotalVolume, "0.00")
                    Else
                        gridMilk.TextMatrix(gridMilk.row, 6) = "0.00"
                    End If
                Else
                    gridMilk.TextMatrix(gridMilk.row, 6) = "0.00"
                End If
                                
                If IsNull(!ActualVolume) = False Then
                    gridMilk.TextMatrix(gridMilk.row, 7) = Format(!ActualVolume, "0.00")
                Else
                    gridMilk.TextMatrix(gridMilk.row, 7) = "0.00"
                End If
                
                gridMilk.TextMatrix(gridMilk.row, 8) = Format(!TestedLMR, "0.0")
                
                
                
                gridMilk.TextMatrix(gridMilk.row, 9) = Format(!TestedFAT, "0.00")
                
                gridMilk.TextMatrix(gridMilk.row, 10) = Format(Round(!actualValue, 10), "0.00")
                
                If IsNull(!ActualVolume) = False Then
                    If !ActualVolume <> 0 Then
                        gridMilk.TextMatrix(gridMilk.row, 11) = Format(Round(!actualValue / !ActualVolume, 10), "0.00")
                    Else
                        gridMilk.TextMatrix(gridMilk.row, 11) = "0.00"
                    End If
                    temGRNVolume = temGRNVolume + !ActualVolume
                Else
                    gridMilk.TextMatrix(gridMilk.row, 11) = "0.00"
                End If
                temGRNValue = temGRNValue + Round(!actualValue, 2)
                
                gridMilk.TextMatrix(gridMilk.row, 12) = Format(Round(!ValueDifference, 2), "0.00")
                
                If Not IsNull(!ValueDifference) Then
                
                    temDifference = temDifference + Round(!ValueDifference, 2)
                
                End If
                
            End If
            .Close
        End With
        ThisDate = MinDate + i
    Next

    With gridMilk
        .Rows = .Rows + 1
        .row = .Rows - 1
        
        .col = 1
        .Text = "Total"
        
        .col = 5
        .Text = temDCMRValue
        
        .col = 12
        .Text = temDifference
        
        .col = 2
        .Text = temDCMRVolume
        
        .col = 10
        .Text = temGRNValue
        
        .col = 7
        .Text = temGRNVolume
        
    End With

'     0: "No."
'     1: "Date"
'     2: "Total Leters"
'     3:"C. LMR"
'     4: "C. FAT%"
'     5:"DCMR Value"
'     6:"T. LMR"
'     7: "T. FAT%"
'     8:"T. Value"
'     9: "Difference"
'     10:"ID"

    txtDCMRLiters.Text = Format(temDCMRVolume, "0.00")
    txtDCMRValue.Text = Format(temDCMRValue, "0.00")
    txtGRNLiters.Text = Format(temGRNVolume, "0.00")
    txtGRNValue.Text = Format(temGRNValue, "0.00")
    txtValueDifference.Text = Format(temDCMRValue - temGRNValue, "0.00")
    txtVolumeDifference.Text = Format(temDCMRVolume - temGRNVolume, "0.00")
End Sub


Private Sub FillCombos()
    Dim Center As New clsFind
    Center.FillCombo cmbCollectingCenter, "tblCollectingCenter", "CollectingCenter", "CollectingCenterID", True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, Me.Name, "Printer", cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, "Paper", cmbPaper.Text
End Sub
