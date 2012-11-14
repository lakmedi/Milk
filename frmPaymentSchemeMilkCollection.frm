VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmPaymentSchemeMilkCollection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Scheme Milk Supply"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
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
   ScaleHeight     =   8670
   ScaleWidth      =   9075
   Begin VB.ComboBox cmbSummeryPrinter 
      Height          =   360
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   7680
      Width           =   2655
   End
   Begin VB.ComboBox cmbSUmmeryPaper 
      Height          =   360
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   8160
      Width           =   2655
   End
   Begin VB.TextBox txtCommision 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6960
      TabIndex        =   21
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox txtDeductions 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6960
      TabIndex        =   20
      Top             =   7170
      Width           =   1935
   End
   Begin VB.TextBox txtTotalPayment 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6960
      TabIndex        =   19
      Top             =   7605
      Width           =   1935
   End
   Begin VB.TextBox txtMilkPayments 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6960
      TabIndex        =   9
      Top             =   6360
      Width           =   1935
   End
   Begin VB.TextBox txtVolume 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   3360
      TabIndex        =   8
      Top             =   6360
      Width           =   1935
   End
   Begin VB.TextBox txtAverage 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   3360
      TabIndex        =   7
      Top             =   6840
      Width           =   1935
   End
   Begin MSDataListLib.DataCombo cmbCollectingCenter 
      Height          =   360
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbCCPS 
      Height          =   360
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbPS 
      Height          =   360
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid gridSummery 
      Height          =   4455
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7858
      _Version        =   393216
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   7800
      TabIndex        =   13
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16776960
      Caption         =   "Pr&ocess"
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
      Left            =   6120
      TabIndex        =   14
      Top             =   8160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16711935
      Caption         =   "Print"
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
      Left            =   7800
      TabIndex        =   15
      Top             =   8160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16711680
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
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   7440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   142475267
      CurrentDate     =   39682
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   8160
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   122683395
      CurrentDate     =   39682
   End
   Begin MSComctlLib.ProgressBar pgb1 
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   6000
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Commisions"
      Height          =   360
      Left            =   5520
      TabIndex        =   24
      Top             =   6765
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Deductions"
      Height          =   360
      Left            =   5520
      TabIndex        =   23
      Top             =   7170
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payment"
      Height          =   360
      Left            =   5520
      TabIndex        =   22
      Top             =   7605
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Milk Payments"
      Height          =   360
      Left            =   5520
      TabIndex        =   12
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Vloume"
      Height          =   360
      Left            =   1920
      TabIndex        =   11
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Average"
      Height          =   360
      Left            =   1920
      TabIndex        =   10
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Center"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Collecting Center"
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment C&ycle"
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmPaymentSchemeMilkCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
   
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
    Private CSetPrinter As New cSetDfltPrinter
    
    Private Const BIF_RETURNONLYFSDIRS = 1
    Private Const BIF_DONTGOBELOWDOMAIN = 2
    Private Const MAX_PATH = 260
    
    Private Declare Function SHBrowseForFolder Lib "shell32" _
                                      (lpbi As BrowseInfo) As Long
    
    Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                      (ByVal pidList As Long, _
                                      ByVal lpBuffer As String) As Long
    
    Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                      (ByVal lpString1 As String, ByVal _
                                      lpString2 As String) As Long
    
    Private Type BrowseInfo
       hWndOwner      As Long
       pIDLRoot       As Long
       pszDisplayName As Long
       lpszTitle      As Long
       ulFlags        As Long
       lpfnCallback   As Long
       lparam         As Long
       iImage         As Long
    End Type
    
    Dim rsCCPS As New ADODB.Recordset

Private Sub btnPrint_Click()
    CSetPrinter.SetPrinterAsDefault cmbSummeryPrinter.Text
    If SelectForm(cmbSUmmeryPaper.Text, Me.hdc) = 1 Then
        Dim i As Integer
        Dim tabNo As Integer
        Dim tabName As Integer
        Dim tabLMR As Integer
        Dim tabFAT As Integer
        Dim tabLiters As Integer
        Dim tabSNF As Integer
        Dim tabMilk As Integer
        Dim tabCommision As Integer
        
        Dim tabAverage As Integer
        
        Dim tabDeduction As Integer
        Dim tabNetPayment As Integer
        Dim tabCenter As Integer
        
        tabNo = 6
        tabName = 11
        tabLMR = 63
        tabFAT = 70
        tabSNF = 77
        tabLiters = 84
        tabCommision = 96
        tabDeduction = 110
        tabAverage = 121
        tabNetPayment = 131

        tabCenter = 45
        
        On Error Resume Next
        Printer.Orientation = cdlLandscape
        Printer.PrintQuality = vbPRPQHigh
        On Error GoTo 0
        
        Printer.CurrentY = 720
        
        Printer.Font.Name = "Arial Black"
        Printer.Font.Size = 14
        Printer.Font.Bold = True
        
        Printer.Print Tab(tabCenter - (Len(InstitutionName) / 2)); InstitutionName
        
        Printer.Font.Name = "Arial"
        Printer.Font.Size = 12
        Printer.Font.Bold = False

        
        Printer.Print Tab(tabCenter - (Len(InstitutionAddressLine1) / 2)); InstitutionAddressLine1
        Printer.Print Tab(tabCenter - (Len(InstitutionAddressLine2) / 2)); InstitutionAddressLine2
        Printer.Print Tab(tabCenter - (Len(InstitutionAddressLine3) / 2)); InstitutionAddressLine3
        
        
        Printer.Print Tab(tabNo); "Payment Summery - " & cmbCollectingCenter.Text
        Printer.Print Tab(tabNo); "From " & Format(dtpFrom.Value, LongDateFormat) & " to " & Format(dtpTo.Value, LongDateFormat)
        
        Printer.Font.Name = "Verdana"
        Printer.Font.Size = 11
        Printer.Font.Bold = False
        Printer.Font.Underline = True
        
        tabCenter = 45
        
        Printer.Print
        Printer.Print Tab(tabNo); "No.";
        Printer.Print Tab(tabName); "Name";
        Printer.Print Tab(tabLMR - (Len("LMR"))); "LMR";
        Printer.Print Tab(tabFAT - (Len("Fat"))); "Fat";
        Printer.Print Tab(tabSNF - (Len("SNF"))); "SNF";
        Printer.Print Tab(tabLiters - (Len("Vol"))); "Vol";
        Printer.Print Tab(tabCommision - (Len("Comm"))); "Com.";
        Printer.Print Tab(tabDeduction - (Len("Ded"))); "Ded.";
        Printer.Print Tab(tabAverage - (Len("Avg."))); "Avg.";
        Printer.Print Tab(tabNetPayment - (Len("Net"))); "Net"
        
        Printer.Font.Name = "Verdana"
        Printer.Font.Size = 11
        Printer.Font.Bold = False
        Printer.Font.Underline = False

        Dim AllRows As Integer
        Dim PrintPageNo As Integer
        PrintPageNo = 1
        With gridSummery
            AllRows = 0
            For i = 1 To .Rows - 1
                AllRows = AllRows + 1
                Printer.Print Tab(tabNo); .TextMatrix(i, 0);
                Printer.Print Tab(tabName); .TextMatrix(i, 1);
                Printer.Print Tab(tabLMR - Len(.TextMatrix(i, 3))); .TextMatrix(i, 3);
                Printer.Print Tab(tabFAT - Len(.TextMatrix(i, 4))); .TextMatrix(i, 4);
                Printer.Print Tab(tabSNF - Len(.TextMatrix(i, 6))); .TextMatrix(i, 6);
                Printer.Print Tab(tabLiters - Len(.TextMatrix(i, 5))); .TextMatrix(i, 5);
                'Printer.Print Tab(tabMilk - Len(.TextMatrix(i, 8))); .TextMatrix(i, 8);
                Printer.Print Tab(tabCommision - Len(.TextMatrix(i, 2))); .TextMatrix(i, 2);
                Printer.Print Tab(tabDeduction - Len(.TextMatrix(i, 9))); .TextMatrix(i, 9);
                Printer.Print Tab(tabAverage - Len(.TextMatrix(i, 7))); .TextMatrix(i, 7);
                Printer.Print Tab(tabNetPayment - Len(.TextMatrix(i, 10))); .TextMatrix(i, 10)
                If AllRows > 27 And PrintPageNo = 1 Then
                    Printer.NewPage
                    Printer.CurrentY = 720
                    Printer.Print
                    AllRows = 0
                    PrintPageNo = PrintPageNo + 1
                ElseIf AllRows > 34 And PrintPageNo > 1 Then
                    Printer.NewPage
                    Printer.CurrentY = 720
                    Printer.Print
                    AllRows = 0
                    PrintPageNo = PrintPageNo + 1
                End If
            Next
        End With
        
        tabMilk = 10
        tabNetPayment = 100
        
        Printer.Print Tab
        Printer.Print Tab(tabMilk); "Total Milk Payment";
        Printer.Print Tab(tabNetPayment - (Len(txtMilkPayments))); txtMilkPayments.Text
        Printer.Print Tab(tabMilk); "Total Commision";
        Printer.Print Tab(tabNetPayment - (Len(txtCommision))); txtCommision.Text
        Printer.Print Tab(tabMilk); "Total Deduction";
        Printer.Print Tab(tabNetPayment - (Len(txtDeductions))); txtDeductions.Text
        Printer.Print Tab(tabMilk); "Total Payment";
        Printer.Print Tab(tabNetPayment - (Len(txtTotalPayment))); txtTotalPayment.Text
        Printer.Print
        Printer.Print Tab(tabMilk); "Total Volume";
        Printer.Print Tab(tabNetPayment - (Len(txtVolume.Text))); txtVolume.Text
        Printer.Print
        Printer.Print Tab(tabMilk); "Total Volume";
        Printer.Print Tab(tabNetPayment - (Len(txtAverage.Text))); txtAverage.Text
        
        Printer.Print
        Printer.Print
        
        Printer.Print Tab(tabMilk); ".....................                           .....................                 ....................."
        Printer.Print Tab(tabMilk); "Prepaired By                               Checked by                      Autherized by"
        Printer.Print
        
        
        Printer.EndDoc
    Else
        MsgBox "Printer Error"
    End If
End Sub

Private Sub btnProcess_Click()
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then
        MsgBox "Please select a Collecting center"
        cmbCollectingCenter.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPS.BoundText) = False Then
        MsgBox "Please select a Payment Scheme"
        cmbPS.SetFocus
        Exit Sub
    End If
    
    
    DoEvents
    
    Dim temDays As Integer
    Dim TemDay As Date
    
    Dim i As Integer
    
    Dim rsCenter As New ADODB.Recordset
    
    Dim Supplier As String
    Dim SupplierID As Long
    Dim SupplierDeduction As Double
    
    Dim temLMR As Double
    Dim temFAT As Double
    Dim temLiters As Double
    Dim temValue As Double
    Dim temPrice As Double
    
    Dim temOwnCommision As Double
    Dim temOthersCommision As Double
    Dim temAdditionalCommision As Double
    
    Dim TotalMilkPayments As Double
    Dim TotalCommisions As Double
    Dim TotalDeductions As Double
    Dim TotalNetPayment As Double
    Dim TotalVolume As Double
    
    Dim MyMilkCollectionM As MilkCollection
    Dim MyMilkCollectionE As MilkCollection
    
    temDays = DateDiff("d", dtpFrom.Value, dtpTo.Value) + 1
    
    If temDays > 31 Then
        MsgBox "No more than 31 days can be processed"
        dtpTo.SetFocus
        Exit Sub
    End If
        
    If temDays < 1 Then
        MsgBox "From and To dates are not correct"
        dtpTo.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass

        
    With gridSummery
        .Clear
        .Rows = 1
        .Cols = 12
        .col = 0
        .Text = "No."
        .col = 1
        .Text = "Name"
        .col = 2
        .Text = "Commision"
        .col = 3
        .Text = "Avg. LMR"
        .col = 4
        .Text = "Avg. FAT%"
        .col = 5
        .Text = "Total Leters"
        .col = 6
        .Text = "Avg. SNF"
        .col = 7
        .Text = "Farmer Avg."
        .col = 8
        .Text = "Milk Payment"
        .col = 9
        .Text = "Deductions"
        .col = 10
        .Text = "Net Payment"
        .col = 11
        .Text = "Supplier ID"
        .ColWidth(11) = 1
        
        End With
            If rsCenter.State = 1 Then rsCenter.Close
            temSQL = "Select * from tblSupplier where CollectingCenterID = " & Val(cmbCollectingCenter.BoundText) & " And Deleted = 0  "
            
            temSQL = "SELECT tblSupplier.Supplier, tblSupplier.SupplierID " & _
                        "FROM tblSupplier LEFT JOIN tblCollectingCenter ON tblSupplier.CollectingCenterID = tblCollectingCenter.CollectingCenterID " & _
                        "WHERE (((tblSupplier.PaymentSchemeID)=" & cmbPS.BoundText & ") AND ((tblSupplier.CollectingCenterID)=" & cmbCollectingCenter.BoundText & ")) OR (((tblSupplier.PaymentSchemeID)=0 Or (tblSupplier.PaymentSchemeID) Is Null) AND ((tblCollectingCenter.PaymentSchemeID)=" & cmbPS.BoundText & ") AND ((tblSupplier.CollectingCenterID)=" & cmbCollectingCenter.BoundText & "))"

            
            
            
            
            rsCenter.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            
            If rsCenter.RecordCount > 0 Then
                rsCenter.MoveLast
                rsCenter.MoveFirst
                
                With pgb1
                    .Value = 0
                    .Min = 0
                    .Max = rsCenter.RecordCount
                End With
                
                While rsCenter.EOF = False
                    Supplier = rsCenter!Supplier
                    SupplierID = rsCenter!SupplierID
                                            
                    pgb1.Value = pgb1.Value + 1
                    DoEvents
                        
                    temLMR = 0
                    temFAT = 0
                    temLiters = 0
                    temValue = 0
                    temPrice = 0

                    MyMilkCollectionM = PeriodMilkSupply(dtpFrom.Value, dtpTo.Value, SupplierID, 1)
                    MyMilkCollectionE = PeriodMilkSupply(dtpFrom.Value, dtpTo.Value, SupplierID, 2)
                    
                    temOwnCommision = OwnCommision(SupplierID, dtpFrom.Value, dtpTo.Value, 0)
                    temOthersCommision = OthersCommision(SupplierID, dtpFrom.Value, dtpTo.Value, 0)
                    temAdditionalCommision = AdditionalCommision(SupplierID, dtpFrom.Value, dtpTo.Value)
                    
                    SupplierDeduction = PeriodDeductions(SupplierID, dtpFrom.Value, dtpTo.Value)
                    
                    If MyMilkCollectionM.Value <> 0 Or temOwnCommision <> 0 Or MyMilkCollectionE.Value <> 0 Or temOthersCommision <> 0 Or SupplierDeduction <> 0 Or temAdditionalCommision <> 0 Then
                        With gridSummery
                            .Rows = .Rows + 1
                            .row = .Rows - 1
                            .col = 0
                            .Text = .row
                            .col = 1
                            .Text = Supplier
                            .col = 11
                            .Text = SupplierID
                            .col = 2
    
                            If temOwnCommision <> 0 Or temAdditionalCommision <> 0 Or temOthersCommision <> 0 Then .Text = Format(temOwnCommision + temOthersCommision + temAdditionalCommision, "0.00")
                            TotalCommisions = TotalCommisions + temOwnCommision + temOthersCommision + temAdditionalCommision
                            If MyMilkCollectionM.Supplied = True Or MyMilkCollectionE.Supplied = True Then
                                .col = 3
                                temLMR = ((MyMilkCollectionM.LMR * MyMilkCollectionM.Liters) + (MyMilkCollectionE.LMR * MyMilkCollectionE.Liters)) / (MyMilkCollectionE.Liters + MyMilkCollectionM.Liters)
                                .Text = Format(temLMR, "0.00")
                                .col = 4
                                temFAT = ((MyMilkCollectionM.FAT * MyMilkCollectionM.Liters) + (MyMilkCollectionE.FAT * MyMilkCollectionE.Liters)) / (MyMilkCollectionE.Liters + MyMilkCollectionM.Liters)
                                .Text = Format(temFAT, "0.00")
                                .col = 5
                                temLiters = MyMilkCollectionM.Liters + MyMilkCollectionE.Liters
                                .Text = Format(temLiters, "0.00")
                                .col = 6
                                .Text = Format(SNF(temLMR, temFAT), "0.0")
                                .col = 7
                                temValue = MyMilkCollectionM.Value + MyMilkCollectionE.Value
                                temPrice = temValue / temLiters
                                .Text = Format(temPrice, "0.00")
                                .col = 8
                                .Text = Format(temValue, "0.00")
                            End If
                            .col = 9
                            TotalDeductions = TotalDeductions + SupplierDeduction
                            If SupplierDeduction <> 0 Then .Text = Format(SupplierDeduction, "0.00")
                            .col = 10
                            .Text = Format(MyMilkCollectionM.Value + MyMilkCollectionE.Value - SupplierDeduction + temOwnCommision + temOthersCommision + temAdditionalCommision, "0.00")
                            TotalMilkPayments = TotalMilkPayments + temValue
                            TotalNetPayment = TotalNetPayment + temValue - SupplierDeduction + temOwnCommision + temOthersCommision + temAdditionalCommision
                            TotalVolume = TotalVolume + temLiters
                        End With
                    End If
                    rsCenter.MoveNext
                    DoEvents
                Wend
            End If
            rsCenter.Close
            txtVolume.Text = Format(TotalVolume, "0.00")
            txtMilkPayments.Text = Format(TotalMilkPayments, "0.00")
            txtCommision.Text = Format(TotalCommisions, "0.00")
            txtDeductions.Text = Format(TotalDeductions, "0.00")
            txtTotalPayment.Text = Format(TotalNetPayment, "0.00")
            If TotalVolume <> 0 Then
                txtAverage.Text = Format(TotalMilkPayments / TotalVolume, "0.000")
            Else
                txtAverage.Text = "0.00"
            End If
            
'            btnProcess.Enabled = False
            pgb1.Value = 0
            Screen.MousePointer = vbDefault

End Sub

Private Sub cmbCCPS_Change()
    Dim rstemCCPS As New ADODB.Recordset
    With rstemCCPS
        If .State = 1 Then .Close
        temSQL = "Select * from tblCollectingCenterPaymentSummery where CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            dtpFrom.Value = !FromDate
            dtpTo.Value = !ToDate
            .Close
        Else
            .Close
        End If
    End With
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FillPrinters
    Call GetSettings
End Sub



Private Sub GetSettings()
    On Error Resume Next
    cmbSummeryPrinter.Text = GetSetting(App.EXEName, Me.Name, "SummeryPrinter", "")
    cmbSUmmeryPaper.Text = GetSetting(App.EXEName, Me.Name, "SummeryPaper", "")
End Sub

Private Sub FillPrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbSummeryPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub FillPapers(PaperCombo As ComboBox, MyPrinterName As String)
    PaperCombo.Clear
    CSetPrinter.SetPrinterAsDefault (MyPrinterName)
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
                PaperCombo.AddItem PtrCtoVbString(.pName)
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, Me.Name, "SummeryPrinter", cmbSummeryPrinter.Text
    SaveSetting App.EXEName, Me.Name, "SummeryPaper", cmbSUmmeryPaper.Text
End Sub

Private Sub FillCombos()
    Dim Centers As New clsFillCombos
    Centers.FillAnyCombo cmbCollectingCenter, "CollectingCenter", True
    Dim PaymentScheme As New clsFillCombos
    PaymentScheme.FillAnyCombo cmbPS, "PaymentScheme", True
End Sub

Private Sub cmbCollectingCenter_Change()
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then Exit Sub
    With rsCCPS
        temSQL = "SELECT ('From ' +  convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) + ' To ' +  convert(varchar, tblCollectingCenterPaymentSummery.ToDate, 102)) AS Display, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID " & _
                    "FROM tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID " & _
                    "Where (((tblCollectingCenterPaymentSummery.CollectingCenterID) = " & Val(cmbCollectingCenter.BoundText) & ")) " & _
                    "ORDER BY tblCollectingCenterPaymentSummery.FromDate DESC"
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCCPS
        Set .RowSource = rsCCPS
        .ListField = "Display"
        .BoundColumn = "CollectingCenterPaymentSummeryID"
    End With
End Sub



