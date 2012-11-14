VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAllCollectingCentrePaymentSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate Collecting Center Payment Advice"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13425
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
   Picture         =   "frmAllCollectingCentrePaymentSummery.frx":0000
   ScaleHeight     =   9870
   ScaleWidth      =   13425
   Begin VB.TextBox txtAverage 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   1800
      TabIndex        =   29
      Top             =   7440
      Width           =   2775
   End
   Begin MSComctlLib.ProgressBar pgb1 
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   9480
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtVolume 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   1800
      TabIndex        =   26
      Top             =   6960
      Width           =   2775
   End
   Begin VB.ComboBox cmbSUmmeryPaper 
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   8880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox cmbMilkPrinter 
      Height          =   360
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   8400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox cmbMilkPaper 
      Height          =   360
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   8880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox cmbCommisionPrinter 
      Height          =   360
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   8400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox cmbCommisionPaper 
      Height          =   360
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   8880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox cmbSummeryPrinter 
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   8400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtCCPS 
      Height          =   495
      Left            =   9240
      TabIndex        =   18
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.TextBox txtTotalPayment 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10440
      TabIndex        =   17
      Top             =   8280
      Width           =   2775
   End
   Begin VB.TextBox txtDeductions 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10440
      TabIndex        =   16
      Top             =   7845
      Width           =   2775
   End
   Begin VB.TextBox txtCommision 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10440
      TabIndex        =   15
      Top             =   7395
      Width           =   2775
   End
   Begin VB.TextBox txtMilkPayments 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10440
      TabIndex        =   14
      Top             =   6960
      Width           =   2775
   End
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   10680
      TabIndex        =   4
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BackColor       =   16711680
      Caption         =   "&Save"
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
   Begin btButtonEx.ButtonEx btnPrintMilkPayments 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   7920
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16711935
      Caption         =   "Print &Milk Payments"
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
   Begin btButtonEx.ButtonEx btnPrintSummery 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16711935
      Caption         =   "&Print Summery"
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
      Left            =   12000
      TabIndex        =   5
      Top             =   8880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin MSFlexGridLib.MSFlexGrid gridSummery 
      Height          =   6255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   11033
      _Version        =   393216
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   1
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   122683395
      CurrentDate     =   39682
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   142213123
      CurrentDate     =   39682
   End
   Begin btButtonEx.ButtonEx btnPrintCommisions 
      Height          =   375
      Left            =   5640
      TabIndex        =   19
      Top             =   7920
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16711935
      Caption         =   "Print &Commisions"
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
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Average Per Liter"
      Height          =   360
      Left            =   120
      TabIndex        =   30
      Top             =   7440
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Vloume"
      Height          =   360
      Left            =   120
      TabIndex        =   27
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payment"
      Height          =   360
      Left            =   8760
      TabIndex        =   13
      Top             =   8280
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Deductions"
      Height          =   360
      Left            =   8760
      TabIndex        =   12
      Top             =   7845
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Commisions"
      Height          =   360
      Left            =   8760
      TabIndex        =   11
      Top             =   7440
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Milk Payments"
      Height          =   360
      Left            =   8760
      TabIndex        =   10
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&From"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&To"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmAllCollectingCentrePaymentSummery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 011 2
    
    Dim temSQL As String
   
    Dim MinusValue As Boolean
   
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
    
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrintAll_Click()
    Dim i As Integer
    With gridSummery
        For i = 1 To .Rows - 1
            SelectPrinter
            SelectPaper
            PrintSinglePayment Val(.TextMatrix(i, 11)), Val(txtCCPS.Text)
            Printer.NewPage
        Next
    End With
    Printer.EndDoc
End Sub

Private Sub btnPrintCommisions_Click()
    CSetPrinter.SetPrinterAsDefault cmbSummeryPrinter.Text
    If SelectForm(cmbSUmmeryPaper.Text, Me.hdc) = 1 Then
        Dim rsCenter As New ADODB.Recordset
        Dim i As Integer
        Dim tabCenter As Integer
        Dim tabNo As Integer
        Dim tabName As Integer
        Dim tabLiters As Integer
        Dim rsS As New ADODB.Recordset
        Dim MyMilkCollection As MilkCollection
        Dim TotalMilk As Double
        Dim temDays As Integer
        Dim AvgMilk As Double
        Dim temCommsisionRate As Double
        Dim temCommsision As Double
        
        Dim temSupplierID As Long
        Dim temSupplier As String
        
        temDays = DateDiff("d", dtpFrom.Value, dtpTo.Value) + 1
        
        If rsCenter.State = 1 Then rsCenter.Close
        
        temSQL = "SELECT DISTINCT tblSupplier.Supplier, tblSupplier.SupplierID, tblSupplier.SupplierCode " & _
                    "FROM tblSupplier AS tblSupplierThroughCollector LEFT JOIN tblSupplier ON tblSupplierThroughCollector.CommisionCollectorID = tblSupplier.SupplierID " & _
                    "Where (((tblSupplier.Supplier) Is Not Null) And ((tblSupplierThroughCollector.Deleted) = 0 ))"
        
        
        rsCenter.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If rsCenter.RecordCount > 0 Then
            rsCenter.MoveLast
            rsCenter.MoveFirst
            
            tabNo = 5
            tabName = 10
            tabLiters = 70
            
            While rsCenter.EOF = False
                temSupplier = rsCenter!Supplier
                temSupplierID = rsCenter!SupplierID
            
                Printer.Print Tab(tabCenter - (Len(InstitutionName) / 2)); InstitutionName
                Printer.Print Tab(tabCenter - (Len(InstitutionAddressLine1) / 2)); InstitutionAddressLine1
                Printer.Print Tab(tabCenter - (Len(InstitutionAddressLine2) / 2)); InstitutionAddressLine2
                Printer.Print Tab(tabCenter - (Len(InstitutionAddressLine3) / 2)); InstitutionAddressLine3
                Printer.Print "Commisions - " & rsCenter!Supplier & " (" & rsCenter!SupplierCode & ")"
                Printer.Print "From " & Format(dtpFrom.Value, LongDateFormat) & " to " & Format(dtpTo.Value, LongDateFormat)
                Printer.Print
                With rsS
                    If .State = 1 Then .Close
                    temSQL = "SELECT tblSupplier.SupplierID, tblSupplier.Supplier, tblSupplier.SupplierCode " & _
                                "From tblSupplier " & _
                                "Where (((tblSupplier.CommisionCollectorID) = " & temSupplierID & ")) " & _
                                "ORDER BY tblSupplier.Supplier"
                    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                    If .RecordCount > 0 Then
                        .MoveLast
                        .MoveFirst
                        TotalMilk = 0
                        For i = 1 To .RecordCount
                            Printer.Print Tab(tabNo); i;
                            Printer.Print Tab(tabName); !Supplier & " " & !SupplierCode;
                            MyMilkCollection = PeriodMilkSupply(dtpFrom.Value, dtpTo.Value, !SupplierID, 0)
                            TotalMilk = TotalMilk + MyMilkCollection.Liters
                            Printer.Print Tab(tabLiters - Len(Format(MyMilkCollection.Liters, "0.00"))); Format(MyMilkCollection.Liters, "0.00")
                            .MoveNext
                        Next
                    End If
                    .Close
                End With
                
                Printer.Print
                Printer.Print Tab(tabName); "Milk Collection for " & temDays & " days";
                Printer.Print Tab(tabLiters - Len(Format(TotalMilk, "0.00"))); Format(TotalMilk, "0.00")
                
                AvgMilk = TotalMilk / temDays
                Printer.Print Tab(tabName); "Average Day Collection";
                Printer.Print Tab(tabLiters - Len(Format(AvgMilk, "0.00"))); Format(AvgMilk, "0.00")
                
                temCommsisionRate = OthersCommisionRate(temSupplierID, AvgMilk)
                Printer.Print Tab(tabName); "Commision Rate";
                Printer.Print Tab(tabLiters - Len(Format(temCommsisionRate, "0.00"))); Format(temCommsisionRate, "0.00")
                
                temCommsision = temCommsisionRate * TotalMilk
                Printer.Print Tab(tabName); "Commision";
                Printer.Print Tab(tabLiters - Len(Format(temCommsision, "0.00"))); Format(temCommsision, "0.00")
                rsCenter.MoveNext
                If rsCenter.EOF = False Then Printer.NewPage
            Wend
        End If
        Printer.EndDoc
        rsCenter.Close
    Else
        MsgBox "Printer Error"
    End If

End Sub

Private Sub btnPrintMilkPayments_Click()
'    Unload frmPrintMilkPayments
'    frmPrintMilkPayments.Show
'    frmPrintMilkPayments.ZOrder 0
'    frmPrintMilkPayments.cmbCollectingCenter.BoundText = cmbCollectingCenter.BoundText
'    frmPrintMilkPayments.dtpFrom.Value = dtpFrom.Value
'    frmPrintMilkPayments.dtpTo.Value = dtpTo.Value
'    With gridSummery
'        Dim i As Integer
'        For i = 1 To .Rows - 1
'            frmPrintMilkPayments.lstAll.AddItem .TextMatrix(i, 1)
'            frmPrintMilkPayments.lstAllIDs.AddItem .TextMatrix(i, 11)
'        Next
'    End With
End Sub

Private Sub btnPrintSummery_Click()
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
        
        tabNo = 5
        tabName = 10
        tabLMR = 60
        tabFAT = 70
        tabSNF = 80
        tabLiters = 90
        
        'tabMilk = 100
        
        tabCommision = 100
        tabDeduction = 110
        
        tabAverage = 120
        
        tabNetPayment = 130
        tabAverage = 130
        tabCenter = 45
        
        On Error Resume Next
        Printer.Orientation = cdlLandscape
        Printer.PrintQuality = vbPRPQHigh
        On Error GoTo 0
        
        Printer.Font.Name = "Arial Black"
        Printer.Font.Size = 14
        Printer.Font.Bold = True
        
        Printer.Print Tab(tabCenter - (Len(InstitutionName) / 2)); InstitutionName
        
        Printer.Font.Name = "Arial Black"
        Printer.Font.Size = 12
        Printer.Font.Bold = True

        
        Printer.Print Tab(tabCenter - (Len(InstitutionAddressLine1) / 2)); InstitutionAddressLine1
        Printer.Print Tab(tabCenter - (Len(InstitutionAddressLine2) / 2)); InstitutionAddressLine2
        Printer.Print Tab(tabCenter - (Len(InstitutionAddressLine3) / 2)); InstitutionAddressLine3
        
        
        Printer.Print Tab(tabNo); "Payment Summery "
        Printer.Print Tab(tabNo); "From " & Format(dtpFrom.Value, LongDateFormat) & " to " & Format(dtpTo.Value, LongDateFormat)
        
        Printer.Font.Name = "Verdana"
        Printer.Font.Size = 11
        Printer.Font.Bold = False
        Printer.Font.Underline = True
        
        
        Printer.Print
        Printer.Print Tab(tabNo); "No.";
        Printer.Print Tab(tabName); "Name";
        Printer.Print Tab(tabLMR - (Len("LMR"))); "LMR";
        Printer.Print Tab(tabFAT - (Len("Fat"))); "Fat";
        Printer.Print Tab(tabLiters - (Len("Vol"))); "Vol";
        Printer.Print Tab(tabSNF - (Len("SNF"))); "SNF";
        'Printer.Print Tab(tabMilk - (Len("Milk"))); "Milk";
        Printer.Print Tab(tabCommision - (Len("Comm"))); "Com.";
        Printer.Print Tab(tabDeduction - (Len("Ded"))); "Ded.";
        Printer.Print Tab(tabNetPayment - (Len("Net"))); "Net"
        
        Printer.Print Tab(tabAverage - (Len("Avg."))); "Avg."

        Printer.Font.Name = "Verdana"
        Printer.Font.Size = 11
        Printer.Font.Bold = False
        Printer.Font.Underline = False

        With gridSummery
            For i = 1 To .Rows - 1
                Printer.Print Tab(tabNo); .TextMatrix(i, 0);
                Printer.Print Tab(tabName); .TextMatrix(i, 1);
                Printer.Print Tab(tabLMR - Len(.TextMatrix(i, 3))); .TextMatrix(i, 3);
                Printer.Print Tab(tabFAT - Len(.TextMatrix(i, 4))); .TextMatrix(i, 4);
                Printer.Print Tab(tabLiters - Len(.TextMatrix(i, 5))); .TextMatrix(i, 5);
                Printer.Print Tab(tabSNF - Len(.TextMatrix(i, 6))); .TextMatrix(i, 6);
                'Printer.Print Tab(tabMilk - Len(.TextMatrix(i, 8))); .TextMatrix(i, 8);
                Printer.Print Tab(tabCommision - Len(.TextMatrix(i, 2))); .TextMatrix(i, 2);
                Printer.Print Tab(tabDeduction - Len(.TextMatrix(i, 9))); .TextMatrix(i, 9);
                
                Printer.Print Tab(tabAverage - Len(.TextMatrix(i, 7))); .TextMatrix(i, 7);
                
                Printer.Print Tab(tabNetPayment - Len(.TextMatrix(i, 10))); .TextMatrix(i, 10)
                
            Next
        End With
        
        
        Printer.Print Tab
        Printer.Print Tab(tabMilk); "Total Milk Payment";
        Printer.Print Tab(tabNetPayment - (Len(txtMilkPayments))); txtMilkPayments
        Printer.Print Tab(tabMilk); "Total Commision";
        Printer.Print Tab(tabNetPayment - (Len(txtCommision))); txtCommision
        Printer.Print Tab(tabMilk); "Total Deduction";
        Printer.Print Tab(tabNetPayment - (Len(txtDeductions))); txtDeductions
        Printer.Print Tab(tabMilk); "Total Payment";
        Printer.Print Tab(tabNetPayment - (Len(txtTotalPayment))); txtTotalPayment
        Printer.Print
        Printer.Print Tab(tabMilk); "Total Volume";
        Printer.Print Tab(tabNetPayment - (Len(txtVolume.Text))); txtVolume.Text
        
        Printer.Print
        Printer.Print Tab(tabMilk); "Average";
        Printer.Print Tab(tabNetPayment - (Len(txtVolume.Text))); txtVolume.Text
        
        Printer.Print
        Printer.Print
        
        Printer.Print ".....................                               .....................                   ....................."
        Printer.Print "Prepaired By                               Checked by                      Autherized by"
        Printer.Print
        
        
        Printer.EndDoc
    Else
        MsgBox "Printer Error"
    End If
End Sub

Private Sub btnProcess_Click()
    
    Dim rsAP As New ADODB.Recordset
    
'    With rsAP
'        If .State = 1 Then .Close
'        temSql = "Select * from tblAdditionalDeduction where DeductionDate between '" & dtpFrom.Value & "' AND '" & dtpTo.Value & "' AND Approved = 0 AND Deleted = 0 "
'        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'        If .RecordCount > 0 Then
'            MsgBox "There are some Additional Deductions awaiting for approval. Please approve them or delete before proceed"
'            .Close
'            Exit Sub
'        End If
'        .Close
'    End With

    With rsAP
        If .State = 1 Then .Close
        temSQL = "Select * from tblAdditionalCommision where CommisionDate between '" & dtpFrom.Value & "' AND '" & dtpTo.Value & "' AND Approved = 0 AND Deleted = 0 "
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            MsgBox "There are some Additional Commisions awaiting for approval. Please approve them or delete before proceed"
            .Close
            Exit Sub
        End If
        .Close
    End With
    
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
    MinusValue = False
    
    
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
            temSQL = "Select * from tblSupplier where Deleted = 0  Order by Supplier"
            rsCenter.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            
            If rsCenter.RecordCount > 0 Then
                rsCenter.MoveLast
                rsCenter.MoveFirst
                
                pgb1.Min = 0
                pgb1.Value = 0
                pgb1.Max = rsCenter.RecordCount
                
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
                    
                    
                    If MyMilkCollectionM.Value <> 0 Or temOwnCommision <> 0 Or MyMilkCollectionE.Value <> 0 Or temOthersCommision <> 0 Or SupplierDeduction <> 0 Or temOthersCommision <> 0 Or temAdditionalCommision <> 0 Then
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
                            
                            If MyMilkCollectionM.Value + MyMilkCollectionE.Value - SupplierDeduction + temOwnCommision + temOthersCommision + temAdditionalCommision < 0 Then
                                MinusValue = True
                                .CellBackColor = vbRed
                            End If
                            
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
                txtAverage.Text = Format(TotalMilkPayments / TotalVolume, "0.0000")
            Else
                txtAverage.Text = "0.00"
            End If
            
            If MinusValue = False Then
                btnProcess.Enabled = False
                btnSave.Enabled = True
            Else
                btnProcess.Enabled = True
                btnSave.Enabled = False
            End If
            pgb1.Value = 0
            Screen.MousePointer = vbDefault
End Sub

Private Sub btnSave_Click()
'    Dim CollectingCenterPaymentSummeryID As Long
'    If IsNumeric(cmbCollectingCenter.BoundText) = False Then
'        MsgBox "Please select a Collecting center"
'        cmbCollectingCenter.SetFocus
'        Exit Sub
'    End If
'    If btnProcess.Enabled = True Then
'        MsgBox "After changing Collecting Center or the dates, you must click the process button before saving"
'        btnProcess.SetFocus
'        Exit Sub
'    End If
'    Dim i As Integer
'    i = MsgBox("Are you sure you want to save the details?" & vbNewLine & "After saving you will not be able to change any detail.", vbYesNo, "Save?")
'    If i = vbNo Then Exit Sub
'    Dim rsCCMilkPayment As New ADODB.Recordset
'    With rsCCMilkPayment
'        If .State = 1 Then .Close
'        temSql = "SELECT * FROM tblCollectingCenterPaymentSummery"
'        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'        .AddNew
'        !CollectingCenterID = Val(cmbCollectingCenter.BoundText)
'        !FromDate = dtpFrom.Value
'        !ToDate = dtpTo.Value
'        !MilkPayment = Val(txtMilkPayments.Text)
'        !Commisions = Val(txtCommision.Text)
'        !Deductions = Val(txtDeductions.Text)
'        !TotalPayment = Val(txtTotalPayment.Text)
'        .Update
'        Dim ColletingCenterPaymentSummeryID As Long
'        CollectingCenterPaymentSummeryID = !CollectingCenterPaymentSummeryID
'        .Close
'    End With
'    txtCCPS.Text = CollectingCenterPaymentSummeryID
'    GeneratePayments (CollectingCenterPaymentSummeryID)
'    btnSave.Enabled = False
'    cmbCollectingCenter.Enabled = False
'    dtpFrom.Enabled = False
'    dtpTo.Enabled = False
'    btnProcess.Enabled = False
End Sub

Private Sub GeneratePayments(CollectingCenterPaymentSummeryID As Long)
    Dim i As Integer
    Dim temSupplierID As Long
    Dim temAmount As Double
    With gridSummery
        For i = 1 To .Rows - 1
            temAmount = Val(.TextMatrix(i, 10))
            temSupplierID = Val(.TextMatrix(i, 11))
            GenerateIndividualPayments CollectingCenterPaymentSummeryID, temSupplierID, temAmount
        Next
    End With
End Sub

Private Sub GenerateIndividualPayments(CollectingCenterPaymentSummeryID As Long, SupplierID As Long, PaymentValue As Double)
    Dim rsPayments As New ADODB.Recordset
    With rsPayments
        If .State = 1 Then .Close
        temSQL = "Select * from tblSupplierPayments where SupplierPaymentsID =0 "
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !SupplierID = SupplierID
        !CollectingCenterPaymentSummeryID = CollectingCenterPaymentSummeryID
        !FromDate = Format(dtpFrom.Value, "dd MMMM yyyy")
        !ToDate = Format(dtpTo.Value, "dd MMMM yyyy")
        !GeneratedPaymentMethodID = SupplierPaymentMethodID(SupplierID)
        !Generated = True
        !GeneratedUserID = UserID
        !GeneratedDate = Date
        !GeneratedTime = Time
        !Value = PaymentValue
        .Update
        .Close
    End With
End Sub


'Private Sub cmbCollectingCenter_Change()
'    If IsNumeric(cmbCollectingCenter.BoundText) = False Then Exit Sub
'    Dim rsCCSummery As New ADODB.Recordset
'    With rsCCSummery
'        If .State = 1 Then .Close
'        temSql = "SELECT Max(tblCollectingCenterPaymentSummery.ToDate) AS MaxOfToDate " & _
'                    "From tblCollectingCenterPaymentSummery " & _
'                    "WHERE tblCollectingCenterPaymentSummery.CollectingCenterID =" & Val(cmbCollectingCenter.BoundText)
'               .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If IsNull(!MaxOfToDate) = False Then
'            dtpFrom.Enabled = False
'            dtpFrom.Value = !MaxOfToDate + 1
'            dtpTo.Value = !MaxOfToDate + 15
'        Else
'            dtpFrom.Enabled = True
'            dtpFrom.Value = Date
'        End If
'    End With
'    btnProcess.Enabled = True
'End Sub

Private Sub cmbCommisionPrinter_Change()
    Call FillPapers(cmbCommisionPaper, cmbCommisionPrinter.Text)
End Sub

Private Sub cmbCommisionPrinter_Click()
    Call FillPapers(cmbCommisionPaper, cmbCommisionPrinter.Text)
End Sub

Private Sub cmbMilkPrinter_Change()
    Call FillPapers(cmbMilkPaper, cmbMilkPrinter.Text)
End Sub

Private Sub cmbMilkPrinter_Click()
    Call FillPapers(cmbMilkPaper, cmbMilkPrinter.Text)
End Sub

Private Sub cmbSummeryPrinter_Change()
    Call FillPapers(cmbSUmmeryPaper, cmbSummeryPrinter.Text)
End Sub

Private Sub cmbSummeryPrinter_Click()
    Call FillPapers(cmbSUmmeryPaper, cmbSummeryPrinter.Text)
End Sub

Private Sub dtpFrom_Change()
    If dtpFrom.Value >= dtpTo.Value Then
        dtpTo.Value = dtpFrom.Value + 1
    End If
    dtpTo.MinDate = dtpFrom.Value
End Sub


Private Sub dtpTo_Change()
    btnProcess.Enabled = True
End Sub

Private Sub Form_Load()
        
    If UserAuthorityLevel = SuperUser Or UserAuthorityLevel = Administrator Then
    
    Else
        btnSave.Visible = False
    End If
        
    Call FillCombos
    Call FillPrinters
    Call GetSettings
    dtpFrom.Value = Date
    dtpTo.Value = Date + 15
    pgb1.Value = 0
    
End Sub

Private Sub GetSettings()
    On Error Resume Next
    cmbSummeryPrinter.Text = GetSetting(App.EXEName, Me.Name, "SummeryPrinter", "")
    cmbMilkPrinter.Text = GetSetting(App.EXEName, Me.Name, "MilkPrinter", "")
    cmbCommisionPrinter.Text = GetSetting(App.EXEName, Me.Name, "CommisionPrinter", "")
    cmbSUmmeryPaper.Text = GetSetting(App.EXEName, Me.Name, "SummeryPaper", "")
    cmbMilkPaper.Text = GetSetting(App.EXEName, Me.Name, "MilkPaper", "")
    cmbCommisionPaper.Text = GetSetting(App.EXEName, Me.Name, "CommisionPaper", "")
End Sub

Private Sub FillPrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbSummeryPrinter.AddItem MyPrinter.DeviceName
        cmbMilkPrinter.AddItem MyPrinter.DeviceName
        cmbCommisionPrinter.AddItem MyPrinter.DeviceName
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


Private Sub FillCombos()
'    Dim Centers As New clsFillCombos
'    Centers.FillAnyCombo cmbCollectingCenter, "CollectingCenter", True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, Me.Name, "SummeryPrinter", cmbSummeryPrinter.Text
    SaveSetting App.EXEName, Me.Name, "SummeryPaper", cmbSUmmeryPaper.Text
    SaveSetting App.EXEName, Me.Name, "MilkPrinter", cmbMilkPrinter.Text
    SaveSetting App.EXEName, Me.Name, "MilkPaper", cmbMilkPaper.Text
    SaveSetting App.EXEName, Me.Name, "CommisionPrinter", cmbCommisionPrinter.Text
    SaveSetting App.EXEName, Me.Name, "CommisionPaper", cmbCommisionPaper.Text
End Sub

Private Sub gridSummery_Click()
'    Dim temSupplierID As Long
'    Dim temRow As Long
'    Dim temCol As Long
'
''    If btnProcess.Enabled = True Then
''        MsgBox "After changing any value, you have to click PROCESS button to go into details"
''        btnProcess.SetFocus
''        Exit Sub
''    End If
'
'    temRow = gridSummery.row
'    temCol = gridSummery.col
'
'    If IsNumeric(gridSummery.TextMatrix(temRow, 11)) = False Then Exit Sub
'    temSupplierID = Val(gridSummery.TextMatrix(temRow, 11))
'
'    Unload frmMilkPayAdviceDisplay
'
'    If Val(gridSummery.TextMatrix(temRow, 8)) <> 0 Then
'        With frmMilkPayAdviceDisplay
'            .Show
'            .cmbCollectingCenter.BoundText = Val(cmbCollectingCenter.BoundText)
'            .dtpFrom.Value = dtpFrom.Value
'            .dtpTo.Value = dtpTo.Value
'            .cmbSupplierName.BoundText = temSupplierID
'            .Top = 0
'            .Left = 0
'            .ZOrder 0
'        End With
'    End If
'
'    Unload frmDisplayCommision
'
'    If Val(gridSummery.TextMatrix(temRow, 2)) <> 0 Then
'        With frmDisplayCommision
'            .Show
'            .cmbCollectingCenter = Val(cmbCollectingCenter.BoundText)
'            .dtpFrom.Value = dtpFrom.Value
'            .dtpTo.Value = dtpTo.Value
'            .cmbSupplierName.BoundText = temSupplierID
'            .FillGrid
'            '.ZOrder 0
'        End With
'    End If
'
'    If temCol = 9 Or temCol = 10 Then
'        Unload frmAddAdditionalDeductions
'        frmAddAdditionalDeductions.Show
'        frmAddAdditionalDeductions.ZOrder 0
'        frmAddAdditionalDeductions.cmbCC.BoundText = cmbCollectingCenter.BoundText
'        frmAddAdditionalDeductions.cmbSupplier.BoundText = temSupplierID
'        frmAddAdditionalDeductions.dtpDate.Value = dtpFrom.Value
'        frmAddAdditionalDeductions.dtpFrom.Value = dtpFrom.Value
'        frmAddAdditionalDeductions.dtpTo.Value = dtpTo.Value
'
'        Unload frmAddDeductions
'        frmAddDeductions.Show
'        frmAddDeductions.ZOrder 0
'        frmAddDeductions.cmbCollectingCenter.BoundText = cmbCollectingCenter.BoundText
'        frmAddDeductions.cmbSupplierName.BoundText = temSupplierID
'        frmAddDeductions.dtpFrom.Value = dtpFrom.Value
'        frmAddDeductions.dtpTo.Value = dtpTo.Value
'    End If
    
End Sub

Private Sub SelectPrinter()
    Dim CSetPrinter As New cSetDfltPrinter
    CSetPrinter.SetPrinterAsDefault (BillPrinterName)
End Sub

Private Sub SelectPaper()
    Dim TemResponce As Long
    Dim RetVal As Integer
    RetVal = SelectForm(BillPaperName, Me.hwnd)
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
        Case FORM_SELECTED   ' 1
            Printer.Print
        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
    End Select
End Sub

Private Sub PrintSinglePayment(SupplierID As Long, CollectingCenterPaymentSummeryID As Long)
    Dim rsCCPS As New ADODB.Recordset
    Dim rsSupplier As New ADODB.Recordset
    Dim rsDay As New ADODB.Recordset
    
    Dim Name As String
    Dim Address As String
    Dim PayMonth As String
    Dim CodeNo As String
    Dim ReceiptNo As String
    
    Dim MilkPayments As Double
    Dim Commisions As Double
    Dim Deductions As Double
    Dim TotalPayment As Double
    
    With rsSupplier
        If .State = 1 Then .Close
        temSQL = "Select * from tblSupplier where supplierID = " & SupplierID
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Name = !Supplier
            If IsNull(!SupplierCode) = False Then
                CodeNo = !SupplierCode
            End If
            If IsNull(!Address) = False Then
                Address = !Address
            End If
            .Close
        Else
            .Close
            MsgBox "No Supplier Found Error. Please contact Lakmedipro"
            Exit Sub
        End If
    End With
    With rsCCPS
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblCOllectingCenterPaymentSummery where CollectingCenterPaymentSummeryID = " & CollectingCenterPaymentSummeryID
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            PayMonth = Format(!FromDate, "MMMM")
            .Close
        Else
            .Close
            MsgBox "No Collecting Center Summery ID Error. Please contact Lakmedipro"
            Exit Sub
        End If
    End With
    
    
    
    Dim NameX As Long
    Dim NameY As Long
    Dim AddressX As Long
    Dim AddressY As Long
    Dim MonthX As Long
    Dim MonthY As Long
    Dim CodeX As Long
    Dim CodeY As Long
    Dim ReceiptNoX As Long
    Dim ReceiptNoY As Long
        
    Dim ValueY As Long
    Dim LineHeight As Long

    Dim DateX As Long
    Dim DateCX As Long
    
    Dim LMRX As Long
    Dim LMRCX As Long
    
    Dim FATX As Long
    Dim FATCX As Long
    
    Dim LitersX As Long
    Dim LitersCX As Long
    
    Dim SNFX As Long
    Dim SNFCX As Long
    
    Dim PriceX As Long
    Dim PriceCX As Long
    
    Dim ValueX As Long
    Dim ValueCX As Long
    
    Dim CommisionX As Long
    
    Dim MilkPaymentY As Long
    Dim CommisionLabelX As Long
    Dim CommisionY As Long
    Dim DiductionLabelX As Long
    Dim DiductionY As Long
    Dim TotalPayY As Long



     NameX = 300
     NameY = 1000
     
     AddressX = 300
     AddressY = 1000
     
     MonthX = 900
     MonthY = 1000
     
     CodeX = 900
     CodeY = 1500
     
     ReceiptNoX = 500
     ReceiptNoY = 150
        
     ValueY = 200
     LineHeight = 50

     DateX = 10
     DateCX = 10
    
     LMRX = 50
     LMRCX = 50
    
     FATX = 75
     FATCX = 75
    
     LitersX = 100
     LitersCX = 100
    
     SNFX = 135
     SNFCX = 135
    
     PriceX = 145
     PriceCX = 145
    
     ValueX = 500
     ValueCX = 550
    
     CommisionX = 600
    
     MilkPaymentY = 600
     CommisionLabelX = 500
     CommisionY = 650
     DiductionLabelX = 500
     DiductionY = 700
     TotalPayY = 750
    
    Dim rsCenter As New ADODB.Recordset
    Dim Supplier As String
    Dim SupplierDeduction As Double
    Dim SupplierCommision As Double
    
    Dim MyMilkCollection As MilkCollection
            
    Printer.Font = "Arial"
    Printer.Font.Size = 12
    Printer.Font.Bold = True
    
    Printer.CurrentX = NameX
    Printer.CurrentY = NameY
    Printer.Print Name
            
    Printer.CurrentX = AddressX
    Printer.CurrentY = AddressY
    Printer.Print Address
    
    Printer.CurrentX = MonthX
    Printer.CurrentY = MonthY
    Printer.Print PayMonth
    
    Printer.CurrentX = CodeX
    Printer.CurrentY = CodeY
    Printer.Print CodeNo
    
    Printer.CurrentX = ValueX
    Printer.CurrentY = TotalPayY
    Printer.Print Format(TotalPayment, "#,##0.00")
    
'    If rsCenter.State = 1 Then rsCenter.Close
'    temSql = "Select * from tblSupplier where CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
'    rsCenter.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'    If rsCenter.RecordCount > 0 Then
'        While rsCenter.EOF = False
'            Supplier = rsCenter!Supplier
'            SupplierID = rsCenter!SupplierID
'
'
'            Dim i As Integer
'            Dim SecessionFAT As Double
'            Dim SecessionLMR As Double
'            Dim SecessionLiters As Double
'            Dim SecessionPrice As Double
'            Dim SecessionValue As Double
'            Dim SecessionCount As Long
'            Dim SecessionLMRXLiters As Double
'            Dim SecessionFATXLIters As Double
'
'            Dim AvgFAT As Double
'            Dim AvgLMR As Double
'            Dim AvgSNF As Double
'            Dim AvgLiters As Double
'            Dim AvgPrice As Double
'
'
'
'            Dim TemDays As Integer
'            Dim TemDay As Date
'
'            Dim SecessionMilk As MilkCollection
'
'            TemDays = DateDiff("d", dtpFrom.Value, dtpTo.Value)
'            TemDay = dtpFrom.Value
'
'            For i = 1 To TemDays + 1
'                With Printer
'                    MyMilkCollection = DailyMilkSupply(TemDay, SupplierID, 1)
'                    SupplierCommision = Commision(SupplierID, TemDay, TemDay)
'                    SupplierDeduction = PeriodDeductions(SupplierID, TemDay, TemDay)
'
'                    If SupplierCommision <> 0 Then
'
'                    .CurrentX = DateX
'                    .CurrentY = DateCX
'                    Printer.Print Format(TemDay, LongDateFormat)

'                If MyMilkCollection.Supplied = True Then
'                    Grid.TextMatrix(i, 1) = Format(MyMilkCollection.LMR, "0.00")
'                    Grid.TextMatrix(i, 2) = Format(MyMilkCollection.FAT, "0.00")
'                    Grid.TextMatrix(i, 3) = Format(MyMilkCollection.Liters, "0.000")
'                    Grid.TextMatrix(i, 4) = Format(MyMilkCollection.SNF, "0.00")
'                    Grid.TextMatrix(i, 5) = Format(MyMilkCollection.Price, "0.00")
'                    Grid.TextMatrix(i, 6) = Format(MyMilkCollection.Value, "0.00")
'                    SecessionFAT = SecessionFAT + MyMilkCollection.FAT
'                    SecessionFATXLIters = SecessionFATXLIters + (MyMilkCollection.FAT * MyMilkCollection.Liters)
'                    SecessionLiters = SecessionLiters + MyMilkCollection.Liters
'                    SecessionLMRXLiters = SecessionLMRXLiters + (MyMilkCollection.LMR * MyMilkCollection.Liters)
'                    SecessionLMR = SecessionLMR + MyMilkCollection.LMR
'                    SecessionPrice = SecessionPrice + MyMilkCollection.Price
'                    SecessionValue = SecessionValue + MyMilkCollection.Value
'                    SecessionCount = SecessionCount + 1
'                End If
'                TemDay = FromDate + i
'                    End If
'                End With
'            Next
'            With Grid
'                .TextMatrix(i + 1, 0) = "Total"
'                .TextMatrix(i + 1, 3) = Format(SecessionLiters, "0.000")
'                .TextMatrix(i + 1, 6) = Format(SecessionValue, "0.00")
'
'                .TextMatrix(i + 2, 0) = "Averages"
'                If SecessionLiters <= 0 Then
'                    SecessionMilk.Supplied = True
'                    AvgLiters = 0
'                    AvgFAT = 0
'                    AvgLiters = 0
'                    AvgSNF = 0
'                    AvgPrice = 0
'                    SecessionMilk.FAT = AvgFAT
'                    SecessionMilk.Liters = AvgLiters
'                    SecessionMilk.LMR = AvgLMR
'                    SecessionMilk.Price = AvgPrice
'                    SecessionMilk.SNF = AvgSNF
'                    SecessionMilk.Value = SecessionValue
'                Else
'                    SecessionMilk.Supplied = False
'                    AvgLMR = SecessionLMRXLiters / SecessionLiters
'                    AvgFAT = SecessionFATXLIters / SecessionLiters
'                    AvgLiters = SecessionLiters / SecessionCount
'                    AvgSNF = SNF(AvgLMR, AvgFAT)
'                    AvgPrice = Price(AvgFAT, AvgSNF)
'                    SecessionMilk.FAT = AvgFAT
'                    SecessionMilk.Liters = AvgLiters
'                    SecessionMilk.LMR = AvgLMR
'                    SecessionMilk.Price = AvgPrice
'                    SecessionMilk.SNF = AvgSNF
'                    SecessionMilk.Value = SecessionValue
'                End If
'                .TextMatrix(i + 2, 1) = Format(AvgLMR, "0.00")
'                .TextMatrix(i + 2, 2) = Format(AvgFAT, "0.00")
'                .TextMatrix(i + 2, 3) = Format(AvgLiters, "0.000")
'                .TextMatrix(i + 2, 4) = Format(AvgSNF, "0.00")
'                .TextMatrix(i + 2, 5) = Format(AvgPrice, "0.00")
'            End With
'
'
'
'
'
'
'
''
'
'            rsCenter.MoveNext
'        Wend
'    End If
'    rsCenter.Close

End Sub
