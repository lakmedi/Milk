VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmPrintCollectingCentrePaymentSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Collecting Center Payment Advice"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15000
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
   Picture         =   "frmPrintCollectingCentrePaymentSummery.frx":0000
   ScaleHeight     =   9915
   ScaleWidth      =   15000
   Begin VB.TextBox txtAverageSNF 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   13200
      TabIndex        =   42
      Top             =   9360
      Width           =   1575
   End
   Begin VB.TextBox txtAverageLMR 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   13200
      TabIndex        =   39
      Top             =   8400
      Width           =   1575
   End
   Begin VB.TextBox txtAverageFAT 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   13200
      TabIndex        =   38
      Top             =   8880
      Width           =   1575
   End
   Begin VB.TextBox txtValues 
      Height          =   1455
      Left            =   11640
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton btnToText 
      Caption         =   "Command1"
      Height          =   375
      Left            =   10560
      TabIndex        =   36
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAverage 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   13200
      TabIndex        =   34
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox txtPath 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   120
      TabIndex        =   31
      Top             =   7440
      Width           =   4455
   End
   Begin VB.TextBox txtVolume 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   13200
      TabIndex        =   28
      Top             =   6960
      Width           =   1575
   End
   Begin VB.ComboBox cmbSUmmeryPaper 
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   8880
      Width           =   2655
   End
   Begin VB.ComboBox cmbMilkPrinter 
      Height          =   360
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   8400
      Width           =   2655
   End
   Begin VB.ComboBox cmbMilkPaper 
      Height          =   360
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   8880
      Width           =   2655
   End
   Begin VB.ComboBox cmbCommisionPrinter 
      Height          =   360
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   8400
      Width           =   2655
   End
   Begin VB.ComboBox cmbCommisionPaper 
      Height          =   360
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   8880
      Width           =   2655
   End
   Begin VB.ComboBox cmbSummeryPrinter 
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   8400
      Width           =   2655
   End
   Begin VB.TextBox txtCCPS 
      Height          =   495
      Left            =   9240
      TabIndex        =   20
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   600
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
      Left            =   9720
      TabIndex        =   19
      Top             =   8280
      Width           =   1935
   End
   Begin VB.TextBox txtDeductions 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   9720
      TabIndex        =   18
      Top             =   7845
      Width           =   1935
   End
   Begin VB.TextBox txtCommision 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   9720
      TabIndex        =   17
      Top             =   7395
      Width           =   1935
   End
   Begin VB.TextBox txtMilkPayments 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   9720
      TabIndex        =   16
      Top             =   6960
      Width           =   1935
   End
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   8400
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   7920
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
      TabIndex        =   3
      Top             =   7920
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
      Left            =   10440
      TabIndex        =   6
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
      Height          =   5775
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   10186
      _Version        =   393216
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   146210819
      CurrentDate     =   39682
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   145686531
      CurrentDate     =   39682
   End
   Begin btButtonEx.ButtonEx btnPrintCommisions 
      Height          =   375
      Left            =   5640
      TabIndex        =   21
      Top             =   7920
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
   Begin btButtonEx.ButtonEx btnTextFile 
      Height          =   375
      Left            =   5640
      TabIndex        =   30
      Top             =   7440
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16711935
      Caption         =   "Summery to Text File"
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
   Begin btButtonEx.ButtonEx btnChangePath 
      Height          =   375
      Left            =   4680
      TabIndex        =   32
      Top             =   7440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16711935
      Caption         =   "Change"
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
      TabIndex        =   33
      Top             =   9480
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Average SNF"
      Height          =   360
      Left            =   11880
      TabIndex        =   43
      Top             =   9360
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Average LMR"
      Height          =   360
      Left            =   11880
      TabIndex        =   41
      Top             =   8400
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Average FAT"
      Height          =   360
      Left            =   11880
      TabIndex        =   40
      Top             =   8880
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Average Price"
      Height          =   360
      Left            =   11880
      TabIndex        =   35
      Top             =   7440
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Vloume"
      Height          =   360
      Left            =   11880
      TabIndex        =   29
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Payment"
      Height          =   360
      Left            =   8400
      TabIndex        =   15
      Top             =   8280
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Deductions"
      Height          =   360
      Left            =   8400
      TabIndex        =   14
      Top             =   7845
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Commisions"
      Height          =   360
      Left            =   8400
      TabIndex        =   13
      Top             =   7440
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Milk Payments"
      Height          =   360
      Left            =   8400
      TabIndex        =   12
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Collecting Center"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&From"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&To"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmPrintCollectingCentrePaymentSummery"
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
    
    
Private Sub btnChangePath_Click()
         Dim lpIDList As Long
         Dim sBuffer As String
         Dim szTitle As String
         Dim tBrowseInfo As BrowseInfo
         szTitle = "Select Backup Directory"
         With tBrowseInfo
            .hWndOwner = Me.hwnd
            .lpszTitle = lstrcat(szTitle, "")
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
         End With
         lpIDList = SHBrowseForFolder(tBrowseInfo)
         If (lpIDList) Then
            sBuffer = Space(MAX_PATH)
            SHGetPathFromIDList lpIDList, sBuffer
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            txtPath.Text = sBuffer
         End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrintAll_Click()
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then
        MsgBox "Please select a Collecting center"
        cmbCollectingCenter.SetFocus
        Exit Sub
    End If
    If btnSave.Enabled = True Then
        MsgBox "You must save before taking print outs"
        btnSave.SetFocus
        Exit Sub
    End If
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
                    "Where (((tblSupplier.Supplier) Is Not Null) And ((tblSupplierThroughCollector.CollectingCenterID) = " & Val(cmbCollectingCenter.BoundText) & ") And ((tblSupplierThroughCollector.Deleted) = 0 ))"
        
        
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

    If IsNumeric(cmbCollectingCenter.BoundText) = False Then Exit Sub

    Unload frmPrintMilkPayments
    frmPrintMilkPayments.Show
    frmPrintMilkPayments.ZOrder 0
    frmPrintMilkPayments.cmbCollectingCenter.BoundText = cmbCollectingCenter.BoundText
    frmPrintMilkPayments.dtpFrom.Value = dtpFrom.Value
    frmPrintMilkPayments.dtpTo.Value = dtpTo.Value
    With gridSummery
        Dim i As Integer
        For i = 1 To .Rows - 1
            frmPrintMilkPayments.lstAll.AddItem .TextMatrix(i, 1)
            frmPrintMilkPayments.lstAllIDs.AddItem .TextMatrix(i, 11)
        Next
    End With
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
    
    DoEvents
    
    Dim temDays As Integer
    Dim TemDay As Date
    
    Dim i As Integer
    
    Dim rsCenter As New ADODB.Recordset
    
    Dim temNetPayment As Double
    
    Dim Supplier As String
    Dim SupplierID As Long
    Dim SupplierDeduction As Double
    
    Dim temLMR As Double
    Dim temFAT As Double
    Dim temLiters As Double
    Dim temValue As Double
    Dim temPrice As Double
    
    Dim TotalLMRXLiters As Double
    Dim TotalFATXLiters As Double
    
    
    Dim temOwnCommision As Double
    Dim temOthersCommision As Double
    Dim temAdditionalCommision As Double
    
    Dim TotalMilkPayments As Double
    Dim TotalCommisions As Double
    Dim TotalDeductions As Double
    Dim TotalNetPayment As Double
    Dim TotalVolume As Double
    
    Dim Diduction1 As Double
    Dim Diduction2 As Double
    Dim Diduction3 As Double
    
    Dim MyMilkCollectionM As MilkCollection
    Dim MyMilkCollectionE As MilkCollection
    
    temDays = DateDiff("d", dtpFrom.Value, dtpTo.Value) + 1
    
'    If temDays > 31 Then
'        MsgBox "No more than 31 days can be processed"
'        dtpTo.SetFocus
'        Exit Sub
'    End If
        
    If temDays < 1 Then
        MsgBox "From and To dates are not correct"
        dtpTo.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass

        
    With gridSummery
        .Clear
        .Rows = 1
        .Cols = 15
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
        
        .col = 12
        .Text = "Cattle Feed Deductions"
        
        .col = 13
        .Text = "Additional Deductions"
        
        .col = 14
        .Text = "Volume Deductions"
        
        End With
            If rsCenter.State = 1 Then rsCenter.Close
            temSQL = "Select * from tblSupplier where CollectingCenterID = " & Val(cmbCollectingCenter.BoundText) & " And Deleted = 0  "
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
                    Supplier = rsCenter!Supplier & "(" & rsCenter!SupplierCode & ")"
                    SupplierID = rsCenter!SupplierID
                                            
                    pgb1.Value = pgb1.Value + 1
                    DoEvents
                        
'                    If Supplier = "D. Samarathunga" Then
'                        DoEvents
'                        DoEvents
'                    End If
                        
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
                    
                    Diduction1 = PeriodDeductionsComponent1(SupplierID, dtpFrom.Value, dtpTo.Value)
                    Diduction2 = PeriodDeductionsComponent2(SupplierID, dtpFrom.Value, dtpTo.Value)
                    Diduction3 = PeriodDeductionsComponent3(SupplierID, dtpFrom.Value, dtpTo.Value)
                    
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
                            
                            temNetPayment = temNetPayment + Val(.Text)
                            
                            .col = 12
                            .Text = Format(Diduction1, "0.00")
                            .col = 13
                            .Text = Format(Diduction2, "0.00")
                            .col = 14
                            .Text = Format(Diduction3, "0.00")
                            
                            
                            TotalLMRXLiters = TotalLMRXLiters + (MyMilkCollectionE.LMR * MyMilkCollectionE.Liters) + (MyMilkCollectionM.LMR * MyMilkCollectionM.Liters)
                            TotalFATXLiters = TotalFATXLiters + (MyMilkCollectionE.FAT * MyMilkCollectionE.Liters) + (MyMilkCollectionM.FAT * MyMilkCollectionM.Liters)
                            
                            
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
            txtTotalPayment.Text = Format(temNetPayment, "0.00")
            
            
            If TotalVolume <> 0 Then
                txtAverage.Text = Format(TotalMilkPayments / TotalVolume, "0.00")
            Else
                txtAverage.Text = "0.00"
            End If
            If TotalVolume <> 0 Then
            
            txtAverageLMR.Text = Format(TotalLMRXLiters / TotalVolume, "0.00")
            txtAverageFAT.Text = Format(TotalFATXLiters / TotalVolume, "0.00")
            
            End If
            txtAverageSNF.Text = Format(SNF(Val(txtAverageLMR.Text), Val(txtAverageFAT.Text)), "0.00")
            
            btnProcess.Enabled = False
            pgb1.Value = 0
            Screen.MousePointer = vbDefault
End Sub

Private Sub btnSave_Click()
    Dim CollectingCenterPaymentSummeryID As Long
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then
        MsgBox "Please select a Collecting center"
        cmbCollectingCenter.SetFocus
        Exit Sub
    End If
    If btnProcess.Enabled = True Then
        MsgBox "After changing Collecting Center or the dates, you must click the process button before saving"
        btnProcess.SetFocus
        Exit Sub
    End If
    Dim i As Integer
    i = MsgBox("Are you sure you want to save the details?" & vbNewLine & "After saving you will not be able to change any detail.", vbYesNo, "Save?")
    If i = vbNo Then Exit Sub
    Dim rsCCMilkPayment As New ADODB.Recordset
    With rsCCMilkPayment
        If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblCollectingCenterPaymentSummery where CollectingCentrePaymentSummeryID =0"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !CollectingCenterID = Val(cmbCollectingCenter.BoundText)
        !FromDate = dtpFrom.Value
        !ToDate = dtpTo.Value
        !MilkPayment = Val(txtMilkPayments.Text)
        !Commisions = Val(txtCommision.Text)
        !Deductions = Val(txtDeductions.Text)
        !TotalPayment = Val(txtTotalPayment.Text)
        .Update
        Dim ColletingCenterPaymentSummeryID As Long
        CollectingCenterPaymentSummeryID = !CollectingCenterPaymentSummeryID
        .Close
    End With
    txtCCPS.Text = CollectingCenterPaymentSummeryID
    GeneratePayments (CollectingCenterPaymentSummeryID)
    btnSave.Enabled = False
    cmbCollectingCenter.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    btnProcess.Enabled = False
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
        temSQL = "Select * from tblSupplierPayments where SupplierPaymentsID =0"
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


Private Sub btnTextFile_Click()
    Dim FSys As New FileSystemObject
    If FSys.FolderExists(txtPath.Text) = False Then
        MsgBox "Please select a valid folder to create a text file"
        btnChangePath.SetFocus
        Exit Sub
    End If
'    If IsNumeric(cmbCCPS.BoundText) = False Then
'        MsgBox "Please select a payment cycle"
'        cmbCCPS.SetFocus
'        Exit Sub
'    End If
    Dim temFile As String
    temFile = txtPath.Text & "\" & cmbCollectingCenter.Text & " From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy") & ".txt"
    While FSys.FileExists(temFile & ".txt") = True
        temFile = temFile & " "
    Wend
    temFile = temFile & ".txt"
    Open temFile For Append As #1
    
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
    
    
    Print #1, Tab(tabCenter - (Len(InstitutionName) / 2)); InstitutionName
    
    
    Print #1, Tab(tabCenter - (Len(InstitutionAddressLine1) / 2)); InstitutionAddressLine1
    Print #1, Tab(tabCenter - (Len(InstitutionAddressLine2) / 2)); InstitutionAddressLine2
    Print #1, Tab(tabCenter - (Len(InstitutionAddressLine3) / 2)); InstitutionAddressLine3
    
    
    Print #1, Tab(tabNo); "Payment Summery - " & cmbCollectingCenter.Text
    Print #1, Tab(tabNo); "From " & Format(dtpFrom.Value, LongDateFormat) & " to " & Format(dtpTo.Value, LongDateFormat)
    
    tabCenter = 45
    
    Print #1,
    Print #1, Tab(tabNo); "No.";
    Print #1, Tab(tabName); "Name";
    Print #1, Tab(tabLMR - (Len("LMR"))); "LMR";
    Print #1, Tab(tabFAT - (Len("Fat"))); "Fat";
    Print #1, Tab(tabSNF - (Len("SNF"))); "SNF";
    Print #1, Tab(tabLiters - (Len("Vol"))); "Vol";
    Print #1, Tab(tabCommision - (Len("Comm"))); "Com.";
    Print #1, Tab(tabDeduction - (Len("Ded"))); "Ded.";
    Print #1, Tab(tabAverage - (Len("Avg."))); "Avg.";
    Print #1, Tab(tabNetPayment - (Len("Net"))); "Net"
    

    Dim AllRows As Integer
    Dim PrintPageNo As Integer
    PrintPageNo = 1
    With gridSummery
        AllRows = 0
        For i = 1 To .Rows - 1
            AllRows = AllRows + 1
            Print #1, Tab(tabNo); .TextMatrix(i, 0);
            Print #1, Tab(tabName); .TextMatrix(i, 1);
            Print #1, Tab(tabLMR - Len(.TextMatrix(i, 3))); .TextMatrix(i, 3);
            Print #1, Tab(tabFAT - Len(.TextMatrix(i, 4))); .TextMatrix(i, 4);
            Print #1, Tab(tabSNF - Len(.TextMatrix(i, 6))); .TextMatrix(i, 6);
            Print #1, Tab(tabLiters - Len(.TextMatrix(i, 5))); .TextMatrix(i, 5);
            'Print #1,  Tab(tabMilk - Len(.TextMatrix(i, 8))); .TextMatrix(i, 8);
            Print #1, Tab(tabCommision - Len(.TextMatrix(i, 2))); .TextMatrix(i, 2);
            Print #1, Tab(tabDeduction - Len(.TextMatrix(i, 9))); .TextMatrix(i, 9);
            Print #1, Tab(tabAverage - Len(.TextMatrix(i, 7))); .TextMatrix(i, 7);
            Print #1, Tab(tabNetPayment - Len(.TextMatrix(i, 10))); .TextMatrix(i, 10)
        Next
    End With
    
    tabMilk = 10
    tabNetPayment = 100
    
    Print #1, Tab
    Print #1, Tab(tabMilk); "Total Milk Payment";
    Print #1, Tab(tabNetPayment - (Len(txtMilkPayments))); txtMilkPayments.Text
    Print #1, Tab(tabMilk); "Total Commision";
    Print #1, Tab(tabNetPayment - (Len(txtCommision))); txtCommision.Text
    Print #1, Tab(tabMilk); "Total Deduction";
    Print #1, Tab(tabNetPayment - (Len(txtDeductions))); txtDeductions.Text
    Print #1, Tab(tabMilk); "Total Payment";
    Print #1, Tab(tabNetPayment - (Len(txtTotalPayment))); txtTotalPayment.Text
    Print #1, ""
    Print #1, Tab(tabMilk); "Total Volume";
    Print #1, Tab(tabNetPayment - (Len(txtVolume.Text))); txtVolume.Text
    Print #1, ""
    Print #1, Tab(tabMilk); "Average";
    Print #1, Tab(tabNetPayment - (Len(txtAverage.Text))); txtAverage.Text
    
    Print #1,
    Print #1,
    
    Print #1, Tab(tabMilk); ".....................                           .....................                 ....................."
    Print #1, Tab(tabMilk); "Prepaired By                               Checked by                      Autherized by"
    Print #1,
    
    
    Close #1

End Sub

Private Sub cmbCCPS_Click(Area As Integer)
'    Dim rstemCCPS As New ADODB.Recordset
'    With rstemCCPS
'        If .State = 1 Then .Close
'        temSql = "Select * from tblCollectingCenterPaymentSummery where CollectingCenterPaymentSummeryID = " & Val(cmbCCPS.BoundText)
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            dtpFrom.Value = !FromDate
'            dtpTo.Value = !ToDate
'            .Close
'            btnProcess_Click
'        Else
'            .Close
'        End If
'    End With
End Sub

Private Sub btnToText_Click()
    Dim i As Integer
    txtValues.Text = Empty
    For i = 0 To gridSummery.Rows - 1
        txtValues.Text = txtValues.Text & vbNewLine & gridSummery.TextMatrix(i, 10)
    Next i
End Sub

Private Sub cmbCollectingCenter_Change()
'    If IsNumeric(cmbCollectingCenter.BoundText) = False Then Exit Sub
'    With rsCCPS
'        temSql = "SELECT convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) AS Display, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID " & _
'                    "FROM tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID " & _
'                    "Where (((tblCollectingCenterPaymentSummery.CollectingCenterID) = " & Val(cmbCollectingCenter.BoundText) & ")) " & _
'                    "ORDER BY tblCollectingCenterPaymentSummery.FromDate DESC"
'        If .State = 1 Then .Close
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'    With cmbCCPS
'        Set .RowSource = rsCCPS
'        .ListField = "Display"
'        .BoundColumn = "CollectingCenterPaymentSummeryID"
'    End With
End Sub

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
    Call FillCombos
    Call FillPrinters
    Call GetSettings
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = DateSerial(Year(Date), Month(Date), 1)
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
    txtPath.Text = GetSetting(App.EXEName, Me.Name, "txtPath", App.Path)
    
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
    Dim Centers As New clsFillCombos
    Centers.FillAnyCombo cmbCollectingCenter, "CollectingCenter", True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, Me.Name, "SummeryPrinter", cmbSummeryPrinter.Text
    SaveSetting App.EXEName, Me.Name, "SummeryPaper", cmbSUmmeryPaper.Text
    SaveSetting App.EXEName, Me.Name, "MilkPrinter", cmbMilkPrinter.Text
    SaveSetting App.EXEName, Me.Name, "MilkPaper", cmbMilkPaper.Text
    SaveSetting App.EXEName, Me.Name, "CommisionPrinter", cmbCommisionPrinter.Text
    SaveSetting App.EXEName, Me.Name, "CommisionPaper", cmbCommisionPaper.Text
    SaveSetting App.EXEName, Me.Name, "txtPath", txtPath.Text
End Sub

Private Sub gridSummery_Click()
    Dim temSupplierID As Long
    Dim temRow As Long
    
    If btnProcess.Enabled = True Then
        MsgBox "After changing any value, you have to click PROCESS button to go into details"
        btnProcess.SetFocus
        Exit Sub
    End If
    
    temRow = gridSummery.row
    
    If IsNumeric(gridSummery.TextMatrix(temRow, 11)) = False Then Exit Sub
    temSupplierID = Val(gridSummery.TextMatrix(temRow, 11))
    
    Unload frmMilkPayAdviceDisplay
    
    If Val(gridSummery.TextMatrix(temRow, 8)) <> 0 Then
        With frmMilkPayAdviceDisplay
            .Show
            .cmbCollectingCenter.BoundText = Val(cmbCollectingCenter.BoundText)
            .dtpFrom.Value = dtpFrom.Value
            .dtpTo.Value = dtpTo.Value
            .cmbSupplierName.BoundText = temSupplierID
            .Top = 0
            .Left = 0
            .ZOrder 0
        End With
    End If
    
    Unload frmDisplayCommision
    
    If Val(gridSummery.TextMatrix(temRow, 2)) <> 0 Then
        With frmDisplayCommision
            .Show
            .cmbCollectingCenter = Val(cmbCollectingCenter.BoundText)
            .dtpFrom.Value = dtpFrom.Value
            .dtpTo.Value = dtpTo.Value
            .cmbSupplierName.BoundText = temSupplierID
            .FillGrid
            '.ZOrder 0
        End With
    End If
    
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
