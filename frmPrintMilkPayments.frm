VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmPrintMilkPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Milk Payments"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10050
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
   ScaleHeight     =   8640
   ScaleWidth      =   10050
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   7320
      Visible         =   0   'False
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ComboBox cmbMilkPaper 
      Height          =   360
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   8160
      Width           =   2895
   End
   Begin VB.ComboBox cmbMilkPrinter 
      Height          =   360
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   7680
      Width           =   2895
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8760
      TabIndex        =   16
      Top             =   8040
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
   Begin VB.ListBox lstPrintIDs 
      Height          =   5580
      Left            =   8520
      MultiSelect     =   2  'Extended
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstAllIDs 
      Height          =   5580
      Left            =   3840
      MultiSelect     =   2  'Extended
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstPrint 
      Height          =   5580
      Left            =   4920
      MultiSelect     =   2  'Extended
      TabIndex        =   7
      Top             =   1320
      Width           =   3975
   End
   Begin VB.ListBox lstAll 
      Height          =   5580
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   1320
      Width           =   3975
   End
   Begin MSDataListLib.DataCombo cmbCollectingCenter 
      Height          =   360
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   137691139
      CurrentDate     =   39682
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   137691139
      CurrentDate     =   39682
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   3360
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   ">"
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
   Begin btButtonEx.ButtonEx btnRemove 
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   3840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "<"
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
   Begin btButtonEx.ButtonEx btnCashAll 
      Height          =   255
      Left            =   7560
      TabIndex        =   10
      Top             =   6960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "All"
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
   Begin btButtonEx.ButtonEx btnCashNone 
      Height          =   255
      Left            =   8280
      TabIndex        =   11
      Top             =   6960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "None"
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
   Begin btButtonEx.ButtonEx btnAllAll 
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   6960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "All"
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
   Begin btButtonEx.ButtonEx btnAllNone 
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   6960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "None"
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
      Height          =   495
      Left            =   7440
      TabIndex        =   17
      Top             =   8040
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
   Begin btButtonEx.ButtonEx btnExcelM 
      Height          =   495
      Left            =   6120
      TabIndex        =   22
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Excel"
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
      Caption         =   "Paper"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   8160
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   7680
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&To"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&From"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Collecting Center"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmPrintMilkPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim CSetPrinter As New cSetDfltPrinter
    Dim temSQL As String
    Dim rsMilk As New ADODB.Recordset
    
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

    
    
Private Sub btnAdd_Click()
    Dim i As Integer
    If lstAll.ListCount = 0 Then Exit Sub
    For i = lstAll.ListCount - 1 To 0 Step -1
        If lstAll.Selected(i) = True Then
            lstPrint.AddItem lstAll.List(i)
            lstPrintIDs.AddItem lstAllIDs.List(i)
            lstAll.RemoveItem (i)
            lstAllIDs.RemoveItem (i)
        End If
    Next i
End Sub

Private Sub btnAllAll_Click()
    Dim i As Integer
    With lstAll
        For i = 0 To .ListCount - 1
            .Selected(i) = True
        Next
    End With
    With lstAllIDs
        For i = 0 To .ListCount - 1
            .Selected(i) = True
        Next
    End With
End Sub

Private Sub btnAllNone_Click()
    Dim i As Integer
    With lstAll
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next
    End With
    With lstAllIDs
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next
    End With
End Sub

Private Sub btnCashAll_Click()
    Dim i As Integer
    With lstPrint
        For i = 0 To .ListCount - 1
            .Selected(i) = True
        Next
    End With
    With lstPrintIDs
        For i = 0 To .ListCount - 1
            .Selected(i) = True
        Next
    End With

End Sub

Private Sub btnCashNone_Click()
    Dim i As Integer
    With lstPrint
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next
    End With
    With lstPrintIDs
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next
    End With
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcelM_Click()
    pb1.Visible = True
    MilkPaymentsToExcelM
    pb1.Visible = False
End Sub

Private Sub btnPrint_Click()
    Call PrintMilkPayments
End Sub


Private Sub MilkPaymentsToExcelM()

    Dim App As Excel.Application
    Dim WB As Excel.Workbook
    Dim WS As Excel.Worksheet
    Dim Chart As Excel.Chart
    
    Dim MyRow As Integer
    Dim MyCOl As Integer
    
    Dim sheetRows As Integer
    Dim temRow As Integer

    Set App = New Excel.Application
    Set WB = App.Workbooks.Add
    Set WS = WB.WorkSheets(1)
    App.Visible = True
    App.UserControl = True

    Dim TemTopic As String
    Dim temSubTopic As String
    Dim temSQL As String
    

    MyRow = 3
    
'    WS.Cells(myRow, 1).Value = "Name"
'    WS.Cells(myRow, 2).Value = "Unplanned Leave"


    Dim MySupplier As New clsSupplier
    Dim MyMilkCollection As MilkCollection
    Dim SecessionMilk As MilkCollection
    
    Dim SecessionFAT As Double
    Dim SecessionLMR As Double
    Dim SecessionLiters As Double
    Dim SecessionPrice As Double
    Dim SecessionValue As Double
    Dim SecessionCount As Long
    Dim SecessionLMRXLiters As Double
    Dim SecessionFATXLIters As Double
    Dim SecessionOwnCommision As Double
    
    Dim AdditionalCommision As Double
    Dim AdditionalDeductions As Double
    
    
    Dim OthersMilkCommision As Double
    Dim PureAdditionalCommision As Double
    
    Dim MorningOwnCommision As Double
    Dim EveningOwnCommision As Double
    
    Dim CattleFeedDeduction As Double
    Dim PureAdditionalDedduction As Double
    Dim VolumeDeductions As Double
    
    Dim MorningMilkValue As Double
    Dim EveningMilkValue As Double
    
    Dim temDeductions As Double
    Dim TemTotal As Double
    
    Dim FromDate As Date
    Dim ToDate As Date
    Dim temDays As Integer
    Dim TemDay As Date
    
    Dim NameX As Long
    Dim DateX As Long
    Dim LMRX As Long
    Dim FATX As Long
    Dim LitersX As Long
    Dim SNFX As Long
    Dim CommisionX As Long
    Dim CommisionRateX As Long
    Dim PriceX As Long
    Dim ValueX As Long
    
'    Dim NameY As Long
'    Dim AddressY As Long
'    Dim MonthY As Long
'    Dim CodeY As Long
'    Dim ReceiptY As Long
'    Dim RowY As Long
'    Dim LineHeight As Long
    
'    Dim GrossY As Long
'    Dim CommisionY As Long
'    Dim DeductionY As Long
'    Dim TotalY As Long
    
    Dim MilkCommisionX As Long
'    Dim MilkCommisionY As Long
'    Dim lblMilkCommisionX As Long
    
    Dim AdditionalCommisionX As Long
'    Dim AdditionalCommisionY As Long
'    Dim lblAdditionalCommisionX As Long
    
    Dim AdditionalDeductionX As Long
'    Dim AdditionalDeductionY As Long
'    Dim lblAdditionalDeductionX As Long
    
    Dim CattleFeedDeductionX As Long
'    Dim CattleFeedDeductionY As Long
'    Dim lblCattleFeedDeductionX As Long

    Dim VolumeDeductionX As Long
'    Dim VolumeDeductionY As Long
'    Dim lblVolumeFeedDeductionX As Long


    Dim NetPaymentX As Long
'    Dim NetPaymentY As Long
'    Dim lblNetPaymentX As Long

    Dim AvgFAT As Double
    Dim AvgLMR As Double
    Dim AvgSNF As Double
    Dim AvgLiters As Double
    Dim AvgPrice As Double
    Dim Date1X As Double
    Dim AddressX As Double
    Dim CodesX As Double
    
    Dim i As Integer
    Dim ii As Integer
    
    
    
'    CSetPrinter.SetPrinterAsDefault cmbMilkPrinter.Text
    
    MyRow = 1
    
        pb1.Min = 0
        pb1.Max = lstPrint.ListCount + 10

        
        For i = 0 To lstPrint.ListCount - 1
            
            pb1.Value = i
            
            MySupplier.ID = Val(lstPrintIDs.List(i))
            
            NameX = 2
            AddressX = 2
            
            Date1X = 1
            DateX = 1
            LMRX = 2
            FATX = 3
            LitersX = 4
            SNFX = 5
            CommisionRateX = 6
            PriceX = 7
            CommisionX = 8
            CodesX = 9
            ValueX = 10
            
            
            
            
            MilkCommisionX = 10

            
            AdditionalCommisionX = 10

            
            AdditionalDeductionX = 10

            
            CattleFeedDeductionX = 10

                                            
            VolumeDeductionX = 10


                
            NetPaymentX = 10

    
            SecessionFAT = 0
            SecessionFATXLIters = 0
            SecessionLiters = 0
            SecessionLMRXLiters = 0
            SecessionLMR = 0
            SecessionPrice = 0
            SecessionValue = 0
            SecessionCount = 0
            SecessionOwnCommision = 0
            
            
             OthersMilkCommision = 0
             PureAdditionalCommision = 0
             CattleFeedDeduction = 0
             PureAdditionalDedduction = 0
             VolumeDeductions = 0
            
            AdditionalCommision = 0
            AdditionalDeductions = 0
            
            MorningMilkValue = 0
            EveningMilkValue = 0
            
            FromDate = dtpFrom.Value
            ToDate = dtpTo.Value
            
            temDays = DateDiff("d", FromDate, ToDate) + 1
            TemDay = FromDate
            
            temRow = MyRow
            
            WS.Cells(MyRow, 1).Value = "Name"
            WS.Cells(MyRow, 2).Value = MySupplier.Name
            
            MyRow = MyRow + 1
            
            WS.Cells(MyRow, 1).Value = "Session"
            WS.Cells(MyRow, 2).Value = "Morning"
            
            WS.Cells(MyRow, 4).Value = "Month"
            WS.Cells(MyRow, 5).Value = Format(dtpFrom.Value, "MMMM")
            
            MyRow = MyRow + 1
            
            WS.Cells(MyRow, 1).Value = "Address"
            WS.Cells(MyRow, 2).Value = MySupplier.Address
            
            WS.Cells(MyRow, 4).Value = "Centre"
            WS.Cells(MyRow, 5).Value = cmbCollectingCenter.Text
            
            
            MyRow = MyRow + 1
            
            WS.Cells(MyRow, 1).Value = "Code"
            WS.Cells(MyRow, 2).Value = MySupplier.Code
            
            WS.Cells(MyRow, 4).Value = "A/C No"
            WS.Cells(MyRow, 5).Value = MySupplier.AccountNo
            
            Dim temMonth As String
            temMonth = Format(dtpFrom.Value, "MMMM")
            
            MyRow = MyRow + 1
            
            WS.Cells(MyRow, DateX).Value = "Date"
            WS.Cells(MyRow, LMRX).Value = "LMR"
            WS.Cells(MyRow, FATX).Value = "FAT"
            WS.Cells(MyRow, LitersX).Value = "Liters"
            WS.Cells(MyRow, SNFX).Value = "SNF"
            WS.Cells(MyRow, PriceX).Value = "Price"
            WS.Cells(MyRow, ValueX).Value = "Value"
            WS.Cells(MyRow, CommisionX).Value = "Commission"
            WS.Cells(MyRow, CommisionRateX).Value = "Com. Rate"
            WS.Cells(MyRow, LMRX).Value = "LMR"
            
            For ii = 1 To temDays
                MyMilkCollection = DailyMilkSupply(TemDay, MySupplier.ID, 1, True)
                MyRow = MyRow + 1
                WS.Cells(MyRow, DateX).Value = Format(TemDay, ShortDateFormat)
                If MyMilkCollection.Value > 0 Then
                    WS.Cells(MyRow, LMRX).Value = Format(MyMilkCollection.LMR, "0.00")
                    WS.Cells(MyRow, FATX).Value = Format(MyMilkCollection.FAT, "0.00")
                    WS.Cells(MyRow, LitersX).Value = Format(MyMilkCollection.Liters, "0.00")
                    WS.Cells(MyRow, SNFX).Value = Format(MyMilkCollection.SNF, "0.00")
                    WS.Cells(MyRow, PriceX).Value = Format(MyMilkCollection.Price, "0.00")
                    WS.Cells(MyRow, ValueX).Value = Format(MyMilkCollection.Value, "0.00")
                    If MyMilkCollection.OwnCommision <> 0 Then
                        WS.Cells(MyRow, CommisionX).Value = Format(MyMilkCollection.OwnCommision, "0.00")
                        WS.Cells(MyRow, CommisionRateX).Value = Format(MyMilkCollection.OwnCommisionRate, "0.00")
                    End If
                    
                    SecessionFAT = SecessionFAT + MyMilkCollection.FAT
                    SecessionFATXLIters = SecessionFATXLIters + (MyMilkCollection.FAT * MyMilkCollection.Liters)
                    SecessionLiters = SecessionLiters + MyMilkCollection.Liters
                    SecessionLMRXLiters = SecessionLMRXLiters + (MyMilkCollection.LMR * MyMilkCollection.Liters)
                    SecessionLMR = SecessionLMR + MyMilkCollection.LMR
                    SecessionPrice = SecessionPrice + MyMilkCollection.Price
                    SecessionValue = SecessionValue + MyMilkCollection.Value
                    SecessionOwnCommision = SecessionOwnCommision + MyMilkCollection.OwnCommision
                    SecessionCount = SecessionCount + 1
                
                    MorningOwnCommision = SecessionOwnCommision
                
                End If
                
                TemDay = FromDate + ii
                
            Next
            
            MorningMilkValue = SecessionValue + SecessionOwnCommision
            OthersMilkCommision = OthersCommision(MySupplier.ID, dtpFrom.Value, dtpTo.Value, 0)
            PureAdditionalCommision = PeriodAdditionalCommision(MySupplier.ID, dtpFrom.Value, dtpTo.Value)
            CattleFeedDeduction = PeriodCattleFeedDeductions(MySupplier.ID, dtpFrom.Value, dtpTo.Value)
            PureAdditionalDedduction = PeriodAdditionalDeductions(MySupplier.ID, dtpFrom.Value, dtpTo.Value)
            
            VolumeDeductions = PeriodVolumeDeductions(MySupplier.ID, dtpFrom.Value, dtpTo.Value)
            
            temDeductions = PeriodDeductions(MySupplier.ID, dtpFrom.Value, dtpTo.Value)
            
            MorningMilkValue = SecessionValue
            
            
            
            ' ***********************
            
'            Printer.CurrentX = ValueX - Printer.TextWidth(Format(SecessionValue, "0.00"))
'            Printer.CurrentY = GrossY
'            Printer.Print Format(SecessionValue, "0.00")

            ' ***********************
            
            MyRow = MyRow + 1
            
            sheetRows = MyRow - temRow

            WS.Cells(MyRow, LitersX).Value = Format(SecessionLiters, "0.00")
        
            
'            Printer.CurrentX = Date1X
'            'Printer.CurrentY = GrossY - 100
'            Printer.CurrentY = MonthY + 800
'            Printer.Print MySupplier.AccountNo
     
'            Printer.CurrentX = LitersX + (1.5 * 1440)
'            Printer.CurrentY = CommisionY - 50
'            Printer.Print "Commisions"
'
'            Printer.CurrentX = ValueX - Printer.TextWidth(Format(SecessionOwnCommision, "0.00"))
'            Printer.CurrentY = CommisionY
'            Printer.Print Format(SecessionOwnCommision, "0.00")
                
            If PeriodMilkSupply(dtpFrom.Value - 1, dtpTo.Value - 1, MySupplier.ID, 2).Liters <= 0 Then
                
'                Printer.CurrentX = LitersX + (1.5 * 1440)
'                Printer.CurrentY = DeductionY - 50
'                Printer.Print "Deductions"
'
'                Printer.CurrentX = ValueX - Printer.TextWidth(Format(TemDeductions, "0.00"))
'                Printer.CurrentY = DeductionY
'                Printer.Print Format(TemDeductions, "0.00")
                
                TemTotal = SecessionValue + SecessionOwnCommision ' - TemDeductions
                
                    
                WS.Cells(MyRow, 1).Value = "Milk Commision"
                WS.Cells(MyRow, 2).Value = Format(MorningOwnCommision, "0.00")
                
                WS.Cells(MyRow + 1, 1).Value = "Additional Commisions"
                WS.Cells(MyRow + 1, 2).Value = Format(PureAdditionalCommision, "0.00")
            
                WS.Cells(MyRow + 2, 1).Value = "Additional Deductions"
                WS.Cells(MyRow + 2, 2).Value = Format(PureAdditionalDedduction, "0.00")
                
                WS.Cells(MyRow + 3, 1).Value = "Vitamin & Cattle Feed Deductions"
                WS.Cells(MyRow + 3, 2).Value = Format(CattleFeedDeduction, "0.00")
                
                WS.Cells(MyRow + 4, 1).Value = "Welfare Deductions"
                WS.Cells(MyRow + 4, 2).Value = Format(VolumeDeductions, "0.00")
                
'                Printer.CurrentX = lblNetPaymentX
'                Printer.CurrentY = NetPaymentY
'                Printer.Print "Net Payment"
                'Printer.CurrentX = NetPaymentX - Printer.TextWidth(Format(SecessionValue + SecessionOwnCommision + PureAdditionalCommision + OthersMilkCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00"))
'                Printer.CurrentX = ValueX - Printer.TextWidth(Format(TemTotal, "0.00"))
'                Printer.CurrentY = TotalY
'                Printer.Print Format(TemTotal, "0.00")
                
                
                WS.Cells(MyRow, ValueX - 1).Value = "Gross"
                WS.Cells(MyRow, ValueX).Value = Format(SecessionValue, "0.00")
                
                WS.Cells(MyRow + 1, ValueX - 1).Value = "Commissions"
                WS.Cells(MyRow + 1, ValueX).Value = Format(MorningOwnCommision + EveningOwnCommision + PureAdditionalCommision, "0.00")
                
                WS.Cells(MyRow + 2, ValueX - 1).Value = "Deductions"
                WS.Cells(MyRow + 2, ValueX).Value = Format(AdditionalDeductions + CattleFeedDeduction + VolumeDeductions, "0.00")

                WS.Cells(MyRow + 3, ValueX - 1).Value = "Net"
                WS.Cells(MyRow + 3, ValueX).Value = Format(SecessionValue + SecessionOwnCommision + PureAdditionalCommision + OthersMilkCommision - PureAdditionalDedduction - CattleFeedDeduction - VolumeDeductions, "0.00")
'                Printer.Print Format(SecessionValue + MorningOwnCommision + PureAdditionalCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00")
            Else
            
            
'                Printer.Font.Name = "Verdana"
'                Printer.Font.Size = 11
'                Printer.Font.Bold = False
'
'                Printer.CurrentX = ValueX - Printer.TextWidth(Format(SecessionValue + SecessionOwnCommision + PureAdditionalCommision + OthersMilkCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00"))
'                Printer.CurrentY = NetPaymentY
'                Printer.Print Format(SecessionValue + SecessionOwnCommision + PureAdditionalCommision + OthersMilkCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00")
            
            
                TemTotal = SecessionValue ' + SecessionOwnCommision ' - TemDeductions
                
                WS.Cells(MyRow, 1).Value = "Total"
                WS.Cells(MyRow, 2).Value = Format(TemTotal, "0.00")
                
'                Printer.NewPage
                
                
                sheetRows = sheetRows + 6
                
                WS.Cells(MyRow + sheetRows, 1).Value = "Milk Commision"
 '               WS.Cells(MyRow + sheetRows, 2).Value = Format(MorningOwnCommision, "0.00")
                
                WS.Cells(MyRow + 1 + sheetRows, 1).Value = "Additional Commisions"
                WS.Cells(MyRow + 1 + sheetRows, 2).Value = Format(PureAdditionalCommision, "0.00")
            
                WS.Cells(MyRow + 2 + sheetRows, 1).Value = "Additional Deductions"
                WS.Cells(MyRow + 2 + sheetRows, 2).Value = Format(PureAdditionalDedduction, "0.00")
                
                WS.Cells(MyRow + 3 + sheetRows, 1).Value = "Vitamin & Cattle Feed Deductions"
                WS.Cells(MyRow + 3 + sheetRows, 2).Value = Format(CattleFeedDeduction, "0.00")
                
                WS.Cells(MyRow + 4 + sheetRows, 1).Value = "Welfare Deductions"
                WS.Cells(MyRow + 4 + sheetRows, 2).Value = Format(VolumeDeductions, "0.00")
                
                
 
'                Printer.CurrentX = MilkCommisionX - Printer.TextWidth(Format(MorningOwnCommision + EveningOwnCommision, "0.00"))
'                Printer.CurrentY = MilkCommisionY
'                Printer.Print Format(MorningOwnCommision + EveningOwnCommision, "0.00")
                
            End If
        
            SecessionFAT = 0
            SecessionFATXLIters = 0
            SecessionLiters = 0
            SecessionLMRXLiters = 0
            SecessionLMR = 0
            SecessionPrice = 0
            SecessionValue = 0
            SecessionCount = 0
            SecessionOwnCommision = 0
        
'            RowY = 2000
            

            If PeriodMilkSupply(dtpFrom.Value - 1, dtpTo.Value - 1, MySupplier.ID, 2).Liters > 0 Then
                
                
                MyRow = MyRow + 6
                
                
                
                FromDate = dtpFrom.Value - 1
                ToDate = dtpTo.Value - 1
                
                temDays = DateDiff("d", FromDate, ToDate) + 1
                TemDay = FromDate
                
                temRow = MyRow
                
                WS.Cells(MyRow, 1).Value = "Name"
                WS.Cells(MyRow, 2).Value = MySupplier.Name
                
                MyRow = MyRow + 1
                
                WS.Cells(MyRow, 1).Value = "Session"
                WS.Cells(MyRow, 2).Value = "Evening"
                
                WS.Cells(MyRow, 4).Value = "Month"
                WS.Cells(MyRow, 5).Value = Format(dtpFrom.Value, "MMMM")
                
                MyRow = MyRow + 1
                
                WS.Cells(MyRow, 1).Value = "Address"
                WS.Cells(MyRow, 2).Value = MySupplier.Address
                
                WS.Cells(MyRow, 4).Value = "Centre"
                WS.Cells(MyRow, 5).Value = cmbCollectingCenter.Text
                
                
                MyRow = MyRow + 1
                
                WS.Cells(MyRow, 1).Value = "Code"
                WS.Cells(MyRow, 2).Value = MySupplier.Code
                
                WS.Cells(MyRow, 4).Value = "A/C No"
                WS.Cells(MyRow, 5).Value = MySupplier.AccountNo
                
'                Dim temMonth As String
                temMonth = Format(dtpFrom.Value, "MMMM")
                
                MyRow = MyRow + 1
                
                WS.Cells(MyRow, DateX).Value = "Date"
                WS.Cells(MyRow, LMRX).Value = "LMR"
                WS.Cells(MyRow, FATX).Value = "FAT"
                WS.Cells(MyRow, LitersX).Value = "Liters"
                WS.Cells(MyRow, SNFX).Value = "SNF"
                WS.Cells(MyRow, PriceX).Value = "Price"
                WS.Cells(MyRow, ValueX).Value = "Value"
                WS.Cells(MyRow, CommisionX).Value = "Commission"
                WS.Cells(MyRow, CommisionRateX).Value = "Com. Rate"
                WS.Cells(MyRow, LMRX).Value = "LMR"
                
                For ii = 1 To temDays
                    
                    MyMilkCollection = DailyMilkSupply(TemDay, MySupplier.ID, 2, True)
                    
'                    RowY = RowY + LineHeight
                    
                    MyRow = MyRow + 1
                    WS.Cells(MyRow, DateX).Value = Format(TemDay, ShortDateFormat)
                    
                    If MyMilkCollection.Value <> 0 Then
                        WS.Cells(MyRow, LMRX).Value = Format(MyMilkCollection.LMR, "0.00")
                        WS.Cells(MyRow, FATX).Value = Format(MyMilkCollection.FAT, "0.00")
                        WS.Cells(MyRow, LitersX).Value = Format(MyMilkCollection.Liters, "0.00")
                        WS.Cells(MyRow, SNFX).Value = Format(MyMilkCollection.SNF, "0.00")
                        WS.Cells(MyRow, PriceX).Value = Format(MyMilkCollection.Price, "0.00")
                        WS.Cells(MyRow, ValueX).Value = Format(MyMilkCollection.Value, "0.00")
                        If MyMilkCollection.OwnCommision <> 0 Then
                            WS.Cells(MyRow, CommisionX).Value = Format(MyMilkCollection.OwnCommision, "0.00")
                            WS.Cells(MyRow, CommisionRateX).Value = Format(MyMilkCollection.OwnCommisionRate, "0.00")
                        End If
                    
                    
                    
'                        Printer.CurrentY = RowY
'                        Printer.CurrentX = LMRX - Printer.TextWidth(Format(MyMilkCollection.LMR, "0.00"))
'                        Printer.Print Format(MyMilkCollection.LMR, "0.00")
'                        Printer.CurrentY = RowY
'                        Printer.CurrentX = FATX - Printer.TextWidth(Format(MyMilkCollection.FAT, "0.00"))
'                        Printer.Print Format(MyMilkCollection.FAT, "0.00")
'                        Printer.CurrentY = RowY
'                        Printer.CurrentX = LitersX - Printer.TextWidth(Format(MyMilkCollection.Liters, "0.00"))
'                        Printer.Print Format(MyMilkCollection.Liters, "0.00")
'                        Printer.CurrentY = RowY
'                        Printer.CurrentX = SNFX - Printer.TextWidth(Format(MyMilkCollection.SNF, "0.00"))
'                        Printer.Print Format(MyMilkCollection.SNF, "0.00")
'                        Printer.CurrentY = RowY
'                        Printer.CurrentX = PriceX - Printer.TextWidth(Format(MyMilkCollection.Price, "0.00"))
'                        Printer.Print Format(MyMilkCollection.Price, "0.00")
'                        Printer.CurrentY = RowY
'                        Printer.CurrentX = ValueX - Printer.TextWidth(Format(MyMilkCollection.Value, "0.00"))
'                        Printer.Print Format(MyMilkCollection.Value, "0.00")
'                        Printer.CurrentY = RowY
'                        Printer.CurrentX = CommisionX
'                        Printer.Print Format(MyMilkCollection.OwnCommision, "0.00")
                    
                    End If
                    
                    SecessionFAT = SecessionFAT + MyMilkCollection.FAT
                    SecessionFATXLIters = SecessionFATXLIters + (MyMilkCollection.FAT * MyMilkCollection.Liters)
                    SecessionLiters = SecessionLiters + MyMilkCollection.Liters
                    SecessionLMRXLiters = SecessionLMRXLiters + (MyMilkCollection.LMR * MyMilkCollection.Liters)
                    SecessionLMR = SecessionLMR + MyMilkCollection.LMR
                    SecessionPrice = SecessionPrice + MyMilkCollection.Price
                    SecessionValue = SecessionValue + MyMilkCollection.Value
                    SecessionOwnCommision = SecessionOwnCommision + MyMilkCollection.OwnCommision
                    EveningOwnCommision = SecessionOwnCommision
                    SecessionCount = SecessionCount + 1
                    TemDay = TemDay + 1
                    
                Next
                
                EveningMilkValue = SecessionValue '+ SecessionOwnCommision
                
                MyRow = MyRow + 1
                
                WS.Cells(MyRow, LitersX).Value = Format(SecessionLiters, "0.00")
                
                WS.Cells(MyRow, ValueX - 1).Value = "Gross"
                WS.Cells(MyRow, ValueX).Value = Format(SecessionValue, "0.00")
                
                WS.Cells(MyRow + 1, ValueX - 1).Value = "Commissions"
                WS.Cells(MyRow + 1, ValueX).Value = Format(MorningOwnCommision + EveningOwnCommision + PureAdditionalCommision, "0.00")
                
                WS.Cells(MyRow + 2, ValueX - 1).Value = "Deductions"
                WS.Cells(MyRow + 2, ValueX).Value = Format(AdditionalDeductions + CattleFeedDeduction + VolumeDeductions, "0.00")

                
                
'                Printer.Font.Name = "Verdana"
'                Printer.Font.Size = 11
'                Printer.Font.Bold = False
'
'                Printer.CurrentX = ValueX - Printer.TextWidth(Format(SecessionValue, "0.00"))
'                Printer.CurrentY = GrossY
'                Printer.Print Format(SecessionValue, "0.00")
    
'                Printer.CurrentX = ValueX - Printer.TextWidth(Format(MorningOwnCommision + EveningOwnCommision + PureAdditionalCommision, "0.00"))
'                Printer.CurrentY = CommisionY
'                Printer.Print Format(MorningOwnCommision + EveningOwnCommision + PureAdditionalCommision, "0.00")
                
'                Printer.CurrentX = ValueX - Printer.TextWidth(Format(AdditionalDeductions + CattleFeedDeduction + VolumeDeductions, "0.00"))
'                Printer.CurrentY = DeductionY
'                Printer.Print Format(AdditionalDeductions + CattleFeedDeduction + VolumeDeductions, "0.00")
                
'                Printer.CurrentX = LitersX - Printer.TextWidth(Format(SecessionLiters, "0.00"))
'                Printer.CurrentY = GrossY
'                Printer.Print Format(SecessionLiters, "0.00")
                
                TemTotal = SecessionValue + SecessionOwnCommision
                
                WS.Cells(MyRow + 3, ValueX - 1).Value = "Net"
                WS.Cells(MyRow + 3, ValueX).Value = Format(MorningMilkValue + EveningMilkValue + MorningOwnCommision + EveningOwnCommision + PureAdditionalCommision - AdditionalDeductions - CattleFeedDeduction - VolumeDeductions, "0.00")
                
                
'                Printer.CurrentX = ValueX - Printer.TextWidth(Format(MorningMilkValue + EveningMilkValue + MorningOwnCommision + EveningOwnCommision + PureAdditionalCommision - AdditionalDeductions - CattleFeedDeduction - VolumeDeductions, "0.00"))
'                Printer.CurrentY = TotalY
'                Printer.Print Format(MorningMilkValue + EveningMilkValue + MorningOwnCommision + EveningOwnCommision + PureAdditionalCommision - AdditionalDeductions - CattleFeedDeduction - VolumeDeductions, "0.00")
                
                
'                Printer.Font.Name = "Verdana"
'                Printer.Font.Size = 11
'                Printer.Font.Bold = False
'
'                Printer.CurrentX = Date1X
'                'Printer.CurrentY = GrossY
'                Printer.CurrentY = MonthY + 800
'                Printer.Print MySupplier.AccountNo
'
'
'                Printer.CurrentX = SNFX + 1440 * 0.6
'                Printer.CurrentY = CommisionY
'                Printer.Print "Commisions"
'
'                Printer.CurrentX = SNFX + 1440 * 0.6
'                Printer.CurrentY = DeductionY
'                Printer.Print "Deductions"
'
'
'                Printer.Font.Name = "Verdana"
'                Printer.Font.Size = 8
'                Printer.Font.Bold = False
            
            
'                WS.Cells(MyRow, 1).Value = "Milk Commision"
'               WS.Cells(MyRow, 2).Value = Format(MorningOwnCommision + EveningOwnCommision, "0.00")
'
            
'                Printer.CurrentX = lblMilkCommisionX
'                Printer.CurrentY = MilkCommisionY
'                Printer.Print "Milk Commision"
'
'                Printer.CurrentX = MilkCommisionX - Printer.TextWidth(Format(MorningOwnCommision + EveningOwnCommision, "0.00"))
'                Printer.CurrentY = MilkCommisionY
'                Printer.Print Format(MorningOwnCommision + EveningOwnCommision, "0.00")
            
            
            
'                If PeriodMilkSupply(dtpFrom.Value - 1, dtpTo.Value - 1, MySupplier.ID, 2).Liters <= 0 Then
'
'                Else
'                    Printer.Font.Name = "Verdana"
'                    Printer.Font.Size = 8
'                    Printer.Font.Bold = False
'
'                    'Printer.CurrentX = NetPaymentX - Printer.TextWidth(Format(SecessionValue + SecessionOwnCommision + PureAdditionalCommision + OthersMilkCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00"))
'                    Printer.CurrentX = ValueX - Printer.TextWidth(Format(MorningMilkValue + MorningOwnCommision + EveningOwnCommision + SecessionValue + SecessionOwnCommision + PureAdditionalCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00"))
'                    Printer.CurrentY = NetPaymentY
'                    'Printer.Print Format(SecessionValue + SecessionOwnCommision + PureAdditionalCommision + OthersMilkCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00")
'                    Printer.Print Format(MorningMilkValue + MorningOwnCommision + EveningOwnCommision + SecessionValue + SecessionOwnCommision + PureAdditionalCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00")
'
'                    Printer.CurrentX = MilkCommisionX - Printer.TextWidth(Format(OthersMilkCommision, "0.00"))
'                    Printer.CurrentY = MilkCommisionY
'                    Printer.Print Format(EveningOwnCommision + MorningOwnCommision, "0.00")
'
'
'                End If
            
            
            End If
            If i < lstPrint.ListCount - 1 Then
                'Printer.NewPage
                MyRow = MyRow + 6
            End If
        
        
        Next i

    Set Chart = Nothing
    Set WS = Nothing
    Set WB = Nothing
    Set App = Nothing



End Sub




Private Sub PrintMilkPayments()

    Dim MySupplier As New clsSupplier
    Dim MyMilkCollection As MilkCollection
    Dim SecessionMilk As MilkCollection
    
    Dim SecessionFAT As Double
    Dim SecessionLMR As Double
    Dim SecessionLiters As Double
    Dim SecessionPrice As Double
    Dim SecessionValue As Double
    Dim SecessionCount As Long
    Dim SecessionLMRXLiters As Double
    Dim SecessionFATXLIters As Double
    Dim SecessionOwnCommision As Double
    
    Dim AdditionalCommision As Double
    Dim AdditionalDeductions As Double
    
    
    Dim OthersMilkCommision As Double
    Dim PureAdditionalCommision As Double
    
    Dim MorningOwnCommision As Double
    Dim EveningOwnCommision As Double
    
    Dim CattleFeedDeduction As Double
    Dim PureAdditionalDedduction As Double
    Dim VolumeDeductions As Double
    
    Dim MorningMilkValue As Double
    Dim EveningMilkValue As Double
    
    Dim temDeductions As Double
    Dim TemTotal As Double
    
    Dim FromDate As Date
    Dim ToDate As Date
    Dim temDays As Integer
    Dim TemDay As Date
    
    Dim NameX As Long
    Dim DateX As Long
    Dim LMRX As Long
    Dim FATX As Long
    Dim LitersX As Long
    Dim SNFX As Long
    Dim CommisionX As Long
    Dim CommisionRateX As Long
    Dim PriceX As Long
    Dim ValueX As Long
    
    Dim NameY As Long
    Dim AddressY As Long
    Dim MonthY As Long
    Dim CodeY As Long
    Dim ReceiptY As Long
    Dim RowY As Long
    Dim LineHeight As Long
    
    Dim GrossY As Long
    Dim CommisionY As Long
    Dim DeductionY As Long
    Dim TotalY As Long
    
    Dim MilkCommisionX As Long
    Dim MilkCommisionY As Long
    Dim lblMilkCommisionX As Long
    
    Dim AdditionalCommisionX As Long
    Dim AdditionalCommisionY As Long
    Dim lblAdditionalCommisionX As Long
    
    Dim AdditionalDeductionX As Long
    Dim AdditionalDeductionY As Long
    Dim lblAdditionalDeductionX As Long
    
    Dim CattleFeedDeductionX As Long
    Dim CattleFeedDeductionY As Long
    Dim lblCattleFeedDeductionX As Long

    Dim VolumeDeductionX As Long
    Dim VolumeDeductionY As Long
    Dim lblVolumeFeedDeductionX As Long


    Dim NetPaymentX As Long
    Dim NetPaymentY As Long
    Dim lblNetPaymentX As Long

    Dim AvgFAT As Double
    Dim AvgLMR As Double
    Dim AvgSNF As Double
    Dim AvgLiters As Double
    Dim AvgPrice As Double
    Dim Date1X As Double
    Dim AddressX As Double
    Dim CodesX As Double
    
    Dim i As Integer
    Dim ii As Integer
    
    CSetPrinter.SetPrinterAsDefault cmbMilkPrinter.Text
    
    If SelectForm(cmbMilkPaper.Text, Me.hwnd) = 1 Then
        
        For i = 0 To lstPrint.ListCount - 1
            
            MySupplier.ID = Val(lstPrintIDs.List(i))
            
            NameX = 2500
            
            Date1X = 100
            DateX = 400
            LMRX = 2800
            FATX = 4500
            LitersX = 5900
            SNFX = 7100
            CommisionRateX = 8300 - 50
            PriceX = 9150 - 50
            CommisionX = 9500
            CodesX = 11000
            ValueX = 11400
            AddressX = 75
            
            NameY = 1000 - 250
            AddressY = 1200 - 150
            MonthY = 800
            CodeY = 1300
            ReceiptY = 1600
            
            RowY = 2050
            
            MilkCommisionX = 4700
            MilkCommisionY = 6300
            lblMilkCommisionX = 500
            
            AdditionalCommisionX = 4700
            AdditionalCommisionY = 6500
            lblAdditionalCommisionX = 500
            
            AdditionalDeductionX = 4700
            AdditionalDeductionY = 6700
            lblAdditionalDeductionX = 500
            
            CattleFeedDeductionX = 4700
            CattleFeedDeductionY = 6900
            lblCattleFeedDeductionX = 500
                                            
            VolumeDeductionX = 4700
            VolumeDeductionY = 7100
            lblVolumeFeedDeductionX = 500

                
            NetPaymentX = 5000
            NetPaymentY = 7100
            lblNetPaymentX = 500
            
            LineHeight = 240
            
            GrossY = 6300
            CommisionY = 6550
            DeductionY = 6800
            TotalY = 7050
    
            SecessionFAT = 0
            SecessionFATXLIters = 0
            SecessionLiters = 0
            SecessionLMRXLiters = 0
            SecessionLMR = 0
            SecessionPrice = 0
            SecessionValue = 0
            SecessionCount = 0
            SecessionOwnCommision = 0
            
            
             OthersMilkCommision = 0
             PureAdditionalCommision = 0
             CattleFeedDeduction = 0
             PureAdditionalDedduction = 0
             VolumeDeductions = 0
            
            AdditionalCommision = 0
            AdditionalDeductions = 0
            
            MorningMilkValue = 0
            EveningMilkValue = 0
            
            FromDate = dtpFrom.Value
            ToDate = dtpTo.Value
            
            temDays = DateDiff("d", FromDate, ToDate) + 1
            TemDay = FromDate
            
            Printer.Font.Name = "Verdana"
            Printer.Font.Size = 14
            Printer.Font.Bold = False
            
            Printer.CurrentX = NameX - 1440 / 2
            Printer.CurrentY = NameY
            Printer.Print MySupplier.Name & " - Morning"
            
            
            Printer.Font.Name = "Verdana"
            Printer.Font.Size = 10
            Printer.Font.Bold = False
            
            Printer.CurrentX = AddressX
            Printer.CurrentY = AddressY
            Printer.Print MySupplier.Address
            
            Printer.Font.Name = "Verdana"
            Printer.Font.Size = 12
            Printer.Font.Bold = False
            
            
            Dim temMonth As String
            temMonth = Format(dtpFrom.Value, "MMMM")
            
            Printer.CurrentX = CodesX - Printer.TextWidth(cmbCollectingCenter.Text)
            Printer.CurrentY = MonthY
            Printer.Print temMonth
            
            Printer.CurrentX = CodesX - Printer.TextWidth(cmbCollectingCenter.Text)
            Printer.CurrentY = MonthY + 250
            Printer.Print MySupplier.Code
            
            Printer.CurrentX = CodesX - Printer.TextWidth(cmbCollectingCenter.Text)
            Printer.CurrentY = MonthY + 500
            Printer.Print cmbCollectingCenter.Text
            
            
            
            Printer.Font.Name = "Verdana"
            Printer.Font.Size = 10
            Printer.Font.Bold = False
            
            For ii = 1 To temDays
                
                
                MyMilkCollection = DailyMilkSupply(TemDay, MySupplier.ID, 1, True)
                
                RowY = RowY + LineHeight
                
                
                Printer.CurrentY = RowY
                Printer.CurrentX = DateX
                Printer.Print Format(TemDay, ShortDateFormat)
                
                If MyMilkCollection.Value > 0 Then
                
                    Printer.CurrentY = RowY
                    Printer.CurrentX = LMRX - Printer.TextWidth(Format(MyMilkCollection.LMR, "0.00"))
                    Printer.Print Format(MyMilkCollection.LMR, "0.00")
                    Printer.CurrentY = RowY
                    Printer.CurrentX = FATX - Printer.TextWidth(Format(MyMilkCollection.FAT, "0.00"))
                    Printer.Print Format(MyMilkCollection.FAT, "0.00")
                    Printer.CurrentY = RowY
                    Printer.CurrentX = LitersX - Printer.TextWidth(Format(MyMilkCollection.Liters, "0.00"))
                    Printer.Print Format(MyMilkCollection.Liters, "0.00")
                    Printer.CurrentY = RowY
                    Printer.CurrentX = SNFX - Printer.TextWidth(Format(MyMilkCollection.SNF, "0.00"))
                    Printer.Print Format(MyMilkCollection.SNF, "0.00")
                    Printer.CurrentY = RowY
                    Printer.CurrentX = PriceX - Printer.TextWidth(Format(MyMilkCollection.Price, "0.00"))
                    Printer.Print Format(MyMilkCollection.Price, "0.00")
                    Printer.CurrentY = RowY
                    Printer.CurrentX = ValueX - Printer.TextWidth(Format(MyMilkCollection.Value, "0.00"))
                    Printer.Print Format(MyMilkCollection.Value, "0.00")
                    If MyMilkCollection.OwnCommision <> 0 Then
                        Printer.CurrentY = RowY
                        Printer.CurrentX = CommisionX
                        Printer.Print Format(MyMilkCollection.OwnCommision, "0.00")
                        Printer.CurrentY = RowY
                        Printer.CurrentX = CommisionRateX - Printer.TextWidth(Format(MyMilkCollection.OwnCommisionRate, "0.00"))
                        Printer.Print Format(MyMilkCollection.OwnCommisionRate, "0.00")
                    End If
                    
                    SecessionFAT = SecessionFAT + MyMilkCollection.FAT
                    SecessionFATXLIters = SecessionFATXLIters + (MyMilkCollection.FAT * MyMilkCollection.Liters)
                    SecessionLiters = SecessionLiters + MyMilkCollection.Liters
                    SecessionLMRXLiters = SecessionLMRXLiters + (MyMilkCollection.LMR * MyMilkCollection.Liters)
                    SecessionLMR = SecessionLMR + MyMilkCollection.LMR
                    SecessionPrice = SecessionPrice + MyMilkCollection.Price
                    SecessionValue = SecessionValue + MyMilkCollection.Value
                    SecessionOwnCommision = SecessionOwnCommision + MyMilkCollection.OwnCommision
                    SecessionCount = SecessionCount + 1
                
                    MorningOwnCommision = SecessionOwnCommision
                
                End If
                
                TemDay = FromDate + ii
                
            Next
            
            MorningMilkValue = SecessionValue + SecessionOwnCommision
            OthersMilkCommision = OthersCommision(MySupplier.ID, dtpFrom.Value, dtpTo.Value, 0)
            PureAdditionalCommision = PeriodAdditionalCommision(MySupplier.ID, dtpFrom.Value, dtpTo.Value)
            CattleFeedDeduction = PeriodCattleFeedDeductions(MySupplier.ID, dtpFrom.Value, dtpTo.Value)
            PureAdditionalDedduction = PeriodAdditionalDeductions(MySupplier.ID, dtpFrom.Value, dtpTo.Value)
            
            VolumeDeductions = PeriodVolumeDeductions(MySupplier.ID, dtpFrom.Value, dtpTo.Value)
            
            temDeductions = PeriodDeductions(MySupplier.ID, dtpFrom.Value, dtpTo.Value)
            
            MorningMilkValue = SecessionValue
            
            Printer.Font.Name = "Verdana"
            Printer.Font.Size = 11
            Printer.Font.Bold = False
            
            
            ' ***********************
            
'            Printer.CurrentX = ValueX - Printer.TextWidth(Format(SecessionValue, "0.00"))
'            Printer.CurrentY = GrossY
'            Printer.Print Format(SecessionValue, "0.00")

            ' ***********************
            

            Printer.CurrentX = LitersX - Printer.TextWidth(Format(SecessionLiters, "0.00"))
            Printer.CurrentY = GrossY
            Printer.Print Format(SecessionLiters, "0.00")
        
            Printer.Font.Name = "Verdana"
            Printer.Font.Size = 11
            Printer.Font.Bold = False
            
            Printer.CurrentX = Date1X
            'Printer.CurrentY = GrossY - 100
            Printer.CurrentY = MonthY + 800
            Printer.Print MySupplier.AccountNo
     
'            Printer.CurrentX = LitersX + (1.5 * 1440)
'            Printer.CurrentY = CommisionY - 50
'            Printer.Print "Commisions"
'
'            Printer.CurrentX = ValueX - Printer.TextWidth(Format(SecessionOwnCommision, "0.00"))
'            Printer.CurrentY = CommisionY
'            Printer.Print Format(SecessionOwnCommision, "0.00")
                
            If PeriodMilkSupply(dtpFrom.Value - 1, dtpTo.Value - 1, MySupplier.ID, 2).Liters <= 0 Then
                
'                Printer.CurrentX = LitersX + (1.5 * 1440)
'                Printer.CurrentY = DeductionY - 50
'                Printer.Print "Deductions"
'
'                Printer.CurrentX = ValueX - Printer.TextWidth(Format(TemDeductions, "0.00"))
'                Printer.CurrentY = DeductionY
'                Printer.Print Format(TemDeductions, "0.00")
                
                TemTotal = SecessionValue + SecessionOwnCommision ' - TemDeductions
                
                    
                Printer.Font.Name = "Verdana"
                Printer.Font.Size = 8
                Printer.Font.Bold = False
            
                Printer.CurrentX = lblMilkCommisionX
                Printer.CurrentY = MilkCommisionY
                Printer.Print "Milk Commision"

                Printer.CurrentX = MilkCommisionX - Printer.TextWidth(Format(MorningOwnCommision, "0.00"))
                Printer.CurrentY = MilkCommisionY
                Printer.Print Format(MorningOwnCommision, "0.00")
                
                Printer.CurrentX = lblAdditionalCommisionX
                Printer.CurrentY = AdditionalCommisionY
                Printer.Print "Additional Commisions"
                
                Printer.CurrentX = AdditionalCommisionX - Printer.TextWidth(Format(PureAdditionalCommision, "0.00"))
                Printer.CurrentY = AdditionalCommisionY
                Printer.Print Format(PureAdditionalCommision, "0.00")
            
                Printer.CurrentX = lblAdditionalDeductionX
                Printer.CurrentY = AdditionalDeductionY
                Printer.Print "Additional Deductions"
                
                Printer.CurrentX = AdditionalDeductionX - Printer.TextWidth(Format(PureAdditionalDedduction, "0.00"))
                Printer.CurrentY = AdditionalDeductionY
                Printer.Print Format(PureAdditionalDedduction, "0.00")
                
                Printer.CurrentX = CattleFeedDeductionX - Printer.TextWidth(Format(CattleFeedDeduction, "0.00"))
                Printer.CurrentY = CattleFeedDeductionY
                Printer.Print Format(CattleFeedDeduction, "0.00")
                
                Printer.CurrentX = lblCattleFeedDeductionX
                Printer.CurrentY = CattleFeedDeductionY
                Printer.Print "Vitamin & Cattle Feed Deductions"
                
                Printer.CurrentX = VolumeDeductionX - Printer.TextWidth(Format(VolumeDeductions, "0.00"))
                Printer.CurrentY = VolumeDeductionY
                Printer.Print Format(VolumeDeductions, "0.00")
                
                Printer.CurrentX = lblVolumeFeedDeductionX
                Printer.CurrentY = VolumeDeductionY
                Printer.Print "Welfare Deductions"
                
                
'                Printer.CurrentX = lblNetPaymentX
'                Printer.CurrentY = NetPaymentY
'                Printer.Print "Net Payment"
'
                'Printer.CurrentX = NetPaymentX - Printer.TextWidth(Format(SecessionValue + SecessionOwnCommision + PureAdditionalCommision + OthersMilkCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00"))
                
                Printer.Font.Name = "Verdana"
                Printer.Font.Size = 11
                Printer.Font.Bold = False
                
                
                
                
'                Printer.CurrentX = ValueX - Printer.TextWidth(Format(TemTotal, "0.00"))
'                Printer.CurrentY = TotalY
'                Printer.Print Format(TemTotal, "0.00")
                
                
                


                Printer.CurrentX = ValueX - Printer.TextWidth(Format(MorningOwnCommision + EveningOwnCommision + PureAdditionalCommision, "0.00"))
                Printer.CurrentY = CommisionY
                Printer.Print Format(MorningOwnCommision + EveningOwnCommision + PureAdditionalCommision, "0.00")
                
                Printer.CurrentX = ValueX - Printer.TextWidth(Format(AdditionalDeductions + CattleFeedDeduction + VolumeDeductions, "0.00"))
                Printer.CurrentY = DeductionY
                Printer.Print Format(AdditionalDeductions + CattleFeedDeduction + VolumeDeductions, "0.00")

                Printer.CurrentX = ValueX - Printer.TextWidth(Format(SecessionValue + SecessionOwnCommision + PureAdditionalCommision + OthersMilkCommision - PureAdditionalDedduction - CattleFeedDeduction - VolumeDeductions, "0.00"))
                Printer.CurrentY = NetPaymentY
                Printer.Print Format(SecessionValue + SecessionOwnCommision + PureAdditionalCommision + OthersMilkCommision - PureAdditionalDedduction - CattleFeedDeduction - VolumeDeductions, "0.00")

                Printer.CurrentX = SNFX + 1440 * 0.6
                Printer.CurrentY = CommisionY
                Printer.Print "Commisions"
                
                Printer.CurrentX = SNFX + 1440 * 0.6
                Printer.CurrentY = DeductionY
                Printer.Print "Deductions"
                
                ' ****************************
                
                Printer.CurrentX = ValueX - Printer.TextWidth(Format(SecessionValue, "0.00"))
                Printer.CurrentY = GrossY
                Printer.Print Format(SecessionValue, "0.00")
                
                ' ******************************
                

'                Printer.Print Format(SecessionValue + MorningOwnCommision + PureAdditionalCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00")
            
            Else
            
            
'                Printer.Font.Name = "Verdana"
'                Printer.Font.Size = 11
'                Printer.Font.Bold = False
'
'                Printer.CurrentX = ValueX - Printer.TextWidth(Format(SecessionValue + SecessionOwnCommision + PureAdditionalCommision + OthersMilkCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00"))
'                Printer.CurrentY = NetPaymentY
'                Printer.Print Format(SecessionValue + SecessionOwnCommision + PureAdditionalCommision + OthersMilkCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00")
            
            
                TemTotal = SecessionValue ' + SecessionOwnCommision ' - TemDeductions
                Printer.CurrentX = ValueX - Printer.TextWidth(Format(TemTotal, "0.00"))
                Printer.CurrentY = TotalY
                Printer.Print Format(TemTotal, "0.00")
            
            
            
            
                Printer.NewPage
                Printer.Font.Name = "Verdana"
                Printer.Font.Size = 8
                Printer.Font.Bold = False
            
                Printer.CurrentX = lblMilkCommisionX
                Printer.CurrentY = MilkCommisionY
                Printer.Print "Milk Commision"

'                Printer.CurrentX = MilkCommisionX - Printer.TextWidth(Format(MorningOwnCommision + EveningOwnCommision, "0.00"))
'                Printer.CurrentY = MilkCommisionY
'                Printer.Print Format(MorningOwnCommision + EveningOwnCommision, "0.00")

                Printer.CurrentX = lblAdditionalCommisionX
                Printer.CurrentY = AdditionalCommisionY
                Printer.Print "Additional Commisions"
                
                Printer.CurrentX = AdditionalCommisionX - Printer.TextWidth(Format(PureAdditionalCommision, "0.00"))
                Printer.CurrentY = AdditionalCommisionY
                Printer.Print Format(PureAdditionalCommision, "0.00")
            
                Printer.CurrentX = lblAdditionalDeductionX
                Printer.CurrentY = AdditionalDeductionY
                Printer.Print "Additional Deductions"
                
                Printer.CurrentX = AdditionalDeductionX - Printer.TextWidth(Format(PureAdditionalDedduction, "0.00"))
                Printer.CurrentY = AdditionalDeductionY
                Printer.Print Format(PureAdditionalDedduction, "0.00")
                
                Printer.CurrentX = CattleFeedDeductionX - Printer.TextWidth(Format(CattleFeedDeduction, "0.00"))
                Printer.CurrentY = CattleFeedDeductionY
                Printer.Print Format(CattleFeedDeduction, "0.00")
                
                Printer.CurrentX = lblCattleFeedDeductionX
                Printer.CurrentY = CattleFeedDeductionY
                Printer.Print "Vitamin & Cattle Feed Deductions"
                
                
                Printer.CurrentX = VolumeDeductionX - Printer.TextWidth(Format(VolumeDeductions, "0.00"))
                Printer.CurrentY = VolumeDeductionY
                Printer.Print Format(VolumeDeductions, "0.00")
                
                Printer.CurrentX = lblVolumeFeedDeductionX
                Printer.CurrentY = VolumeDeductionY
                Printer.Print "Welfare Deductions"
                
            End If
        
            SecessionFAT = 0
            SecessionFATXLIters = 0
            SecessionLiters = 0
            SecessionLMRXLiters = 0
            SecessionLMR = 0
            SecessionPrice = 0
            SecessionValue = 0
            SecessionCount = 0
            SecessionOwnCommision = 0
        
            RowY = 2000


            If PeriodMilkSupply(dtpFrom.Value - 1, dtpTo.Value - 1, MySupplier.ID, 2).Liters > 0 Then
                
                Printer.Font.Name = "Verdana"
                Printer.Font.Size = 14
                Printer.Font.Bold = False
                
                FromDate = dtpFrom.Value - 1
                ToDate = dtpTo.Value - 1
                
                temDays = DateDiff("d", FromDate, ToDate) + 1
                TemDay = FromDate
                
                Printer.CurrentX = NameX
                Printer.CurrentY = NameY
                Printer.Print MySupplier.Name & " - Evening"
                
                Printer.Font.Name = "Verdana"
                Printer.Font.Size = 10
                Printer.Font.Bold = False
                
                
                Printer.CurrentX = AddressX
                Printer.CurrentY = AddressY
                Printer.Print MySupplier.Address
                
                temMonth = Format(dtpFrom.Value, "MMMM")
                
                Printer.CurrentX = CodesX - Printer.TextWidth(cmbCollectingCenter.Text)
                Printer.CurrentY = MonthY
                Printer.Print temMonth
                
                Printer.CurrentX = CodesX - Printer.TextWidth(cmbCollectingCenter.Text)
                Printer.CurrentY = MonthY + 250
                Printer.Print MySupplier.Code
                
                Printer.CurrentX = CodesX - Printer.TextWidth(cmbCollectingCenter.Text)
                Printer.CurrentY = MonthY + 500
                Printer.Print cmbCollectingCenter.Text
                
                Printer.Font.Name = "Verdana"
                Printer.Font.Size = 10
                Printer.Font.Bold = False
                
                For ii = 1 To temDays
                    
                    MyMilkCollection = DailyMilkSupply(TemDay, MySupplier.ID, 2, True)
                    
                    RowY = RowY + LineHeight
                    
                    
                    Printer.CurrentY = RowY
                    Printer.CurrentX = DateX
                    Printer.Print Format(TemDay, ShortDateFormat)
                    
                    If MyMilkCollection.Value <> 0 Then
                        Printer.CurrentY = RowY
                        Printer.CurrentX = LMRX - Printer.TextWidth(Format(MyMilkCollection.LMR, "0.00"))
                        Printer.Print Format(MyMilkCollection.LMR, "0.00")
                        Printer.CurrentY = RowY
                        Printer.CurrentX = FATX - Printer.TextWidth(Format(MyMilkCollection.FAT, "0.00"))
                        Printer.Print Format(MyMilkCollection.FAT, "0.00")
                        Printer.CurrentY = RowY
                        Printer.CurrentX = LitersX - Printer.TextWidth(Format(MyMilkCollection.Liters, "0.00"))
                        Printer.Print Format(MyMilkCollection.Liters, "0.00")
                        Printer.CurrentY = RowY
                        Printer.CurrentX = SNFX - Printer.TextWidth(Format(MyMilkCollection.SNF, "0.00"))
                        Printer.Print Format(MyMilkCollection.SNF, "0.00")
                        Printer.CurrentY = RowY
                        Printer.CurrentX = PriceX - Printer.TextWidth(Format(MyMilkCollection.Price, "0.00"))
                        Printer.Print Format(MyMilkCollection.Price, "0.00")
                        Printer.CurrentY = RowY
                        Printer.CurrentX = ValueX - Printer.TextWidth(Format(MyMilkCollection.Value, "0.00"))
                        Printer.Print Format(MyMilkCollection.Value, "0.00")
                        Printer.CurrentY = RowY
                        Printer.CurrentX = CommisionX
                        Printer.Print Format(MyMilkCollection.OwnCommision, "0.00")
                    
                    End If
                    
                    SecessionFAT = SecessionFAT + MyMilkCollection.FAT
                    SecessionFATXLIters = SecessionFATXLIters + (MyMilkCollection.FAT * MyMilkCollection.Liters)
                    SecessionLiters = SecessionLiters + MyMilkCollection.Liters
                    SecessionLMRXLiters = SecessionLMRXLiters + (MyMilkCollection.LMR * MyMilkCollection.Liters)
                    SecessionLMR = SecessionLMR + MyMilkCollection.LMR
                    SecessionPrice = SecessionPrice + MyMilkCollection.Price
                    SecessionValue = SecessionValue + MyMilkCollection.Value
                    SecessionOwnCommision = SecessionOwnCommision + MyMilkCollection.OwnCommision
                    EveningOwnCommision = SecessionOwnCommision
                    SecessionCount = SecessionCount + 1
                    TemDay = TemDay + 1
                    
                Next
                
                EveningMilkValue = SecessionValue '+ SecessionOwnCommision
                
                Printer.Font.Name = "Verdana"
                Printer.Font.Size = 11
                Printer.Font.Bold = False
                
                Printer.CurrentX = ValueX - Printer.TextWidth(Format(SecessionValue, "0.00"))
                Printer.CurrentY = GrossY
                Printer.Print Format(SecessionValue, "0.00")
    
                Printer.CurrentX = ValueX - Printer.TextWidth(Format(MorningOwnCommision + EveningOwnCommision + PureAdditionalCommision, "0.00"))
                Printer.CurrentY = CommisionY
                Printer.Print Format(MorningOwnCommision + EveningOwnCommision + PureAdditionalCommision, "0.00")
                
                Printer.CurrentX = ValueX - Printer.TextWidth(Format(AdditionalDeductions + CattleFeedDeduction + VolumeDeductions, "0.00"))
                Printer.CurrentY = DeductionY
                Printer.Print Format(AdditionalDeductions + CattleFeedDeduction + VolumeDeductions, "0.00")
                
                Printer.CurrentX = LitersX - Printer.TextWidth(Format(SecessionLiters, "0.00"))
                Printer.CurrentY = GrossY
                Printer.Print Format(SecessionLiters, "0.00")
                
                TemTotal = SecessionValue + SecessionOwnCommision
                
                Printer.CurrentX = ValueX - Printer.TextWidth(Format(MorningMilkValue + EveningMilkValue + MorningOwnCommision + EveningOwnCommision + PureAdditionalCommision - AdditionalDeductions - CattleFeedDeduction - VolumeDeductions, "0.00"))
                Printer.CurrentY = TotalY
                Printer.Print Format(MorningMilkValue + EveningMilkValue + MorningOwnCommision + EveningOwnCommision + PureAdditionalCommision - AdditionalDeductions - CattleFeedDeduction - VolumeDeductions, "0.00")
                
                
                Printer.Font.Name = "Verdana"
                Printer.Font.Size = 11
                Printer.Font.Bold = False
            
                Printer.CurrentX = Date1X
                'Printer.CurrentY = GrossY
                Printer.CurrentY = MonthY + 800
                Printer.Print MySupplier.AccountNo
            
            
                Printer.CurrentX = SNFX + 1440 * 0.6
                Printer.CurrentY = CommisionY
                Printer.Print "Commisions"
                
                Printer.CurrentX = SNFX + 1440 * 0.6
                Printer.CurrentY = DeductionY
                Printer.Print "Deductions"
            
            
                Printer.Font.Name = "Verdana"
                Printer.Font.Size = 8
                Printer.Font.Bold = False
            
                Printer.CurrentX = lblMilkCommisionX
                Printer.CurrentY = MilkCommisionY
                Printer.Print "Milk Commision"

                Printer.CurrentX = MilkCommisionX - Printer.TextWidth(Format(MorningOwnCommision + EveningOwnCommision, "0.00"))
                Printer.CurrentY = MilkCommisionY
                Printer.Print Format(MorningOwnCommision + EveningOwnCommision, "0.00")
            
            
            
'                If PeriodMilkSupply(dtpFrom.Value - 1, dtpTo.Value - 1, MySupplier.ID, 2).Liters <= 0 Then
'
'                Else
'                    Printer.Font.Name = "Verdana"
'                    Printer.Font.Size = 8
'                    Printer.Font.Bold = False
'
'                    'Printer.CurrentX = NetPaymentX - Printer.TextWidth(Format(SecessionValue + SecessionOwnCommision + PureAdditionalCommision + OthersMilkCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00"))
'                    Printer.CurrentX = ValueX - Printer.TextWidth(Format(MorningMilkValue + MorningOwnCommision + EveningOwnCommision + SecessionValue + SecessionOwnCommision + PureAdditionalCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00"))
'                    Printer.CurrentY = NetPaymentY
'                    'Printer.Print Format(SecessionValue + SecessionOwnCommision + PureAdditionalCommision + OthersMilkCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00")
'                    Printer.Print Format(MorningMilkValue + MorningOwnCommision + EveningOwnCommision + SecessionValue + SecessionOwnCommision + PureAdditionalCommision - PureAdditionalDedduction - CattleFeedDeduction, "0.00")
'
'                    Printer.CurrentX = MilkCommisionX - Printer.TextWidth(Format(OthersMilkCommision, "0.00"))
'                    Printer.CurrentY = MilkCommisionY
'                    Printer.Print Format(EveningOwnCommision + MorningOwnCommision, "0.00")
'
'
'                End If
            
            
            End If
            If i < lstPrint.ListCount - 1 Then Printer.NewPage
        
        
        Next i
        Printer.EndDoc
    Else
        MsgBox "Printer Error"
    End If
End Sub




Private Sub btnRemove_Click()
    Dim i As Integer
    If lstPrint.ListCount = 0 Then Exit Sub
    For i = lstPrint.ListCount - 1 To 0 Step -1
        If lstPrint.Selected(i) = True Then
            lstAll.AddItem lstPrint.List(i)
            lstAllIDs.AddItem lstPrintIDs.List(i)
            lstPrint.RemoveItem (i)
            lstPrintIDs.RemoveItem (i)
        End If
    Next i
End Sub

Private Sub cmbMilkPrinter_Click()
    Call FillPapers(cmbMilkPaper, cmbMilkPrinter.Text)
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FillPrinters
    Call GetSettings
End Sub

Private Sub FillCombos()
    Dim Centers As New clsFillCombos
    Centers.FillAnyCombo cmbCollectingCenter, "CollectingCenter", True
End Sub

Private Sub FillPrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbMilkPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub cmbMilkPrinter_Change()
    Call FillPapers(cmbMilkPaper, cmbMilkPrinter.Text)
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
    SaveSetting App.EXEName, Me.Name, "MilkPrinter", cmbMilkPrinter.Text
    SaveSetting App.EXEName, Me.Name, "MilkPaper", cmbMilkPaper.Text
End Sub

Private Sub GetSettings()
    On Error Resume Next
    cmbMilkPrinter.Text = GetSetting(App.EXEName, Me.Name, "MilkPrinter", "")
    cmbMilkPaper.Text = GetSetting(App.EXEName, Me.Name, "MilkPaper", "")
End Sub
