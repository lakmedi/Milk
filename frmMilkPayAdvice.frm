VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMilkPayAdvice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Milk Pay Advice"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11040
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
   ScaleHeight     =   9285
   ScaleWidth      =   11040
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   9600
      TabIndex        =   15
      Top             =   8520
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
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "Process"
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
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      TabIndex        =   9
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Caption         =   "Command1"
      Height          =   495
      Left            =   4920
      TabIndex        =   7
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "Morning"
      TabPicture(0)   =   "frmMilkPayAdvice.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Evening"
      TabPicture(1)   =   "frmMilkPayAdvice.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid2"
      Tab(1).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   5775
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   10186
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid Grid2 
         Height          =   5775
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   10186
         _Version        =   393216
      End
   End
   Begin MSDataListLib.DataCombo cmbSupplierName 
      Height          =   360
      Left            =   2040
      TabIndex        =   1
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
      TabIndex        =   2
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
      TabIndex        =   10
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMM yyyy"
      Format          =   57016323
      CurrentDate     =   39682
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMM yyyy"
      Format          =   57016323
      CurrentDate     =   39682
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Farmer &Name"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmMilkPayAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Centre As New clsFind
    Dim Supplier As New clsFind
    Dim rsMil As New ADODB.Recordset
    Dim temSql As String
    

Private Sub btnProcess_Click()
    FormatGrid Grid1
    FormatGrid Grid2
    
    Dim temSupplierID As Long
    Dim MorningCollection As MilkCollection
    Dim EveningCollection As MilkCollection
    Dim TotalCollection As MilkCollection
    
    temSupplierID = Val(cmbSupplierName.BoundText)
    
    
    MorningCollection = FillGrid(temSupplierID, 1, Grid1, dtpFrom.Value, dtpTo.Value)
    EveningCollection = FillGrid(temSupplierID, 2, Grid2, dtpFrom.Value, dtpTo.Value)

End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub cmbCollectingCenter_Change()
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then Exit Sub
    Supplier.FillCombo cmbSupplierName, "tblSupplier", "Supplier", "SupplierID", True, "CollectingCenterID", Val(cmbCollectingCenter.BoundText)
End Sub

Private Sub Command1_Click()
    Grid1.ColWidth(Val(Text1.Text)) = Val(Text2.Text)
End Sub


Private Sub Form_Load()
    If Day(Date) > 15 Then
        dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dtpTo.Value = DateSerial(Year(Date), Month(Date), 15)
    Else
    
    End If
    Call FillCombos
    Call FormatGrid(Grid1)
    Call FormatGrid(Grid2)
End Sub

Private Sub FillCombos()
    Centre.FillCombo cmbCollectingCenter, "tblCollectingCenter", "CollectingCenter", "CollectingCenterID", True
    Supplier.FillCombo cmbSupplierName, "tblSupplier", "Supplier", "SupplierID", True
End Sub

Private Function FillGrid(ByVal SupplierID As Long, ByVal SecessionID As Long, Grid As MSFlexGrid, ByVal FromDate As Date, ByVal ToDate As Date) As MilkCollection
    Dim i As Integer
    Dim SecessionFAT As Double
    Dim SecessionLMR As Double
    Dim SecessionLiters As Double
    Dim SecessionPrice As Double
    Dim SecessionValue As Double
    Dim SecessionCount As Long
    Dim SecessionLMRXLiters As Double
    Dim SecessionFATXLIters As Double
    
    Dim AvgFAT As Double
    Dim AvgLMR As Double
    Dim AvgSNF As Double
    Dim AvgLiters As Double
    Dim AvgPrice As Double
    
    
    
    Dim temDays As Integer
    Dim TemDay As Date
    Dim MyMilkCollection As MilkCollection
    
    Dim SecessionMilk As MilkCollection
    
    temDays = DateDiff("d", FromDate, ToDate) + 1
    TemDay = FromDate
    Grid.Rows = temDays + 5
    For i = 1 To temDays
        Grid.TextMatrix(i, 0) = Format(TemDay, ShortDateFormat)
        MyMilkCollection = DailyMilkSupply(TemDay, SupplierID, SecessionID)
        If MyMilkCollection.Supplied = True Then
            Grid.TextMatrix(i, 1) = Format(MyMilkCollection.LMR, "0.00")
            Grid.TextMatrix(i, 2) = Format(MyMilkCollection.FAT, "0.00")
            Grid.TextMatrix(i, 3) = Format(MyMilkCollection.Liters, "0.000")
            Grid.TextMatrix(i, 4) = Format(MyMilkCollection.SNF, "0.00")
            Grid.TextMatrix(i, 5) = Format(MyMilkCollection.Price, "0.00")
            Grid.TextMatrix(i, 6) = Format(MyMilkCollection.Value, "0.00")
            SecessionFAT = SecessionFAT + MyMilkCollection.FAT
            SecessionFATXLIters = SecessionFATXLIters + (MyMilkCollection.FAT * MyMilkCollection.Liters)
            SecessionLiters = SecessionLiters + MyMilkCollection.Liters
            SecessionLMRXLiters = SecessionLMRXLiters + (MyMilkCollection.LMR * MyMilkCollection.Liters)
            SecessionLMR = SecessionLMR + MyMilkCollection.LMR
            SecessionPrice = SecessionPrice + MyMilkCollection.Price
            SecessionValue = SecessionValue + MyMilkCollection.Value
            SecessionCount = SecessionCount + 1
        End If
        TemDay = FromDate + i
    Next
    With Grid
        .TextMatrix(i + 1, 0) = "Total"
        .TextMatrix(i + 1, 3) = Format(SecessionLiters, "0.000")
        .TextMatrix(i + 1, 6) = Format(SecessionValue, "0.00")
        
        .TextMatrix(i + 2, 0) = "Averages"
        If SecessionLiters <= 0 Then
            SecessionMilk.Supplied = True
            AvgLiters = 0
            AvgFAT = 0
            AvgLiters = 0
            AvgSNF = 0
            AvgPrice = 0
            SecessionMilk.FAT = AvgFAT
            SecessionMilk.Liters = AvgLiters
            SecessionMilk.LMR = AvgLMR
            SecessionMilk.Price = AvgPrice
            SecessionMilk.SNF = AvgSNF
            SecessionMilk.Value = SecessionValue
        Else
            SecessionMilk.Supplied = False
            AvgLMR = SecessionLMRXLiters / SecessionLiters
            AvgFAT = SecessionFATXLIters / SecessionLiters
            AvgLiters = SecessionLiters / SecessionCount
            AvgSNF = SNF(AvgLMR, AvgFAT)
            
            'Check
            AvgPrice = SecessionValue / SecessionLiters  '  Price(AvgFAT, AvgSNF, Val(cmbCollectingCenter.BoundText))
            'Check
            SecessionMilk.FAT = AvgFAT
            SecessionMilk.Liters = AvgLiters
            SecessionMilk.LMR = AvgLMR
            SecessionMilk.Price = AvgPrice
            SecessionMilk.SNF = AvgSNF
            SecessionMilk.Value = SecessionValue
        End If
        .TextMatrix(i + 2, 1) = Format(AvgLMR, "0.00")
        .TextMatrix(i + 2, 2) = Format(AvgFAT, "0.00")
        .TextMatrix(i + 2, 3) = Format(AvgLiters, "0.000")
        .TextMatrix(i + 2, 4) = Format(AvgSNF, "0.00")
        .TextMatrix(i + 2, 5) = Format(AvgPrice, "0.00")
    End With
End Function

Private Sub FormatGrid(Grid As MSFlexGrid)
    With Grid
        .Cols = 8
        .Rows = 1
        .row = 0
        .col = 0
        .CellAlignment = 4
        .Text = "Date"
        
        .col = 1
        .CellAlignment = 4
        .Text = "LMR"
        
        .col = 2
        .CellAlignment = 4
        .Text = "FAT%"
        
        .col = 3
        .CellAlignment = 4
        .Text = "Liters"
        
        .col = 4
        .CellAlignment = 4
        .Text = "SNF"
        
        .col = 5
        .CellAlignment = 4
        .Text = "Price"
        
        .col = 6
        .CellAlignment = 4
        .Text = "Value"
        
        .col = 7
        .CellAlignment = 4
        .Text = "ID"
    End With
End Sub
