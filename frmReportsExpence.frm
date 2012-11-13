VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmReportsExpence 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expences"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
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
   ScaleHeight     =   7680
   ScaleWidth      =   6960
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
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
   Begin btButtonEx.ButtonEx bttnCreate 
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Graph"
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
   Begin btButtonEx.ButtonEx bttnReport 
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Report"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Analysis Details"
      TabPicture(0)   =   "frmReportsExpence.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dtpFrom"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dtpTO"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Graph Details"
      TabPicture(1)   =   "frmReportsExpence.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(5)=   "Frame11"
      Tab(1).Control(6)=   "Frame4"
      Tab(1).ControlCount=   7
      Begin VB.Frame Frame3 
         Height          =   3615
         Left            =   1320
         TabIndex        =   37
         Top             =   3120
         Width           =   5175
         Begin VB.ListBox lstCCIDs 
            Height          =   2700
            Left            =   4800
            TabIndex        =   39
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ListBox lstCC 
            Height          =   3030
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   38
            Top             =   240
            Width           =   4935
         End
      End
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -72120
         TabIndex        =   24
         Top             =   360
         Width           =   3615
         Begin VB.OptionButton optStandardChart 
            Caption         =   "Standared Chart type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Value           =   -1  'True
            Width           =   3375
         End
         Begin VB.OptionButton optOtherCharts 
            Caption         =   "Other Chart Types"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   3375
         End
         Begin VB.ComboBox cmbChartType 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1080
            Width           =   3135
         End
      End
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   21
         Top             =   4320
         Width           =   2535
         Begin VB.OptionButton optDisplayValues 
            Caption         =   "Display values"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optDoNotDisplayValues 
            Caption         =   "Do not display values"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   18
         Top             =   3120
         Width           =   2535
         Begin VB.OptionButton optNoTitle 
            Caption         =   "No title"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton optDisplayTitle 
            Caption         =   "Display title"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   15
         Top             =   1920
         Width           =   2535
         Begin VB.OptionButton opt2D 
            Caption         =   "2 D"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton opt3D 
            Caption         =   "3 D"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   1320
         TabIndex        =   14
         Top             =   1320
         Width           =   5175
         Begin VB.CheckBox chkTotal 
            Caption         =   "Total"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CheckBox chkOther 
            Caption         =   "Other Expences"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   2535
         End
         Begin VB.CheckBox chkCommisions 
            Caption         =   "Commisions"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox chkMilk 
            Caption         =   "Milk Payments"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -72120
         TabIndex        =   11
         Top             =   2040
         Width           =   3615
         Begin VB.OptionButton optYAxis 
            Caption         =   "Plot By Rows"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   3255
         End
         Begin VB.OptionButton optXAxis 
            Caption         =   "Plot By Colmns"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   2895
         End
      End
      Begin VB.Frame Frame11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -72120
         TabIndex        =   8
         Top             =   3240
         Width           =   3615
         Begin VB.OptionButton optDisplayLegend 
            Caption         =   "Display Ligend"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optNoLegend 
            Caption         =   "Do not display Legend"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   2535
         Begin VB.OptionButton optPie 
            Caption         =   "Pie Chart"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton optLine 
            Caption         =   "Line Chart"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton optBar 
            Caption         =   "Bar Chart"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin MSComCtl2.DTPicker dtpTO 
         Height          =   375
         Left            =   1320
         TabIndex        =   28
         Top             =   960
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dddd, dd MMMM yyyy"
         Format          =   285343747
         CurrentDate     =   39576
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   1320
         TabIndex        =   29
         Top             =   480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dddd, dd MMMM yyyy"
         Format          =   285343747
         CurrentDate     =   39576
      End
      Begin VB.Label Label3 
         Caption         =   "Collecting Center"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Calculate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmReportsExpence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim myworkbook As Excel.Workbook
    Dim myworksheet As Excel.Worksheet
    Dim mychart As Excel.Chart
    Dim tempath As String
    Dim FSys As New Scripting.FileSystemObject
    Dim rsViewAgent As New ADODB.Recordset
    Dim rsTemReport As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim temSQL As String
    Dim rsProduction As New ADODB.Recordset
    Dim rsViewItem As New ADODB.Recordset
    Dim rsAgent As New ADODB.Recordset
    Dim TemTopic As String
    Dim temSubTopic As String
    
Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnCreate_Click()
        
    Dim temValue As Double
    Dim i As Integer
    Dim tr As Integer
    Dim temDay1 As Date
    Dim temDay2 As Date
    Dim temday3 As Date
    Dim CCCount As Long
    Dim ii As Integer
    Dim CCIDs() As Long
    Dim CCs() As String
    Dim temnum1 As Long
    Dim Flag As Boolean
    
    If dtpFrom.Value > dtpTo.Value Then
        temDay1 = dtpTo.Value
        dtpTo.Value = dtpFrom.Value
        dtpFrom.Value = temDay1
    Else
        temDay1 = dtpFrom.Value
        temDay2 = dtpTo.Value
    End If
    
    If chkCommisions.Value = 0 And chkTotal.Value = 0 And chkMilk.Value = 0 And chkOther.Value = False Then
        tr = MsgBox("Please select at least one", vbCritical)
        chkMilk.SetFocus
        Exit Sub
    End If
    
    CCCount = 0
    
    For i = 0 To lstCCIDs.ListCount - 1
        If lstCC.Selected(i) = True Then
            CCCount = CCCount + 1
        End If
    Next
   
    ReDim CCIDs(CCCount) As Long
    ReDim CCs(CCCount) As String
    
    ii = 0
    For i = 0 To lstCCIDs.ListCount - 1
        If lstCC.Selected(i) = True Then
            CCIDs(ii) = Val(lstCCIDs.List(i))
            CCs(ii) = lstCC.List(i)
            ii = ii + 1
        End If
    Next
    
    
    tempath = App.Path
    If FSys.FileExists(tempath & "\Lucky1.xls") = False Then
        tr = MsgBox("There are no graphs on the specified location")
        Exit Sub
    End If
    
    frmPleaseWait.Show
    Me.MousePointer = vbHourglass
    DoEvents
    
    Set myworkbook = GetObject(tempath & "\Lucky1.xls")
    Set myworksheet = myworkbook.WorkSheets.Item(1)
    Set mychart = myworkbook.Charts.Item(1)
    
    myworksheet.UsedRange.Clear
    myworksheet.Cells(1, 1) = "Expences"
    
    Dim MyRow  As Integer
    Dim rsSupp As New ADODB.Recordset
    Dim TotalMilkPay As Double
    Dim TotalCommi As Double
    Dim TotalOther As Double
    Dim MyCOl As MilkCollection
    
    For temnum1 = 0 To CCCount - 1
        myworksheet.Cells(2, temnum1 + 2) = CCs(temnum1)
        MyRow = 2
        TotalMilkPay = 0
        TotalCommi = 0
        TotalOther = 0
        With rsSupp
            If .State = 1 Then .Close
            temSQL = "Select * from tblSupplier where Deleted = 0  AND CollectingCenterID = " & CCIDs(temnum1)
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            While .EOF = False
                MyCOl = PeriodMilkSupply(dtpFrom.Value, dtpTo.Value, !SupplierID, 0)
                TotalMilkPay = TotalMilkPay + MyCOl.Value ' + MyCOl.OwnCommision
                TotalCommi = TotalCommi + OthersCommision(!SupplierID, dtpFrom.Value, dtpTo.Value, 0) + MyCOl.OwnCommision
                .MoveNext
            Wend
            .Close
        End With
        
        Dim rsExpence As New ADODB.Recordset
        With rsExpence
            If .State = 1 Then .Close
            temSQL = "SELECT tblExpence.* From tblExpence where tblExpence.ExpenceDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "'  AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' And tblExpence.CollectingCenterID = " & CCIDs(temnum1) & " And tblExpence.Deleted = 0 "
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            While .EOF = False
                TotalOther = TotalOther + !ExpenceValue
                .MoveNext
            Wend
            .Close
        End With
        
        
        MyRow = 2
        If chkMilk.Value = 1 Then
            MyRow = MyRow + 1
            myworksheet.Cells(MyRow, temnum1 + 2) = TotalMilkPay
        End If
        If chkCommisions.Value = 1 Then
            MyRow = MyRow + 1
            myworksheet.Cells(MyRow, temnum1 + 2) = TotalCommi
        End If
        If chkOther.Value = 1 Then
            MyRow = MyRow + 1
            myworksheet.Cells(MyRow, temnum1 + 2) = TotalOther
        End If
        If chkTotal.Value = 1 Then
            MyRow = MyRow + 1
            myworksheet.Cells(MyRow, temnum1 + 2) = TotalMilkPay + TotalCommi + TotalOther
        End If
    Next
    
    MyRow = 2
    If chkMilk.Value = 1 Then
        MyRow = MyRow + 1
        myworksheet.Cells(MyRow, 1) = "Milk Payments"
    End If
    If chkCommisions.Value = 1 Then
        MyRow = MyRow + 1
        myworksheet.Cells(MyRow, 1) = "Commisions"
    End If
    If chkOther.Value = 1 Then
        MyRow = MyRow + 1
        myworksheet.Cells(MyRow, 1) = "Other Expences"
    End If
    If chkTotal.Value = 1 Then
        MyRow = MyRow + 1
        myworksheet.Cells(MyRow, 1) = "Total"
    End If
    
    
    mychart.SetSourceData myworksheet.Range("a2:" & GetColumnName(CCCount + 2) & MyRow)
    
    Call SetGraph
    If optDisplayTitle.Value = True Then
        TemTopic = "Expences From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " To " & Format(dtpTo.Value, "dd MMMM yyyy")
        mychart.HasTitle = True
        mychart.ChartTitle.Caption = TemTopic & vbNewLine & temSubTopic
    Else
        mychart.HasTitle = False
    End If
    
    mychart.Activate
    Unload frmPleaseWait
    Me.MousePointer = vbDefault
    frmGraph.Show
    frmGraph.Caption = TemTopic & " - " & temSubTopic

End Sub



Private Sub bttnReport_Click()
        
    Dim temValue As Double
    Dim i As Integer
    Dim tr As Integer
    Dim temDay1 As Date
    Dim temDay2 As Date
    Dim temday3 As Date
    Dim CCCount As Long
    Dim ii As Integer
    Dim CCIDs() As Long
    Dim CCs() As String
    Dim temnum1 As Long
    Dim Flag As Boolean
    
'    If optAll.Value = False And optStore.Value = False And optOutlet.Value = False And optAgent.Value = False And optRep.Value = False Then
'        tr = MsgBox("You have not selected the place", vbCritical, "Place?")
'        optStore.SetFocus
'        Exit Sub
'    End If
'    If optStore.Value = True And IsNumeric(dtcStore.BoundText) = False Then
'        tr = MsgBox("You have not selected the department", vbCritical, "Department")
'        dtcStore.SetFocus
'        Exit Sub
'    End If
'    If optAgent.Value = True And IsNumeric(dtcAgent.BoundText) = False Then
'        tr = MsgBox("You have not selected the agent", vbCritical, "agent")
'        dtcAgent.SetFocus
'        Exit Sub
'    End If
'    If optRep.Value = True And IsNumeric(dtcRep.BoundText) = False Then
'        tr = MsgBox("You have not selected the Rep", vbCritical, "Rep")
'        dtcRep.SetFocus
'        Exit Sub
'    End If
'    If optOutlet.Value = True And IsNumeric(dtcOutlet.BoundText) = False Then
'        tr = MsgBox("You have not selected the outlet", vbCritical, "Outlet")
'        dtcOutlet.SetFocus
'        Exit Sub
'    End If
'
'
'    If optAllItems.Value = False Then
'        Flag = False
'        CCCount = 0
'        For i = 0 To lstItemIDs.ListCount - 1
'            If lstItemIDs.Selected(i) = True Then
'                Flag = True
'                CCCount = CCCount + 1
'            End If
'        Next
'        If Flag = False Then
'            tr = MsgBox("You have not selected a Category", vbCritical, "Select Category")
'            optAllItems.SetFocus
'            Exit Sub
'        End If
'    Else
'        For i = 0 To lstItemIDs.ListCount - 1
'            lstItemIDs.Selected(i) = True
'            CCCount = CCCount + 1
'        Next
'    End If
'
'    Me.MousePointer = vbHourglass
'
'    ReDim CCIDs(CCCount) As Long
'    ReDim CCs(CCCount) As String
'    ii = 0
'
'    For i = 0 To lstItemIDs.ListCount - 1
'        If lstItemIDs.Selected(i) = True Then
'            lstItemIDs.ListIndex = i
'            lstItems.ListIndex = i
'            CCIDs(ii) = Val(lstItemIDs.Text)
'            CCs(ii) = lstItems.Text
'            ii = ii + 1
'        End If
'    Next
'
'    If dtpFrom.Value > dtpTO.Value Then
'        temDay1 = dtpTO.Value
'        dtpTO.Value = dtpFrom.Value
'        dtpFrom.Value = temDay1
'    Else
'        temDay1 = dtpFrom.Value
'        temDay2 = dtpTO.Value
'    End If
'
'    With rsTemReport
'        If .State = 1 Then .Close
'        temSQL = "Delete * from tblTemReport1 where UserID = " & UserID
'        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
'        If .State = 1 Then .Close
'        temSQL = "SELECT * from tblTemReport1"
'        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
'    End With
'
'
'    If rsTemReport.State = 1 Then rsTemReport.Close
'    temSQL = "SELECT * from tblTemReport1 where UserID = " & UserID
'    rsTemReport.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
'    rsTemReport.AddNew
'    rsTemReport!UserID = UserID
'    rsTemReport!txt1 = " "
'    rsTemReport!txt2 = "zz - Software By Lakmedipro  -  zz"
'    rsTemReport.Update
'    If rsTemReport.State = 1 Then rsTemReport.Close
'
'    temSQL = "SELECT * from tblTemReport1"
'    rsTemReport.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
'    i = i + 1
'    For temnum1 = 0 To CCCount - 1
'        If rsProduction.State = 1 Then rsProduction.Close
'        If optByQty.Value = True Then
'            temSQL = "SELECT Sum(Quentity) AS DisplayValue "
'        ElseIf optByVal.Value = True Then
'            temSQL = "SELECT Sum(Value) AS DisplayValue "
'        ElseIf optByVol.Value = True Then
'            temSQL = "SELECT Sum(Volume) AS DisplayValue "
'        End If
'        temSQL = temSQL & "FROM tblDiscard  "
'        temSQL = temSQL & "WHERE DOD Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTO.Value, "dd MMMM yyyy") & "' "
'        temSQL = temSQL & "AND CategoryID = " & CCIDs(temnum1)
'        If optAgent.Value = True Then
'            temSQL = temSQL & " AND AgentID = " & dtcAgent.BoundText
'        ElseIf optStore.Value = True Then
'            temSQL = temSQL & " And StoreID = " & dtcStore.BoundText
'        ElseIf optRep.Value = True Then
'            temSQL = temSQL & " AND RepID = " & dtcRep.BoundText
'        ElseIf optOutlet.Value = True Then
'            temSQL = temSQL & " And OutletID = " & dtcOutlet.BoundText
'        End If
'        rsProduction.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'        If rsProduction.RecordCount > 0 Then
'            If Not IsNull(rsProduction!DisplayValue) Then
'                rsTemReport.AddNew
'                rsTemReport!UserID = UserID
'                rsTemReport!txt1 = "Production"
'                rsTemReport!txt2 = CCs(temnum1)
'                rsTemReport!dbl1 = rsProduction!DisplayValue
'                rsTemReport.Update
'            End If
'        End If
'        If rsProduction.State = 1 Then rsProduction.Close
'    Next
'    If rsTemReport.State = 1 Then rsTemReport.Close
'    temSQL = "SELECT * from tblTemReport1 where UserID = " & UserID
'    rsTemReport.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
'    rsTemReport.AddNew
'    rsTemReport!UserID = UserID
'    rsTemReport!txt1 = " "
'    rsTemReport!txt2 = "zz - Software By Lakmedipro  -  zz"
'    rsTemReport.Update
'    If rsTemReport.State = 1 Then rsTemReport.Close
'    TemTopic = ""
'    If optByQty.Value = True Then
'        TemTopic = TemTopic & "Quentity-wise "
'    ElseIf optByQty.Value = True Then
'        TemTopic = TemTopic & "Value-wise "
'    End If
'    TemTopic = TemTopic & "Discard "
'    If optAllItems.Value = True Then
'        TemTopic = TemTopic & "of all Categories "
'    Else
'        TemTopic = TemTopic & "of Selected Categories "
'    End If
'    If dtpFrom.Value = dtpTO.Value Then
'        temSubTopic = "On " & Format(dtpFrom.Value, LongDateFormat)
'    Else
'        temSubTopic = "From " & Format(dtpFrom.Value, LongDateFormat) & " to " & Format(dtpTO.Value, LongDateFormat)
'    End If
'    Const PreSHape = "SHAPE {"
'    Const Sql = "SELECT tblTemReport1.* FROM tblTemReport1 "
'    Const PostSHape = " }  AS cmmdTemReportCatogerised COMPUTE cmmdTemReportCatogerised BY 'txt1'"
'
'    With DataEnvironment1
'        If .rscmmdTemReportCatogerised_Grouping.State = 1 Then .rscmmdTemReportCatogerised_Grouping.Close
'        .Commands!cmmdTemReportCatogerised_Grouping.CommandText = PreSHape & Sql & " WHERE UserID =" & UserID & " " & PostSHape
'        .cmmdTemReportCatogerised_Grouping
'    End With
'    If optByVal.Value = True Then
'        With dtrTemReportCatogerised2
'            Set .DataSource = DataEnvironment1
'            .Sections("PageHeader").Controls.Item("lblDbl1").Caption = "Value"
'            .Sections("ReportHeader").Controls.Item("lblTopic").Caption = TemTopic
'            .Sections("ReportHeader").Controls.Item("lblSubTopic").Caption = temSubTopic
'            .Caption = TemTopic & " - " & temSubTopic
'            .Show
'        End With
'    ElseIf optByQty.Value = True Then
'        With dtrTemReportCatogerised1
'            Set .DataSource = DataEnvironment1
'            .Sections("PageHeader").Controls.Item("lblDbl1").Caption = "Quentity"
'            .Sections("ReportHeader").Controls.Item("lblTopic").Caption = TemTopic
'            .Sections("ReportHeader").Controls.Item("lblSubTopic").Caption = temSubTopic
'            .Caption = TemTopic & " - " & temSubTopic
'            .Show
'        End With
'    Else
'        With dtrTemReportCatogerised1
'            Set .DataSource = DataEnvironment1
'            .Sections("PageHeader").Controls.Item("lblDbl1").Caption = "Volume"
'            .Sections("ReportHeader").Controls.Item("lblTopic").Caption = TemTopic
'            .Sections("ReportHeader").Controls.Item("lblSubTopic").Caption = temSubTopic
'            .Caption = TemTopic & " - " & temSubTopic
'            .Show
'        End With
'    End If
'    Me.MousePointer = vbDefault
'    Unload frmPleaseWait
End Sub

'Private Sub cmbCC_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then
'        cmbCC.Text = Empty
'    End If
'End Sub

Private Sub cmbChartType_Click()
    On Error Resume Next
    If cmbChartType.ListIndex > 0 Then
        mychart.ChartType = cmbChartType.ItemData(cmbChartType.ListIndex)
    End If
End Sub

Private Sub cmbChartType_Scroll()
    On Error Resume Next
    If cmbChartType.ListIndex > 0 Then
        mychart.ChartType = cmbChartType.ItemData(cmbChartType.ListIndex)
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    dtpFrom.Value = Date
    dtpTo.Value = Date
    SSTab1.Tab = 0
    cmbChartType.Enabled = False
    With cmbChartType
        .AddItem "3D Area"
        .AddItem "3D Area Stacked"
        .AddItem "3D Area Stacked 100"
        .AddItem "xl3DBar"
        .AddItem "3D Bar Clustered"
        .AddItem "3DBarStacked"
        .AddItem "3DBarStacked100"
        .AddItem "3DColumn"
        .AddItem "3DColumnClustered"
        .AddItem "3DColumnStacked"
        .AddItem "3DColumnStacked100"
        .AddItem "3DLine"
        .AddItem "3DPie"
        .AddItem "3DPieExploded"
        .AddItem "Area"
        .AddItem "AreaStacked"
        .AddItem "AreaStacked100"
        .AddItem "BarClustered"
        .AddItem "BarOfPie"
        .AddItem "BarStacked"
        .AddItem "BarStacked"
        .AddItem "BarStacked100"
        .AddItem "Bubble"
        .AddItem "Bubble3DEffect"
        .AddItem "Column"
        .AddItem "ColumnClustered"
        .AddItem "ColumnStacked"
        .AddItem "ColumnStacked"
        .AddItem "ColumnStacked100"
        .AddItem "ConeBarClustered"
        .AddItem "ConeBarStacked"
        .AddItem "ConeBarStacked100"
        .AddItem "ConeCol"
        .AddItem "ConeColClustered"
        .AddItem "ConeColStacked"
        .AddItem "ConeColStacked100"
        .AddItem "Cylinder"
        .AddItem "CylinderBarClustered"
        .AddItem "CylinderBarStacked"
        .AddItem "CylinderBarStacked100"
        .AddItem "CylinderCol"
        .AddItem "CylinderColClustered"
        .AddItem "CylinderColStacked"
        .AddItem "CylinderColStacked100"
        .AddItem "Doughnut"
        .AddItem "DoughnutExploded"
        .AddItem "Line"
        .AddItem "LineMarkers"
        .AddItem "LineMarkersStacked"
        .AddItem "LineMarkersStacked100"
        .AddItem "LineStacked"
        .AddItem "LineStacked100"
        .AddItem "Pie"
        .AddItem "PieExploded"
        .AddItem "PieOfPie"
        .AddItem "PyramidBarClustered"
        .AddItem "PyramidBarStacked"
        .AddItem "PyramidBarStacked100"
        .AddItem "PyramidCol"
        .AddItem "PyramidColClustered"
        .AddItem "PyramidColStacked"
        .AddItem "PyramidColStacked100"
        .AddItem "Radar"
        .AddItem "RadarFilled"
        .AddItem "RadarMarkers"
        .AddItem "Surface"
        .AddItem "SurfaceTopView"
        .AddItem "SurfaceTopViewWireframe"
        .AddItem "SurfaceWireframe"
        .AddItem "XYScatter"
        .AddItem "XYScatterLines"
        .AddItem "XYScatterLinesNoMarkers"
        .AddItem "XYScatterSmooth"
        .AddItem "XYScatterSmoothNoMarkers"
        
        .ItemData(0) = xl3DArea
        .ItemData(1) = xl3DAreaStacked
        .ItemData(2) = xl3DAreaStacked
        .ItemData(3) = xl3DBarClustered
        .ItemData(4) = xl3DBarClustered
        .ItemData(5) = xl3DBarStacked
        .ItemData(6) = xl3DBarStacked100
        .ItemData(7) = xl3DColumn
        .ItemData(8) = xl3DColumnClustered
        .ItemData(9) = xl3DColumnStacked
        .ItemData(10) = xl3DColumnStacked100
        .ItemData(11) = xl3DLine
        .ItemData(12) = xl3DPie
        .ItemData(13) = xl3DPieExploded
        .ItemData(14) = xlArea
        .ItemData(15) = xlAreaStacked
        .ItemData(16) = xlAreaStacked100
        .ItemData(17) = xlBarClustered
        .ItemData(18) = xlBarOfPie
        .ItemData(19) = xlBarStacked
        .ItemData(20) = xlBarStacked
        .ItemData(21) = xlBarStacked100
        .ItemData(22) = xlBubble
        .ItemData(23) = xlBubble3DEffect
        .ItemData(24) = xlColumnClustered
        .ItemData(25) = xlColumnClustered
        .ItemData(26) = xlColumnStacked
        .ItemData(27) = xlColumnStacked
        .ItemData(28) = xlColumnStacked100
        .ItemData(29) = xlConeBarClustered
        .ItemData(30) = xlConeBarStacked
        .ItemData(31) = xlConeBarStacked100
        .ItemData(32) = xlConeCol
        .ItemData(33) = xlConeColClustered
        .ItemData(34) = xlConeColStacked
        .ItemData(35) = xlConeColStacked100
        .ItemData(36) = xlCylinderBarClustered
        .ItemData(37) = xlCylinderBarClustered
        .ItemData(38) = xlCylinderBarStacked
        .ItemData(39) = xlCylinderBarStacked100
        .ItemData(40) = xlCylinderCol
        .ItemData(41) = xlCylinderColClustered
        .ItemData(42) = xlCylinderColStacked
        .ItemData(43) = xlCylinderColStacked100
        .ItemData(44) = xlDoughnut
        .ItemData(45) = xlDoughnutExploded
        .ItemData(46) = xlLine
        .ItemData(47) = xlLineMarkers
        .ItemData(48) = xlLineMarkersStacked
        .ItemData(49) = xlLineMarkersStacked100
        .ItemData(50) = xlLineStacked
        .ItemData(51) = xlLineStacked100
        .ItemData(52) = xlPie
        .ItemData(53) = xlPieExploded
        .ItemData(54) = xlPieOfPie
        .ItemData(55) = xlPyramidBarClustered
        .ItemData(56) = xlPyramidBarStacked
        .ItemData(57) = xlPyramidBarStacked100
        .ItemData(58) = xlPyramidCol
        .ItemData(59) = xlPyramidColClustered
        .ItemData(60) = xlPyramidColStacked
        .ItemData(61) = xlPyramidColStacked100
        .ItemData(62) = xlRadar
        .ItemData(63) = xlRadarFilled
        .ItemData(64) = xlRadarMarkers
        .ItemData(65) = xlSurface
        .ItemData(66) = xlSurfaceTopView
        .ItemData(67) = xlSurfaceTopViewWireframe
        .ItemData(68) = xlSurfaceWireframe
        .ItemData(69) = xlXYScatter
        .ItemData(70) = xlXYScatterLines
        .ItemData(71) = xlXYScatterLinesNoMarkers
        .ItemData(72) = xlXYScatterSmooth
        .ItemData(73) = xlXYScatterSmoothNoMarkers
    End With
End Sub


Private Sub opt2D_Click()
    Call SetGraph
End Sub

Private Sub opt3D_Click()
    Call SetGraph
End Sub

Private Sub optBar_Click()
    Call SetGraph
End Sub

Private Sub optDisplayLegend_Click()
    Call SetGraph
End Sub

Private Sub optDisplayTitle_Click()
    Call SetGraph
End Sub

Private Sub optDisplayValues_Click()
    SetGraph
End Sub

Private Sub optDoNotDisplayValues_Click()
    SetGraph
End Sub

Private Sub optLine_Click()
    Call SetGraph
End Sub

Private Sub optNoLegend_Click()
    Call SetGraph
End Sub

Private Sub optNoTitle_Click()
    Call SetGraph
End Sub

Private Sub optPie_Click()
    Call SetGraph
End Sub

Private Sub optStandardChart_Click()
    If optStandardChart.Value = True Then
        cmbChartType.Enabled = False
    Else
        cmbChartType.Enabled = True
    End If
End Sub


Private Sub optOtherCharts_Click()
    If optOtherCharts.Value = True Then
        cmbChartType.Enabled = True
    Else
        cmbChartType.Enabled = False
    End If
End Sub

Private Sub cmbChartType_Change()
    On Error Resume Next
    If cmbChartType.ListIndex > 0 Then
        mychart.ChartType = cmbChartType.ItemData(cmbChartType.ListIndex)
    End If
End Sub


Private Sub FillCombos()
    With rsItem
        If .State = 1 Then .Close
        temSQL = "Select * From tblCollectingCenter Where Deleted = 0  Order By CollectingCenter"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            lstCCIDs.AddItem !CollectingCenterID
            lstCC.AddItem !CollectingCenter
            .MoveNext
        Wend
    End With
End Sub

Private Function GetColumnName(ColumnNo As Long) As String
    Dim temnum As Integer
    Dim temnum1 As Integer
    
    If ColumnNo < 27 Then
        GetColumnName = Chr(ColumnNo + 64)
    Else
        temnum = ColumnNo \ 26
        temnum1 = ColumnNo Mod 26
        GetColumnName = Chr(temnum + 64) & Chr(temnum1 + 64)
    End If
End Function

Private Sub optXAxis_Click()
    If optXAxis.Value = True Then
        mychart.PlotBy = xlColumns
    ElseIf optYAxis.Value = True Then
        mychart.PlotBy = xlRows
    End If
End Sub

Private Sub optYAxis_Click()
    If optXAxis.Value = True Then
        mychart.PlotBy = xlColumns
    ElseIf optYAxis.Value = True Then
        mychart.PlotBy = xlRows
    End If
End Sub

Private Sub SetGraph()
    If optBar.Value = True Then
        If opt2D.Value = True Then
            mychart.ChartType = xlColumnClustered
        ElseIf opt3D.Value = True Then
            mychart.ChartType = xl3DColumn
        End If
    ElseIf optLine.Value = True Then
        If opt2D.Value = True Then
            mychart.ChartType = xlLine
        ElseIf opt3D.Value = True Then
            mychart.ChartType = xl3DLine
        End If
    ElseIf optPie.Value = True Then
        If opt2D.Value = True Then
            mychart.ChartType = xlPie
        Else
            mychart.ChartType = xl3DPie
        End If
    End If
    optStandardChart.Value = True
    optOtherCharts.Value = False
    cmbChartType.Enabled = False
    If optDisplayTitle.Value = True Then
        mychart.HasTitle = True
    Else
        mychart.HasTitle = False
    End If
    If optDisplayLegend.Value = True Then
        mychart.HasLegend = True
    Else
        mychart.HasLegend = False
    End If
    If optDisplayValues.Value = True Then
        mychart.ApplyDataLabels xlDataLabelsShowValue
    Else
        mychart.ApplyDataLabels xlDataLabelsShowNone
    End If
End Sub


