VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmReportsExpencesCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expences"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
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
   ScaleHeight     =   6855
   ScaleWidth      =   6915
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   6240
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
      Top             =   6240
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
      Top             =   6240
      Visible         =   0   'False
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
      Height          =   6015
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Analysis Details"
      TabPicture(0)   =   "frmReportsExpencesCategories.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dtpFrom"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dtpTO"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Graph Details"
      TabPicture(1)   =   "frmReportsExpencesCategories.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(4)=   "Frame4"
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(6)=   "Frame11"
      Tab(1).ControlCount=   7
      Begin VB.Frame Frame11 
         Height          =   1095
         Left            =   -72120
         TabIndex        =   32
         Top             =   3240
         Width           =   3615
         Begin VB.OptionButton optDisplayLegend 
            Caption         =   "Display Ligend"
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optNoLegend 
            Caption         =   "Do not display Legend"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   -72120
         TabIndex        =   29
         Top             =   2040
         Width           =   3615
         Begin VB.OptionButton optXAxis 
            Caption         =   "Plot By Colmns"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Value           =   -1  'True
            Width           =   2895
         End
         Begin VB.OptionButton optYAxis 
            Caption         =   "Plot By Rows"
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   3255
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4455
         Left            =   1560
         TabIndex        =   26
         Top             =   1440
         Width           =   3855
         Begin VB.ListBox lstItems 
            Height          =   4020
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   28
            Top             =   240
            Width           =   3615
         End
         Begin VB.ListBox lstItemIDs 
            Height          =   1680
            Left            =   3240
            Style           =   1  'Checkbox
            TabIndex        =   27
            Top             =   1800
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   2535
         Begin VB.OptionButton optBar 
            Caption         =   "Bar Chart"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optLine 
            Caption         =   "Line Chart"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton optPie 
            Caption         =   "Pie Chart"
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   14
         Top             =   1920
         Width           =   2535
         Begin VB.OptionButton opt3D 
            Caption         =   "3 D"
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton opt2D 
            Caption         =   "2 D"
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   11
         Top             =   3120
         Width           =   2535
         Begin VB.OptionButton optDisplayTitle 
            Caption         =   "Display title"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optNoTitle 
            Caption         =   "No title"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   8
         Top             =   4320
         Width           =   2535
         Begin VB.OptionButton optDoNotDisplayValues 
            Caption         =   "Do not display values"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton optDisplayValues 
            Caption         =   "Display values"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1695
         Left            =   -72120
         TabIndex        =   4
         Top             =   360
         Width           =   3615
         Begin VB.ComboBox cmbChartType 
            Height          =   360
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1080
            Width           =   3135
         End
         Begin VB.OptionButton optOtherCharts 
            Caption         =   "Other Chart Types"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   3375
         End
         Begin VB.OptionButton optStandardChart 
            Caption         =   "Standared Chart type"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   3375
         End
      End
      Begin MSComCtl2.DTPicker dtpTO 
         Height          =   375
         Left            =   1560
         TabIndex        =   21
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dddd, dd MMMM yyyy"
         Format          =   120717315
         CurrentDate     =   39576
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   1560
         TabIndex        =   22
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dddd, dd MMMM yyyy"
         Format          =   120717315
         CurrentDate     =   39576
      End
      Begin VB.Label Label5 
         Caption         =   "Center Selection"
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmReportsExpencesCategories"
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
    Dim rsShape As New ADODB.Recordset
    Dim TemTopic As String
    Dim temSubTopic As String
    Dim rsTem As New ADODB.Recordset
    Dim temSELECT As String
    Dim temWHERE As String
    Dim temFROM As String
    Dim temOrderBy As String
    Dim temGROUPBY As String


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
    Dim ItemCount As Long
    Dim ii As Integer
    Dim ArrayItemIDs() As Long
    Dim ArrayItems() As String
    Dim temnum1 As Long
    Dim Flag As Boolean
    
        Flag = False
        ItemCount = 0
        For i = 0 To lstItemIDs.ListCount - 1
            If lstItemIDs.Selected(i) = True Then
                Flag = True
                ItemCount = ItemCount + 1
            End If
        Next
        If Flag = False Then
            tr = MsgBox("You have not selected an item", vbCritical, "Select Item")
            lstItems.SetFocus
            Exit Sub
        End If
    
    ReDim ArrayItemIDs(ItemCount) As Long
    ReDim ArrayItems(ItemCount) As String
    ii = 0
    
    For i = 0 To lstItemIDs.ListCount - 1
        If lstItemIDs.Selected(i) = True Then
            lstItemIDs.ListIndex = i
            lstItems.ListIndex = i
            ArrayItemIDs(ii) = Val(lstItemIDs.Text)
            ArrayItems(ii) = lstItems.Text
            ii = ii + 1
        End If
    Next
    
    If dtpFrom.Value > dtpTo.Value Then
        temDay1 = dtpTo.Value
        dtpTo.Value = dtpFrom.Value
        dtpFrom.Value = temDay1
    Else
        temDay1 = dtpFrom.Value
        temDay2 = dtpTo.Value
    End If
    
    tempath = App.Path
    If FSys.FileExists(tempath & "\Lucky1.xls") = False Then
        tr = MsgBox("There are no graphs on the specified location")
        Exit Sub
    End If
    
    frmPleaseWait.Show
    
    DoEvents
    
    Set myworkbook = GetObject(tempath & "\Lucky1.xls")
    Set myworksheet = myworkbook.WorkSheets.Item(1)
    Set mychart = myworkbook.Charts.Item(1)
    
    myworksheet.UsedRange.Clear
        
    myworksheet.Cells(1, 1) = "Collecting Center"
        
    
    For i = 0 To ItemCount - 1
        myworksheet.Cells(1, i + 2) = ArrayItems(i)
    Next i
    
    Dim rsTem1 As New ADODB.Recordset
    
    
    With rsTem
        temSQL = "Select * from tblExpenceCategory Where Deleted = 0  order by ExpenceCategory"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For ii = 0 To .RecordCount - 1
            
                myworksheet.Cells(ii + 2, 1) = !expenceCategory
                
                For i = 0 To ItemCount - 1
                    myworksheet.Cells(1, i + 2) = ArrayItems(i)
                    If rsProduction.State = 1 Then rsProduction.Close
                    temSQL = "SELECT Sum(ExpenceValue) AS DisplayValue "
                    temSQL = temSQL & "From tblExpence "
                    temSQL = temSQL & "WHERE Deleted = 0  AND CollectingCenterID = " & ArrayItemIDs(i) & " AND ExpenceCategoryID = " & !ExpenceCategoryID & " AND ExpenceDate Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
                    If rsProduction.State = 1 Then rsProduction.Close
                    rsProduction.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                    If rsProduction.RecordCount > 0 Then
                        If Not IsNull(rsProduction!DisplayValue) Then
                            myworksheet.Cells(ii + 2, i + 2) = rsProduction!DisplayValue
                        End If
                    End If
                    If rsProduction.State = 1 Then rsProduction.Close
                    DoEvents
                Next i
                .MoveNext
            
            Next ii
        End If
        .Close
    End With
    
    
    
    
    mychart.SetSourceData myworksheet.Range("A1:" & GetColumnName(i + 1) & ii + 1)
    
    Call SetGraph
        
    If optDisplayTitle.Value = True Then
        TemTopic = ""
        TemTopic = TemTopic & "Total Expences "
        TemTopic = TemTopic & "of Collecting Centers "
        If dtpFrom.Value = dtpTo.Value Then
            temSubTopic = "On " & Format(dtpFrom.Value, LongDateFormat)
        Else
            temSubTopic = "From " & Format(dtpFrom.Value, LongDateFormat) & " to " & Format(dtpTo.Value, LongDateFormat)
        End If
        mychart.HasTitle = True
        mychart.ChartTitle.Caption = TemTopic & vbNewLine & temSubTopic
    Else
        mychart.HasTitle = False
    End If
    
    mychart.HasLegend = True
    myworkbook.Save
    mychart.Activate
    frmGraph.Caption = TemTopic & " - " & temSubTopic
    Unload frmPleaseWait
    frmGraph.Show
End Sub



Private Sub bttnReport_Click()
'
'    Dim TemValue As Double
'    Dim i As Integer
'    Dim tr As Integer
'    Dim temDay1 As Date
'    Dim temDay2 As Date
'    Dim temday3 As Date
'    Dim ItemCount As Long
'    Dim ii As Integer
'    Dim ArrayItemIDs() As Long
'    Dim ArrayItems() As String
'    Dim temnum1 As Long
'    Dim Flag As Boolean
'
'    If optSelectdeItem.Value = True Then
'        Flag = False
'        ItemCount = 0
'        For i = 0 To lstItemIDs.ListCount - 1
'            If lstItemIDs.Selected(i) = True Then
'                Flag = True
'                ItemCount = ItemCount + 1
'            End If
'        Next
'        If Flag = False Then
'            tr = MsgBox("You have not selected an item", vbCritical, "Select Item")
'            lstItems.SetFocus
'            Exit Sub
'        End If
'    ElseIf optSingleItem.Value = True Then
'        If IsNumeric(dtcItem.BoundText) = False Then
'            tr = MsgBox("You have not selected an item", vbCritical, "Select Item")
'            dtcItem.SetFocus
'            Exit Sub
'        End If
'    End If
'
'    ReDim ArrayItemIDs(ItemCount) As Long
'    ReDim ArrayItems(ItemCount) As String
'    ii = 0
'
'    For i = 0 To lstItemIDs.ListCount - 1
'        If lstItemIDs.Selected(i) = True Then
'            lstItemIDs.ListIndex = i
'            lstItems.ListIndex = i
'            ArrayItemIDs(ii) = Val(lstItemIDs.Text)
'            ArrayItems(ii) = lstItems.Text
'            ii = ii + 1
'        End If
'    Next
'
'    If dtpFrom.Value > dtpTo.Value Then
'        temDay1 = dtpTo.Value
'        dtpTo.Value = dtpFrom.Value
'        dtpFrom.Value = temDay1
'    Else
'        temDay1 = dtpFrom.Value
'        temDay2 = dtpTo.Value
'    End If
'
'    With rsTemReport
'        If .State = 1 Then .Close
'        temSql = "Delete * from tblTemReport1 where UserID = " & UserID
'        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'        If .State = 1 Then .Close
'        temSql = "SELECT * from tblTemReport1"
'        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'    End With
'
'
'    If rsTemReport.State = 1 Then rsTemReport.Close
'    temSql = "SELECT * from tblTemReport1"
'    rsTemReport.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'
'    With rsAgent
'        If .State = 1 Then .Close
'        temSql = "SELECT * from tblAgent order by Agent"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            i = 0
'            While .EOF = False
'                i = i + 1
'                If optAllItems.Value = True Then
'                    If rsProduction.State = 1 Then rsProduction.Close
'                    If optByVal.Value = True Then
'                        temSql = "SELECT Sum(tblAgentIssue.Price) AS DisplayValue "
'                    ElseIf optByQty.Value = True Then
'                        temSql = "SELECT Sum(tblAgentIssue.Quentity) AS DisplayValue "
'                    ElseIf optByVol.Value = True Then
'                        temSql = "SELECT Sum(tblAgentIssue.Volume) AS DisplayValue "
'                    End If
'                    temSql = temSql & "From tblAgentIssue "
'                    temSql = temSql & "WHERE tblAgentIssue.AgentID = " & !AgentID & " AND tblAgentIssue.Date Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
'                    rsProduction.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                    If rsProduction.RecordCount > 0 Then
'                        rsTemReport.AddNew
'                        rsTemReport!UserID = UserID
'                        rsTemReport!txt1 = !Agent
'                        If Not IsNull(rsProduction!DisplayValue) Then
'                            rsTemReport!dbl1 = rsProduction!DisplayValue
'                        End If
'                        rsTemReport.Update
'                    End If
'                    If rsProduction.State = 1 Then rsProduction.Close
'              ElseIf optSingleItem.Value = True Then
'                    If rsProduction.State = 1 Then rsProduction.Close
'                    If optByVal.Value = True Then
'                        temSql = "SELECT Sum(tblAgentIssue.Price) AS DisplayValue "
'                    ElseIf optByQty.Value = True Then
'                        temSql = "SELECT Sum(tblAgentIssue.Quentity) AS DisplayValue "
'                    ElseIf optByVol.Value = True Then
'                        temSql = "SELECT Sum(tblAgentIssue.Volume) AS DisplayValue "
'                    End If
'                    temSql = temSql & "From tblAgentIssue "
'                    temSql = temSql & "WHERE tblAgentIssue.AgentID = " & !AgentID & " AND tblAgentIssue.Date Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'  And tblAgentIssue.ItemID = " & dtcItem.BoundText
'                    rsProduction.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                    If rsProduction.RecordCount > 0 Then
'                        rsTemReport.AddNew
'                        rsTemReport!UserID = UserID
'                        rsTemReport!txt1 = !Agent
'                        If Not IsNull(rsProduction!DisplayValue) Then
'                            rsTemReport!dbl1 = rsProduction!DisplayValue
'                        End If
'                        rsTemReport.Update
'                    End If
'                    If rsProduction.State = 1 Then rsProduction.Close
'              ElseIf optSelectdeItem.Value = True Then
'                    For temnum1 = 0 To ItemCount
'                        If rsProduction.State = 1 Then rsProduction.Close
'                        If optByVal.Value = True Then
'                            temSql = "SELECT Sum(tblAgentIssue.Price) AS DisplayValue "
'                        ElseIf optByQty.Value = True Then
'                            temSql = "SELECT Sum(tblAgentIssue.Quentity) AS DisplayValue "
'                        ElseIf optByVol.Value = True Then
'                            temSql = "SELECT Sum(tblAgentIssue.Volume) AS DisplayValue "
'                        End If
'                        temSql = temSql & "From tblAgentIssue "
'                        temSql = temSql & "WHERE tblAgentIssue.AgentID = " & !AgentID & " AND tblAgentIssue.Date Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
'                        temSql = temSql & " AND tblAgentIssue.ItemID = " & ArrayItemIDs(temnum1)
'                        rsProduction.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                        If rsProduction.RecordCount > 0 Then
'                            rsTemReport.AddNew
'                            rsTemReport!UserID = UserID
'                            rsTemReport!txt1 = !Agent
'                            rsTemReport!txt2 = ArrayItems(temnum1)
'                            If Not IsNull(rsProduction!DisplayValue) Then
'                                rsTemReport!dbl1 = rsProduction!DisplayValue
'                            End If
'                            rsTemReport.Update
'                        End If
'                        If rsProduction.State = 1 Then rsProduction.Close
'                    Next
'                End If
'                .MoveNext
'                DoEvents
'            Wend
'        End If
'    End With
'    If rsTemReport.State = 1 Then rsTemReport.Close
'    temSql = "SELECT * from tblTemReport1 where UserID = " & UserID
'    rsTemReport.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'    rsTemReport.AddNew
'    rsTemReport!UserID = UserID
'    rsTemReport!txt1 = " "
'    rsTemReport!txt2 = "zz - Software By Lakmedipro  -  zz"
'    rsTemReport.Update
'    If rsTemReport.State = 1 Then rsTemReport.Close
'
'
'    TemTopic = ""
'    If optByQty.Value = True Then
'        TemTopic = TemTopic & " Quentity-wise "
'    ElseIf optByVal.Value = True Then
'        TemTopic = TemTopic & " Value-wise "
'    ElseIf optByVol.Value = True Then
'        TemTopic = TemTopic & " Volume-wise "
'    End If
'    TemTopic = TemTopic & "Issue "
'    If optAllItems.Value = True Then
'        TemTopic = TemTopic & "of all Items "
'    ElseIf optSingleItem.Value = True Then
'        TemTopic = TemTopic & " of " & dtcItem.Text & " "
'    Else
'        TemTopic = TemTopic & "of selected Items "
'    End If
'    TemTopic = TemTopic & " To All Agents"
'    If dtpFrom.Value = dtpTo.Value Then
'        temSubTopic = "On " & Format(dtpFrom.Value, LongDateFormat)
'    Else
'        temSubTopic = "From " & Format(dtpFrom.Value, LongDateFormat) & " to " & Format(dtpTo.Value, LongDateFormat)
'    End If
'
'
'    If optAllItems.Value = True Or optSingleItem.Value = True Then
'        With rsTemReport
'            If .State = 1 Then .Close
'            temSql = "SELECT * from tblTemReport1 where UserID = " & UserID & " Order by temReportID"
'            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        End With
'        If optByVal.Value = True Then
'            With dtrTemReport1c
'                Set .DataSource = rsTemReport
'                .Sections("Section1").Controls("txtTxt1").DataField = "txt1"
'                .Sections("Section1").Controls("txtdbl1").DataField = "dbl1"
'                .Sections("Section5").Controls("fundbl1").DataField = "dbl1"
'                .Sections("Section2").Controls("lbldbl1").Caption = "Sale"
'                .Sections("Section2").Controls("lbltxt1").Caption = "Time period"
'                .Sections("Section4").Controls("lblTopic").Caption = TemTopic
'                .Sections("Section4").Controls("lblSubTopic").Caption = temSubTopic
'                .Caption = TemTopic & " - " & temSubTopic
'                .Show
'            End With
'        Else
'            With dtrTemReport1d
'                Set .DataSource = rsTemReport
'                .Sections("Section1").Controls("txtTxt1").DataField = "txt1"
'                .Sections("Section1").Controls("txtdbl1").DataField = "dbl1"
'                .Sections("Section5").Controls("fundbl1").DataField = "dbl1"
'                .Sections("Section2").Controls("lbldbl1").Caption = "Sale"
'                .Sections("Section2").Controls("lbltxt1").Caption = "Time period"
'                .Sections("Section4").Controls("lblTopic").Caption = TemTopic
'                .Sections("Section4").Controls("lblSubTopic").Caption = temSubTopic
'                .Caption = TemTopic & " - " & temSubTopic
'                .Show
'            End With
'        End If
'
'    ElseIf optSelectdeItem.Value = True Then
'
'        Const PreSHape = "SHAPE {"
'        Const Sql = "SELECT tblTemReport1.* FROM tblTemReport1 "
'        Const PostSHape = " }  AS cmmdTemReportCatogerised COMPUTE cmmdTemReportCatogerised BY 'txt1'"
'
'        With DataEnvironment1
'            If .rscmmdTemReportCatogerised_Grouping.State = 1 Then .rscmmdTemReportCatogerised_Grouping.Close
'            .Commands!cmmdTemReportCatogerised_Grouping.CommandText = PreSHape & Sql & " WHERE UserID =" & UserID & " " & PostSHape
'            .cmmdTemReportCatogerised_Grouping
'        End With
'
'            If optByVal.Value = True Then
'                With dtrTemReportCatogerised1
'                    Set .DataSource = DataEnvironment1
'                    .Sections("PageHeader").Controls.Item("lblDbl1").Caption = "Value"
'                    .Sections("ReportHeader").Controls.Item("lblTopic").Caption = temSql
'                    .Sections("ReportHeader").Controls.Item("lblSubTopic").Caption = Empty
'                    .Caption = TemTopic & " - " & temSubTopic
'                    .Show
'                End With
'            Else
'                With dtrTemReportCatogerised1
'                    Set .DataSource = DataEnvironment1
'                    .Sections("PageHeader").Controls.Item("lblDbl1").Caption = "Value"
'                    .Sections("ReportHeader").Controls.Item("lblTopic").Caption = temSql
'                    .Sections("ReportHeader").Controls.Item("lblSubTopic").Caption = Empty
'                    .Caption = TemTopic & " - " & temSubTopic
'                    .Show
'                End With
'            End If
'    End If
'
'    Unload frmPleaseWait
End Sub

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
    optAllItems_Click
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
    Call SetGraph
End Sub

Private Sub optDoNotDisplayValues_Click()
    Call SetGraph
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
        temSQL = "Select * From tblCollectingCenter Where Deleted = 0  Order By CollectingCenter "
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        lstItemIDs.Clear
        lstItems.Clear
        If .RecordCount > 0 Then
            While .EOF = False
                lstItemIDs.AddItem !CollectingCenterID
                lstItems.AddItem !CollectingCenter
                .MoveNext
            Wend
        End If
    End With
End Sub

Private Sub lstItems_Click()
    lstItemIDs.ListIndex = lstItems.ListIndex
    lstItemIDs.Selected(lstItems.ListIndex) = lstItems.Selected(lstItems.ListIndex)
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

Private Sub optAllItems_Click()
End Sub


Private Sub optXAxis_Click()
    Call SetGraph
End Sub

Private Sub optYAxis_Click()
    Call SetGraph
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
    If optXAxis.Value = True Then
        mychart.PlotBy = xlColumns
    ElseIf optYAxis.Value = True Then
        mychart.PlotBy = xlRows
    End If
End Sub

