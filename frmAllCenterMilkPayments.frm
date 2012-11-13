VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAllCenterMilkPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
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
   ScaleHeight     =   3090
   ScaleWidth      =   5295
   Begin VB.TextBox txtPath 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   4815
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Close"
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
   Begin btButtonEx.ButtonEx btnGraph 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Save"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   224657411
      CurrentDate     =   39861
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   224657411
      CurrentDate     =   39861
   End
   Begin btButtonEx.ButtonEx btnChangePath 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1320
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
   Begin VB.Label Label3 
      Caption         =   "Path to Save"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmAllCenterMilkPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
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
    
    
'    Dim myworkbook As Excel.Workbook
'    Dim myworksheet As Excel.Worksheet
'    Dim myworksheet1 As Excel.Worksheet
'
'    Dim mychart As Excel.Chart
    
    Dim tempath As String
    Dim FSys As New Scripting.FileSystemObject
    Dim rsViewDriver As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim rsShape As New ADODB.Recordset
    
    Dim TemTopic As String
    Dim temSubTopic As String
    
    Dim rsTem As New ADODB.Recordset
        
    Dim rsTemReport As New ADODB.Recordset

    Dim temSql As String
    Dim temSELECT As String
    Dim temWHERE As String
    Dim temFROM As String
    Dim temOrderBy As String
    Dim temGROUPBY As String
    
    Dim rsProduction As New ADODB.Recordset
    Dim rsViewItem As New ADODB.Recordset


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

Private Sub btnChangePath_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtPath.SetFocus
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnGraph_Click()

    Dim MyExcelApplication As Excel.Application
    
    Dim myworkbook As Excel.Workbook
    Dim MyWorksheetDMCRVAL As Excel.Worksheet
    Dim MyWorksheetDMCRVOL As Excel.Worksheet
    Dim MyWorksheetGRNVAL As Excel.Worksheet
    Dim MyWorksheetGRNVOL As Excel.Worksheet
    
    
    
    Dim temDays As Integer
    Dim temDay1 As Date
    Dim temDay2 As Date
    
    Dim TotalDMCRLiters As Double
    Dim TotalDMCRValue As Double
    Dim TotalGRNLiters As Double
    Dim TotalGRNValue As Double
    
    
    Dim TemDMCRLiters As Double
    Dim TemDMCRValue As Double
    Dim TemGRNLiters As Double
    Dim temGRNValue As Double
    
    
    Dim i As Integer
    
    Dim ThisDay As Date
    
    Dim CollectingCenterCOunt As Long
    Dim DayCount As Long
    
    Dim rsCC As New ADODB.Recordset
    
    DayCount = DateDiff("d", dtpFrom.Value, dtpTo.Value)
    
    If dtpFrom.Value > dtpTo.Value Then
        temDay1 = dtpTo.Value
        dtpTo.Value = dtpFrom.Value
        dtpFrom.Value = temDay1
    Else
        temDay1 = dtpFrom.Value
        temDay2 = dtpTo.Value
    End If
    
    frmPleaseWait.Show
    DoEvents
    
    Set MyExcelApplication = New Excel.Application
    
    Set myworkbook = MyExcelApplication.Workbooks.Add
    
    Set MyWorksheetDMCRVAL = myworkbook.Worksheets.Add
    Set MyWorksheetDMCRVAL = myworkbook.Worksheets("Sheet4")
    Set MyWorksheetDMCRVOL = myworkbook.Worksheets("Sheet1")
    Set MyWorksheetGRNVAL = myworkbook.Worksheets("Sheet2")
    Set MyWorksheetGRNVOL = myworkbook.Worksheets("Sheet3")
    
    MyWorksheetDMCRVAL.Name = "DMCR Value"
    MyWorksheetDMCRVOL.Name = "DMCR Volume"
    MyWorksheetGRNVAL.Name = "GRN Value"
    MyWorksheetGRNVOL.Name = "GRN Volume"
    
    CollectingCenterCOunt = 1
    

    
    With rsCC
        If .State = 1 Then .Close
        temSql = "Select * from tblCOllectingCenter where Deleted = 0  order by COllectingCenter"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            
            CollectingCenterCOunt = CollectingCenterCOunt + 1
            
            MyWorksheetDMCRVAL.Cells(3, CollectingCenterCOunt) = !CollectingCenter
            MyWorksheetDMCRVOL.Cells(3, CollectingCenterCOunt) = !CollectingCenter
            MyWorksheetGRNVAL.Cells(3, CollectingCenterCOunt) = !CollectingCenter
            MyWorksheetGRNVOL.Cells(3, CollectingCenterCOunt) = !CollectingCenter
            

            
            TotalGRNLiters = 0
            TotalGRNValue = 0
            TotalDMCRLiters = 0
            TotalDMCRValue = 0
            
            
            For i = 0 To DayCount
                ThisDay = dtpFrom.Value + i
                temGRNValue = 0
                TemGRNLiters = 0
                TemDMCRValue = 0
                TemDMCRLiters = 0
                
                TemDMCRLiters = CCMilkSupply(ThisDay, !CollectingCenterID).Liters
                TemDMCRValue = CCMilkSupply(ThisDay, !CollectingCenterID).Value
                TemGRNLiters = CCGRN(ThisDay, !CollectingCenterID).Liters
                temGRNValue = CCGRN(ThisDay, !CollectingCenterID).Value

                TotalDMCRLiters = TotalDMCRLiters + TemDMCRLiters
                TotalDMCRValue = TotalDMCRValue + TemDMCRValue
                TotalGRNLiters = TotalGRNLiters + TemGRNLiters
                TotalGRNValue = TotalGRNValue + temGRNValue

                MyWorksheetDMCRVAL.Cells(i + 4, 1) = ThisDay
                MyWorksheetDMCRVOL.Cells(i + 4, 1) = ThisDay
                MyWorksheetGRNVAL.Cells(i + 4, 1) = ThisDay
                MyWorksheetGRNVOL.Cells(i + 4, 1) = ThisDay
                
                MyWorksheetDMCRVAL.Cells(i + 4, CollectingCenterCOunt) = TemDMCRValue
                MyWorksheetDMCRVOL.Cells(i + 4, CollectingCenterCOunt) = TemDMCRLiters
                MyWorksheetGRNVAL.Cells(i + 4, CollectingCenterCOunt) = temGRNValue
                MyWorksheetGRNVOL.Cells(i + 4, CollectingCenterCOunt) = TemGRNLiters

            Next

            i = i + 1

            MyWorksheetDMCRVAL.Cells(i + 3, CollectingCenterCOunt) = TotalDMCRValue
            MyWorksheetDMCRVOL.Cells(i + 3, CollectingCenterCOunt) = TotalDMCRLiters
            MyWorksheetGRNVAL.Cells(i + 3, CollectingCenterCOunt) = TotalGRNValue
            MyWorksheetGRNVOL.Cells(i + 3, CollectingCenterCOunt) = TotalGRNLiters
            
            
            .MoveNext
            
        Wend
        
    
    
    End With
    
    
    TemTopic = InstitutionName
    temSubTopic = "Milk Collection from " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    
    MyWorksheetDMCRVAL.Activate
    
    
    myworkbook.Close True, txtPath.Text & "\" & temSubTopic & ".xls"
    
    Set MyWorksheetDMCRVAL = Nothing
    Set MyWorksheetDMCRVOL = Nothing
    Set MyWorksheetGRNVAL = Nothing
    Set MyWorksheetGRNVOL = Nothing
    
    Set myworkbook = Nothing
    Set MyExcelApplication = Nothing
    
    Unload frmPleaseWait

    ShellExecute 0&, vbNullString, txtPath.Text & "\" & temSubTopic & ".xls", vbNullString, vbNullString, vbNormalFocus
    

End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpTo.SetFocus
    End If
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnChangePath_Click
    End If
End Sub

Private Sub Form_Load()
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = Date
    txtPath.Text = GetSetting(App.EXEName, Me.Name, "SavePath", App.Path)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, Me.Name, "SavePath", txtPath.Text
End Sub
