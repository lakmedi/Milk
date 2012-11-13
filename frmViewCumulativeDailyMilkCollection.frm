VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmViewCumulativeDailyMilkCollection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cumulative Report For daily milk collection"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
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
   ScaleHeight     =   7170
   ScaleWidth      =   11910
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   10680
      TabIndex        =   7
      Top             =   6720
      Width           =   1095
      _ExtentX        =   1931
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
      Format          =   100335619
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
      Format          =   100335619
      CurrentDate     =   39682
   End
   Begin MSFlexGridLib.MSFlexGrid gridMilk 
      Height          =   5415
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9551
      _Version        =   393216
      SelectionMode   =   1
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
Attribute VB_Name = "frmViewCumulativeDailyMilkCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsDailyCollection As New ADODB.Recordset
    Dim i As Integer

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub cmbCollectingCenter_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpTo_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpFrom_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    dtpFrom.Value = Date
    dtpTo.Value = Date
    Call FillCombos
    Call FormatGrid
    Call FillGrid
    
End Sub

Private Sub FormatGrid()
    With gridMilk
        
        .Rows = 1
        .Cols = 12
        
        .row = 0
        
        For i = 0 To .Cols - 1
            Select Case i
                Case 0:
                    .ColWidth(i) = 400
                    .col = i
                    .CellAlignment = 4
                    .Text = "No."
                Case 1:
                    .ColWidth(i) = 1600
                    .col = i
                    .CellAlignment = 4
                    .Text = "Date"
                Case 2:
                    .ColWidth(i) = 1100
                    .col = i
                    .CellAlignment = 4
                    .Text = "Total Leters"
                Case 3:
                    .ColWidth(i) = 1000
                    .col = i
                    .CellAlignment = 4
                    .Text = "C. LMR"
                Case 4:
                    .ColWidth(i) = 1000
                    .col = i
                    .CellAlignment = 4
                    .Text = "C. FAT%"
                Case 5:
                    .ColWidth(i) = 1300
                    .col = i
                    .CellAlignment = 4
                    .Text = "DCMR Value"
                Case 6:
                    .ColWidth(i) = 1300
                    .col = i
                    .CellAlignment = 4
                    .Text = "T. LMR"
                Case 7:
                    .ColWidth(i) = 1000
                    .col = i
                    .CellAlignment = 4
                    .Text = "T. FAT%"
                Case 8:
                    .ColWidth(i) = 1000
                    .col = i
                    .CellAlignment = 4
                    .Text = "T. Value"
                Case 9:
                    .ColWidth(i) = 1000
                    .col = i
                    .CellAlignment = 4
                    .Text = "Difference"
                Case 10:
                    .ColWidth(i) = 1
                    .col = i
                    .CellAlignment = 4
                    .Text = "ID"
                Case 11:
                    .ColWidth(i) = 1
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
    

End Sub

Private Sub FillGrid()
    If IsNumeric(cmbCollectingCenter.BoundText) = False Then Exit Sub
   
    Dim NoOfDays As Long
    Dim MaxDate As Date
    Dim MinDate As Date
    Dim ThisDate As Date
    
    Dim temTotalLeters As Double
    Dim temDCMR As Double
    Dim temTValue As Double
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
            temSql = "SELECT tblDailyCollection.* FROM tblDailyCollection Where ProgramDate = '" & Format(ThisDate, "dd MMMM yyyy") & "'  And CollectingCenterID = " & Val(cmbCollectingCenter.BoundText) & " order by tblDailyCollection.DailyCollectionID DESC"
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                gridMilk.Rows = gridMilk.Rows + 1
                gridMilk.row = gridMilk.Rows - 1
                gridMilk.TextMatrix(gridMilk.row, 0) = gridMilk.row
                gridMilk.TextMatrix(gridMilk.row, 1) = Format(ThisDate, "dd MMM yyyy")
                gridMilk.TextMatrix(gridMilk.row, 2) = !TotalVolume
                temTotalLeters = temTotalLeters + !TotalVolume
                gridMilk.TextMatrix(gridMilk.row, 3) = Round(!CalculatedLMR, 2)
                
                gridMilk.TextMatrix(gridMilk.row, 4) = Round(!CalculatedFat, 2)
                gridMilk.TextMatrix(gridMilk.row, 5) = Format(!TotalValue, "0.00")
                
                temDCMR = temDCMR + !TotalValue
                
                gridMilk.TextMatrix(gridMilk.row, 6) = !TestedLMR
                gridMilk.TextMatrix(gridMilk.row, 7) = !TestedFAT
                gridMilk.TextMatrix(gridMilk.row, 8) = Format(Round(!actualValue, 2), "0.00")
                
                temTValue = temTValue + Round(!actualValue, 2)
                
                gridMilk.TextMatrix(gridMilk.row, 9) = Format(Round(!ValueDifference, 2), "0.00")
                
                temDifference = temDifference + Round(!ValueDifference, 2)
                
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
        .Text = temDCMR
        
        .col = 9
        .Text = temDifference
        
        .col = 2
        .Text = temTotalLeters
        
        .col = 8
        .Text = temTValue
        
    
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

End Sub


Private Sub FillCombos()
    Dim Center As New clsFind
    Center.FillCombo cmbCollectingCenter, "tblCollectingCenter", "CollectingCenter", "CollectingCenterID", True
End Sub

