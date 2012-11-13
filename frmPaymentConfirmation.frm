VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPaymentConfirmation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Confirmation"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11385
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
   ScaleHeight     =   7050
   ScaleWidth      =   11385
   Begin VB.ListBox lstCompleteValue 
      Height          =   780
      Left            =   9000
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox lstToCompleteValue 
      Height          =   780
      Left            =   9240
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "To Complete"
      TabPicture(0)   =   "frmPaymentConfirmation.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstToCompleteID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstToComplete"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnComplete"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtToComplete"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Completed"
      TabPicture(1)   =   "frmPaymentConfirmation.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstCompletedID"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lstCompleted"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "btnIncomplete"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtCompleted"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.TextBox txtCompleted 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   4320
         Width           =   2535
      End
      Begin VB.TextBox txtToComplete 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   -72360
         TabIndex        =   16
         Top             =   4320
         Width           =   2535
      End
      Begin btButtonEx.ButtonEx btnComplete 
         Height          =   495
         Left            =   -66000
         TabIndex        =   9
         Top             =   4200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "Mark as Complete"
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
      Begin btButtonEx.ButtonEx btnIncomplete 
         Height          =   495
         Left            =   9000
         TabIndex        =   11
         Top             =   4200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "Mark as Incomplete"
         ForeColor       =   0
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
      Begin VB.ListBox lstToComplete 
         Height          =   3660
         Left            =   -74880
         MultiSelect     =   2  'Extended
         TabIndex        =   8
         Top             =   480
         Width           =   10935
      End
      Begin VB.ListBox lstToCompleteID 
         Height          =   3660
         Left            =   -68280
         MultiSelect     =   2  'Extended
         TabIndex        =   13
         Top             =   480
         Width           =   495
      End
      Begin VB.ListBox lstCompleted 
         Height          =   3660
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   10
         Top             =   480
         Width           =   10935
      End
      Begin VB.ListBox lstCompletedID 
         Height          =   3660
         Left            =   6720
         MultiSelect     =   2  'Extended
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Total Value to Complete"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   4320
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "Total Value to Complete"
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   4320
         Width           =   3975
      End
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   276627459
      CurrentDate     =   39748
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   276627459
      CurrentDate     =   39748
   End
   Begin MSDataListLib.DataCombo cmbCollectingCenter 
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9120
      TabIndex        =   12
      Top             =   6480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "C&lose"
      ForeColor       =   0
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
   Begin VB.Label Label4 
      Caption         =   "Pay Orders"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmPaymentConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsCollectingCenters As New ADODB.Recordset
    
    Dim temSql As String
    

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnComplete_Click()
    Dim i As Integer
    Dim rsSP As New ADODB.Recordset
    For i = 0 To lstToComplete.ListCount - 1
        If lstToComplete.Selected(i) = True Then
            With rsSP
                If .State = 1 Then .Close
                temSql = "SELECT * from tblSupplierPayments where SupplierPaymentsID = " & Val(lstToCompleteID.List(i))
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !Completed = True
                    !CompletedUserID = UserID
                    !CompletedDate = Date
                    !CompletedTime = Time
                    .Update
                End If
                .Close
            End With
        End If
    Next i
    FillLists
End Sub

Private Sub btnIncomplete_Click()
    Dim i As Integer
    Dim rsSP As New ADODB.Recordset
    For i = 0 To lstCompleted.ListCount - 1
        If lstCompleted.Selected(i) = True Then
            With rsSP
                If .State = 1 Then .Close
                temSql = "SELECT * from tblSupplierPayments where SupplierPaymentsID = " & Val(lstCompletedID.List(i))
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !Completed = False
                    !InCompletedUserID = UserID
                    !InCompletedDate = Date
                    !InCompletedTime = Time
                    .Update
                End If
                .Close
            End With
        End If
    Next i
    FillLists

End Sub

Private Sub cmbCollectingCenter_Change()
    Call FillLists
End Sub



Private Sub dtpFrom_Change()
    Call FillLists
End Sub

Private Sub dtpTo_Change()
    Call FillLists
End Sub

Private Sub Form_Load()
    dtpFrom.Value = Date
    dtpTo.Value = Date
    Call FillCombos
End Sub

Private Sub FillCombos()
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCollectingCenter, "CollectingCenter", True
End Sub

Private Sub FillLists()
    lstCompleted.Clear
    lstCompletedID.Clear
    lstCompleteValue.Clear
    lstToComplete.Clear
    lstToCompleteID.Clear
    lstToCompleteValue.Clear
    
    Dim ToCompleteValue As Double
    Dim CompletedValue As Double
    
    Dim rsSP As New ADODB.Recordset
    With rsSP
        temSql = "SELECT (tblSupplier.Supplier + ' - From ' +  convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) + ' To ' +  convert(varchar, tblCollectingCenterPaymentSummery.ToDate, 102) + ' -  Rs. ' +  Format(tblSupplierPayments.Value,'#,##0.00')) AS Display, tblSupplierPayments.SupplierPaymentsID, tblSupplierPayments.Value " & _
                    "FROM ((tblSupplierPayments LEFT JOIN tblSupplier ON tblSupplierPayments.SupplierID = tblSupplier.SupplierID) LEFT JOIN tblCollectingCenterPaymentSummery ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblCollectingCenter ON tblSupplier.CollectingCenterID = tblCollectingCenter.CollectingCenterID " & _
                    "WHERE (((tblCollectingCenter.CollectingCenterID)=" & Val(cmbCollectingCenter.BoundText) & ") AND ((tblSupplierPayments.Completed) = 0) AND ((tblSupplierPayments.GeneratedDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'))"
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                lstToComplete.AddItem !Display
                lstToCompleteID.AddItem !SupplierPaymentsID
                lstToCompleteValue.AddItem !Value
                ToCompleteValue = ToCompleteValue + !Value
                .MoveNext
            Wend
        End If
        .Close
        
        
        ' convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102)
        ' convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102)
        
        temSql = "SELECT (tblSupplier.Supplier + ' - From ' +  convert(varchar, tblCollectingCenterPaymentSummery.FromDate, 102) + ' To ' +  convert(varchar, tblCollectingCenterPaymentSummery.ToDate, 102) + ' -  Rs. ' +  Format(tblSupplierPayments.Value,'#,##0.00')) AS Display, tblSupplierPayments.SupplierPaymentsID, tblSupplierPayments.Value " & _
                    "FROM ((tblSupplierPayments LEFT JOIN tblSupplier ON tblSupplierPayments.SupplierID = tblSupplier.SupplierID) LEFT JOIN tblCollectingCenterPaymentSummery ON tblSupplierPayments.CollectingCenterPaymentSummeryID = tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID) LEFT JOIN tblCollectingCenter ON tblSupplier.CollectingCenterID = tblCollectingCenter.CollectingCenterID " & _
                    "WHERE (((tblCollectingCenter.CollectingCenterID)=" & Val(cmbCollectingCenter.BoundText) & ") AND ((tblSupplierPayments.Completed)=True) AND ((tblSupplierPayments.GeneratedDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                lstCompleted.AddItem !Display
                lstCompletedID.AddItem !SupplierPaymentsID
                lstCompleteValue.AddItem !Value
                CompletedValue = CompletedValue + !Value
                .MoveNext
            Wend
        End If
        .Close
    End With
    
    
    txtToComplete.Text = Format(ToCompleteValue, "#,##0.00")
    txtCompleted.Text = Format(CompletedValue, "#,##0.00")
    
End Sub
