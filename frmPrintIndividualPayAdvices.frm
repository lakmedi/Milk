VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPrintIndividualPayAdvices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Individual Pay Advices"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9645
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
   ScaleHeight     =   6900
   ScaleWidth      =   9645
   Begin VB.ListBox lstSPID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   6600
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.ListBox lstSP 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   2160
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   1080
      Width           =   5295
   End
   Begin MSDataListLib.DataCombo cmbCollectingCenter 
      Height          =   360
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbCCPS 
      Height          =   360
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "&Payment Period"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmPrintIndividualPayAdvices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsCCPS As New ADODB.Recordset
    Dim temSql As String
    
Private Sub cmbCCPS_Change()
    Call FillLists
End Sub

Private Sub FillLists()
    Dim rsTemPS As New ADODB.Recordset
    lstSP.Clear
    lstSPID.Clear
    With rsTemPS
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblSupplier where CollectingCenterID = " & Val(cmbCollectingCenter.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                lstSP.AddItem !Supplier
                lstSP.AddItem !SupplierID
                .MoveNext
            Wend
        End If
        .Close
    End With
End Sub

Private Sub cmbCollectingCenter_Change()
    With rsCCPS
        temSql = "SELECT ('From ' & Format(tblCollectingCenterPaymentSummery.FromDate,'dd mmmm yyyy',1,1) & ' To ' & Format(tblCollectingCenterPaymentSummery.ToDate,'dd mmmm yyyy',1,1)) AS Display, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID " & _
                    "FROM tblCollectingCenter RIGHT JOIN tblCollectingCenterPaymentSummery ON tblCollectingCenter.CollectingCenterID = tblCollectingCenterPaymentSummery.CollectingCenterID " & _
                    "Where (((tblCollectingCenterPaymentSummery.CollectingCenterID) = " & Val(cmbCollectingCenter.BoundText) & ")) " & _
                    "ORDER BY tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID DESC"
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCCPS
        Set .RowSource = rsCCPS
        .ListField = "Display"
        .BoundColumn = "CollectingCenterPaymentSummeryID"
    End With
End Sub

Private Sub Form_Load()
    FillCombos
End Sub

Private Sub FillCombos()
    Dim Centers As New clsFillCombos
    Centers.FillAnyCombo cmbCollectingCenter, "CollectingCenter", True
End Sub

Private Sub PrintIndividualPayAdvice(SupplierID As Long, CollectingCenterPaymentSummeryID As Long)

End Sub
