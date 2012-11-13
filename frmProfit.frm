VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmProfit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profit"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
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
   ScaleHeight     =   5895
   ScaleWidth      =   9375
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   7800
      TabIndex        =   7
      Top             =   5160
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
   Begin VB.TextBox txtProfit 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox txtOtherExpences 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   13
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtMilkPayments 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox txtIncome 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   1560
      Width           =   2295
   End
   Begin btButtonEx.ButtonEx btnCalculate 
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&calculate"
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
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMM yyyy"
      Format          =   226557955
      CurrentDate     =   39682
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMM yyyy"
      Format          =   226557955
      CurrentDate     =   39682
   End
   Begin VB.Label Label7 
      Caption         =   "Total Profit"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Total Income"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Other Expences"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Total Milk Payments"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub btnCalculate_Click()
    Screen.MousePointer = vbHourglass
    Call ClearValues
    Call CalculateMilkPayments
    Call CalculateExpences
    Call CalculateIncome
    Call CalculateProfit
    Screen.MousePointer = vbDefault
    cmbCollectingCenter.SetFocus
End Sub

Private Sub ClearValues()
    txtMilkPayments.Text = "0.00"
    txtIncome.Text = "0.00"
    txtOtherExpences.Text = "0.00"
    txtProfit.Text = "0.00"
    
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


Private Sub cmbCollectingCenter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpFrom.SetFocus
    End If
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpTo.SetFocus
    End If
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnCalculate.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = Date
End Sub

Private Sub FillCombos()
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCollectingCenter, "CollectingCenter", True
End Sub

Private Sub CalculateMilkPayments()
    Dim rsPayments As New ADODB.Recordset
    With rsPayments
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblSupplierPayments.Value) AS SumOfValue " & _
                    "FROM tblSupplierPayments LEFT JOIN tblSupplier ON tblSupplierPayments.SupplierID = tblSupplier.SupplierID "
        If IsNumeric(cmbCollectingCenter.BoundText) = True Then
            temSql = temSql & "WHERE (((tblSupplierPayments.CompletedDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND ((tblSupplier.CollectingCenterID)=" & Val(cmbCollectingCenter.BoundText) & ") AND ((tblSupplierPayments.Completed)=True) AND ((tblSupplierPayments.Deleted) = 0))"

        Else
            temSql = temSql & "WHERE (((tblSupplierPayments.CompletedDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND ((tblSupplierPayments.Completed)=True) AND ((tblSupplierPayments.Deleted) = 0))"
        End If
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then
                txtMilkPayments.Text = Format(!SumOfValue, "0.00")
            End If
        End If
        .Close
    End With
    
End Sub

Private Sub CalculateIncome()
    Dim rsIncome As New ADODB.Recordset
    With rsIncome
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblIncome.IncomeValue) AS SumOfIncomeValue FROM tblIncome "
        If IsNumeric(cmbCollectingCenter.BoundText) = True Then
            temSql = temSql & "WHERE (((tblIncome.IncomeDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND ((tblIncome.Deleted) = 0) AND ((tblIncome.CollectingCenterID)=" & Val(cmbCollectingCenter.BoundText) & "))"
        Else
            temSql = temSql & "WHERE (((tblIncome.IncomeDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND ((tblIncome.Deleted) = 0))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfIncomeValue) = False Then
                txtIncome.Text = Format(!SumOfIncomeValue, "0.00")
            End If
        End If
        .Close
    End With
End Sub

Private Sub CalculateExpences()
    Dim rsExpence As New ADODB.Recordset
    With rsExpence
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblExpence.ExpenceValue) AS SumOfExpenceValue FROM tblExpence "
        If IsNumeric(cmbCollectingCenter.BoundText) = True Then
            temSql = temSql & "WHERE (((tblExpence.ExpenceDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND ((tblExpence.Deleted) = 0) AND ((tblExpence.CollectingCenterID)=" & Val(cmbCollectingCenter.BoundText) & "))"
        Else
            temSql = temSql & "WHERE (((tblExpence.ExpenceDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND ((tblExpence.Deleted) = 0))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfExpenceValue) = False Then
                txtOtherExpences.Text = Format(!SumOfExpenceValue, "0.00")
            End If
        End If
        .Close
    End With
End Sub

Private Sub CalculateProfit()
    txtProfit.Text = Format(Val(txtIncome.Text) - Val(txtOtherExpences.Text) - Val(txtMilkPayments.Text), "0.00")
End Sub
