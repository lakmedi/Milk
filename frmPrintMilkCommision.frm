VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPrintMilkCommision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Milk Commsion"
   ClientHeight    =   7920
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
   ScaleHeight     =   7920
   ScaleWidth      =   10050
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8760
      TabIndex        =   16
      Top             =   7320
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
   Begin VB.ListBox lstPrintIDs 
      Height          =   5580
      Left            =   8520
      MultiSelect     =   2  'Extended
      TabIndex        =   15
      Top             =   1320
      Width           =   375
   End
   Begin VB.ListBox lstAllIDs 
      Height          =   5580
      Left            =   3840
      MultiSelect     =   2  'Extended
      TabIndex        =   14
      Top             =   1320
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
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   61145091
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
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   61145091
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
      Top             =   7320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
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
Attribute VB_Name = "frmPrintMilkCommision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim CSetPrinter As New cSetDfltPrinter
    Dim temSql As String
    Dim rsMilk As New ADODB.Recordset
    
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

Private Sub btnPrint_Click()
    Dim MySupplier As New clsSupplier
    CSetPrinter.SetPrinterAsDefault (BillPrinterName)
    Dim TemResponce As Long
    Dim RetVal As Integer
    RetVal = SelectForm(BillPaperName, Me.hwnd)
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
        Case FORM_SELECTED   ' 1
        
            Dim i As Integer
            For i = 0 To lstPrint.ListCount - 1
                MySupplier.ID = Val(lstPrintIDs.List(i))
                With rsMilk
                    If .State = 1 Then .Close
                    temSql = "SELECT tblCollection.Date, tblCollection.LMR, tblCollection.FAT, tblCollection.Liters, tblCollection.SNF, tblCollection.Commision, tblCollection.commisionRate, tblCollection.Price, tblCollection.Value " & _
                                "From tblCollection " & _
                                "WHERE (((tblCollection.Date) Between #" & Format(dtpFrom.Value, "Dd MMMM yyyy") & "# And #" & Format(dtpTo.Value, "dd MMMM yyyy") & "#) AND ((tblCollection.SupplierID)=" & Val(lstPrintIDs.List(i)) & ") AND ((tblCollection.SecessionID)=2) AND ((tblCollection.Deleted)=False)) " & _
                                "ORDER BY tblCollection.Date"
                    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                End With
                If rsMilk.RecordCount > 0 Then
                    If MySupplier.HasCommision = True Then
                        With dtrIndividualWithCommission
                            MySupplier.ID = Val(lstPrintIDs.List(i))
                            Set .DataSource = rsMilk
                            .Sections("Section2").Controls("lblName").Caption = MySupplier.Name
                            .Sections("Section2").Controls("lblAddress").Caption = MySupplier.Address
                            .Sections("Section2").Controls("lblMonth").Caption = Format(dtpFrom.Value, "MMMM")
                            .Sections("Section2").Controls("lblCode").Caption = MySupplier.Code
                            .Sections("Section2").Controls("lblSecession").Caption = "EVENING"
                            
                            .Sections("Section3").Controls("lblMilkPayment").Caption = MySupplier.Address
                            .Sections("Section3").Controls("lblMonth").Caption = Format(dtpFrom.Value, "MMMM")
                            .Sections("Section3").Controls("lblCode").Caption = MySupplier.Code
                            .Sections("Section3").Controls("lblSecession").Caption = "EVENING"
                            
                            .Refresh
'                            .Show
'                            MsgBox "OK"
                            .PrintReport False
                        End With
                    Else
                        With dtrIndividualMilkPaymentAdvice
                            MySupplier.ID = Val(lstPrintIDs.List(i))
                            Set .DataSource = rsMilk
                            .Sections("Section2").Controls("lblName").Caption = MySupplier.Name
                            .Sections("Section2").Controls("lblAddress").Caption = MySupplier.Address
                            .Sections("Section2").Controls("lblMonth").Caption = Format(dtpFrom.Value, "MMMM")
                            .Sections("Section2").Controls("lblCode").Caption = MySupplier.Code
                            .Sections("Section2").Controls("lblSecession").Caption = "EVENING"
                            .Refresh
'                            .Show
'                            MsgBox "OK"
                            .PrintReport False
                        End With
                    End If
                End If
                    
                With rsMilk
                    If .State = 1 Then .Close
                    temSql = "SELECT tblCollection.Date, tblCollection.LMR, tblCollection.FAT, tblCollection.Liters, tblCollection.SNF, tblCollection.Commision, tblCollection.commisionRate, tblCollection.Price, tblCollection.Value " & _
                                "From tblCollection " & _
                                "WHERE (((tblCollection.Date) Between #" & Format(dtpFrom.Value, "Dd MMMM yyyy") & "# And #" & Format(dtpTo.Value, "dd MMMM yyyy") & "#) AND ((tblCollection.SupplierID)=" & Val(lstPrintIDs.List(i)) & ") AND ((tblCollection.SecessionID)=1) AND ((tblCollection.Deleted)=False)) " & _
                                "ORDER BY tblCollection.Date"
                    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                End With
                If rsMilk.RecordCount > 0 Then
                    If MySupplier.HasCommision = True Then
                        With dtrIndividualWithCommission
                            Set .DataSource = rsMilk
                            .Sections("Section2").Controls("lblName").Caption = MySupplier.Name
                            .Sections("Section2").Controls("lblAddress").Caption = MySupplier.Address
                            .Sections("Section2").Controls("lblMonth").Caption = Format(dtpFrom.Value, "MMMM")
                            .Sections("Section2").Controls("lblCode").Caption = MySupplier.Code
                            .Sections("Section2").Controls("lblSecession").Caption = "MORNING"
                            .Refresh
'                            .Show
'                            MsgBox "OK"
                            .PrintReport False
                        End With
                    Else
                        With dtrIndividualMilkPaymentAdvice
                            Set .DataSource = rsMilk
                            .Sections("Section2").Controls("lblName").Caption = MySupplier.Name
                            .Sections("Section2").Controls("lblAddress").Caption = MySupplier.Address
                            .Sections("Section2").Controls("lblMonth").Caption = Format(dtpFrom.Value, "MMMM")
                            .Sections("Section2").Controls("lblCode").Caption = MySupplier.Code
                            .Sections("Section2").Controls("lblSecession").Caption = "MORNING"
                            .Refresh
'                            .Show
'                            MsgBox "OK"
                            .PrintReport False
                        End With
                    End If
                End If
                    
                    
            Next i
        
        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
    End Select

    
    
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

Private Sub Form_Load()
    Call FillCombos

End Sub

Private Sub FillCombos()
    Dim Centers As New clsFillCombos
    Centers.FillAnyCombo cmbCollectingCenter, "CollectingCenter", True
End Sub


