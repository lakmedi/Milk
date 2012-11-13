VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPrintingSlips 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing Slips"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9225
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
   ScaleHeight     =   6840
   ScaleWidth      =   9225
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   6360
      TabIndex        =   14
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.ListBox lstPrintID 
      Height          =   4380
      Left            =   8280
      MultiSelect     =   2  'Extended
      TabIndex        =   10
      Top             =   1320
      Width           =   375
   End
   Begin VB.ListBox lstAllID 
      Height          =   4380
      Left            =   4200
      MultiSelect     =   2  'Extended
      TabIndex        =   8
      Top             =   1320
      Width           =   375
   End
   Begin VB.ListBox lstPrint 
      Height          =   4380
      Left            =   5280
      MultiSelect     =   2  'Extended
      TabIndex        =   11
      Top             =   1320
      Width           =   3375
   End
   Begin VB.ListBox lstAll 
      Height          =   4380
      Left            =   1080
      MultiSelect     =   2  'Extended
      TabIndex        =   9
      Top             =   1320
      Width           =   3495
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   6000
      Width           =   4095
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   6360
      Width           =   4095
   End
   Begin MSDataListLib.DataCombo cmbCC 
      Height          =   360
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
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
   Begin MSDataListLib.DataCombo cmbPayments 
      Height          =   360
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
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
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   2880
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
      Left            =   4680
      TabIndex        =   13
      Top             =   3360
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   7800
      TabIndex        =   15
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Close"
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
      TabIndex        =   7
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Payment Period"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmPrintingSlips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    Dim rsPayments As New ADODB.Recordset
    
    Dim CSetPrinter As New cSetDfltPrinter
    
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
    
Private Sub FillCombos()
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCC, "CollectingCenter", True
End Sub

Private Sub GetSettings()
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, "Printer", "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, "Paper", "")
End Sub


Private Sub ListPrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub ListPapers()
    cmbPaper.Clear
    CSetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
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
                cmbPaper.AddItem PtrCtoVbString(.pName)
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub FillPayments(ByVal CCID As Long)
    temSQL = "SELECT ('From ' +  Format([tblCollectingCenterPaymentSummery].[FromDate],'dd mmmm yyyy',1,1) + ' To ' +  Format([tblCollectingCenterPaymentSummery].[ToDate],'dd mmmm yyyy',1,1)) AS Display, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID, tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID " & _
                "From tblCollectingCenterPaymentSummery " & _
                "Where (((tblCollectingCenterPaymentSummery.CollectingCenterID) = " & CCID & "))" & _
                "ORDER BY tblCollectingCenterPaymentSummery.CollectingCenterPaymentSummeryID DESC"
    With rsPayments
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbPayments
        Set .RowSource = rsPayments
        .ListField = "Display"
        .BoundColumn = "CollectingCenterPaymentSummeryID"
    End With
End Sub

Private Sub btnAdd_Click()
    Dim i As Integer
    Dim ii As Integer
    Dim AlreadySelected As Boolean
    For i = 0 To lstAllID.ListCount - 1
        If lstAll.Selected(i) = True Then
            AlreadySelected = False
            For ii = 0 To lstPrintID.ListCount - 1
                If Val(lstPrintID.List(ii)) = Val(lstAllID.List(i)) Then
                    AlreadySelected = True
                    MsgBox "Supplier " & lstAll.List(i) & " is already listed"
                End If
            Next
            If AlreadySelected = False Then
                lstPrint.AddItem lstAll.List(i)
                lstPrintID.AddItem lstAllID.List(i)
            End If
        End If
    Next
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnRemove_Click()
    Dim i As Integer
    For i = lstPrint.ListCount - 1 To 0 Step -1
        If lstPrint.Selected(i) = True Then
            lstPrint.RemoveItem (i)
            lstPrintID.RemoveItem (i)
        End If
    Next
End Sub

Private Sub cmbCC_Change()
    Call FillPayments(Val(cmbCC.BoundText))
End Sub

Private Sub cmbPayments_Change()
    Call FillLists
End Sub

Private Sub FillLists()
    
    Dim rsSuppliers As New ADODB.Recordset
    
    lstAll.Clear
    lstAllID.Clear
    
    lstPrint.Clear
    lstPrintID.Clear
    
    With rsSuppliers
        If .State = 1 Then .Close
        temSQL = "SELECT tblSupplier.Supplier, tblSupplier.SupplierCode, tblSupplierPayments.Value, tblSupplierPayments.SupplierPaymentsID " & _
                    "FROM (tblSupplierPayments LEFT JOIN tblSupplier ON tblSupplierPayments.SupplierID = tblSupplier.SupplierID) LEFT JOIN tblPaymentMethod ON tblSupplierPayments.PrintedPaymentMethodID = tblPaymentMethod.PaymentMethodID " & _
                    "WHERE (((tblPaymentMethod.PaymentMethod)='To Bank') AND ((tblSupplierPayments.CollectingCenterPaymentSummeryID)=" & Val(cmbPayments.BoundText) & ")) " & _
                    "ORDER BY tblSupplier.Supplier"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                lstAll.AddItem !Supplier & " (" & !SupplierCode & ")"
                lstAllID.AddItem !SupplierPaymentsID
                .MoveNext
            Wend
        End If
        .Close
    End With
    
End Sub

Private Sub cmbPrinter_Change()
    Call ListPapers
End Sub

Private Sub cmbPrinter_Click()
    Call ListPapers
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call ListPrinters
    Call GetSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, Me.Name, "Printer", cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, "Paper", cmbPaper.Text
End Sub
