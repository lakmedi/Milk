VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPrintingPreferances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing Preferances"
   ClientHeight    =   4710
   ClientLeft      =   4440
   ClientTop       =   1680
   ClientWidth     =   8205
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
   ScaleHeight     =   4710
   ScaleWidth      =   8205
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "Receipt Printer"
      TabPicture(0)   =   "frmPrintingPreferances.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Report Printer"
      TabPicture(1)   =   "frmPrintingPreferances.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Add New Form"
      TabPicture(2)   =   "frmPrintingPreferances.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frameAddForm"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Other"
      TabPicture(3)   =   "frmPrintingPreferances.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).ControlCount=   0
      Begin VB.Frame frameAddForm 
         Caption         =   "Add New Form"
         Height          =   3375
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   7695
         Begin VB.TextBox txtFormWidth 
            Height          =   360
            Left            =   5520
            TabIndex        =   21
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtFormHeight 
            Height          =   360
            Left            =   2520
            TabIndex        =   20
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtFormName 
            Height          =   360
            Left            =   1800
            TabIndex        =   19
            Top             =   360
            Width           =   5055
         End
         Begin VB.TextBox txtFormTop 
            Height          =   360
            Left            =   2520
            TabIndex        =   18
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtFormBottom 
            Height          =   360
            Left            =   5520
            TabIndex        =   17
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtFormLeft 
            Height          =   360
            Left            =   2520
            TabIndex        =   16
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtFormRight 
            Height          =   360
            Left            =   5520
            TabIndex        =   15
            Top             =   1920
            Width           =   1335
         End
         Begin btButtonEx.ButtonEx bttnAddForm 
            Height          =   495
            Left            =   5040
            TabIndex        =   22
            Top             =   2520
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            Appearance      =   3
            BorderColor     =   16384
            Caption         =   "Add Form"
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
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   255
            Left            =   6960
            TabIndex        =   35
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   255
            Left            =   3960
            TabIndex        =   34
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Width"
            Height          =   375
            Left            =   4920
            TabIndex        =   33
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Height"
            Height          =   375
            Left            =   1800
            TabIndex        =   32
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Form Name"
            Height          =   375
            Left            =   480
            TabIndex        =   31
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Top"
            Height          =   375
            Left            =   1800
            TabIndex        =   30
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Bottom"
            Height          =   375
            Left            =   4920
            TabIndex        =   29
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   255
            Left            =   3960
            TabIndex        =   28
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   255
            Left            =   6960
            TabIndex        =   27
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Left"
            Height          =   375
            Left            =   1800
            TabIndex        =   26
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Right"
            Height          =   375
            Left            =   4920
            TabIndex        =   25
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   255
            Left            =   3960
            TabIndex        =   24
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   255
            Left            =   6960
            TabIndex        =   23
            Top             =   1920
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Report Printer"
         Height          =   3375
         Left            =   -74880
         TabIndex        =   8
         Top             =   360
         Width           =   7695
         Begin VB.ComboBox ComboReportPrinter 
            Height          =   360
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   360
            Width           =   6615
         End
         Begin VB.ListBox ListReportPrinterPapers 
            Height          =   2220
            Left            =   960
            TabIndex        =   9
            Top             =   840
            Width           =   6615
         End
         Begin VB.ListBox ListReportPrinterPapers1 
            Height          =   300
            Left            =   960
            TabIndex        =   10
            Top             =   1680
            Width           =   5295
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Printer"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Paper"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Bill Printing"
         Height          =   3375
         Left            =   -74880
         TabIndex        =   2
         Top             =   360
         Width           =   7695
         Begin VB.ComboBox ComboBillPrinter 
            Height          =   360
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   360
            Width           =   6615
         End
         Begin VB.ListBox ListBillPrinterPapers 
            Height          =   2220
            Left            =   960
            TabIndex        =   3
            Top             =   840
            Width           =   6615
         End
         Begin VB.ListBox ListBillPrinterPapers1 
            Height          =   300
            Left            =   960
            TabIndex        =   4
            Top             =   2160
            Width           =   3135
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "&Printer"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "P&aper"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   3015
         End
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   4080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Save And Exit"
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
End
Attribute VB_Name = "frmPrintingPreferances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
    Dim SuppliedWord As String
    Dim FSys As New Scripting.FileSystemObject
    Private CSetPrinter As New cSetDfltPrinter
    
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

Private Sub Form_Load()
    Dim ingRet As Long
    Dim TabPrinter(2) As Long
'    TabPrinter(0) = 48
'    TabPrinter(1) = 78
'    ingRet = SendMessage(ListBillPrinterPapers.hwnd, LB_SETTABSTOPS, 2, TabPrinter(0))
'    ingRet = SendMessage(ListReportPrinterPapers.hwnd, LB_SETTABSTOPS, 2, TabPrinter(0))
    Call PopulatePrinters
    Call PopulateBillPrinterPapers
    Call PopulateReportPrinterPapers
    Call SetPreferances
    SSTab1.Tab = 0
End Sub

Private Sub PopulatePrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        ComboBillPrinter.AddItem MyPrinter.DeviceName
        ComboReportPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub PopulateBillPrinterPapers()
    ListBillPrinterPapers.Clear
    ListBillPrinterPapers1.Clear
    SetPrinter = False
    CSetPrinter.SetPrinterAsDefault (BillPrinterName)
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
                'FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
                'ComboBillPrinterPapers.AddItem FormItem
                ListBillPrinterPapers1.AddItem PtrCtoVbString(.pName)
                ListBillPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub PopulateReportPrinterPapers()
    ListReportPrinterPapers.Clear
    ListReportPrinterPapers1.Clear
    SetPrinter = False
    CSetPrinter.SetPrinterAsDefault (ReportPaperName)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        With FormSize
            .cx = ReportPaperWidth
            .cy = ReportPaperHeight
        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                'FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
                'ComboReportPrinterPapers.AddItem FormItem
                ListReportPrinterPapers1.AddItem PtrCtoVbString(.pName)
                ListReportPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub bttnAddForm_Click()
    Dim TemResponce As Long
    If Trim(txtFormName.Text) = "" Then
        TemResponce = MsgBox("You have not enter a valid name for the form", vbCritical, "No name")
        txtFormName.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtFormHeight.Text) Then
        TemResponce = MsgBox("You have not entered a valid height in millimeters for the height of the form", vbCritical, "No Height")
        txtFormHeight.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtFormWidth.Text) Then
        TemResponce = MsgBox("You have not entered a valid width in millimeters for the width of the form", vbCritical, "No Width")
        txtFormWidth.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtFormTop.Text) Then
        TemResponce = MsgBox("You have not entered a valid top margin in millimeters for the height of the form", vbCritical, "No Top Margin")
        txtFormTop.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtFormBottom.Text) Then
        TemResponce = MsgBox("You have not entered a valid bottom margin in millimeters for the width of the form", vbCritical, "No Bottom Margin")
        txtFormBottom.SetFocus
        Exit Sub
    End If
    
     If Not IsNumeric(txtFormRight.Text) Then
        TemResponce = MsgBox("You have not entered a valid right margin in millimeters for the height of the form", vbCritical, "No Right Margin")
        txtFormRight.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtFormLeft.Text) Then
        TemResponce = MsgBox("You have not entered a valid left margin in millimeters for the width of the form", vbCritical, "No Left Margin")
        txtFormLeft.SetFocus
        Exit Sub
    End If
   
    
    Dim TemFormName As String
    Dim PrinterHandle As Long   ' Handle to printer
    
    If OpenPrinter(Printer.DeviceName, PrinterHandle, 0&) Then
        TemFormName = AddMyNewForm(PrinterHandle, Trim(txtFormName.Text), Val(txtFormHeight.Text) * 1000, Val(txtFormWidth.Text) * 1000, Val(txtFormBottom.Text) * 1000, Val(txtFormTop.Text) * 1000, Val(txtFormLeft.Text) * 1000, Val(txtFormRight.Text) * 1000)
        
        If TemFormName <> "none" Then
            TemResponce = MsgBox("The new form was added", vbInformation, "Added")
            Call PopulatePrinters
            Call PopulateBillPrinterPapers
            Call PopulateReportPrinterPapers
        End If
    End If
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub SetPreferances()
    Dim TemResponce As Integer
    
    On Error GoTo ErrBillPrinter
    ComboBillPrinter.Text = BillPrinterName
    
    On Error GoTo ErrReportPrinter
    ComboReportPrinter.Text = ReportPrinterName
    
    On Error GoTo ErrPrescreptionPrinter
    
    On Error GoTo ErrBillPrinterPaper
    ListBillPrinterPapers1.Text = BillPaperName
    ListBillPrinterPapers.ListIndex = ListBillPrinterPapers1.ListIndex
    
    On Error GoTo ErrReportPrinterPaper
    ListReportPrinterPapers1.Text = ReportPaperName
    ListReportPrinterPapers.ListIndex = ListReportPrinterPapers1.ListIndex
    
    On Error GoTo ErrPrescreptionPrinterPaper
    
    Exit Sub
    
    
ErrBillPrinter:
    If Err.Number = 383 Then
        TemResponce = MsgBox("The Bill printer you have selected is not available now. Please select another printer", vbCritical, "New Bill Printer")
        If ComboBillPrinter.ListCount <> 0 Then ComboBillPrinter.ListIndex = 0
    Else
        TemResponce = MsgBox("An unknown error occured. Please contact Lakmedipro (077 3177874) with following details" & vbNewLine & Err.Description & vbNewLine & Err.Number & vbNewLine & Me.Caption, vbCritical, "Unknown Error")
    End If
    Exit Sub

ErrPrescreptionPrinter:
    If Err.Number = 383 Then
        TemResponce = MsgBox("The Prescreption printer you have selected is not available now. Please select another printer", vbCritical, "New Bill Printer")
    Else
        TemResponce = MsgBox("An unknown error occured. Please contact Lakmedipro (077 3177874) with following details" & vbNewLine & Err.Description & vbNewLine & Err.Number & vbNewLine & Me.Caption, vbCritical, "Unknown Error")
    End If
    Exit Sub

ErrBillPrinterPaper:
    If Err.Number = 383 Then
        TemResponce = MsgBox("The Bill printer paper you have selected is not available now. Please select another printer", vbCritical, "New Bill Printer Paper")
        If ListBillPrinterPapers.ListCount <> 0 Then ListBillPrinterPapers.ListIndex = 0
    Else
        TemResponce = MsgBox("An unknown error occured. Please contact Lakmedipro (077 3177874) with following details" & vbNewLine & Err.Description & vbNewLine & Err.Number & vbNewLine & Me.Caption, vbCritical, "Unknown Error")
    End If
    Exit Sub


ErrReportPrinter:
    If Err.Number = 383 Then
        TemResponce = MsgBox("The Report printer you have selected is not available now. Please select another printer", vbCritical, "New Bill Printer")
        If ComboReportPrinter.ListCount <> 0 Then ComboReportPrinter.ListIndex = 0
    Else
        TemResponce = MsgBox("An unknown error occured. Please contact Lakmedipro (077 3177874) with following details" & vbNewLine & Err.Description & vbNewLine & Err.Number & vbNewLine & Me.Caption, vbCritical, "Unknown Error")
    End If
    Exit Sub
    
ErrReportPrinterPaper:
    If Err.Number = 383 Then
        TemResponce = MsgBox("The Report printer paper you have selected is not available now. Please select another printer", vbCritical, "New Bill Printer Paper")
        If ListReportPrinterPapers.ListCount <> 0 Then ListReportPrinterPapers.ListIndex = 0
    Else
        TemResponce = MsgBox("An unknown error occured. Please contact Lakmedipro (077 3177874) with following details" & vbNewLine & Err.Description & vbNewLine & Err.Number & vbNewLine & Me.Caption, vbCritical, "Unknown Error")
    End If
    Exit Sub
    
ErrPrescreptionPrinterPaper:
    If Err.Number = 383 Then
        TemResponce = MsgBox("The Prescreption printer paper you have selected is not available now. Please select another printer", vbCritical, "New Prescreption Printer Paper")
    Else
        TemResponce = MsgBox("An unknown error occured. Please contact Lakmedipro (077 3177874) with following details" & vbNewLine & Err.Description & vbNewLine & Err.Number & vbNewLine & Me.Caption, vbCritical, "Unknown Error")
    End If
    Exit Sub
    
    
    
End Sub

Private Sub SavePreferancesToFile()
    SaveSetting App.EXEName, "Options", "BillPrinterName", ComboBillPrinter.Text
    SaveSetting App.EXEName, "Options", "BillPaperName", ListBillPrinterPapers1.Text
    SaveSetting App.EXEName, "Options", "ReportPrinterName", ComboReportPrinter.Text
    SaveSetting App.EXEName, "Options", "ReportPaperName", ListReportPrinterPapers1.Text
    SaveSetting App.EXEName, "Options", "BillPrinterName", ComboBillPrinter.Text
    SaveSetting App.EXEName, "Options", "BillPaperName", ListBillPrinterPapers1.Text         ' Mid(ComboBillPrinterPapers.Text, 1, InStr(1, ComboBillPrinterPapers.Text, " -") - 1)                                          '   ComboBillPrinterPapers.Text
    SaveSetting App.EXEName, "Options", "ReportPrinterName", ComboReportPrinter.Text
    SaveSetting App.EXEName, "Options", "ReportPaperName", ListReportPrinterPapers1.Text   ' Mid(ComboReportPrinterPapers.Text, 1, InStr(1, ComboReportPrinterPapers.Text, " -") - 1)                                          '   ComboBillPrinterPapers.Text
End Sub

Private Sub SavePreferancesToMemory()
    BillPrinterName = ComboBillPrinter.Text
    BillPaperName = ListBillPrinterPapers1.Text
    ReportPrinterName = ComboReportPrinter.Text
    ReportPaperName = ListReportPrinterPapers1.Text
End Sub

Private Sub bttnPrintingPositions_Click()
'    frmPrintingPositions.Show
'    frmPrintingPositions.ZOrder 0
End Sub
Private Sub ComboBillPrinter_Change()
    CSetPrinter.SetPrinterAsDefault (ComboBillPrinter.Text)
    Call PopulateBillPrinterPapers
End Sub

Private Sub ComboReportPrinter_Change()
    CSetPrinter.SetPrinterAsDefault (ComboReportPrinter.Text)
    Call PopulateReportPrinterPapers
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SavePreferancesToFile
    Call SavePreferancesToMemory
End Sub

Private Sub ListBillPrinterPapers_Click()
    ListBillPrinterPapers1.ListIndex = ListBillPrinterPapers.ListIndex
End Sub

Private Sub ListReportPrinterPapers_Click()
    ListReportPrinterPapers1.ListIndex = ListReportPrinterPapers.ListIndex
End Sub

