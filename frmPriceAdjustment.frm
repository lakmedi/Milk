VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmPriceAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price Adjustments"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
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
   ScaleHeight     =   9210
   ScaleWidth      =   11430
   Begin btButtonEx.ButtonEx btnUpdateDate 
      Height          =   375
      Left            =   7080
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "Update Date"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   7320
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   3201
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "New Prices"
      TabPicture(0)   =   "frmPriceAdjustment.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(5)=   "btnAdd"
      Tab(0).Control(6)=   "txtPrice"
      Tab(0).Control(7)=   "txtFATFrom"
      Tab(0).Control(8)=   "txtFATTo"
      Tab(0).Control(9)=   "txtSNFFrom"
      Tab(0).Control(10)=   "txtSNFTo"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Increase Prices"
      TabPicture(1)   =   "frmPriceAdjustment.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(3)=   "Label13"
      Tab(1).Control(4)=   "Label14"
      Tab(1).Control(5)=   "btnAAdd"
      Tab(1).Control(6)=   "txtAPrice"
      Tab(1).Control(7)=   "txtAFATFrom"
      Tab(1).Control(8)=   "txtAFATTo"
      Tab(1).Control(9)=   "txtASNFFrom"
      Tab(1).Control(10)=   "txtASNFTo"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Copy"
      TabPicture(2)   =   "frmPriceAdjustment.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label15"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label16"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "btnCopy"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmbCPriceCycle"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmbCPaymentScheme"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.TextBox txtASNFTo 
         Height          =   360
         Left            =   -71880
         TabIndex        =   30
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtASNFFrom 
         Height          =   360
         Left            =   -73200
         TabIndex        =   29
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtAFATTo 
         Height          =   360
         Left            =   -71880
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtAFATFrom 
         Height          =   360
         Left            =   -73200
         TabIndex        =   27
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtAPrice 
         Height          =   360
         Left            =   -69360
         TabIndex        =   26
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtSNFTo 
         Height          =   360
         Left            =   -71760
         TabIndex        =   19
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtSNFFrom 
         Height          =   360
         Left            =   -73080
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtFATTo 
         Height          =   360
         Left            =   -71760
         TabIndex        =   17
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtFATFrom 
         Height          =   360
         Left            =   -73080
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtPrice 
         Height          =   360
         Left            =   -69240
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin btButtonEx.ButtonEx btnAdd 
         Height          =   375
         Left            =   -67920
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "&Add"
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
      Begin btButtonEx.ButtonEx btnAAdd 
         Height          =   375
         Left            =   -68040
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "&Add"
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
      Begin MSDataListLib.DataCombo cmbCPaymentScheme 
         Height          =   360
         Left            =   2040
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   600
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCPriceCycle 
         Height          =   360
         Left            =   2040
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1080
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin btButtonEx.ButtonEx btnCopy 
         Height          =   375
         Left            =   7560
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "&Copy"
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
      Begin VB.Label Label16 
         Caption         =   "Payment Scheme"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label15 
         Caption         =   "Price Cycle"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label14 
         Caption         =   "SNF"
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label13 
         Caption         =   "FAT"
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label12 
         Caption         =   "From"
         Height          =   255
         Left            =   -73200
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "To"
         Height          =   255
         Left            =   -71880
         TabIndex        =   33
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Price Increment"
         Height          =   495
         Left            =   -70440
         TabIndex        =   32
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "SNF"
         Height          =   255
         Left            =   -74760
         TabIndex        =   25
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "FAT"
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "From"
         Height          =   255
         Left            =   -73080
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "To"
         Height          =   255
         Left            =   -71760
         TabIndex        =   22
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "New Price"
         Height          =   255
         Left            =   -70200
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.TextBox txtPCID 
      Height          =   360
      Left            =   1680
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPSID 
      Height          =   360
      Left            =   480
      TabIndex        =   12
      Top             =   1215
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtCellText 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   9600
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid gridPrice 
      Height          =   5655
      Left            =   480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1560
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9975
      _Version        =   393216
      ScrollTrack     =   -1  'True
   End
   Begin MSDataListLib.DataCombo cmbPaymentScheme 
      Height          =   360
      Left            =   1920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   10080
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   8640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin btButtonEx.ButtonEx btnFill 
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Fill"
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
   Begin MSDataListLib.DataCombo cmbPriceCycle 
      Height          =   360
      Left            =   1920
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   720
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   16515075
      CurrentDate     =   39877
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   16515075
      CurrentDate     =   39877
   End
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   375
      Left            =   8520
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Enabled         =   0   'False
      BorderColor     =   16711680
      Caption         =   "&Excel"
      Enabled         =   0   'False
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   8520
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Delete"
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
      Caption         =   "Price Cycle"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "FAT"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "SNF"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Payment Scheme"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmPriceAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTemPrice As New ADODB.Recordset
    Dim rsPrice  As New ADODB.Recordset
    Dim temSQL As String
    Dim temRow As Long
    Dim temCol As Long
    Dim temText As String
    Dim temCellText As String
    Dim temBoxText As String
    
    
Private Sub btnAAdd_Click()
    If IsNumeric(txtASNFFrom.Text) = False Then
        Exit Sub
    End If
    If IsNumeric(txtASNFTo.Text) = False Then
        Exit Sub
    End If
    If IsNumeric(txtAFATFrom.Text) = False Then
        Exit Sub
    End If
    If Val(txtAFATFrom.Text) > Val(txtAFATTo.Text) Then
        Exit Sub
    End If
    If Val(txtASNFFrom.Text) > Val(txtASNFTo.Text) Then
        Exit Sub
    End If
    Dim temSNF As Double
    Dim temFAT As Double
    For temSNF = Val(txtASNFFrom.Text) To Val(txtASNFTo.Text) Step 0.1
        For temFAT = Val(txtAFATFrom.Text) To Val(txtAFATTo.Text) Step 0.1
            With rsTemPrice
                If .State = 1 Then .Close
                temSQL = "Select * from tblPrice where PriceCycleID = " & Val(cmbPriceCycle.BoundText) & " AND PaymentSchemeID = " & Val(cmbPaymentScheme.BoundText) & " AND SNF = " & Format(temSNF, "0.0") & " AND FAT = " & Format(temFAT, "0.0")
                .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !Price = !Price + Val(txtAPrice.Text)
                Else
                    .AddNew
                    !PriceCycleID = Val(cmbPriceCycle.BoundText)
                    !PaymentSchemeID = Val(cmbPaymentScheme.BoundText)
                    !SNF = Format(temSNF, "0.0")
                    !FAT = Format(temFAT, "0.0")
                    !FromDate = dtpFrom.Value
                    !ToDate = dtpTo.Value
                    !Price = Val(txtAPrice.Text)
                End If
                .Update
            End With
        Next
    Next
Call FillGrid

End Sub

Private Sub btnAdd_Click()
    If IsNumeric(txtSNFFrom.Text) = False Then
        Exit Sub
    End If
    If IsNumeric(txtSNFTo.Text) = False Then
        Exit Sub
    End If
    If IsNumeric(txtFATFrom.Text) = False Then
        Exit Sub
    End If
    If Val(txtFATFrom.Text) > Val(txtFATTo.Text) Then
        Exit Sub
    End If
    If Val(txtSNFFrom.Text) > Val(txtSNFTo.Text) Then
        Exit Sub
    End If
    Dim temSNF As Double
    Dim temFAT As Double
    For temSNF = Val(txtSNFFrom.Text) To Val(txtSNFTo.Text) Step 0.1
        For temFAT = Val(txtFATFrom.Text) To Val(txtFATTo.Text) Step 0.1
            With rsTemPrice
                If .State = 1 Then .Close
                temSQL = "Select * from tblPrice where PriceCycleID = " & Val(cmbPriceCycle.BoundText) & " AND PaymentSchemeID = " & Val(cmbPaymentScheme.BoundText) & " AND SNF = " & Format(temSNF, "0.0") & " AND FAT = " & Format(temFAT, "0.0")
                .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !Price = Val(txtPrice.Text)
                Else
                    .AddNew
                    !PriceCycleID = Val(cmbPriceCycle.BoundText)
                    !PaymentSchemeID = Val(cmbPaymentScheme.BoundText)
                    !SNF = Format(temSNF, "0.0")
                    !FAT = Format(temFAT, "0.0")
                    !FromDate = dtpFrom.Value
                    !ToDate = dtpTo.Value
                    !Price = Val(txtPrice.Text)
                End If
                .Update
            End With
        Next
    Next
Call FillGrid
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub FormatGrid()
    Dim i As Integer
    With gridPrice
        .Rows = 58
        .Cols = 52
        For i = 0 To 50
            .TextMatrix(0, i + 1) = Empty
        Next i
        For i = 0 To 55
            .TextMatrix(i + 1, 0) = Empty
        Next i
        For i = 0 To 50
            .TextMatrix(0, i + 1) = Format((6 + 0.1 * i), "0.0")
        Next i
        For i = 0 To 55
            .TextMatrix(i + 1, 0) = Format((2.5 + 0.1 * i), "0.00")
        Next i
    End With
End Sub

Private Sub btnCopy_Click()
    Dim i As Integer
    i = MsgBox("Are you sure?", vbYesNo)
    If i = vbNo Then Exit Sub
    If IsNumeric(cmbPaymentScheme.BoundText) = False Then
        MsgBox "Please select a Payment Scheme"
        cmbPaymentScheme.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbCPriceCycle.BoundText) = False Then
        MsgBox "Please select a Price Cycle"
        cmbCPriceCycle.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbCPaymentScheme.BoundText) = False Then
        MsgBox "Please select a Payment Scheme"
        cmbCPaymentScheme.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPriceCycle.BoundText) = False Then
        MsgBox "Please select a Price Cycle"
        cmbPriceCycle.SetFocus
        Exit Sub
    End If
    If Val(cmbPaymentScheme.BoundText) = Val(cmbCPaymentScheme.BoundText) And Val(cmbPriceCycle.BoundText) = Val(cmbCPriceCycle.BoundText) Then
        MsgBox "Both can not be the same"
        cmbPriceCycle.SetFocus
        Exit Sub
    End If
        
    Dim rsTem As New ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    
    With rsTemPrice
        If .State = 1 Then .Close
        temSQL = "DELETE  from tblPrice where PriceCycleID = " & Val(cmbPriceCycle.BoundText) & " AND PaymentSchemeID = " & Val(cmbPaymentScheme.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblPrice where PriceCycleID = " & Val(cmbCPriceCycle.BoundText) & " AND PaymentSchemeID = " & Val(cmbCPaymentScheme.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If rsTem.State = 1 Then rsTem.Close
            temSQL = "SELECT * from tblPrice where PriceID = 0 "
            rsTem.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            rsTem.AddNew
            rsTem!PriceCycleID = Val(cmbPriceCycle.BoundText)
            rsTem!PaymentSchemeID = Val(cmbPaymentScheme.BoundText)
            rsTem!SNF = !SNF
            rsTem!FAT = !FAT
            rsTem!FromDate = dtpFrom.Value
            rsTem!ToDate = dtpTo.Value
            rsTem!Price = !Price
            rsTem.Update
            rsTem.Close
            .MoveNext
        Wend
        .Close
    End With
    
    btnFill_Click
    
    MsgBox "OK"
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete?", vbYesNo)
    
    If i = vbNo Then Exit Sub
    If IsNumeric(cmbPaymentScheme.BoundText) = False Then
        MsgBox "Please select a Payment Scheme"
        cmbPaymentScheme.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPriceCycle.BoundText) = False Then
        MsgBox "Please select a Price Cycle"
        cmbPriceCycle.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    With rsTemPrice
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblPrice where PriceCycleID = " & Val(cmbPriceCycle.BoundText) & " AND PaymentSchemeID = " & Val(cmbPaymentScheme.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            .Delete adAffectCurrent
            .MoveNext
        Wend
        If .State = 1 Then .Close
    End With
    
    MsgBox "OK"
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub btnExcel_Click()
    GridToExcel gridPrice, cmbPaymentScheme.Text, cmbPriceCycle.Text
End Sub

Private Sub btnFill_Click()
        
    If IsNumeric(cmbPaymentScheme.BoundText) = False Then
        MsgBox "Please select a Payment Scheme"
        cmbPaymentScheme.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPriceCycle.BoundText) = False Then
        MsgBox "Please select a Price Cycle"
        cmbPriceCycle.SetFocus
        Exit Sub
    End If
    
    txtPCID.Text = cmbPriceCycle.BoundText
    txtPSID.Text = cmbPaymentScheme.BoundText
    
    Screen.MousePointer = vbHourglass
    DoEvents
    Call FillGrid
    gridPrice.col = 0
    gridPrice.row = 0
    gridPrice.col = 1
    gridPrice.row = 1
    Screen.MousePointer = vbDefault
    btnExcel.Enabled = True
End Sub

Private Sub btnUpdateDate_Click()
    Dim i As Integer
    
    If i = vbNo Then Exit Sub
    If IsNumeric(cmbPaymentScheme.BoundText) = False Then
        MsgBox "Please select a Payment Scheme"
        cmbPaymentScheme.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPriceCycle.BoundText) = False Then
        MsgBox "Please select a Price Cycle"
        cmbPriceCycle.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    With rsTemPrice
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblPrice where PriceCycleID = " & Val(cmbPriceCycle.BoundText) & " AND PaymentSchemeID = " & Val(cmbPaymentScheme.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            !FromDate = dtpFrom.Value
            !ToDate = dtpTo.Value
            .Update
            .MoveNext
        Wend
        .Close
    End With
    
    MsgBox "OK"
    
    Screen.MousePointer = vbDefault
    

End Sub

Private Sub cmbPaymentScheme_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnFill.SetFocus
    End If
End Sub


Private Sub cmbPriceCycle_Change()
    Dim rsPS As New ADODB.Recordset
    With rsPS
        If .State = 1 Then .Close
        temSQL = "Select * from tblPriceCycle where PriceCycleID = " & Val(cmbPriceCycle.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            dtpFrom.Value = !FromDate
            dtpTo.Value = !ToDate
        End If
        .Close
    End With
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FormatGrid
    txtCellText.Visible = False
End Sub

Private Sub FillCombos()
    Dim PaymentMethod As New clsFillCombos
    PaymentMethod.FillAnyCombo cmbPaymentScheme, "PaymentScheme", True
    Dim PC As New clsFillCombos
    PC.FillAnyCombo cmbPriceCycle, "PriceCycle", True
    Dim CPaymentMethod As New clsFillCombos
    CPaymentMethod.FillAnyCombo cmbCPaymentScheme, "PaymentScheme", True
    Dim CPC As New clsFillCombos
    CPC.FillAnyCombo cmbCPriceCycle, "PriceCycle", True

End Sub

Private Sub FillGrid()
    If IsNumeric(cmbPaymentScheme.BoundText) = False Then Exit Sub
    Dim rsPrice As New ADODB.Recordset
    gridPrice.Visible = False
    Dim row As Integer
    Dim col As Integer
    Dim temFAT As Double
    Dim temSNF As Double
    With gridPrice
'        .Clear
        If rsPrice.State = 1 Then rsPrice.Close
        For row = 1 To .Rows - 1
            For col = 1 To .Cols - 1
                temSNF = Val(.TextMatrix(0, col))
                temFAT = Val(.TextMatrix(row, 0))
                temSQL = "Select * from tblPrice where SNF = " & temSNF & " And FAT = " & temFAT & " And PaymentSchemeID = " & Val(txtPSID.Text) & " AND PriceCycleID = " & Val(txtPCID.Text)
                rsPrice.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                If rsPrice.RecordCount > 0 Then
                    If IsNull(rsPrice!Price) = False Then
                        .TextMatrix(row, col) = Format(rsPrice!Price, "0.00")
                    Else
                        .TextMatrix(row, col) = Format(0, "0.00")
                    End If
                End If
                rsPrice.Close
            Next

        Next
    End With
    gridPrice.Visible = True
End Sub


Private Sub gridPrice_EnterCell()
    temRow = gridPrice.row
    temCol = gridPrice.col
    temCellText = gridPrice.TextMatrix(temRow, temCol)
    txtCellText.Top = gridPrice.Top + gridPrice.CellTop
    txtCellText.Left = gridPrice.Left + gridPrice.CellLeft
    txtCellText.Height = gridPrice.CellHeight
    txtCellText.Width = gridPrice.CellWidth
    txtCellText.Text = temCellText
    txtCellText.Visible = True
    txtCellText.SetFocus
    On Error Resume Next
    SendKeys "{Home}+{end}"
End Sub

Private Sub gridPrice_LeaveCell()
    temBoxText = txtCellText.Text
    Dim temFAT As Double
    Dim temSNF As Double
    Dim temPrice As Double
    temFAT = Val(gridPrice.TextMatrix(temRow, 0))
    temSNF = Val(gridPrice.TextMatrix(0, temCol))
    temPrice = Val(temBoxText)
    If temBoxText <> temCellText Then
        gridPrice.TextMatrix(temRow, temCol) = temBoxText
        With rsPrice
            If .State = 1 Then .Close
            temSQL = "Select * from tblPrice where SNF = " & temSNF & " And FAT = " & temFAT & " And PaymentSchemeID = " & Val(txtPSID.Text) & " AND PriceCycleID = " & Val(txtPCID.Text)
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Price = temPrice
                !FromDate = dtpFrom.Value
                !ToDate = dtpTo.Value
                .Update
            Else
                .AddNew
                !Price = temPrice
                !SNF = temSNF
                !FAT = temFAT
                !PriceCycleID = Val(txtPCID.Text)
                !PaymentSchemeID = (txtPSID.Text)
                !FromDate = dtpFrom.Value
                !ToDate = dtpTo.Value
                .Update
            End If
            .Close
        End With
    End If
End Sub

Private Sub gridPrice_Scroll()
    txtCellText.Visible = False
End Sub

Private Sub txtCellText_KeyDown(KeyCode As Integer, Shift As Integer)
    With gridPrice
        If KeyCode = vbKeyReturn Then
            If temCol < .Cols - 1 Then
                .col = temCol + 1
            Else
                .col = 1
                .row = temRow + 1
            End If
        ElseIf KeyCode = vbKeyEscape Then
            txtCellText.Text = temText
        ElseIf KeyCode = vbKeyTab Then
            If temCol < .Cols - 1 Then
                .col = temCol + 1
            Else
                .col = 1
                .row = temRow + 1
            End If
        ElseIf KeyCode = vbKeyUp Then
            If temRow > 1 Then
                .row = temRow - 1
            End If
        ElseIf KeyCode = vbKeyDown Then
            If temRow < .Rows - 1 Then
                .row = temRow + 1
            End If
        ElseIf KeyCode = vbKeyLeft Then
            If temCol > 1 Then
                .col = temCol - 1
            End If
        ElseIf KeyCode = vbKeyRight Then
            If temCol < .Cols - 1 Then
                .col = temCol + 1
            End If
        End If
    End With
End Sub
