VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmEditAdditionalPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Additional Payments"
   ClientHeight    =   7200
   ClientLeft      =   2505
   ClientTop       =   2325
   ClientWidth     =   9840
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
   ScaleHeight     =   7200
   ScaleWidth      =   9840
   Begin MSFlexGridLib.MSFlexGrid gridAP 
      Height          =   5055
      Left            =   2040
      TabIndex        =   8
      Top             =   1680
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8916
      _Version        =   393216
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
   Begin VB.TextBox txtComments 
      Height          =   1215
      Left            =   7080
      TabIndex        =   10
      Top             =   1920
      Width           =   2535
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8520
      TabIndex        =   12
      Top             =   6600
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   495
      Left            =   7080
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   20774915
      CurrentDate     =   39749
   End
   Begin MSDataListLib.DataCombo cmbCC 
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   635
      _Version        =   393216
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
   Begin MSDataListLib.DataCombo cmbSupplier 
      Height          =   360
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   635
      _Version        =   393216
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
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   20774915
      CurrentDate     =   39749
   End
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "&Supplier"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "&To"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Co&mments"
      Height          =   255
      Left            =   7080
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "&From"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frmEditAdditionalPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsSup As New ADODB.Recordset
    Dim i As Integer
    Dim rsViewSuppliers As New ADODB.Recordset
    
Private Sub FillCombos()
    Dim cc As New clsFillCombos
    cc.FillAnyCombo cmbCC, "CollectingCenter", True
End Sub
Private Sub FillGrid()
    With gridAP
        .Clear
        
        .Rows = 1
        .Cols = 5
        
        .row = 0
        
        .col = 0
        .Text = "No."
        
        .col = 1
        .Text = "Date"
        
        .col = 2
        .Text = "Comments"
        
        .col = 3
        .Text = "Value"
        
        .col = 4
        .Text = "ID"
        .ColWidth(4) = 1
        
        
    End With
    With rsSup
        If .State = 1 Then .Close
        temSql = "Select * from tblAdditionalCommision where CommisionDate between #" & Format(dtpFrom.Value, "dd MMMM yyyy") & "# And #" & Format(dtpTo.Value, "dd MMMM yyyy") & " And Deleted = false And SupplierID = " & Val(cmbSupplier.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            gridAP.Rows = .RecordCount + 1
            For i = 1 To .RecordCount
                gridAP.TextMatrix(i, 0) = i
                gridAP.TextMatrix(i, 1) = Format(!CommisionDate, ShortDateFormat)
                gridAP.TextMatrix(i, 2) = !Comments
                gridAP.TextMatrix(i, 3) = Format(!Value, "0.00")
                gridAP.TextMatrix(i, 4) = !AdditionalCommisionID
                .MoveNext
            Next
        End If
        .Close
    End With
End Sub

Private Sub btnDelete_Click()
    Dim temRow As Integer
    Dim temID As Long
    Dim rsAP As New ADODB.Recordset
    With gridAP
        temRow = .row
        If IsNumeric(.TextMatrix(temRow, 4)) = False Then
            MsgBox "Nothing to Delete"
            Exit Sub
        Else
            temID = Val(.TextMatrix(temRow, 4))
        End If
    End With
    If Trim(txtComments.Text) = "" Then
        MsgBox "Enter some comments"
        txtComments.SetFocus
        Exit Sub
    End If
    With rsAP
        If .State = 1 Then .Close
        temSql = "SELECT * from tblAdditionalCommision where AdditionalCommisionID = " & temID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedDate = Date
            !DeletedTime = Time
            !DeletedUserID = UserID
            !DeleteComments = txtComments.Text
            .Update
        End If
        .Close
    End With
    Call FillGrid
End Sub

Private Sub cmbCC_Change()
    With rsViewSuppliers
        If .State = 1 Then .Close
        temSql = "Select * from tblSupplier where farmer = true and deleted = false and CollectingCenterID = " & Val(cmbCC.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbSupplier
        Set .RowSource = rsViewSuppliers
        .ListField = "Supplier"
        .BoundColumn = "SupplierID"
    End With
End Sub
Private Sub Form_Load()
    FillCombos
    dtpFrom.Value = Date
    dtpTo.Value = Date
End Sub
