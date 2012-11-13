VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmErrorDetection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error Detection"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
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
   ScaleHeight     =   6270
   ScaleWidth      =   7935
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   5640
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
   Begin MSFlexGridLib.MSFlexGrid gridError 
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6165
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbCC 
      Height          =   360
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   115998723
      CurrentDate     =   39861
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   115998723
      CurrentDate     =   39861
   End
   Begin btButtonEx.ButtonEx btnDisplay 
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Display"
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
      Caption         =   "Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmErrorDetection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDisplay_Click()
    With gridError
        .Clear
        
        .Cols = 3
        
        .Rows = 1
        
        .row = 0
        
        .col = 0
        .Text = "Date"
        
        .col = 1
        .Text = "Secession"
        
        .col = 2
        .Text = "Deleted Supplier"
        
        .ColWidth(0) = 1800
        .ColWidth(1) = 1800
        .ColWidth(2) = .Width - 3900
        
        
    End With
    
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblCollection.Date, tblCollection.Date, tblSupplier.Supplier, tblSupplier.Deleted, tblCollectingCenter.CollectingCenter, tblStaff.Staff, tblCollection.SecessionID " & _
            "FROM tblStaff RIGHT JOIN ((tblCollection LEFT JOIN tblSupplier ON tblCollection.SupplierID = tblSupplier.SupplierID) LEFT JOIN tblCollectingCenter ON tblCollection.CollectingCenterID = tblCollectingCenter.CollectingCenterID) ON tblStaff.StaffID = tblCollection.DeletedUserID " & _
            "WHERE tblCollection.Deleted = 0 AND tblCollection.ProgramDate Between '" & DTPicker1.Value & "' And '" & DTPicker2.Value & "' AND tblSupplier.Deleted = True AND tblCollection.CollectingCenterID = " & Val(cmbCC.BoundText) & " "
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridError.Rows = gridError.Rows + 1
            gridError.row = gridError.Rows - 1
            gridError.col = 0
            gridError.Text = !Date
            If !SecessionID = 1 Then
                gridError.col = 1
                gridError.Text = "Morning"
            Else
                gridError.col = 1
                gridError.Text = "Evening"
            End If
            gridError.col = 2
            gridError.Text = !Supplier
            .MoveNext
        Wend
    End With
End Sub

Private Sub Form_Load()
    Dim CC As New clsFillCombos
    CC.FillAnyCombo cmbCC, "CollectingCenter", True
End Sub
