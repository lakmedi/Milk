VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmItemSuppliers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Suppliers"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   ClipControls    =   0   'False
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
   Moveable        =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   Begin VB.Frame Frame4 
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4800
      TabIndex        =   26
      Top             =   6720
      Width           =   6975
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   375
         Left            =   4920
         TabIndex        =   13
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "Ca&ncel"
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
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "&Save"
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
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "&Save"
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
   Begin VB.Frame Frame3 
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   6720
      Width           =   4575
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "&Edit"
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
      Begin btButtonEx.ButtonEx bttnAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "Item Suppliers"
      ForeColor       =   &H00FF0000&
      Height          =   6495
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   4575
      Begin MSDataListLib.DataCombo dtcItemSupplier 
         Height          =   5940
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   10478
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
         Text            =   ""
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item Supplier Details"
      ForeColor       =   &H00FF0000&
      Height          =   6495
      Left            =   4800
      TabIndex        =   15
      Top             =   120
      Width           =   6975
      Begin MSDataListLib.DataCombo dtcCity 
         Height          =   360
         Left            =   2520
         TabIndex        =   5
         Top             =   2160
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txtAddress 
         Height          =   1215
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtTelephone 
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   2640
         Width           =   4095
      End
      Begin VB.TextBox txtFax 
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   3120
         Width           =   4095
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   3600
         Width           =   4095
      End
      Begin VB.TextBox txtWebsite 
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   4080
         Width           =   4095
      End
      Begin VB.TextBox txtOther 
         Height          =   975
         Left            =   2520
         TabIndex        =   10
         Top             =   4560
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Item Supplier Name"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Item Supplier Address"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Item Supplier City"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Item Supplier Telephone"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Item Supplier Fax"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Item Supplier Email"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Item Supplier Website"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Other Details"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   4560
         Width           =   2295
      End
   End
   Begin btButtonEx.ButtonEx bttnCLose 
      Height          =   375
      Left            =   9720
      TabIndex        =   14
      Top             =   7560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
End
Attribute VB_Name = "frmItemSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsViewItemSupplier As New ADODB.Recordset
    Dim rsViewCity1 As New ADODB.Recordset
    Dim rsItemSupplier As New ADODB.Recordset
    Dim A As Integer
    Dim TemItemSupplierId As Long

Private Sub bttnCancel_Click()
    ClearValues
    BeforeAddEdit
    dtcItemSupplier.Text = Empty
    dtcItemSupplier.SetFocus
End Sub
Private Sub FillCity1()
    With rsViewCity1
        If .State = 1 Then .Close
        .Open "Select* From tblCity Order By City", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Set dtcCity.RowSource = rsViewCity1
        dtcCity.BoundColumn = "CityID"
        dtcCity.ListField = "City"
    End With
End Sub
Private Sub EditItemSupplier()
    If Trim(txtName.Text) = "" Then Call NoName: Exit Sub
'    On Error GoTo ErrorHandler
    With rsItemSupplier
        If .State = 1 Then .Close
        .Open "Select* From tblItemSupplier Where ItemSupplierID = " & TemItemSupplierId & "", cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount = 0 Then Exit Sub
        !ItemSupplier = Trim(txtName.Text)
        !ItemSupplierAddress = txtAddress.Text
        !ItemSupplierCityID = Val(dtcCity.BoundText)
        !ItemSupplierTelephone = txtTelephone.Text
        !ItemSupplierFax = txtFax.Text
        !ItemSupplierEmail = txteMail.Text
        !ItemSupplierWebsite = txtWebsite.Text
        !ItemSupplierComments = txtOther.Text
        .Update
        If .State = 1 Then .Close
        BeforeAddEdit
        ClearValues
        Call FillItemSupplier
        dtcItemSupplier.SetFocus
        dtcItemSupplier.Text = Empty
        Exit Sub
    
ErrorHandler:
        A = MsgBox(Err.Number & vbNewLine & Err.Description & vbTab & Me.Caption, vbCritical + vbOKOnly, "Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
        dtcItemSupplier.Text = Empty
        dtcItemSupplier.SetFocus
    End With
        
End Sub

Private Sub bttnChange_Click()
    Call EditItemSupplier
End Sub

Private Sub bttnSave_Click()
    Call SaveItemSupplier
End Sub

Private Sub dtcItemSupplier_Click(Area As Integer)
    If IsNumeric(dtcItemSupplier.BoundText) = False Then Exit Sub
    Call DisplaySelected
End Sub

Private Sub Form_Load()
    Call FillItemSupplier
    Call FillCity1
    Call BeforeAddEdit
    Call ClearValues
    
Select Case UserAuthorityLevel
    
    
    Case Authority.Administrator '6
    bttnEdit.Visible = True
    bttnAdd.Visible = True
    
    Case Authority.SuperUser '5
    bttnAdd.Visible = True
    bttnEdit.Visible = True
    
    Case Authority.PowerUser '4
    bttnAdd.Visible = True
    bttnEdit.Visible = True

    Case Authority.OrdinaryUser '3
    bttnAdd.Visible = True
    bttnEdit.Visible = False

    Case Else
    
    End Select
    
    If ItemSuppiersEditAllowed = True Then
       bttnEdit.Visible = True
    Else
        bttnEdit.Visible = False
    End If

End Sub

Private Sub FillItemSupplier()
    With rsViewItemSupplier
        If .State = 1 Then .Close
        .Open "Select* From tblItemSupplier Order By ItemSupplier", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Set dtcItemSupplier.RowSource = rsViewItemSupplier
        dtcItemSupplier.BoundColumn = "ItemSupplierID"
        dtcItemSupplier.ListField = "ItemSupplier"
    End With
End Sub

Private Sub FillCity11()
    With rsViewCity1
        If .State = 1 Then .Close
        .Open "Select* From tblCity11 Order By City", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Set dtcCity.RowSource = rsViewCity1
        dtcCity.BoundColumn = "CityID"
        dtcCity.ListField = "City"
    End With
End Sub

Private Sub bttnAdd_Click()
    ClearValues
    AfterAdd
    txtName.SetFocus
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    AfterEdit
    txtName.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub SaveItemSupplier()
    If Trim(txtName.Text) = "" Then Call NoName: Exit Sub
'    On Error GoTo ErrorHandler
    With rsItemSupplier
        If .State = 1 Then .Close
        .Open "Select* From tblItemSupplier where ItemSupplierID =0 ", cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ItemSupplier = Trim(txtName.Text)
        !ItemSupplierAddress = txtAddress.Text
        !ItemSupplierCityID = Val(dtcCity.BoundText)
        !ItemSupplierTelephone = txtTelephone.Text
        !ItemSupplierFax = txtFax.Text
        !ItemSupplierEmail = txteMail.Text
        !ItemSupplierWebsite = txtWebsite.Text
        !ItemSupplierComments = txtOther.Text
        .Update
        If .State = 1 Then .Close
        BeforeAddEdit
        ClearValues
        Call FillItemSupplier
        dtcItemSupplier.SetFocus
        dtcItemSupplier.Text = Empty
        Exit Sub
ErrorHandler:
        A = MsgBox(Err.Number & vbNewLine & Err.Description & vbTab & Me.Caption, vbCritical + vbOKOnly, "Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
        dtcItemSupplier.SetFocus
        dtcItemSupplier.Text = Empty
    End With
End Sub


Private Sub AfterAdd()
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnCancel.Visible = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    Frame1.Enabled = True
    Frame2.Enabled = False
End Sub

Private Sub AfterEdit()
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    Frame1.Enabled = True
    Frame2.Enabled = False
End Sub

Private Sub BeforeAddEdit()
    bttnAdd.Enabled = True
    bttnEdit.Enabled = True
    bttnSave.Visible = False
    bttnCancel.Visible = False
    bttnChange.Visible = False
    Frame1.Enabled = False
    Frame2.Enabled = True
End Sub

Private Sub NoName()
    Dim TemResponce As Integer
    TemResponce = MsgBox("No Such ItemSupplier found among the records", , "No Record")
    Exit Sub
End Sub

Private Sub ClearValues()
    txtName.Text = Empty
    txtAddress.Text = Empty
    dtcCity.Text = Empty
    txtTelephone.Text = Empty
    txtFax.Text = Empty
    txteMail.Text = Empty
    txtWebsite.Text = Empty
    txtOther.Text = Empty
End Sub

Private Sub DisplaySelected()
    If Not IsNumeric(dtcItemSupplier.BoundText) Then Exit Sub
    With rsItemSupplier
        If .State = 1 Then .Close
        .Open "Select* From tblItemSupplier Where ItemSupplierID = " & dtcItemSupplier.BoundText & "", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Call ClearValues
        If Not (!ItemSupplier) = "" Then txtName.Text = !ItemSupplier
        If Not (!ItemSupplierAddress) = "" Then txtAddress.Text = !ItemSupplierAddress
        If Not (!ItemSupplierCityID) = "" Then dtcCity.BoundText = Val(!ItemSupplierCityID)
        If Not (!ItemSupplierTelephone) = "" Then txtTelephone.Text = !ItemSupplierTelephone
        If Not (!ItemSupplierFax) = "" Then txtFax.Text = !ItemSupplierFax
        If Not (!ItemSupplierEmail) = "" Then txteMail.Text = !ItemSupplierEmail
        If Not (!ItemSupplierWebsite) = "" Then txtWebsite.Text = !ItemSupplierWebsite
        If Not (!ItemSupplierComments) = "" Then txtOther.Text = !ItemSupplierComments
        TemItemSupplierId = !ItemSupplierID
        If .RecordCount = 0 Then Exit Sub
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsViewItemSupplier.State = 1 Then rsViewItemSupplier.Close: Set rsViewItemSupplier = Nothing
    If rsViewCity1.State = 1 Then rsViewCity1.Close: Set rsViewCity1 = Nothing
End Sub
