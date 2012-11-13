VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmStaffDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff Details"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14145
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
   ScaleHeight     =   9240
   ScaleWidth      =   14145
   Begin VB.ListBox lstCommentIDs 
      Height          =   1185
      IntegralHeight  =   0   'False
      Left            =   4560
      TabIndex        =   25
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fra2 
      Caption         =   "Staff Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8535
      Left            =   4320
      TabIndex        =   23
      Top             =   120
      Width           =   9735
      Begin VB.CheckBox chkPasswordResetting 
         Caption         =   "Need Password Resetting"
         Height          =   255
         Left            =   1440
         TabIndex        =   45
         Top             =   3480
         Width           =   3135
      End
      Begin VB.Frame frameAuthority 
         Height          =   2415
         Left            =   120
         TabIndex        =   27
         Top             =   3720
         Width           =   9495
         Begin VB.CheckBox chkDErrors 
            Caption         =   "Detect Errors"
            Height          =   255
            Left            =   6600
            TabIndex        =   50
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CheckBox chkReports 
            Caption         =   "Reports"
            Height          =   255
            Left            =   6600
            TabIndex        =   49
            Top             =   960
            Width           =   2055
         End
         Begin VB.CheckBox chkOutstanding 
            Caption         =   "Outstanding"
            Height          =   255
            Left            =   4440
            TabIndex        =   48
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CheckBox chkPIPayments 
            Caption         =   "Print Individual Payments"
            Height          =   255
            Left            =   6600
            TabIndex        =   47
            Top             =   600
            Width           =   2655
         End
         Begin VB.CheckBox chkGIPayments 
            Caption         =   "Gearate Individual Payments"
            Height          =   255
            Left            =   6600
            TabIndex        =   46
            Top             =   240
            Width           =   2775
         End
         Begin VB.CheckBox chkApprovePayments 
            Caption         =   "Approve Payments"
            Height          =   255
            Left            =   4440
            TabIndex        =   44
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox chkItemSuppiers 
            Caption         =   "Item Suppiers"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   2040
            Width           =   1695
         End
         Begin VB.CheckBox chkProfits 
            Caption         =   "Profits"
            Height          =   255
            Left            =   4440
            TabIndex        =   42
            Top             =   1680
            Width           =   2055
         End
         Begin VB.CheckBox chkApprovals 
            Caption         =   "Approvals"
            Height          =   255
            Left            =   4440
            TabIndex        =   41
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CheckBox chkConfirmPayments 
            Caption         =   "Confirm Payments"
            Height          =   255
            Left            =   4440
            TabIndex        =   40
            Top             =   960
            Width           =   2055
         End
         Begin VB.CheckBox chkIssuePayments 
            Caption         =   "Issue Payments"
            Height          =   255
            Left            =   4440
            TabIndex        =   39
            Top             =   600
            Width           =   2055
         End
         Begin VB.CheckBox chkPrintPayAdvice 
            Caption         =   "Print Pay Advice"
            Height          =   255
            Left            =   2160
            TabIndex        =   38
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CheckBox chkGeneratePayAdvice 
            Caption         =   "Generate Pay advice"
            Height          =   255
            Left            =   2160
            TabIndex        =   37
            Top             =   1680
            Width           =   2055
         End
         Begin VB.CheckBox chkCumalativeReport 
            Caption         =   "Cumulative Report"
            Height          =   255
            Left            =   2160
            TabIndex        =   36
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CheckBox chkExpenceCategory 
            Caption         =   "Expence Category"
            Height          =   255
            Left            =   2160
            TabIndex        =   35
            Top             =   960
            Width           =   2055
         End
         Begin VB.CheckBox chkIncomeCategory 
            Caption         =   "Income Category"
            Height          =   255
            Left            =   2160
            TabIndex        =   34
            Top             =   600
            Width           =   2055
         End
         Begin VB.CheckBox chkStaff 
            Caption         =   "Staff"
            Height          =   255
            Left            =   2160
            TabIndex        =   33
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "Item"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CheckBox chkPrices 
            Caption         =   "Prices"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox chkFarmers 
            Caption         =   "Farmers"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox chkCollectingCenter 
            Caption         =   "Collecting Center"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox chkBackUpAllowed 
            Caption         =   "Backup"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1815
         End
      End
      Begin MSDataListLib.DataCombo dtcAuthority 
         Height          =   360
         Left            =   1440
         TabIndex        =   17
         Top             =   3120
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.ListBox lstComments 
         Height          =   1185
         IntegralHeight  =   0   'False
         Left            =   1560
         TabIndex        =   18
         Top             =   6480
         Width           =   5415
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   5535
      End
      Begin VB.TextBox txtAddress 
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   5535
      End
      Begin VB.TextBox txtPhone 
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   1200
         Width           =   5535
      End
      Begin VB.TextBox txtMobile 
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   1680
         Width           =   5535
      End
      Begin VB.TextBox txtUserName 
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   2160
         Width           =   5535
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   2640
         Width           =   5535
      End
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Top             =   7920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "Save"
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
         Left            =   2400
         TabIndex        =   20
         Top             =   7920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "Save"
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
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   375
         Left            =   3600
         TabIndex        =   21
         Top             =   7920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "Cancel"
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
      Begin VB.Label lblAuthority 
         Caption         =   "Authority"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblMobile 
         Caption         =   "Mobile"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblComments 
         Caption         =   "Comments"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label lblUserName 
         Caption         =   "User Name"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   1455
      End
   End
   Begin VB.Frame fra1 
      Caption         =   "Staff"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   7920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "Edit"
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
         TabIndex        =   2
         Top             =   7920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "Add"
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
      Begin MSDataListLib.DataCombo dtcStaffDetails 
         Height          =   7380
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   13018
         _Version        =   393216
         Style           =   1
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
      Begin btButtonEx.ButtonEx btnDelete 
         Height          =   375
         Left            =   2760
         TabIndex        =   26
         Top             =   7920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16711680
         Caption         =   "Delete"
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
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   10080
      TabIndex        =   22
      Top             =   8760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "Close"
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
Attribute VB_Name = "frmStaffDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsStaffDetails As New ADODB.Recordset
    Dim rsViewStaffDetails As New ADODB.Recordset
    Dim rsComments As New ADODB.Recordset
    Dim rsViewAuthority As New ADODB.Recordset
    
    Dim temSQL As String
    
Private Sub AfterAdd()
    
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    dtcStaffDetails.Enabled = False
    
    bttnSave.Enabled = True
    bttnChange.Enabled = False
    bttnCancel.Enabled = True
    txtName.Enabled = True
    txtAddress.Enabled = True
    txtPhone.Enabled = True
    txtMobile.Enabled = True
    txtUserName.Enabled = True
    txtPassword.Enabled = True
    lstComments.Enabled = True
    dtcAuthority.Enabled = True
    
    chkBackUpAllowed.Enabled = True
    chkCollectingCenter.Enabled = True
    chkFarmers.Enabled = True
    chkPrices.Enabled = True
    chkItem.Enabled = True
    chkStaff.Enabled = True
    chkIncomeCategory.Enabled = True
    chkExpenceCategory.Enabled = True
    chkGeneratePayAdvice.Enabled = True
    chkPrintPayAdvice.Enabled = True
    chkCumalativeReport.Enabled = True
    chkIssuePayments.Enabled = True
    chkApprovePayments.Enabled = True
    chkConfirmPayments.Enabled = True
    chkApprovals.Enabled = True
    chkProfits.Enabled = True
    chkItemSuppiers.Enabled = True
    
    chkOutstanding.Enabled = True
    chkGIPayments.Enabled = True
    chkPIPayments.Enabled = True
    chkReports.Enabled = True
    chkDErrors.Enabled = True
    
    
    chkPasswordResetting.Enabled = True
    
    bttnSave.Visible = True
    bttnChange.Visible = False
    
End Sub
Private Sub AfterEdit()
    
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    dtcStaffDetails.Enabled = False
    
    bttnSave.Enabled = False
    bttnChange.Enabled = True
    bttnCancel.Enabled = True
    txtName.Enabled = True
    txtAddress.Enabled = True
    txtPhone.Enabled = True
    txtMobile.Enabled = True
    txtUserName.Enabled = True
    txtPassword.Enabled = True
    lstComments.Enabled = True
    dtcAuthority.Enabled = True
    chkPasswordResetting.Enabled = True
   
    chkBackUpAllowed.Enabled = True
    chkCollectingCenter.Enabled = True
    chkFarmers.Enabled = True
    chkPrices.Enabled = True
    chkItem.Enabled = True
    chkStaff.Enabled = True
    chkIncomeCategory.Enabled = True
    chkExpenceCategory.Enabled = True
    chkGeneratePayAdvice.Enabled = True
    chkPrintPayAdvice.Enabled = True
    chkCumalativeReport.Enabled = True
    chkIssuePayments.Enabled = True
    chkApprovePayments.Enabled = True
    chkConfirmPayments.Enabled = True
    chkApprovals.Enabled = True
    chkProfits.Enabled = True
    chkItemSuppiers.Enabled = True
    
    chkOutstanding.Enabled = True
    chkGIPayments.Enabled = True
    chkPIPayments.Enabled = True
    chkReports.Enabled = True
    chkDErrors.Enabled = True

    
   
    bttnSave.Visible = False
    bttnChange.Visible = True
    
End Sub

Private Sub BeforeAddEdit()
    
    bttnAdd.Enabled = True
    bttnEdit.Enabled = True
    dtcStaffDetails.Enabled = True
    
    bttnSave.Enabled = False
    bttnChange.Enabled = False
    bttnCancel.Enabled = False
    txtName.Enabled = False
    txtAddress.Enabled = False
    txtPhone.Enabled = False
    txtMobile.Enabled = False
    txtUserName.Enabled = False
    txtPassword.Enabled = False
    lstComments.Enabled = False
    dtcAuthority.Enabled = False
    
    chkPasswordResetting.Enabled = False
   
    chkBackUpAllowed.Enabled = False
    chkCollectingCenter.Enabled = False
    chkFarmers.Enabled = False
    chkPrices.Enabled = False
    chkItem.Enabled = False
    chkStaff.Enabled = False
    chkIncomeCategory.Enabled = False
    chkExpenceCategory.Enabled = False
    chkGeneratePayAdvice.Enabled = False
    chkPrintPayAdvice.Enabled = False
    chkCumalativeReport.Enabled = False
    chkIssuePayments.Enabled = False
    chkApprovePayments.Enabled = False
    chkConfirmPayments.Enabled = False
    chkApprovals.Enabled = False
    chkProfits.Enabled = False
    chkItemSuppiers.Enabled = False
    
    chkOutstanding.Enabled = False
    chkGIPayments.Enabled = False
    chkPIPayments.Enabled = False
    chkReports.Enabled = False
    chkDErrors.Enabled = False

    
   
    bttnSave.Visible = True
    bttnChange.Visible = True
    
    On Error Resume Next
    dtcStaffDetails.SetFocus
    
End Sub

Private Sub ClearValues()
    txtName.Text = Empty
    txtAddress.Text = Empty
    txtPhone.Text = Empty
    txtMobile.Text = Empty
    txtUserName.Text = Empty
    txtPassword.Text = Empty
    dtcAuthority.Text = Empty
    
    chkPasswordResetting.Value = 0
    
    chkBackUpAllowed.Value = 0
    chkCollectingCenter.Value = 0
    chkFarmers.Value = 0
    chkPrices.Value = 0
    chkItem.Value = 0
    chkStaff.Value = 0
    chkIncomeCategory.Value = 0
    chkExpenceCategory.Value = 0
    chkGeneratePayAdvice.Value = 0
    chkPrintPayAdvice.Value = 0
    chkCumalativeReport.Value = 0
    chkIssuePayments.Value = 0
    chkApprovePayments.Value = 0
    chkConfirmPayments.Value = 0
    chkApprovals.Value = 0
    chkProfits.Value = 0
    chkItemSuppiers.Value = 0
    
    chkOutstanding.Value = 0
    chkGIPayments.Value = 0
    chkPIPayments.Value = 0
    chkReports.Value = 0
    chkDErrors.Value = 0

   
   
    lstComments.Clear
    lstCommentIDs.Clear
End Sub

Private Sub btnDelete_Click()

    Dim i As Integer
    i = MsgBox("Are you sure?", vbYesNo)
    If i = vbNo Then Exit Sub
    
    If IsNumeric(dtcStaffDetails.BoundText) = False Then
        MsgBox "Please Select a User to delete"
        dtcStaffDetails.SetFocus
        Exit Sub
    End If
    
    i = MsgBox("Are you sure you want to delete the user named " & dtcStaffDetails.Text & "?", vbYesNo)
    If i = vbYes Then
        Dim rsStaff As New ADODB.Recordset
        With rsStaff
            If .State = 1 Then .Close
            temSQL = "SELECT * FROM tblStaff where StaffID = " & Val(dtcStaffDetails.BoundText)
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Deleted = True
                !DeletedUserID = UserID
                !DeletedTime = Now
                .Update
                MsgBox "User Deleted Successfully"
            Else
                MsgBox "Error"
            End If
            .Close
        End With
    End If
    Call ClearValues
    Call FillCombos
    dtcStaffDetails.SetFocus
    dtcStaffDetails.Text = Empty
End Sub

Private Sub bttnAdd_Click()
    ClearValues
    AfterAdd
    txtName.SetFocus
    txtName.Text = dtcStaffDetails.Text
    SendKeys "{home}+{end}"
End Sub

Private Sub bttnCancel_Click()
    Call BeforeAddEdit
    Call ClearValues
    dtcStaffDetails.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub bttnChange_Click()
    If CanAdd = False Then Exit Sub
    Dim TemResponce As Integer
    With rsStaffDetails
    'On Error GoTo ErrorHandler
        If .State = 1 Then .Close
        .Open "Select * From tblStaff Where StaffID = " & Val(dtcStaffDetails.BoundText), cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount = 0 Then Exit Sub
        !Staff = Trim(txtName.Text)
        !Address = txtAddress.Text
        !Phone = txtPhone.Text
        !Mobile = txtMobile.Text
        !UserName = EncreptedWord(txtUserName.Text)
        !Password = EncreptedWord(txtPassword.Text)
        !AuthorityID = dtcAuthority.BoundText
        
        If chkPasswordResetting.Value = 1 Then
            !NeedPasswordReset = True
        Else
            !NeedPasswordReset = False
        End If
        
        If chkBackUpAllowed.Value = 1 Then
            !BackupAllowed = True
        Else
            !BackupAllowed = False
        End If
        
        If chkCollectingCenter.Value = 1 Then
            !CollectingCenterAllowed = True
        Else
            !CollectingCenterAllowed = False
        End If
        
        If chkFarmers.Value = 1 Then
            !FarmersAllowed = True
        Else
            !FarmersAllowed = False
        End If
        
        If chkPrices.Value = 1 Then
            !PricesAllowed = True
        Else
            !PricesAllowed = False
        End If
        
        If chkItem.Value = 1 Then
            !ItemAllowed = True
        Else
            !ItemAllowed = False
        End If
        
        If chkStaff.Value = 1 Then
            !StaffsAllowed = True
        Else
            !StaffsAllowed = False
        End If
        
        If chkIncomeCategory.Value = 1 Then
            !IncomeCategoryAllowed = True
        Else
            !IncomeCategoryAllowed = False
        End If
        
        If chkExpenceCategory.Value = 1 Then
            !ExpenceCategoryAllowed = True
        Else
            !ExpenceCategoryAllowed = False
        End If
        
        If chkGeneratePayAdvice.Value = 1 Then
            !GenarateCollectingCenterPayAdviceAllowed = True
        Else
            !GenarateCollectingCenterPayAdviceAllowed = False
        End If
        
        If chkPrintPayAdvice.Value = 1 Then
            !PrintCollectingCenterPayAdviceAllowed = True
        Else
            !PrintCollectingCenterPayAdviceAllowed = False
        End If
        
        If chkCumalativeReport.Value = 1 Then
            !CumulativeReportPrintAllowed = True
        Else
            !CumulativeReportPrintAllowed = False
        End If
        
        If chkIssuePayments.Value = 1 Then
            !IssuePaymentDeleteAllowed = True
        Else
            !IssuePaymentDeleteAllowed = False
        End If
        
        If chkApprovePayments.Value = 1 Then
            !ApprovePaymentSaveAllowed = True
        Else
            !ApprovePaymentSaveAllowed = False
        End If
        
        If chkConfirmPayments.Value = 1 Then
            !ConfirmPaymentsAllowed = True
        Else
            !ConfirmPaymentsAllowed = False
        End If
        
        If chkApprovals.Value = 1 Then
            !ApprovalsAllowed = True
        Else
            !ApprovalsAllowed = False
        End If
        
        If chkProfits.Value = 1 Then
            !ProfitsAllowed = True
        Else
            !ProfitsAllowed = False
        End If
        
        If chkItemSuppiers.Value = 1 Then
            !ItemSuppiersEditAllowed = True
        Else
            !ItemSuppiersEditAllowed = False
        End If
        
        If chkOutstanding.Value = 1 Then
            !OutstandingAllowed = True
        Else
            !OutstandingAllowed = False
        End If
        
        If chkGIPayments.Value = 1 Then
            !GenerateIndividualPaymentsAllowed = True
        Else
            !GenerateIndividualPaymentsAllowed = False
        End If
        
         If chkPIPayments.Value = 1 Then
            !PrintIndividualPaymentsAllowed = True
        Else
            !PrintIndividualPaymentsAllowed = False
        End If
        
        
         If chkReports.Value = 1 Then
            !ReportsAllowed = True
        Else
            !ReportsAllowed = False
        End If
        
         If chkDErrors.Value = 1 Then
            !DetectErrorsAllowed = True
        Else
            !DetectErrorsAllowed = False
        End If
        
        
        .Update
        If .State = 1 Then .Close
        FillCombos
        BeforeAddEdit
        ClearValues
        dtcStaffDetails.Text = Empty
        dtcStaffDetails.SetFocus
        Exit Sub
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Save Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        dtcStaffDetails.Text = Empty
        dtcStaffDetails.SetFocus
        If .State = 1 Then .Close
    End With
End Sub
    

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    AfterEdit
    txtName.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub bttnSave_Click()
    If CanAdd = False Then Exit Sub
    Dim TemResponce As Integer
    With rsStaffDetails
        'On Error GoTo ErrorHandler
        If .State = 1 Then .Close
        .Open "Select * From tblStaff", cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Staff = Trim(txtName.Text)
        !Address = txtAddress.Text
        !Phone = txtPhone.Text
        !Mobile = txtMobile.Text
        !UserName = EncreptedWord(txtUserName.Text)
        !Password = EncreptedWord(txtPassword.Text)
        !AuthorityID = dtcAuthority.BoundText
        
        If chkPasswordResetting.Value = 1 Then
            !NeedPasswordReset = True
        Else
            !NeedPasswordReset = False
        End If
        
        If chkBackUpAllowed.Value = 1 Then
            !BackupAllowed = True
        Else
            !BackupAllowed = False
        End If
        
        
        If chkCollectingCenter.Value = 1 Then
            !CollectingCenterAllowed = True
        Else
            !CollectingCenterAllowed = False
        End If
        
        If chkFarmers.Value = 1 Then
            !FarmersAllowed = True
        Else
            !FarmersAllowed = False
        End If
        
        If chkPrices.Value = 1 Then
            !PricesAllowed = True
        Else
            !PricesAllowed = False
        End If
        
        If chkItem.Value = 1 Then
            !ItemAllowed = True
        Else
            !ItemAllowed = False
        End If
        
        If chkStaff.Value = 1 Then
            !StaffsAllowed = True
        Else
            !StaffsAllowed = False
        End If
        
        If chkIncomeCategory.Value = 1 Then
            !IncomeCategoryAllowed = True
        Else
            !IncomeCategoryAllowed = False
        End If
        
        If chkExpenceCategory.Value = 1 Then
            !ExpenceCategoryAllowed = True
        Else
            !ExpenceCategoryAllowed = False
        End If
        
        If chkGeneratePayAdvice.Value = 1 Then
            !GenarateCollectingCenterPayAdviceAllowed = True
        Else
            !GenarateCollectingCenterPayAdviceAllowed = False
        End If
        
        If chkPrintPayAdvice.Value = 1 Then
            !PrintCollectingCenterPayAdviceAllowed = True
        Else
            !PrintCollectingCenterPayAdviceAllowed = False
        End If
        
        If chkCumalativeReport.Value = 1 Then
            !CumulativeReportPrintAllowed = True
        Else
            !CumulativeReportPrintAllowed = False
        End If
        
        If chkIssuePayments.Value = 1 Then
            !IssuePaymentDeleteAllowed = True
        Else
            !IssuePaymentDeleteAllowed = False
        End If
        
        If chkApprovePayments.Value = 1 Then
            !ApprovePaymentSaveAllowed = True
        Else
            !ApprovePaymentSaveAllowed = False
        End If
        
        If chkConfirmPayments.Value = 1 Then
            !ConfirmPaymentsAllowed = True
        Else
            !ConfirmPaymentsAllowed = False
        End If
        
        If chkApprovals.Value = 1 Then
            !ApprovalsAllowed = True
        Else
            !ApprovalsAllowed = False
        End If
        
        If chkProfits.Value = 1 Then
            !ProfitsAllowed = True
        Else
            !ProfitsAllowed = False
        End If
        
        If chkItemSuppiers.Value = 1 Then
            !ItemSuppiersEditAllowed = True
        Else
            !ItemSuppiersEditAllowed = False
        End If
        
               
        If chkOutstanding.Value = 1 Then
            !OutstandingAllowed = True
        Else
            !OutstandingAllowed = False
        End If
        
        If chkGIPayments.Value = 1 Then
            !GenerateIndividualPaymentsAllowed = True
        Else
            !GenerateIndividualPaymentsAllowed = False
        End If
        
         If chkPIPayments.Value = 1 Then
            !PrintIndividualPaymentsAllowed = True
        Else
            !PrintIndividualPaymentsAllowed = False
        End If
        
         If chkReports.Value = 1 Then
            !ReportsAllowed = True
        Else
            !ReportsAllowed = False
        End If
        
         If chkDErrors.Value = 1 Then
            !DetectErrorsAllowed = True
        Else
            !DetectErrorsAllowed = False
        End If
     
        .Update
        If .State = 1 Then .Close
        FillCombos
        BeforeAddEdit
        ClearValues
        dtcStaffDetails.Text = Empty
        dtcStaffDetails.SetFocus
        Exit Sub
    
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Save Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        dtcStaffDetails.Text = Empty
        dtcStaffDetails.SetFocus
        If .State = 1 Then .Close
    End With
End Sub

Private Function CanAdd() As Boolean
    CanAdd = False
    Dim tr As Integer
    If Trim(txtName.Text) = Empty Then
        tr = MsgBox("You have not entered the Name", vbCritical, "No Name")
        txtName.SetFocus
        Exit Function
    End If
    If IsNumeric(dtcAuthority.BoundText) = False Then
        tr = MsgBox("You have not selected an authority", vbCritical, "Authority")
        dtcAuthority.SetFocus
        Exit Function
    End If
'    If UserAuthority <> 1 And dtcAuthority.BoundText = 1 Then
'        tr = MsgBox("You have not allowed to set the authority as an administrator", vbCritical, "Not Authorised")
'        dtcAuthority.SetFocus
'        Exit Function
'    End If
    If UserNameAvailable = False Then
        tr = MsgBox("User Name Already Exisits")
        txtUserName.SetFocus
        Exit Function
    End If
    CanAdd = True
End Function

Private Sub dtcAuthority_Click(Area As Integer)
    If Val(dtcAuthority.BoundText) = 6 Then
        chkBackUpAllowed.Value = 1
        chkCollectingCenter.Value = 1
        chkFarmers.Value = 1
        chkPrices.Value = 1
        chkItem.Value = 1
        chkStaff.Value = 1
        chkIncomeCategory.Value = 1
        chkExpenceCategory.Value = 1
        chkGeneratePayAdvice.Value = 1
        chkPrintPayAdvice.Value = 1
        chkCumalativeReport.Value = 1
        chkIssuePayments.Value = 1
        chkApprovePayments.Value = 1
        chkConfirmPayments.Value = 1
        chkApprovals.Value = 1
        chkProfits.Value = 1
        chkItemSuppiers.Value = 1
        
        chkOutstanding.Value = 1
        chkGIPayments.Value = 1
        chkPIPayments.Value = 1
        chkReports.Value = 1
        chkDErrors.Value = 1
    End If

    
    If Val(dtcAuthority.BoundText) = 3 Then
        chkCumalativeReport.Value = 0
        chkBackUpAllowed.Value = 0
        chkCollectingCenter.Value = 0
        chkFarmers.Value = 0
        chkPrices.Value = 0
        chkItem.Value = 0
        chkStaff.Value = 0
        chkIncomeCategory.Value = 0
        chkExpenceCategory.Value = 0
        chkGeneratePayAdvice.Value = 0
        chkPrintPayAdvice.Value = 1
        chkCumalativeReport.Value = 1
        chkIssuePayments.Value = 0
        chkApprovePayments.Value = 0
        chkConfirmPayments.Value = 0
        chkApprovals.Value = 0
        chkProfits.Value = 1
        chkItemSuppiers.Value = 0
        
        chkOutstanding.Value = 0
        chkGIPayments.Value = 0
        chkPIPayments.Value = 0
        chkConfirmPayments.Value = 0
        chkReports.Value = 0
        chkDErrors.Value = 0
    End If
    
    If Val(dtcAuthority.BoundText) = 4 Then
        chkCumalativeReport.Value = 0
        chkBackUpAllowed.Value = 0
        chkCollectingCenter.Value = 0
        chkFarmers.Value = 1
        chkPrices.Value = 0
        chkItem.Value = 1
        chkStaff.Value = 0
        chkIncomeCategory.Value = 1
        chkExpenceCategory.Value = 1
        chkGeneratePayAdvice.Value = 0
        chkPrintPayAdvice.Value = 1
        chkCumalativeReport.Value = 1
        chkIssuePayments.Value = 0
        chkApprovePayments.Value = 0
        chkConfirmPayments.Value = 1
        chkApprovals.Value = 1
        chkProfits.Value = 1
        chkItemSuppiers.Value = 0
        
        chkOutstanding.Value = 1
        chkGIPayments.Value = 1
        chkPIPayments.Value = 1
        chkReports.Value = 1
        chkDErrors.Value = 1
    End If
    
    If Val(dtcAuthority.BoundText) = 5 Then
        chkCumalativeReport.Value = 1
        chkBackUpAllowed.Value = 0
        chkCollectingCenter.Value = 1
        chkFarmers.Value = 1
        chkPrices.Value = 1
        chkItem.Value = 1
        chkStaff.Value = 1
        chkIncomeCategory.Value = 1
        chkExpenceCategory.Value = 1
        chkGeneratePayAdvice.Value = 1
        chkPrintPayAdvice.Value = 1
        chkCumalativeReport.Value = 1
        chkIssuePayments.Value = 1
        chkApprovePayments.Value = 1
        chkConfirmPayments.Value = 1
        chkApprovals.Value = 1
        chkProfits.Value = 1
        chkItemSuppiers.Value = 1
    
        chkOutstanding.Value = 1
        chkGIPayments.Value = 1
        chkPIPayments.Value = 1
        chkReports.Value = 1
        chkDErrors.Value = 1
    
    End If
    
End Sub

Private Sub dtcStaffDetails_Change()
    If IsNumeric(dtcStaffDetails.BoundText) = True Then
        bttnEdit.Enabled = True
        bttnAdd.Enabled = False
        ClearValues
        DisplaySelected
    Else
        bttnEdit.Enabled = False
        bttnAdd.Enabled = True
        ClearValues
    End If
End Sub


Private Sub Form_Load()
    If UserAuthority = Authority.SuperUser Then
        dtcAuthority.Locked = True
        txtUserName.Locked = True
        
        chkPasswordResetting.Visible = False
        
        chkBackUpAllowed.Enabled = False
        chkCollectingCenter.Enabled = False
        chkFarmers.Enabled = False
        chkPrices.Enabled = False
        chkItem.Enabled = False
        chkStaff.Enabled = False
        chkIncomeCategory.Enabled = False
        chkExpenceCategory.Enabled = False
        chkGeneratePayAdvice.Enabled = False
        chkPrintPayAdvice.Enabled = False
        chkCumalativeReport.Enabled = False
        chkIssuePayments.Enabled = False
        chkApprovePayments.Enabled = False
        chkConfirmPayments.Enabled = False
        chkApprovals.Enabled = False
        chkProfits.Enabled = False
        chkItemSuppiers.Enabled = False
        
        chkOutstanding.Visible = False
        chkGIPayments.Visible = False
        chkPIPayments.Visible = False
        chkReports.Visible = False
        chkDErrors.Visible = False
     
        
        btnDelete.Visible = False
        bttnAdd.Visible = False
        
        End If
        
Select Case UserAuthorityLevel
    
        Case Authority.SuperUser '5
            frameAuthority.Visible = False
            dtcAuthority.Visible = False
            txtUserName.Visible = False
            txtPassword.Visible = False
            lblUserName.Visible = False
            lblPassword.Visible = False
            lblAuthority.Visible = False
            chkPasswordResetting.Visible = False
            btnDelete.Visible = False

        Case Else
End Select
    
    FillCombos
    BeforeAddEdit
    ClearValues
End Sub
Private Sub FillCombos()
    With rsViewStaffDetails
        If .State = 1 Then .Close
        temSQL = "Select * From tblStaff Where Deleted = 0  Order By Staff"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcStaffDetails
        Set .RowSource = rsViewStaffDetails
        .ListField = "Staff"
        .BoundColumn = "StaffID"
    End With
    With rsViewAuthority
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblAuthority order by Authority"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcAuthority
        Set .RowSource = rsViewAuthority
        .ListField = "Authority"
        .BoundColumn = "AuthorityID"
    End With
    
End Sub

Private Sub NoName()
    Dim TemResponce As Integer
    TemResponce = MsgBox("You have not entered a Name to save", vbCritical, "No Name")
    txtName.SetFocus
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsStaffDetails.State = 1 Then rsStaffDetails.Close: Set rsStaffDetails = Nothing
    If rsViewStaffDetails.State = 1 Then rsViewStaffDetails.Close: Set rsViewStaffDetails = Nothing
End Sub

Private Sub DisplaySelected()
    If Not IsNumeric(dtcStaffDetails.BoundText) Then Exit Sub
    With rsStaffDetails
        If .State = 1 Then .Close
        .Open "Select * From tblStaff Where (StaffID = " & dtcStaffDetails.BoundText & ")", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Call ClearValues
        txtName.Text = !Staff
        txtAddress.Text = !Address
        txtPhone.Text = !Phone
        txtMobile.Text = !Mobile
        txtUserName.Text = DecreptedWord(!UserName)
        txtPassword.Text = DecreptedWord(!Password)
        If Not IsNull(!AuthorityID) Then dtcAuthority.BoundText = !AuthorityID
        
        If !NeedPasswordReset = 1 Then chkPasswordResetting.Value = 1
        
        If !BackupAllowed = 1 Then chkBackUpAllowed.Value = 1
        If !CollectingCenterAllowed = 1 Then chkCollectingCenter.Value = 1
        If !FarmersAllowed = 1 Then chkFarmers.Value = 1
        If !PricesAllowed = 1 Then chkPrices.Value = 1
        If !ItemAllowed = True Then chkItem.Value = 1
        If !StaffsAllowed = True Then chkStaff.Value = 1
        If !IncomeCategoryAllowed = True Then chkIncomeCategory.Value = 1
        If !ExpenceCategoryAllowed = True Then chkExpenceCategory.Value = 1
        If !GenarateCollectingCenterPayAdviceAllowed = True Then chkGeneratePayAdvice = 1
        If !PrintCollectingCenterPayAdviceAllowed = True Then chkPrintPayAdvice.Value = 1
        If !CumulativeReportPrintAllowed = True Then chkCumalativeReport.Value = 1
        If !IssuePaymentDeleteAllowed = True Then chkIssuePayments.Value = 1
        If !ApprovePaymentSaveAllowed = True Then chkApprovePayments.Value = 1
        If !ConfirmPaymentsAllowed = True Then chkConfirmPayments.Value = 1
        If !ApprovalsAllowed = True Then chkApprovals.Value = 1
        If !ProfitsAllowed = True Then chkProfits.Value = 1
        If !ItemSuppiersEditAllowed = True Then chkItemSuppiers.Value = 1
        
        If !OutstandingAllowed = True Then chkOutstanding.Value = 1
        If !GenerateIndividualPaymentsAllowed = True Then chkGIPayments.Value = 1
        If !PrintIndividualPaymentsAllowed = True Then chkPIPayments.Value = 1
        If !ReportsAllowed = True Then chkReports.Value = 1
        If !DetectErrorsAllowed = True Then chkDErrors.Value = 1
        
        
        
        Call ListComments
        If .State = 1 Then .Close
    End With
End Sub

Private Sub ListComments()
    With rsComments
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblStaffComment where StaffID = " & dtcStaffDetails.BoundText
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        lstComments.Clear
        lstCommentIDs.Clear
        If .RecordCount > 0 Then
            While .EOF = False
                lstComments.AddItem Format(!Date, "dd MMM yyyy") & vbTab & !Comment
                lstCommentIDs.AddItem !StaffCommentID
                .MoveNext
            Wend
        End If
        .Close
    End With
End Sub


Private Sub lstComments_Click()
'    lstCommentIDs.ListIndex = lstComments.ListIndex
'    StaffCommentIDTx = Val(lstCommentIDs.Text)
'    Unload frmStaffCommentDisplay
'    frmStaffCommentDisplay.Show
End Sub

Private Function UserNameAvailable() As Boolean
    UserNameAvailable = True
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblStaff where Deleted = 0  and StaffID <> " & Val(dtcStaffDetails.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If txtUserName.Text = DecreptedWord(!UserName) Then
                UserNameAvailable = False
            End If
            .MoveNext
        Wend
    End With
End Function
