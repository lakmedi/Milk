VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmStaffDetails1 
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
      Height          =   9135
      Left            =   4320
      TabIndex        =   5
      Top             =   0
      Width           =   9735
      Begin VB.Frame FrameOptions 
         Caption         =   "Options"
         Height          =   855
         Left            =   8280
         TabIndex        =   63
         Top             =   3960
         Width           =   1335
         Begin VB.CheckBox chkPrinting 
            Caption         =   "Printing"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox chkDatabase 
            Caption         =   "Database"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   1095
         End
      End
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   375
         Left            =   3480
         TabIndex        =   16
         Top             =   8520
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
      Begin VB.Frame FrameBackOffice 
         Caption         =   "Back Office"
         Height          =   2535
         Left            =   6360
         TabIndex        =   54
         Top             =   4800
         Width           =   3255
         Begin VB.CheckBox chkApprovePaymentsSave 
            Caption         =   "Approve Payments Save"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   2160
            Width           =   2175
         End
         Begin VB.CheckBox chkDetectErrors 
            Caption         =   "Detect errors"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   1920
            Width           =   2535
         End
         Begin VB.CheckBox chkProfits 
            Caption         =   "Profits"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   1680
            Width           =   2535
         End
         Begin VB.CheckBox chkReports 
            Caption         =   "Reports"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CheckBox chkOutstanding 
            Caption         =   "Outstanding"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   1200
            Width           =   2535
         End
         Begin VB.CheckBox chkApprovals 
            Caption         =   "Approvals"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   960
            Width           =   2535
         End
         Begin VB.CheckBox chkConfirmPayments 
            Caption         =   "Confirm Payments"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkPIPayments 
            Caption         =   "Print individual Payments"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   480
            Width           =   2535
         End
         Begin VB.CheckBox chkGIPayments 
            Caption         =   "Generate Individual Payments"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame FrameIssuePayments 
         Caption         =   "Issue && Payments"
         Height          =   855
         Left            =   1440
         TabIndex        =   46
         Top             =   3960
         Width           =   6735
         Begin VB.CheckBox chkIssuePaymentDelete 
            Caption         =   "Issue Payment Delete"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4680
            TabIndex        =   68
            Top             =   480
            Width           =   1935
         End
         Begin VB.CheckBox chkIncome 
            Caption         =   "Incomes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   53
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox chkExpences 
            Caption         =   "Expences"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   52
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkadditionalDeduct 
            Caption         =   "Additional Deductions"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   51
            Top             =   480
            Width           =   1935
         End
         Begin VB.CheckBox chkAdditionalComm 
            Caption         =   "Additional Commisions"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   50
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox chkAddDeductions 
            Caption         =   "Add Deductions"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4680
            TabIndex        =   49
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkItemPurchase 
            Caption         =   "Item Purchase"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox chkItemIssue 
            Caption         =   "Item Issue"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame FrameMilkCollection 
         Caption         =   "Milk Collection"
         Height          =   1335
         Left            =   120
         TabIndex        =   39
         Top             =   6000
         Width           =   6135
         Begin VB.CheckBox chkCumulativeReportPrint 
            Caption         =   "Cumulative Report Print"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   67
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox chkPrintCCPA 
            Caption         =   "Print Collecting Center Pay Advice"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   45
            Top             =   480
            Width           =   2775
         End
         Begin VB.CheckBox chkGenerateCCPA 
            Caption         =   "Generate Collecting Center Pay Advice"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   44
            Top             =   720
            Width           =   3135
         End
         Begin VB.CheckBox chkMilkPayAdvice 
            Caption         =   "Milk Pay Advice"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkCumulativeReport 
            Caption         =   "Cumulative Report"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   42
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox chkGRN 
            Caption         =   "Good Recieve Note"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   480
            Width           =   1815
         End
         Begin VB.CheckBox chkDailyCollection 
            Caption         =   "Daily Collection"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame FrameEdit 
         Caption         =   "Edit"
         Height          =   1215
         Left            =   120
         TabIndex        =   29
         Top             =   4800
         Width           =   6135
         Begin VB.CheckBox chkItemSuppiersEdit 
            Caption         =   "Item Suppliers Edit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   66
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox chkItemSuppliers 
            Caption         =   "Item Suppliers"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   38
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox chkExpenceCategory 
            Caption         =   "Expence Category"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   37
            Top             =   480
            Width           =   1695
         End
         Begin VB.CheckBox chkIncomeCategory 
            Caption         =   "Income Category"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   36
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox chkAuthority 
            Caption         =   "Authority"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkStaff 
            Caption         =   "Staff"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   34
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "Item"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   33
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox chkPrices 
            Caption         =   "Prices"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox chkFarmers 
            Caption         =   "Farmers"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox chkCollectingCenter 
            Caption         =   "Collecting Center"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   30
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame FrameFile 
         Caption         =   "File"
         Height          =   855
         Left            =   120
         TabIndex        =   26
         Top             =   3960
         Width           =   1215
         Begin VB.CheckBox chkRestore 
            Caption         =   "Restore"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox chkBackUp 
            Caption         =   "Backup"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   2640
         Width           =   5535
      End
      Begin VB.TextBox txtUserName 
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   2160
         Width           =   5535
      End
      Begin VB.TextBox txtMobile 
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   1680
         Width           =   5535
      End
      Begin VB.TextBox txtPhone 
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   1200
         Width           =   5535
      End
      Begin VB.TextBox txtAddress 
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   5535
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   5535
      End
      Begin VB.ListBox lstComments 
         Height          =   945
         IntegralHeight  =   0   'False
         Left            =   1440
         TabIndex        =   8
         Top             =   7440
         Width           =   5535
      End
      Begin VB.CheckBox chkPasswordResetting 
         Caption         =   "Need Password Resetting"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   3600
         Width           =   3135
      End
      Begin MSDataListLib.DataCombo dtcAuthority 
         Height          =   360
         Left            =   1440
         TabIndex        =   7
         Top             =   3120
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   3480
         TabIndex        =   15
         Top             =   8520
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
         Left            =   4680
         TabIndex        =   17
         Top             =   8520
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
      Begin VB.Label lblPassword 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblUserName 
         Caption         =   "User Name"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblComments 
         Caption         =   "Comments"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   7440
         Width           =   1215
      End
      Begin VB.Label lblMobile 
         Caption         =   "Mobile"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblAuthority 
         Caption         =   "Authority"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3120
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
      Height          =   9135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   8520
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
         Top             =   8520
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
         Height          =   8100
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   14288
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
         TabIndex        =   4
         Top             =   8520
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
End
Attribute VB_Name = "frmStaffDetails1"
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
    
    chkPasswordResetting.Enabled = True
    
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------
    
'File
    chkBackUp.Enabled = True
    chkRestore.Enabled = True
    
'Edit
    chkCollectingCenter.Enabled = True
    chkFarmers.Enabled = True
    chkPrices.Enabled = True
    chkItem.Enabled = True
    chkStaff.Enabled = True
    chkAuthority.Enabled = True
    chkExpenceCategory.Enabled = True
    chkIncomeCategory.Enabled = True
    chkItemSuppliers.Enabled = True
    
    chkItemSuppiersEdit.Enabled = True
    
'Milk Collection
    chkDailyCollection.Enabled = True
    chkGRN.Enabled = True
    chkCumulativeReport.Enabled = True
    chkGenerateCCPA.Enabled = True
    chkPrintCCPA.Enabled = True
    chkMilkPayAdvice.Enabled = True
    
    chkCumulativeReportPrint.Enabled = True
    
'Issue & Payments
    chkItemIssue.Enabled = True
    chkItemPurchase.Enabled = True
    chkAddDeductions.Enabled = True
    chkAdditionalComm.Enabled = True
    chkadditionalDeduct.Enabled = True
    chkExpences.Enabled = True
    chkIncome.Enabled = True
    
    chkIssuePaymentDelete.Enabled = True
    
'Back Office
    chkGIPayments.Enabled = True
    chkPIPayments.Enabled = True
    chkConfirmPayments.Enabled = True
    chkApprovals.Enabled = True
    chkOutstanding.Enabled = True
    chkReports.Enabled = True
    chkProfits.Enabled = True
    chkDetectErrors.Enabled = True
    
    chkApprovePaymentsSave.Enabled = True
    
'Options
    chkDatabase.Enabled = True
    chkPrinting.Enabled = True
    
'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------

    
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
    
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------

'File
    chkBackUp.Enabled = True
    chkRestore.Enabled = True
    
'Edit
    chkCollectingCenter.Enabled = True
    chkFarmers.Enabled = True
    chkPrices.Enabled = True
    chkItem.Enabled = True
    chkStaff.Enabled = True
    chkAuthority.Enabled = True
    chkExpenceCategory.Enabled = True
    chkIncomeCategory.Enabled = True
    chkItemSuppliers.Enabled = True
    
    chkItemSuppiersEdit.Enabled = True
    
'Milk Collection
    chkDailyCollection.Enabled = True
    chkGRN.Enabled = True
    chkCumulativeReport.Enabled = True
    chkGenerateCCPA.Enabled = True
    chkPrintCCPA.Enabled = True
    chkMilkPayAdvice.Enabled = True
    
    chkCumulativeReportPrint.Enabled = True
    
'Issue & Payments
    chkItemIssue.Enabled = True
    chkItemPurchase.Enabled = True
    chkAddDeductions.Enabled = True
    chkAdditionalComm.Enabled = True
    chkadditionalDeduct.Enabled = True
    chkExpences.Enabled = True
    chkIncome.Enabled = True
    
    chkIssuePaymentDelete.Enabled = True
    
'Back Office
    chkGIPayments.Enabled = True
    chkPIPayments.Enabled = True
    chkConfirmPayments.Enabled = True
    chkApprovals.Enabled = True
    chkOutstanding.Enabled = True
    chkReports.Enabled = True
    chkProfits.Enabled = True
    chkDetectErrors.Enabled = True
    
    chkApprovePaymentsSave.Enabled = True
    
'Options
    chkDatabase.Enabled = True
    chkPrinting.Enabled = True

'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------

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
    
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------

'File
    chkBackUp.Enabled = False
    chkRestore.Enabled = False
    
'Edit
    chkCollectingCenter.Enabled = False
    chkFarmers.Enabled = False
    chkPrices.Enabled = False
    chkItem.Enabled = False
    chkStaff.Enabled = False
    chkAuthority.Enabled = False
    chkExpenceCategory.Enabled = False
    chkIncomeCategory.Enabled = False
    chkItemSuppliers.Enabled = False
    
    chkItemSuppiersEdit.Enabled = False
    
'Milk Collection
    chkDailyCollection.Enabled = False
    chkGRN.Enabled = False
    chkCumulativeReport.Enabled = False
    chkGenerateCCPA.Enabled = False
    chkPrintCCPA.Enabled = False
    chkMilkPayAdvice.Enabled = False
    
    chkCumulativeReportPrint.Enabled = False
    
'Issue & Payments
    chkItemIssue.Enabled = False
    chkItemPurchase.Enabled = False
    chkAddDeductions.Enabled = False
    chkAdditionalComm.Enabled = False
    chkadditionalDeduct.Enabled = False
    chkExpences.Enabled = False
    chkIncome.Enabled = False
    
    chkIssuePaymentDelete.Enabled = False
    
'Back Office
    chkGIPayments.Enabled = False
    chkPIPayments.Enabled = False
    chkConfirmPayments.Enabled = False
    chkApprovals.Enabled = False
    chkOutstanding.Enabled = False
    chkReports.Enabled = False
    chkProfits.Enabled = False
    chkDetectErrors.Enabled = False
    
    chkApprovePaymentsSave.Enabled = False
    
'Options
    chkDatabase.Enabled = False
    chkPrinting.Enabled = False

'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------
  
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
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------

'File
    chkBackUp.Value = 0
    chkRestore.Value = 0
    
'Edit
    chkCollectingCenter.Value = 0
    chkFarmers.Value = 0
    chkPrices.Value = 0
    chkItem.Value = 0
    chkStaff.Value = 0
    chkAuthority.Value = 0
    chkExpenceCategory.Value = 0
    chkIncomeCategory.Value = 0
    chkItemSuppliers.Value = 0
    
    chkItemSuppiersEdit.Value = 0
    
'Milk Collection
    chkDailyCollection.Value = 0
    chkGRN.Value = 0
    chkCumulativeReport.Value = 0
    chkGenerateCCPA.Value = 0
    chkPrintCCPA.Value = 0
    chkMilkPayAdvice.Value = 0
    
    chkCumulativeReportPrint.Value = 0
    
'Issue & Payments
    chkItemIssue.Value = 0
    chkItemPurchase.Value = 0
    chkAddDeductions.Value = 0
    chkAdditionalComm.Value = 0
    chkadditionalDeduct.Value = 0
    chkExpences.Value = 0
    chkIncome.Value = 0
    
    chkIssuePaymentDelete.Value = 0
    
'Back Office
    chkGIPayments.Value = 0
    chkPIPayments.Value = 0
    chkConfirmPayments.Value = 0
    chkApprovals.Value = 0
    chkOutstanding.Value = 0
    chkReports.Value = 0
    chkProfits.Value = 0
    chkDetectErrors.Value = 0
    
    chkApprovePaymentsSave.Value = 0
    
'Options
    chkDatabase.Value = 0
    chkPrinting.Value = 0
    
'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------

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
    On Error Resume Next
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
        
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------

'File
        If chkBackUp.Value = 1 Then
            !BackupAllowed = True
        Else
            !BackupAllowed = False
        End If
        
        If chkRestore.Value = 1 Then
            !RestoreAllowed = True
        Else
            !RestoreAllowed = False
        End If
        
'Edit

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
'
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
        
        If chkAuthority.Value = 1 Then
            !AuthorityAllowed = True
        Else
            !AuthorityAllowed = False
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
        
        If chkItemSuppliers.Value = 1 Then
            !ItemSuppiersAllowed = True
        Else
            !ItemSuppiersAllowed = False
        End If
        
        
        If chkItemSuppiersEdit.Value = 1 Then
            !ItemSuppiersEditAllowed = True
        Else
            !ItemSuppiersEditAllowed = False
        End If
        
'Milk Collection

        If chkDailyCollection.Value = 1 Then
            !DailyCollectionAllowed = True
        Else
            !DailyCollectionAllowed = False
        End If
        
        If chkGRN.Value = 1 Then
            !GoodRecieveNoteAllowed = True
        Else
            !GoodRecieveNoteAllowed = False
        End If
        
        If chkCumulativeReport.Value = 1 Then
            !CumulativeReportAllowed = True
        Else
            !CumulativeReportAllowed = False
        End If
        
        If chkGenerateCCPA.Value = 1 Then
            !GenarateCollectingCenterPayAdviceAllowed = True
        Else
            !GenarateCollectingCenterPayAdviceAllowed = False
        End If

        If chkPrintCCPA.Value = 1 Then
            !PrintCollectingCenterPayAdviceAllowed = True
        Else
            !PrintCollectingCenterPayAdviceAllowed = False
        End If
        
        If chkMilkPayAdvice.Value = 1 Then
            !MilkPayAdviceAllowed = True
        Else
            !MilkPayAdviceAllowed = False
        End If
        

        If chkCumulativeReportPrint.Value = 1 Then
            !CumulativeReportPrintAllowed = True
        Else
            !CumulativeReportPrintAllowed = False
        End If
        
'Issue & Payments

        If chkItemIssue.Value = 1 Then
            !ItemIssueAllowed = True
        Else
            !ItemIssueAllowed = False
        End If
        
        If chkItemPurchase.Value = 1 Then
            !ItemPurchaseAllowed = True
        Else
            !ItemPurchaseAllowed = False
        End If
        
        If chkAddDeductions.Value = 1 Then
            !AddDeductionsAllowed = True
        Else
            !AddDeductionsAllowed = False
        End If

        If chkAdditionalComm.Value = 1 Then
            !AdditionalCommisionsAllowed = True
        Else
            !AdditionalCommisionsAllowed = False
        End If
        
        If chkadditionalDeduct.Value = 1 Then
            !AdditionalDeductionsAllowed = True
        Else
            !AdditionalDeductionsAllowed = False
        End If
        
        If chkExpences.Value = 1 Then
            !ExpencesAllowed = True
        Else
            !ExpencesAllowed = False
        End If
        
        If chkIncome.Value = 1 Then
            !IncomeAllowed = True
        Else
            !IncomeAllowed = False
        End If
        

        If chkIssuePaymentDelete.Value = 1 Then
            !IssuePaymentDeleteAllowed = True
        Else
            !IssuePaymentDeleteAllowed = False
        End If

'Back Office

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
        
        If chkOutstanding.Value = 1 Then
            !OutstandingAllowed = True
        Else
            !OutstandingAllowed = False
        End If
        
         If chkReports.Value = 1 Then
            !ReportsAllowed = True
        Else
            !ReportsAllowed = False
        End If
        
        If chkProfits.Value = 1 Then
            !ProfitsAllowed = True
        Else
            !ProfitsAllowed = False
        End If
    
         If chkDetectErrors.Value = 1 Then
            !DetectErrorsAllowed = True
        Else
            !DetectErrorsAllowed = False
        End If
        

        If chkApprovePaymentsSave.Value = 1 Then
            !ApprovePaymentSaveAllowed = True
        Else
            !ApprovePaymentSaveAllowed = False
        End If

'Options

        If chkDatabase.Value = 1 Then
            !DatabaseAllowed = True
        Else
            !DatabaseAllowed = False
        End If
        
        If chkPrinting.Value = 1 Then
            !PrintingAllowed = True
        Else
            !PrintingAllowed = False
        End If
        
'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------
        
        
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
    On Error Resume Next
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
        !AuthorityID = Val(dtcAuthority.BoundText)
        
        If chkPasswordResetting.Value = 1 Then
            !NeedPasswordReset = True
        Else
            !NeedPasswordReset = False
        End If
        
    
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------

'File
        If chkBackUp.Value = 1 Then
            !BackupAllowed = True
        Else
            !BackupAllowed = False
        End If
        
        If chkRestore.Value = 1 Then
            !RestoreAllowed = True
        Else
            !RestoreAllowed = False
        End If
        
'Edit

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
'
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
        
        If chkAuthority.Value = 1 Then
            !AuthorityAllowed = True
        Else
            !AuthorityAllowed = False
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
        
        If chkItemSuppliers.Value = 1 Then
            !ItemSuppiersAllowed = True
        Else
            !ItemSuppiersAllowed = False
        End If
        
        
        If chkItemSuppiersEdit.Value = 1 Then
            !ItemSuppiersEditAllowed = True
        Else
            !ItemSuppiersEditAllowed = False
        End If
        
'Milk Collection

        If chkDailyCollection.Value = 1 Then
            !DailyCollectionAllowed = True
        Else
            !DailyCollectionAllowed = False
        End If
        
        If chkGRN.Value = 1 Then
            !GoodRecieveNoteAllowed = True
        Else
            !GoodRecieveNoteAllowed = False
        End If
        
        If chkCumulativeReport.Value = 1 Then
            !CumulativeReportAllowed = True
        Else
            !CumulativeReportAllowed = False
        End If
        
        If chkGenerateCCPA.Value = 1 Then
            !GenarateCollectingCenterPayAdviceAllowed = True
        Else
            !GenarateCollectingCenterPayAdviceAllowed = False
        End If

        If chkPrintCCPA.Value = 1 Then
            !PrintCollectingCenterPayAdviceAllowed = True
        Else
            !PrintCollectingCenterPayAdviceAllowed = False
        End If
        
        If chkMilkPayAdvice.Value = 1 Then
            !MilkPayAdviceAllowed = True
        Else
            !MilkPayAdviceAllowed = False
        End If
        

        If chkCumulativeReportPrint.Value = 1 Then
            !CumulativeReportPrintAllowed = True
        Else
            !CumulativeReportPrintAllowed = False
        End If
        
'Issue & Payments

        If chkItemIssue.Value = 1 Then
            !ItemIssueAllowed = True
        Else
            !ItemIssueAllowed = False
        End If
        
        If chkItemPurchase.Value = 1 Then
            !ItemPurchaseAllowed = True
        Else
            !ItemPurchaseAllowed = False
        End If
        
        If chkAddDeductions.Value = 1 Then
            !AddDeductionsAllowed = True
        Else
            !AddDeductionsAllowed = False
        End If

        If chkAdditionalComm.Value = 1 Then
            !AdditionalCommisionsAllowed = True
        Else
            !AdditionalCommisionsAllowed = False
        End If
        
        If chkadditionalDeduct.Value = 1 Then
            !AdditionalDeductionsAllowed = True
        Else
            !AdditionalDeductionsAllowed = False
        End If
        
        If chkExpences.Value = 1 Then
            !ExpencesAllowed = True
        Else
            !ExpencesAllowed = False
        End If
        
        If chkIncome.Value = 1 Then
            !IncomeAllowed = True
        Else
            !IncomeAllowed = False
        End If
        

        If chkIssuePaymentDelete.Value = 1 Then
            !IssuePaymentDeleteAllowed = True
        Else
            !IssuePaymentDeleteAllowed = False
        End If

'Back Office

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
        
        If chkOutstanding.Value = 1 Then
            !OutstandingAllowed = True
        Else
            !OutstandingAllowed = False
        End If
        
         If chkReports.Value = 1 Then
            !ReportsAllowed = True
        Else
            !ReportsAllowed = False
        End If
        
        If chkProfits.Value = 1 Then
            !ProfitsAllowed = True
        Else
            !ProfitsAllowed = False
        End If
    
         If chkDetectErrors.Value = 1 Then
            !DetectErrorsAllowed = True
        Else
            !DetectErrorsAllowed = False
        End If
        

        If chkApprovePaymentsSave.Value = 1 Then
            !ApprovePaymentSaveAllowed = True
        Else
            !ApprovePaymentSaveAllowed = False
        End If

'Options

        If chkDatabase.Value = 1 Then
            !DatabaseAllowed = True
        Else
            !DatabaseAllowed = False
        End If
        
        If chkPrinting.Value = 1 Then
            !PrintingAllowed = True
        Else
            !PrintingAllowed = False
        End If
        
'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------

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
    If dtcAuthority.Visible = True Then
        If IsNumeric(dtcAuthority.BoundText) = False Then
            tr = MsgBox("You have not selected an authority", vbCritical, "Authority")
            dtcAuthority.SetFocus
            Exit Function
        End If
    End If
'    If UserAuthority <> 1 And dtcAuthority.BoundText = 1 Then
'        tr = MsgBox("You have not allowed to set the authority as an administrator", vbCritical, "Not Authorised")
'        dtcAuthority.SetFocus
'        Exit Function
'    End If
     If txtUserName.Visible = True Then
        If UserNameAvailable = False Then
            tr = MsgBox("User Name Already Exisits")
            txtUserName.SetFocus
            Exit Function
        End If
    End If
    CanAdd = True
End Function

Private Sub dtcAuthority_Click(Area As Integer)
    If Val(dtcAuthority.BoundText) = 6 Then
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------
'File
    chkBackUp.Value = 1
    chkRestore.Value = 1
    
'Edit
    chkCollectingCenter.Value = 1
    chkFarmers.Value = 1
    chkPrices.Value = 1
    chkItem.Value = 1
    chkStaff.Value = 1
    chkAuthority.Value = 1
    chkExpenceCategory.Value = 1
    chkIncomeCategory.Value = 1
    chkItemSuppliers.Value = 1
    
    chkItemSuppiersEdit.Value = 1
    
'Milk Collection
    chkDailyCollection.Value = 1
    chkGRN.Value = 1
    chkCumulativeReport.Value = 1
    chkGenerateCCPA.Value = 1
    chkPrintCCPA.Value = 1
    chkMilkPayAdvice.Value = 1
    
    chkCumulativeReportPrint.Value = 1
    
'Issue & Payments
    chkItemIssue.Value = 1
    chkItemPurchase.Value = 1
    chkAddDeductions.Value = 1
    chkAdditionalComm.Value = 1
    chkadditionalDeduct.Value = 1
    chkExpences.Value = 1
    chkIncome.Value = 1
    
    chkIssuePaymentDelete.Value = 1
    
'Back Office
    chkGIPayments.Value = 1
    chkPIPayments.Value = 1
    chkConfirmPayments.Value = 1
    chkApprovals.Value = 1
    chkOutstanding.Value = 1
    chkReports.Value = 1
    chkProfits.Value = 1
    chkDetectErrors.Value = 1
    
    chkApprovePaymentsSave.Value = 1
    
'Options
    chkDatabase.Value = 1
    chkPrinting.Value = 1

'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------

    End If

    
    If Val(dtcAuthority.BoundText) = 3 Then
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------

'File
    chkBackUp.Value = 0
    chkRestore.Value = 0
    
'Edit
    chkCollectingCenter.Value = 0
    chkFarmers.Value = 0
    chkPrices.Value = 0
    chkItem.Value = 0
    chkStaff.Value = 0
    chkAuthority.Value = 0
    chkExpenceCategory.Value = 0
    chkIncomeCategory.Value = 0
    chkItemSuppliers.Value = 1
    
    chkItemSuppiersEdit.Value = 0
    
'Milk Collection
    chkDailyCollection.Value = 1
    chkGRN.Value = 1
    chkCumulativeReport.Value = 1
    chkGenerateCCPA.Value = 1
    chkPrintCCPA.Value = 1
    chkMilkPayAdvice.Value = 1
    
    chkCumulativeReportPrint.Value = 0
    
'Issue & Payments
    chkItemIssue.Value = 1
    chkItemPurchase.Value = 1
    chkAddDeductions.Value = 1
    chkAdditionalComm.Value = 1
    chkadditionalDeduct.Value = 1
    chkExpences.Value = 1
    chkIncome.Value = 1
    
    chkIssuePaymentDelete.Value = 0
    
'Back Office
    chkGIPayments.Value = 0
    chkPIPayments.Value = 0
    chkConfirmPayments.Value = 0
    chkApprovals.Value = 0
    chkOutstanding.Value = 1
    chkReports.Value = 1
    chkProfits.Value = 1
    chkDetectErrors.Value = 0
    
    chkApprovePaymentsSave.Value = 0
    
'Options
    chkDatabase.Value = 0
    chkPrinting.Value = 0


'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------

    End If
    
    If Val(dtcAuthority.BoundText) = 4 Then
    
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------

'File
    chkBackUp.Value = 0
    chkRestore.Value = 0
    
'Edit
    chkCollectingCenter.Value = 0
    chkFarmers.Value = 1
    chkPrices.Value = 0
    chkItem.Value = 1
    chkStaff.Value = 0
    chkAuthority.Value = 0
    chkExpenceCategory.Value = 1
    chkIncomeCategory.Value = 1
    chkItemSuppliers.Value = 0
    
    chkItemSuppiersEdit.Value = 0
    
'Milk Collection
    chkDailyCollection.Value = 1
    chkGRN.Value = 1
    chkCumulativeReport.Value = 1
    chkGenerateCCPA.Value = 1
    chkPrintCCPA.Value = 1
    chkMilkPayAdvice.Value = 1
    
    chkCumulativeReportPrint.Value = 1
    
'Issue & Payments
    chkItemIssue.Value = 1
    chkItemPurchase.Value = 1
    chkAddDeductions.Value = 1
    chkAdditionalComm.Value = 1
    chkadditionalDeduct.Value = 1
    chkExpences.Value = 1
    chkIncome.Value = 1
    
    chkIssuePaymentDelete.Value = 0
    
'Back Office
    chkGIPayments.Value = 0
    chkPIPayments.Value = 0
    chkConfirmPayments.Value = 0
    chkApprovals.Value = 0
    chkOutstanding.Value = 1
    chkReports.Value = 1
    chkProfits.Value = 1
    chkDetectErrors.Value = 0
    
    chkApprovePaymentsSave.Value = 1
    
'Options
    chkDatabase.Value = 0
    chkPrinting.Value = 0


'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------


    End If
    
    If Val(dtcAuthority.BoundText) = 5 Then
    
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------
'File
    chkBackUp.Value = 0
    chkRestore.Value = 0
    
'Edit
    chkCollectingCenter.Value = 1
    chkFarmers.Value = 1
    chkPrices.Value = 1
    chkItem.Value = 1
    chkStaff.Value = 1
    chkAuthority.Value = 0
    chkExpenceCategory.Value = 1
    chkIncomeCategory.Value = 1
    chkItemSuppliers.Value = 1
    
    chkItemSuppiersEdit.Value = 1
    
'Milk Collection
    chkDailyCollection.Value = 1
    chkGRN.Value = 1
    chkCumulativeReport.Value = 1
    chkGenerateCCPA.Value = 1
    chkPrintCCPA.Value = 1
    chkMilkPayAdvice.Value = 1
    
    chkCumulativeReportPrint.Value = 1
    
'Issue & Payments
    chkItemIssue.Value = 1
    chkItemPurchase.Value = 1
    chkAddDeductions.Value = 1
    chkAdditionalComm.Value = 1
    chkadditionalDeduct.Value = 1
    chkExpences.Value = 1
    chkIncome.Value = 1
    
    chkIssuePaymentDelete.Value = 1
    
'Back Office
    chkGIPayments.Value = 1
    chkPIPayments.Value = 1
    chkConfirmPayments.Value = 1
    chkApprovals.Value = 1
    chkOutstanding.Value = 1
    chkReports.Value = 1
    chkProfits.Value = 1
    chkDetectErrors.Value = 1
    
    chkApprovePaymentsSave.Value = 1
    
'Options
    chkDatabase.Value = 0
    chkPrinting.Value = 0

'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------
    
    End If

    If Val(dtcAuthority.BoundText) = 2 Then
    
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------
'File
    chkBackUp.Value = 0
    chkRestore.Value = 0
    
'Edit
    chkCollectingCenter.Value = 0
    chkFarmers.Value = 0
    chkPrices.Value = 0
    chkItem.Value = 0
    chkStaff.Value = 0
    chkAuthority.Value = 0
    chkExpenceCategory.Value = 0
    chkIncomeCategory.Value = 0
    chkItemSuppliers.Value = 0
    
    chkItemSuppiersEdit.Value = 0
    
'Milk Collection
    chkDailyCollection.Value = 1
    chkGRN.Value = 0
    chkCumulativeReport.Value = 1
    chkGenerateCCPA.Value = 0
    chkPrintCCPA.Value = 0
    chkMilkPayAdvice.Value = 1
    
    chkCumulativeReportPrint.Value = 1
    
'Issue & Payments
    chkItemIssue.Value = 0
    chkItemPurchase.Value = 0
    chkAddDeductions.Value = 0
    chkAdditionalComm.Value = 0
    chkadditionalDeduct.Value = 0
    chkExpences.Value = 0
    chkIncome.Value = 0
    
    chkIssuePaymentDelete.Value = 0
    
'Back Office
    chkGIPayments.Value = 1
    chkPIPayments.Value = 1
    chkConfirmPayments.Value = 1
    chkApprovals.Value = 1
    chkOutstanding.Value = 1
    chkReports.Value = 1
    chkProfits.Value = 1
    chkDetectErrors.Value = 1
    
    chkApprovePaymentsSave.Value = 1
    
'Options
    chkDatabase.Value = 0
    chkPrinting.Value = 0

'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------
    
    End If

    If Val(dtcAuthority.BoundText) = 1 Then
    
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------
'File
    chkBackUp.Value = 0
    chkRestore.Value = 0
    
'Edit
    chkCollectingCenter.Value = 1
    chkFarmers.Value = 1
    chkPrices.Value = 1
    chkItem.Value = 1
    chkStaff.Value = 1
    chkAuthority.Value = 0
    chkExpenceCategory.Value = 1
    chkIncomeCategory.Value = 1
    chkItemSuppliers.Value = 1
    
    chkItemSuppiersEdit.Value = 1
    
'Milk Collection
    chkDailyCollection.Value = 1
    chkGRN.Value = 1
    chkCumulativeReport.Value = 1
    chkGenerateCCPA.Value = 1
    chkPrintCCPA.Value = 1
    chkMilkPayAdvice.Value = 1
    
    chkCumulativeReportPrint.Value = 1
    
'Issue & Payments
    chkItemIssue.Value = 1
    chkItemPurchase.Value = 1
    chkAddDeductions.Value = 1
    chkAdditionalComm.Value = 1
    chkadditionalDeduct.Value = 1
    chkExpences.Value = 1
    chkIncome.Value = 1
    
    chkIssuePaymentDelete.Value = 1
    
'Back Office
    chkGIPayments.Value = 1
    chkPIPayments.Value = 1
    chkConfirmPayments.Value = 1
    chkApprovals.Value = 1
    chkOutstanding.Value = 1
    chkReports.Value = 1
    chkProfits.Value = 1
    chkDetectErrors.Value = 1
    
    chkApprovePaymentsSave.Value = 1
    
'Options
    chkDatabase.Value = 0
    chkPrinting.Value = 0

'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------
    
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
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------
'File
    chkBackUp.Enabled = False
    chkRestore.Enabled = False
    
'Edit
    chkCollectingCenter.Enabled = False
    chkFarmers.Enabled = False
    chkPrices.Enabled = False
    chkItem.Enabled = False
    chkStaff.Enabled = False
    chkAuthority.Enabled = False
    chkExpenceCategory.Enabled = False
    chkIncomeCategory.Enabled = False
    chkItemSuppliers.Enabled = False
    
    chkItemSuppiersEdit.Enabled = False
    
'Milk Collection
    chkDailyCollection.Enabled = False
    chkGRN.Enabled = False
    chkCumulativeReport.Enabled = False
    chkGenerateCCPA.Enabled = False
    chkPrintCCPA.Enabled = False
    chkMilkPayAdvice.Enabled = False
    
    chkCumulativeReportPrint.Enabled = False
    
'Issue & Payments
    chkItemIssue.Enabled = False
    chkItemPurchase.Enabled = False
    chkAddDeductions.Enabled = False
    chkAdditionalComm.Enabled = False
    chkadditionalDeduct.Enabled = False
    chkExpences.Enabled = False
    chkIncome.Enabled = False
    
    chkIssuePaymentDelete.Enabled = False
    
'Back Office
    chkGIPayments.Enabled = False
    chkPIPayments.Enabled = False
    chkConfirmPayments.Enabled = False
    chkApprovals.Enabled = False
    chkOutstanding.Enabled = False
    chkReports.Enabled = False
    chkProfits.Enabled = False
    chkDetectErrors.Enabled = False
    
    chkApprovePaymentsSave.Enabled = False
    
'Options
    chkDatabase.Enabled = False
    chkPrinting.Enabled = False

'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------
     
        
        btnDelete.Visible = False
        bttnAdd.Visible = False
        
        End If
        
Select Case UserAuthorityLevel
    
        Case Authority.SuperUser '5
            FrameFile.Visible = False
            FrameEdit.Visible = False
            FrameMilkCollection.Visible = False
            FrameIssuePayments.Visible = False
            FrameBackOffice.Visible = False
            FrameOptions.Visible = False
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
        
        If !NeedPasswordReset = True Then chkPasswordResetting.Value = 1
        
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------

'File
        If !BackupAllowed = True Then chkBackUp.Value = 1
        If !RestoreAllowed = True Then chkRestore.Value = 1
        
'Edit
        If !CollectingCenterAllowed = True Then chkCollectingCenter.Value = 1
        If !FarmersAllowed = True Then chkFarmers.Value = 1
        If !PricesAllowed = True Then chkPrices.Value = 1
        If !ItemAllowed = True Then chkItem.Value = 1
        If !StaffsAllowed = True Then chkStaff.Value = 1
        If !AuthorityAllowed = True Then chkAuthority.Value = 1
        If !IncomeCategoryAllowed = True Then chkIncomeCategory.Value = 1
        If !ExpenceCategoryAllowed = True Then chkExpenceCategory.Value = 1
        If !ItemSuppiersAllowed = True Then chkItemSuppliers.Value = 1
        
        If !ItemSuppiersEditAllowed = True Then chkItemSuppiersEdit.Value = 1
        
'Milk Collection
        If !DailyCollectionAllowed = True Then chkDailyCollection.Value = 1
        If !GoodRecieveNoteAllowed = True Then chkGRN.Value = 1
        If !CumulativeReportAllowed = True Then chkCumulativeReport.Value = 1
        If !GenarateCollectingCenterPayAdviceAllowed = True Then chkGenerateCCPA.Value = 1
        If !PrintCollectingCenterPayAdviceAllowed = True Then chkPrintCCPA.Value = 1
        If !MilkPayAdviceAllowed = True Then chkMilkPayAdvice.Value = 1
        
        If !CumulativeReportPrintAllowed = True Then chkCumulativeReportPrint.Value = 1
        
'Issue & Payments
        If !ItemIssueAllowed = True Then chkItemIssue.Value = 1
        If !ItemPurchaseAllowed = True Then chkItemPurchase.Value = 1
        If !AddDeductionsAllowed = True Then chkAddDeductions.Value = 1
        If !AdditionalCommisionsAllowed = True Then chkAdditionalComm.Value = 1
        If !AdditionalDeductionsAllowed = True Then chkadditionalDeduct.Value = 1
        If !ExpencesAllowed = True Then chkExpences.Value = 1
        If !IncomeAllowed = True Then chkIncome.Value = 1
        
        If !IssuePaymentDeleteAllowed = True Then chkIssuePaymentDelete.Value = 1
        
'Back Office
        If !GenerateIndividualPaymentsAllowed = True Then chkGIPayments.Value = 1
        If !PrintIndividualPaymentsAllowed = True Then chkPIPayments.Value = 1
        If !ConfirmPaymentsAllowed = True Then chkConfirmPayments.Value = 1
        If !ApprovalsAllowed = True Then chkApprovals.Value = 1
        If !OutstandingAllowed = True Then chkOutstanding.Value = 1
        If !ReportsAllowed = True Then chkReports.Value = 1
        If !ProfitsAllowed = True Then chkProfits.Value = 1
        If !DetectErrorsAllowed = True Then chkDetectErrors.Value = 1
        
        If !ApprovePaymentSaveAllowed = True Then chkApprovePaymentsSave.Value = 1
        
'Options
        If !DatabaseAllowed = True Then chkDatabase.Value = 1
        If !PrintingAllowed = True Then chkPrinting.Value = 1

'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------
             
'        Call ListComments
        If .State = 1 Then .Close
    End With
End Sub

'Private Sub ListComments()
'    With rsComments
'        If .State = 1 Then .Close
'        temSql = "SELECT * from tblStaffComment where StaffID = " & dtcStaffDetails.BoundText
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        lstComments.Clear
'        lstCommentIDs.Clear
'        If .RecordCount > 0 Then
'            While .EOF = False
'                lstComments.AddItem Format(!Date, "dd MMM yyyy") & vbTab & !Comment
'                lstCommentIDs.AddItem !StaffCommentID
'                .MoveNext
'            Wend
'        End If
'        .Close
'    End With
'End Sub


'Private Sub lstComments_Click()
'    lstCommentIDs.ListIndex = lstComments.ListIndex
'    StaffCommentIDTx = Val(lstCommentIDs.Text)
'    Unload frmStaffCommentDisplay
'    frmStaffCommentDisplay.Show
'End Sub

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

