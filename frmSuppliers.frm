VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSuppliers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Farmer"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   10800
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   4080
      TabIndex        =   7
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   12091
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "Professional"
      TabPicture(0)   =   "frmSuppliers.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAdditionalCommision"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblFixedCommision"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label33"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label34"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmbCity"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmbPaymentScheme"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbPaymentMethod"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmbBank"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmbCollectingCenter"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtAdditionalCommision"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chkFarmer"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtAccount"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtCode"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtName"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "chkCommision"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "chkThroughCollector"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmbCommision"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmbCollector"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtFixedCommision"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtAccountHolderName"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "chkCollector"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Personal"
      TabPicture(1)   =   "frmSuppliers.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "Label12"
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(4)=   "Label17"
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(6)=   "Label15"
      Tab(1).Control(7)=   "Label18"
      Tab(1).Control(8)=   "Label19"
      Tab(1).Control(9)=   "cmbSex"
      Tab(1).Control(10)=   "cmbTitle"
      Tab(1).Control(11)=   "txtNIC"
      Tab(1).Control(12)=   "txtAddress"
      Tab(1).Control(13)=   "dtpDOB"
      Tab(1).Control(14)=   "txtNOK"
      Tab(1).Control(15)=   "txtPhone"
      Tab(1).Control(16)=   "txteMail"
      Tab(1).Control(17)=   "txtFullName"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Family"
      TabPicture(2)   =   "frmSuppliers.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label20"
      Tab(2).Control(1)=   "Label21"
      Tab(2).Control(2)=   "Label22"
      Tab(2).Control(3)=   "Label23"
      Tab(2).Control(4)=   "Label24"
      Tab(2).Control(5)=   "Label25"
      Tab(2).Control(6)=   "Label26"
      Tab(2).Control(7)=   "Label27"
      Tab(2).Control(8)=   "Label28"
      Tab(2).Control(9)=   "Label29"
      Tab(2).Control(10)=   "Label30"
      Tab(2).Control(11)=   "Label31"
      Tab(2).Control(12)=   "Label32"
      Tab(2).Control(13)=   "dtpFifthChildDOB"
      Tab(2).Control(14)=   "dtpFourthChildDOB"
      Tab(2).Control(15)=   "dtpThirdChildDOB"
      Tab(2).Control(16)=   "dtpSecondChild"
      Tab(2).Control(17)=   "dtpFirstChild"
      Tab(2).Control(18)=   "txtSporseName"
      Tab(2).Control(19)=   "txtFirstChildName"
      Tab(2).Control(20)=   "dtpSporseDOB"
      Tab(2).Control(21)=   "txtSporseNIC"
      Tab(2).Control(22)=   "txtSecondChildName"
      Tab(2).Control(23)=   "txtThirdChildName"
      Tab(2).Control(24)=   "txtFourthChildName"
      Tab(2).Control(25)=   "txtFifthChildName"
      Tab(2).ControlCount=   26
      Begin VB.CheckBox chkCollector 
         Caption         =   "Collector"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtFifthChildName 
         Height          =   375
         Left            =   -72720
         TabIndex        =   77
         Top             =   5760
         Width           =   4095
      End
      Begin VB.TextBox txtFourthChildName 
         Height          =   375
         Left            =   -72720
         TabIndex        =   73
         Top             =   4800
         Width           =   4095
      End
      Begin VB.TextBox txtThirdChildName 
         Height          =   375
         Left            =   -72720
         TabIndex        =   69
         Top             =   3840
         Width           =   4095
      End
      Begin VB.TextBox txtSecondChildName 
         Height          =   375
         Left            =   -72720
         TabIndex        =   66
         Top             =   2880
         Width           =   4095
      End
      Begin VB.TextBox txtSporseNIC 
         Height          =   375
         Left            =   -72720
         TabIndex        =   65
         Top             =   1440
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpSporseDOB 
         Height          =   375
         Left            =   -72720
         TabIndex        =   59
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   393216
         Format          =   31850496
         CurrentDate     =   39812
      End
      Begin VB.TextBox txtFirstChildName 
         Height          =   375
         Left            =   -72720
         TabIndex        =   58
         Top             =   1920
         Width           =   4095
      End
      Begin VB.TextBox txtSporseName 
         Height          =   375
         Left            =   -72720
         TabIndex        =   56
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtFullName 
         Height          =   375
         Left            =   -73560
         TabIndex        =   32
         Top             =   900
         Width           =   4455
      End
      Begin VB.TextBox txtAccountHolderName 
         Height          =   375
         Left            =   1920
         TabIndex        =   28
         Top             =   4980
         Width           =   4095
      End
      Begin VB.TextBox txtFixedCommision 
         Height          =   375
         Left            =   3600
         TabIndex        =   18
         Top             =   3000
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo cmbCollector 
         Height          =   360
         Left            =   2400
         TabIndex        =   21
         Top             =   3480
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCommision 
         Height          =   360
         Left            =   1920
         TabIndex        =   16
         Top             =   2520
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txteMail 
         Height          =   375
         Left            =   -73560
         TabIndex        =   48
         Top             =   4740
         Width           =   4455
      End
      Begin VB.TextBox txtPhone 
         Height          =   375
         Left            =   -73560
         TabIndex        =   46
         Top             =   4260
         Width           =   4455
      End
      Begin VB.TextBox txtNOK 
         Height          =   1095
         Left            =   -73560
         MultiLine       =   -1  'True
         TabIndex        =   50
         Top             =   5220
         Width           =   4455
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   -73560
         TabIndex        =   40
         Top             =   1860
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         _Version        =   393216
         CalendarForeColor=   16711680
         CalendarTitleForeColor=   16711680
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   114753539
         CurrentDate     =   39705
      End
      Begin VB.TextBox txtAddress 
         Height          =   1335
         Left            =   -73560
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   2820
         Width           =   4455
      End
      Begin VB.TextBox txtNIC 
         Height          =   375
         Left            =   -73560
         TabIndex        =   42
         Top             =   2340
         Width           =   4455
      End
      Begin VB.CheckBox chkThroughCollector 
         Caption         =   "Through Collector"
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   3480
         Width           =   2175
      End
      Begin VB.CheckBox chkCommision 
         Caption         =   "Commision"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtCode 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox txtAccount 
         Height          =   375
         Left            =   1920
         TabIndex        =   35
         Top             =   6360
         Width           =   4095
      End
      Begin VB.CheckBox chkFarmer 
         Caption         =   "Farmer"
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtAdditionalCommision 
         Height          =   375
         Left            =   3600
         TabIndex        =   19
         Top             =   3000
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo cmbCollectingCenter 
         Height          =   360
         Left            =   1920
         TabIndex        =   12
         Top             =   1560
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbBank 
         Height          =   360
         Left            =   1920
         TabIndex        =   31
         Top             =   5460
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbPaymentMethod 
         Height          =   360
         Left            =   1920
         TabIndex        =   25
         Top             =   4500
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbTitle 
         Height          =   360
         Left            =   -73560
         TabIndex        =   36
         Top             =   1380
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbSex 
         Height          =   360
         Left            =   -70800
         TabIndex        =   38
         Top             =   1380
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker dtpFirstChild 
         Height          =   375
         Left            =   -72720
         TabIndex        =   67
         Top             =   2400
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   393216
         Format          =   142147584
         CurrentDate     =   39812
      End
      Begin MSComCtl2.DTPicker dtpSecondChild 
         Height          =   375
         Left            =   -72720
         TabIndex        =   68
         Top             =   3360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   393216
         Format          =   142147584
         CurrentDate     =   39812
      End
      Begin MSComCtl2.DTPicker dtpThirdChildDOB 
         Height          =   375
         Left            =   -72720
         TabIndex        =   70
         Top             =   4320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   393216
         Format          =   142147584
         CurrentDate     =   39812
      End
      Begin MSComCtl2.DTPicker dtpFourthChildDOB 
         Height          =   375
         Left            =   -72720
         TabIndex        =   74
         Top             =   5280
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   393216
         Format          =   142147584
         CurrentDate     =   39812
      End
      Begin MSComCtl2.DTPicker dtpFifthChildDOB 
         Height          =   375
         Left            =   -72720
         TabIndex        =   78
         Top             =   6240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   393216
         Format          =   142147584
         CurrentDate     =   39812
      End
      Begin MSDataListLib.DataCombo cmbPaymentScheme 
         Height          =   360
         Left            =   1920
         TabIndex        =   23
         Top             =   3960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCity 
         Height          =   360
         Left            =   1920
         TabIndex        =   84
         Top             =   5880
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label34 
         Caption         =   "Branch"
         Height          =   375
         Left            =   360
         TabIndex        =   83
         Top             =   5880
         Width           =   2775
      End
      Begin VB.Label Label33 
         Caption         =   "Payment Scheme"
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   3960
         Width           =   2775
      End
      Begin VB.Label Label32 
         Caption         =   "Fifth Child Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   80
         Top             =   5760
         Width           =   2055
      End
      Begin VB.Label Label31 
         Caption         =   "Fifth Child DOB"
         Height          =   255
         Left            =   -74760
         TabIndex        =   79
         Top             =   6240
         Width           =   2055
      End
      Begin VB.Label Label30 
         Caption         =   "Fourth Child Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   76
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label Label29 
         Caption         =   "Fourth Child DOB"
         Height          =   255
         Left            =   -74760
         TabIndex        =   75
         Top             =   5280
         Width           =   2055
      End
      Begin VB.Label Label28 
         Caption         =   "Third Child Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   72
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label27 
         Caption         =   "Third Child DOB"
         Height          =   255
         Left            =   -74760
         TabIndex        =   71
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label Label26 
         Caption         =   "Sporse N I C No"
         Height          =   255
         Left            =   -74760
         TabIndex        =   64
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label25 
         Caption         =   "Second Child DOB"
         Height          =   255
         Left            =   -74760
         TabIndex        =   63
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label24 
         Caption         =   "Second Child Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   62
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label23 
         Caption         =   "First Child DOB"
         Height          =   255
         Left            =   -74760
         TabIndex        =   61
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label22 
         Caption         =   "First Child Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   60
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label21 
         Caption         =   "Sporse Date Of Birth"
         Height          =   255
         Left            =   -74760
         TabIndex        =   57
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label20 
         Caption         =   "Sporse Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   55
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label19 
         Caption         =   "Full Name"
         Height          =   375
         Left            =   -74760
         TabIndex        =   30
         Top             =   900
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "Account Holder"
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   4980
         Width           =   2775
      End
      Begin VB.Label lblFixedCommision 
         Caption         =   "Fixed Commision"
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   3000
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label18 
         Caption         =   "eMail"
         Height          =   375
         Left            =   -74760
         TabIndex        =   47
         Top             =   4740
         Width           =   2775
      End
      Begin VB.Label Label15 
         Caption         =   "Phone"
         Height          =   375
         Left            =   -74760
         TabIndex        =   45
         Top             =   4260
         Width           =   2775
      End
      Begin VB.Label Label6 
         Caption         =   "Next of Kin"
         Height          =   375
         Left            =   -74760
         TabIndex        =   49
         Top             =   5220
         Width           =   2775
      End
      Begin VB.Label Label17 
         Caption         =   "Date of Birth"
         Height          =   375
         Left            =   -74760
         TabIndex        =   39
         Top             =   1860
         Width           =   2775
      End
      Begin VB.Label Label16 
         Caption         =   "Address"
         Height          =   375
         Left            =   -74760
         TabIndex        =   43
         Top             =   2820
         Width           =   2775
      End
      Begin VB.Label Label13 
         Caption         =   "Title"
         Height          =   375
         Left            =   -74760
         TabIndex        =   34
         Top             =   1380
         Width           =   2775
      End
      Begin VB.Label Label12 
         Caption         =   "Sex"
         Height          =   375
         Left            =   -71640
         TabIndex        =   37
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "NIC No"
         Height          =   375
         Left            =   -74760
         TabIndex        =   41
         Top             =   2340
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Code"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Collecting Center"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "Bank"
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   5460
         Width           =   2775
      End
      Begin VB.Label Label7 
         Caption         =   "Account"
         Height          =   375
         Left            =   360
         TabIndex        =   33
         Top             =   6360
         Width           =   2775
      End
      Begin VB.Label lblAdditionalCommision 
         Caption         =   "Addion Commision"
         Height          =   375
         Left            =   1800
         TabIndex        =   27
         Top             =   3000
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label10 
         Caption         =   "Payment Method"
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   4500
         Width           =   2775
      End
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9480
      TabIndex        =   54
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "C&lose"
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
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   4920
      TabIndex        =   51
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Save"
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
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   495
      Left            =   6240
      TabIndex        =   52
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Cancel"
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
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Add"
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
   Begin MSDataListLib.DataCombo cmbSupplier 
      Height          =   5460
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   9631
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Edit"
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Delete"
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
   Begin MSDataListLib.DataCombo cmbMainCollectingCenter 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   7560
      TabIndex        =   81
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Print"
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
   Begin btButtonEx.ButtonEx btnPrintAll 
      Height          =   495
      Left            =   120
      TabIndex        =   82
      Top             =   7560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Print All"
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
   Begin btButtonEx.ButtonEx btnBranch 
      Height          =   495
      Left            =   1440
      TabIndex        =   85
      Top             =   7560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "New &Bank Branch"
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
   Begin VB.Label Label8 
      Caption         =   "C&ollecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "&Farmer"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsSupplier As New ADODB.Recordset
    Dim rsViewSupplier As New ADODB.Recordset
    Dim rsViewCollector As New ADODB.Recordset
    Dim rsViewBranch As New ADODB.Recordset
    Dim temSQL As String

Private Sub btnAdd_Click()
    Dim temString As String
    If IsNumeric(cmbSupplier.BoundText) = True Then
        temString = Empty
    Else
        temString = cmbSupplier.Text
    End If
    cmbSupplier.Text = Empty
    ClearValues
    txtName.Text = temString
    EditMode
    txtName.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnAdd_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtName.SetFocus
    End If
End Sub

Private Sub btnBranch_Click()
    frmBankBranch.Show
    frmBankBranch.ZOrder 0
End Sub

Private Sub btnCancel_Click()
    ClearValues
    SelectMode
    cmbMainCollectingCenter_Change
    cmbSupplier.Text = Empty
'    cmbCollectingCenter.SetFocus
End Sub

Private Sub btnCancel_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbCollectingCenter.SetFocus
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
On Error GoTo eh
    Dim i As Integer
    i = MsgBox("Are you sure?", vbYesNo)
    If i = vbNo Then Exit Sub

    If IsNumeric(cmbSupplier.BoundText) = False Then
        MsgBox "Nothing to Delete"
        cmbSupplier.SetFocus
        Exit Sub
    End If
    With rsSupplier
        If .State = 1 Then .Close
        temSQL = "Select * from tblSupplier where SupplierID = " & Val(cmbSupplier.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            .Update
        Else
            MsgBox "Error"
        End If
        .Close
        
        ClearValues
        FillCombos
        cmbSupplier.Text = Empty
        cmbSupplier.SetFocus
        
        
        Exit Sub
eh:
        MsgBox Err.Description
        If .State = 1 Then .CancelUpdate
        If .State = 1 Then .Close
    End With
    ClearValues
    FillCombos
    cmbMainCollectingCenter_Change
    cmbSupplier.Text = Empty
    cmbSupplier.SetFocus
    
End Sub

Private Sub btnDelete_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbCollectingCenter.SetFocus
    End If
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbSupplier.BoundText) = False Then
        MsgBox "Nothing to Edit"
        cmbSupplier.SetFocus
        Exit Sub
    End If
    EditMode
    txtName.SetFocus
    On Error Resume Next
    SendKeys "{home}+{end}"
    
End Sub

Private Sub btnEdit_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtName.SetFocus
    End If
End Sub

Private Sub btnPrintAll_Click()
    
    Dim rsReport As New ADODB.Recordset
    Dim i As Integer
    With rsReport
        If .State = 1 Then .Close
        temSQL = "SELECT tblSupplier.Supplier, tblSupplier.SupplierCode, tblSupplier.Address, tblSupplier.AccountNo, tblSupplier.NICNo, tblSupplier.CollectingCenterID " & _
                    "FROM tblSupplier " & _
                    "WHERE tblSupplier.CollectingCenterID = " & Val(cmbMainCollectingCenter.BoundText) & " AND Deleted = 0  " & _
                    "ORDER BY tblSupplier.Supplier, tblSupplier.SupplierCode;"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtrSuppliers
        Set .DataSource = rsReport
        
        .Sections("Section4").Controls("lblIName").Caption = InstitutionName
        .Sections("Section4").Controls("lblIAddress1").Caption = InstitutionAddressLine1
        .Sections("Section4").Controls("lblIAddress2").Caption = InstitutionAddressLine2
        .Sections("Section4").Controls("lblIAddress3").Caption = InstitutionAddressLine3
        
        .Sections("Section4").Controls("lblCC2").Caption = cmbMainCollectingCenter.Text
        
        .Sections("Section1").Controls("txtSupplierCode").DataField = "SupplierCode"
        .Sections("Section1").Controls("txtSupplierName").DataField = "Supplier"
        .Sections("Section1").Controls("txtSupplierAddress").DataField = "Address"
        .Sections("Section1").Controls("txtNIC").DataField = "NICNo"
        .Sections("Section1").Controls("txtAcc").DataField = "AccountNo"
        
        .Show
    End With

End Sub

Private Sub btnSave_Click()
    If Trim(txtName.Text) = "" Then
        MsgBox "Nothing to Save"
        txtName.SetFocus
        Exit Sub
    End If
    
    If Trim(UCase(cmbPaymentMethod.Text)) = Trim(UCase("To Bank")) And Trim(txtAccount.Text) = Empty Then
        MsgBox "If you have selected the bank, please enter a bank account number"
        txtAccount.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbSupplier.BoundText) = False Then
        If SupplierNameRecords > 0 Then
            MsgBox "The Supplier name " & Trim(txtName.Text) & " already exists under the collecting center " & cmbCollectingCenter.Text & ". Please enter another name"
            txtName.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        If SupplierCodeRecords > 0 Then
            MsgBox "The Supplier Code " & Trim(txtCode.Text) & " already exists under the collecting center " & cmbCollectingCenter.Text & ". Please enter another code"
            txtCode.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        SaveNew
    Else
        If Trim(txtName.Text) <> cmbSupplier.Text Then
            If SupplierNameRecords > 0 Then
                MsgBox "The Supplier name " & Trim(txtName.Text) & " already exists under the collecting center " & cmbCollectingCenter.Text & ". Please enter another name"
                txtName.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            End If
            If SupplierCodeRecords > 0 Then
                MsgBox "The Supplier Code " & Trim(txtCode.Text) & " already exists under the collecting center " & cmbCollectingCenter.Text & ". Please enter another code"
                txtCode.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            End If
        End If
        SaveOld
    End If
    ClearValues
    SelectMode
    Dim temString As String
    Dim temID As Long
    temString = cmbMainCollectingCenter.Text
    temID = Val(cmbMainCollectingCenter.BoundText)
    FillCombos
    cmbMainCollectingCenter.Text = Empty
    cmbMainCollectingCenter_Change
    cmbMainCollectingCenter.BoundText = 0
    cmbMainCollectingCenter.BoundText = temID
    cmbSupplier.Text = Empty
    cmbSupplier.SetFocus
    
End Sub

Private Function SupplierNameRecords() As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblSupplier Where Deleted = 0  AND COllectingCenterID = " & Val(cmbCollectingCenter.BoundText) & " AND Supplier = '" & Trim(txtName.Text) & "'"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        SupplierNameRecords = .RecordCount
        .Close
    End With
    Set rsTem = Nothing
End Function

Private Function SupplierCodeRecords() As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblSupplier Where Deleted = 0  AND COllectingCenterID = " & Val(cmbCollectingCenter.BoundText) & " AND SupplierCode = '" & Trim(txtName.Text) & "'"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        SupplierCodeRecords = .RecordCount
        .Close
    End With
    Set rsTem = Nothing
End Function

Private Sub btnSave_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbCollectingCenter.SetFocus
    End If
End Sub

Private Sub bttnPrint_Click()
    Dim tab1 As Long
    Dim tab2 As Long
    Dim tab3 As Long
    Dim tab4 As Long
    Dim Tab5 As Long
    Dim Tab6 As Long
    
    tab1 = 5
    tab2 = 47
    tab3 = 40
    tab4 = 70
    Tab5 = 52
    Tab6 = 45
    
    Printer.Font.Bold = True
    Printer.Print Tab(tab2); "Lucky lanka Milk Processing Co.Ltd."
    Printer.Font.Bold = False
    Printer.Print Tab(Tab5); "Bibulewela, Karagoda, Uyangoda. "
    Printer.Print Tab(Tab6); "Tel:04122926502/0412293032 Fax:0412292831"
    Printer.Print
    Printer.Font.Bold = True
    Printer.Font.Underline = True
    Printer.Print Tab(tab1); "Farmers Professional Details"
    Printer.Font.Underline = False
    Printer.Font.Bold = False
    Printer.Print
    Printer.Print Tab(tab1); "Farmer Name  :    ";
    Printer.Print Tab(tab3); txtName.Text
    Printer.Print Tab(tab1); "Farmer Code  :    ";
    Printer.Print Tab(tab3); txtCode.Text
    Printer.Print Tab(tab1); "Collecting Center  :  ";
    Printer.Print Tab(tab3); cmbCollectingCenter.Text
    Printer.Print Tab(tab1); "Payment Method  : ";
    Printer.Print Tab(tab3); cmbPaymentMethod.Text
    Printer.Print Tab(tab1); "Account Holder  : ";
    Printer.Print Tab(tab3); txtAccountHolderName.Text
    Printer.Print Tab(tab1); "Bank  :   ";
    Printer.Print Tab(tab3); cmbBank.Text
    Printer.Print Tab(tab1); "Account Number  : ";
    Printer.Print Tab(tab3); txtAccount.Text
    Printer.Print
    Printer.Font.Bold = True
    Printer.Font.Underline = True
    Printer.Print Tab(tab1); "Farmers Personal Details"
    Printer.Font.Underline = False
    Printer.Font.Bold = False
    Printer.Print
    Printer.Print Tab(tab1); "Full Name  :  ";
    Printer.Print Tab(tab3); txtFullName.Text
    Printer.Print Tab(tab1); "Address  :    ";
    Printer.Print Tab(tab3); txtAddress.Text
    Printer.Print Tab(tab1); "City  :   ";
    Printer.Print Tab(tab3); cmbCity.Text
    Printer.Print Tab(tab1); "Date Of Birth  :  ";
    Printer.Print Tab(tab3); dtpDOB.Value
    Printer.Print Tab(tab1); "Sex  :    ";
    Printer.Print Tab(tab3); cmbSex.Text
    Printer.Print Tab(tab1); "NIC No  : ";
    Printer.Print Tab(tab3); txtNIC.Text
    Printer.Print Tab(tab1); "Telephone  :  ";
    Printer.Print Tab(tab3); txtPhone.Text
    Printer.Print Tab(tab1); "eMail Address  :  ";
    Printer.Print Tab(tab3); txtEmail.Text
    Printer.Print
    Printer.Font.Bold = True
    Printer.Font.Underline = True
    Printer.Print Tab(tab1); "Farmers Family Detalis"
    Printer.Font.Underline = False
    Printer.Font.Bold = False
    Printer.Print
    Printer.Print Tab(tab1); "Sporse Name  :    ";
    Printer.Print Tab(tab3); txtSporseName.Text
    Printer.Print Tab(tab1); "Sporse NIC No  :  ";
    Printer.Print Tab(tab3); txtSporseNIC.Text
    Printer.Print Tab(tab1); "Sporse Birthday : ";
    Printer.Print Tab(tab3); dtpSporseDOB.Value
    Printer.Print Tab(tab1); "First Child Name  :   ";
    Printer.Print Tab(tab3); txtFirstChildName.Text
    Printer.Print Tab(tab1); "First Child Birthday  :   ";
    Printer.Print Tab(tab3); dtpFirstChild.Value
    Printer.Print Tab(tab1); "Second Child Name  :  ";
    Printer.Print Tab(tab3); txtSecondChildName.Text
    Printer.Print Tab(tab1); "Second Child Birthday  :  ";
    Printer.Print Tab(tab3); dtpSecondChild
    Printer.Print Tab(tab1); "Third Child Name  :   ";
    Printer.Print Tab(tab3); txtThirdChildName.Text
    Printer.Print Tab(tab1); "Third Child Birthday  :   ";
    Printer.Print Tab(tab3); dtpThirdChildDOB
    Printer.Print Tab(tab1); "Fourth Child Name  :  ";
    Printer.Print Tab(tab3); txtFourthChildName.Text
    Printer.Print Tab(tab1); "Fourth Child Birthday  :  ";
    Printer.Print Tab(tab3); dtpFourthChildDOB.Value
    Printer.Print Tab(tab1); "Fifth Child Name  :   ";
    Printer.Print Tab(tab3); txtFifthChildName.Text
    Printer.Print Tab(tab1); "Fifth Child Birthday  :   ";
    Printer.Print Tab(tab3); dtpFifthChildDOB.Value
    Printer.EndDoc
End Sub

Private Sub chkCollector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        chkCommision.SetFocus
    End If
End Sub

Private Sub chkCommision_Click()
    If chkCommision.Value = 1 Then
        cmbCommision.Visible = True
    Else
        cmbCommision.Visible = False
        cmbCommision.Text = Empty
        lblAdditionalCommision.Visible = False
        txtAdditionalCommision.Visible = False
        txtAdditionalCommision.Text = Empty
    End If
End Sub

Private Sub chkCommision_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmbCommision.Visible = True Then
            cmbCommision.SetFocus
        Else
            chkThroughCollector.SetFocus
        End If
    End If
End Sub

Private Sub chkFarmer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        chkCollector.SetFocus
    End If
End Sub

Private Sub chkThroughCollector_Click()
    If chkThroughCollector.Value = 1 Then
        cmbCollector.Visible = True
    Else
        cmbCollector.Visible = False
        cmbCollector.Text = Empty
    End If
End Sub

Private Sub chkThroughCollector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmbCollector.Visible = True Then
            cmbPaymentMethod.SetFocus
        Else
            cmbPaymentMethod.SetFocus
        End If
        cmbPaymentMethod.SetFocus
    End If
End Sub

Private Sub cmbBank_Change()
    With rsViewBranch
        If .State = 1 Then .Close
        temSQL = "Select * from tblCity where BankID = " & Val(cmbBank.BoundText) & " AND Deleted = 0  order by City"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCity
        Set .RowSource = rsViewBranch
        .ListField = "City"
        .BoundText = "CityID"
    End With
End Sub

Private Sub cmbBank_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbCity.SetFocus
    End If
End Sub

Private Sub cmbBranch_Click(Area As Integer)

End Sub

Private Sub cmbBranch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtAccount.SetFocus
    End If
End Sub

Private Sub cmbCity_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtAccount.SetFocus
    End If
End Sub

Private Sub cmbCollectingCenter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        chkFarmer.SetFocus
    End If
End Sub

Private Sub cmbCollector_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbPaymentMethod.SetFocus
    End If
End Sub

Private Sub cmbCommision_Change()
    If UCase(cmbCommision.Text) = UCase("By Volume And Additional Commission") Then
        lblAdditionalCommision.Visible = True
        txtAdditionalCommision.Visible = True
        lblFixedCommision.Visible = False
        txtFixedCommision.Visible = False
        txtFixedCommision.Text = Empty
    ElseIf UCase(cmbCommision.Text) = UCase("Fixed") Then
        lblAdditionalCommision.Visible = False
        txtAdditionalCommision.Visible = False
        txtAdditionalCommision.Text = Empty
        lblFixedCommision.Visible = True
        txtFixedCommision.Visible = True
    ElseIf UCase(cmbCommision.Text) = UCase("By Volume") Then
        lblAdditionalCommision.Visible = False
        txtAdditionalCommision.Visible = False
        txtAdditionalCommision.Text = Empty
        lblFixedCommision.Visible = False
        txtFixedCommision.Visible = False
        txtFixedCommision.Text = Empty
    ElseIf UCase(cmbCommision.Text) = UCase("By Volume And Fixed Commission") Then
        lblAdditionalCommision.Visible = False
        txtAdditionalCommision.Visible = False
        txtAdditionalCommision.Text = Empty
        lblFixedCommision.Visible = True
        txtFixedCommision.Visible = True
    End If
End Sub

Private Sub cmbMainCollectingCenter_Change()
    With rsViewSupplier
        If .State = 1 Then .Close
        If IsNumeric(cmbMainCollectingCenter.BoundText) = True Then
            temSQL = "Select * from tblSUpplier where CollectingCenterID = " & Val(cmbMainCollectingCenter.BoundText) & " And Deleted = 0  Order by Supplier"
        Else
            temSQL = "Select * from tblSUpplier where Deleted = 0  Order by Supplier"
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbSupplier
        Set .RowSource = rsViewSupplier
        .ListField = "Supplier"
        .BoundColumn = "SupplierID"
    End With
    With rsViewCollector
        If .State = 1 Then .Close
        If IsNumeric(cmbMainCollectingCenter.BoundText) = True Then
            temSQL = "Select * from tblSupplier where Deleted = 0  AND Collector = 1 And CollectingCenterID =" & Val(cmbMainCollectingCenter.BoundText) & " order by Supplier "
        Else
            temSQL = "Select * from tblSupplier where Deleted = 0  AND Collector = 1 order by Supplier"
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCollector
        Set .RowSource = rsViewCollector
        .ListField = "Supplier"
        .BoundColumn = "SupplierID"
    End With
    
End Sub

Private Sub cmbMainCollectingCenter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbMainCollectingCenter.Text = Empty
    ElseIf KeyCode = vbKeyReturn Then
        cmbSupplier.SetFocus
    End If
End Sub


Private Sub cmbPaymentMethod_Change()
    If UCase(cmbPaymentMethod.Text) = UCase("Cash") Then
        Label9.Visible = False
        txtAccountHolderName.Visible = False
        Label5.Visible = False
        cmbBank.Visible = False
        Label7.Visible = False
        txtAccount.Visible = False
    ElseIf UCase(cmbPaymentMethod.Text) = UCase("To Bank") Then
        Label9.Visible = True
        txtAccountHolderName.Visible = True
        Label5.Visible = True
        cmbBank.Visible = True
        Label7.Visible = True
        txtAccount.Visible = True
    ElseIf UCase(cmbPaymentMethod.Text) = UCase("Cheque") Then
        Label9.Visible = False
        txtAccountHolderName.Visible = False
        Label5.Visible = False
        cmbBank.Visible = False
        Label7.Visible = False
        txtAccount.Visible = False
    End If
End Sub

Private Sub cmbPaymentMethod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtAccountHolderName.SetFocus
    End If
End Sub


Private Sub cmbPaymentScheme_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbPaymentMethod.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbPaymentScheme.Text = Empty
    End If
End Sub

Private Sub cmbSupplier_Change()
    If IsNumeric(cmbSupplier.BoundText) = False Then
        ClearValues
    Else
        DisplayDetails
    End If
End Sub

Private Sub Form_Load()
    
    ClearValues
    SelectMode
    FillCombos
    cmbMainCollectingCenter_Change
    
            btnAdd.Enabled = False
            btnEdit.Enabled = False
            btnDelete.Enabled = False
    
    
    Select Case UserAuthorityLevel
        Case Authority.Administrator
            btnAdd.Enabled = True
            btnEdit.Enabled = True
            btnDelete.Enabled = True
    
        Case Authority.PowerUser '4
        btnDelete.Visible = False
        Case Else
    
    End Select

    
End Sub

Private Sub FillCombos()
    Dim MainCC As New clsFillCombos
    MainCC.FillAnyCombo cmbMainCollectingCenter, "CollectingCenter", True
    Dim CollectingCentre As New clsFillCombos
    CollectingCentre.FillAnyCombo cmbCollectingCenter, "CollectingCenter", True
    Dim Bank As New clsFillCombos
    Bank.FillAnyCombo cmbBank, "Bank", True
    Dim City As New clsFillCombos
    City.FillAnyCombo cmbCity, "City", True
    Dim Sex As New clsFillCombos
    Sex.FillAnyCombo cmbSex, "Sex", True
    Dim Title As New clsFillCombos
    Title.FillAnyCombo cmbTitle, "Title", True
    Dim PaymentMethod As New clsFillCombos
    PaymentMethod.FillAnyCombo cmbPaymentMethod, "PaymentMethod", True
    Dim Commision As New clsFillCombos
    Commision.FillAnyCombo cmbCommision, "Commision", True
    Dim PS As New clsFillCombos
    PS.FillAnyCombo cmbPaymentScheme, "PaymentScheme", True
    
    With rsViewCollector
        If .State = 1 Then .Close
        temSQL = "Select * from tblSupplier where Deleted = 0  AND Collector  = 1 And CollectingCenterID =" & Val(cmbCollectingCenter.BoundText) & "  order by Supplier "
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCollector
        Set .RowSource = rsViewCollector
        .ListField = "Supplier"
        .BoundColumn = "SupplierID"
    End With
    cmbSupplier_Change
End Sub

Private Sub EditMode()

    cmbSupplier.Enabled = False
    btnDelete.Enabled = False
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    
    txtName.Enabled = True
    cmbCollectingCenter.Enabled = True
    txtCode.Enabled = True
    txtAccount.Enabled = True
    cmbBank.Enabled = True
    
    txtFixedCommision.Enabled = True
    txtAdditionalCommision.Enabled = True
    txtAddress.Enabled = True
    txtEmail.Enabled = True
    txtNIC.Enabled = True
    txtNOK.Enabled = True
    txtPhone.Enabled = True
    txtAccountHolderName.Enabled = True
    txtFullName.Enabled = True
    
    dtpDOB.Enabled = True
    
    cmbCity.Enabled = True
    cmbCollector.Enabled = True
    cmbCommision.Enabled = True
    cmbPaymentMethod.Enabled = True
    cmbSex.Enabled = True
    cmbTitle.Enabled = True
    cmbPaymentScheme.Enabled = True
    
    chkCollector.Enabled = True
    chkCommision.Enabled = True
    chkFarmer.Enabled = True
    chkThroughCollector.Enabled = True
    
    txtSporseName.Enabled = True
    txtSporseNIC.Enabled = True
    dtpSporseDOB.Enabled = True
    txtFirstChildName.Enabled = True
    dtpFirstChild.Enabled = True
    txtSecondChildName.Enabled = True
    dtpSecondChild.Enabled = True
    txtThirdChildName.Enabled = True
    dtpThirdChildDOB.Enabled = True
    txtFourthChildName.Enabled = True
    dtpFourthChildDOB.Enabled = True
    txtFifthChildName.Enabled = True
    dtpFifthChildDOB.Enabled = True
    
    btnSave.Enabled = True
    btnCancel.Enabled = True

End Sub

Private Sub SelectMode()

    
    cmbSupplier.Enabled = True
    btnDelete.Enabled = True
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    
    txtName.Enabled = False
    txtCode.Enabled = False
    cmbCollectingCenter.Enabled = False
    txtAccount.Enabled = False
    cmbBank.Enabled = False

    txtAdditionalCommision.Enabled = False
    txtFixedCommision.Enabled = False
    txtAddress.Enabled = False
    txtEmail.Enabled = False
    txtNIC.Enabled = False
    txtNOK.Enabled = False
    txtPhone.Enabled = False
    txtAccountHolderName.Enabled = False
    txtFullName.Enabled = False
    
    dtpDOB.Enabled = False
    
    chkCollector.Enabled = False
    chkCommision.Enabled = False
    chkFarmer.Enabled = False
    chkThroughCollector.Enabled = False
    
    cmbPaymentScheme.Enabled = False
    cmbCity.Enabled = False
    cmbCollector.Enabled = False
    cmbCommision.Enabled = False
    cmbPaymentMethod.Enabled = False
    cmbSex.Enabled = False
    cmbTitle.Enabled = False

    txtSporseName.Enabled = False
    txtSporseNIC.Enabled = False
    dtpSporseDOB.Enabled = False
    txtFirstChildName.Enabled = False
    dtpFirstChild.Enabled = False
    txtSecondChildName.Enabled = False
    dtpSecondChild.Enabled = False
    txtThirdChildName.Enabled = False
    dtpThirdChildDOB.Enabled = False
    txtFourthChildName.Enabled = False
    dtpFourthChildDOB.Enabled = False
    txtFifthChildName.Enabled = False
    dtpFifthChildDOB.Enabled = False
      
    btnSave.Enabled = False
    btnCancel.Enabled = False

End Sub

Private Sub SaveNew()
    With rsSupplier
        If .State = 1 Then .Close
        temSQL = "Select * from tblSupplier where SupplierID =0"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Supplier = Trim(txtName.Text)
        !SupplierCode = txtCode.Text
        !CollectingCenterID = Val(cmbCollectingCenter.BoundText)
        !BankID = Val(cmbBank.BoundText)
        !AccountNo = txtAccount.Text
        If chkFarmer.Value = 1 Then
            !Farmer = True
        Else
            !Farmer = False
        End If
        If chkCollector.Value = 1 Then
            !Collector = True
        Else
            !Collector = False
        End If
        If chkCommision.Value = 1 Then
            !Commision = True
        Else
            !Commision = False
        End If
        !CommisionType = Val(cmbCommision.BoundText)
        !AdditionalCommision = Val(txtAdditionalCommision.Text)
        !FixedCommision = Val(txtFixedCommision.Text)
        !FullName = txtFullName.Text
        !AccountHolder = txtAccountHolderName.Text
        !PaymentSchemeID = Val(cmbPaymentScheme.BoundText)
        If chkThroughCollector.Value = 1 Then
            !ThroughCollector = True
        Else
            !ThroughCollector = False
        End If
        !CommisionCollectorID = Val(cmbCollector.BoundText)
        !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        !TitleID = Val(cmbTitle.BoundText)
        !SexID = Val(cmbSex.BoundText)
        !CityID = Val(cmbCity.BoundText)
        !NOK = txtNOK.Text
        !DOB = Format(dtpDOB.Value, "dd MMMM yyyy")
        !NICNo = txtNIC.Text
        !Phone = txtPhone.Text
        !eMail = txtEmail.Text
        !Address = txtAddress.Text
        
        !SporseName = txtSporseName.Text
        !SporseDOB = Format(dtpSporseDOB.Value, "dd MMMM yyyy")
        !SporseNIC = txtSporseNIC.Text
        !FirstChildNAme = txtFirstChildName.Text
        !FirstChildDOB = Format(dtpFirstChild.Value, "dd MMMM yyyy")
        !SecondChildName = txtSecondChildName.Text
        !SecondChildDOB = Format(dtpSecondChild.Value, "dd MMMM yyyy")
        !ThirdChildName = txtThirdChildName.Text
        !ThirdChildDOB = Format(dtpThirdChildDOB.Value, "dd MMMM yyyy")
        !FourthChildName = txtFourthChildName.Text
        !FourthChildDOB = Format(dtpFourthChildDOB.Value, "dd MMMM yyyy")
        !FifthChildName = txtFifthChildName.Text
        !FifthChildDOB = Format(dtpFifthChildDOB.Value, "dd MMMM yyyy")

        .Update
        .Close
    End With
End Sub

Private Sub SaveOld()
    With rsSupplier
        If .State = 1 Then .Close
        temSQL = "Select * from tblSupplier where SupplierID =" & Val(cmbSupplier.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Supplier = Trim(txtName.Text)
            !SupplierCode = txtCode.Text
            !CollectingCenterID = Val(cmbCollectingCenter.BoundText)
            !BankID = Val(cmbBank.BoundText)
            !AccountNo = txtAccount.Text
            !FullName = txtFullName.Text
            !AccountHolder = txtAccountHolderName.Text
            If chkFarmer.Value = 1 Then
                !Farmer = True
            Else
                !Farmer = False
            End If
            If chkCollector.Value = 1 Then
                !Collector = True
            Else
                !Collector = False
            End If
            If chkCommision.Value = 1 Then
                !Commision = True
            Else
                !Commision = False
            End If
            !PaymentSchemeID = Val(cmbPaymentScheme.BoundText)
            !CommisionType = Val(cmbCommision.BoundText)
            !AdditionalCommision = Val(txtAdditionalCommision.Text)
            !FixedCommision = Val(txtFixedCommision.Text)
            If chkThroughCollector.Value = 1 Then
                !ThroughCollector = True
            Else
                !ThroughCollector = False
            End If
            !CommisionCollectorID = Val(cmbCollector.BoundText)
            !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
            !TitleID = Val(cmbTitle.BoundText)
            !SexID = Val(cmbSex.BoundText)
            !CityID = Val(cmbCity.BoundText)
            !NOK = txtNOK.Text
            !DOB = Format(dtpDOB.Value, "dd MMMM yyyy")
            !NICNo = txtNIC.Text
            !Phone = txtPhone.Text
            !eMail = txtEmail.Text
            !Address = txtAddress.Text
            
            !SporseName = txtSporseName.Text
            !SporseDOB = Format(dtpSporseDOB.Value, "dd MMMM yyyy")
            !SporseNIC = txtSporseNIC.Text
            !FirstChildNAme = txtFirstChildName.Text
            !FirstChildDOB = Format(dtpFirstChild.Value, "dd MMMM yyyy")
            !SecondChildName = txtSecondChildName.Text
            !SecondChildDOB = Format(dtpSecondChild.Value, "dd MMMM yyyy")
            !ThirdChildName = txtThirdChildName.Text
            !ThirdChildDOB = Format(dtpThirdChildDOB.Value, "dd MMMM yyyy")
            !FourthChildName = txtFourthChildName.Text
            !FourthChildDOB = Format(dtpFourthChildDOB.Value, "dd MMMM yyyy")
            !FifthChildName = txtFifthChildName.Text
            !FifthChildDOB = Format(dtpFifthChildDOB.Value, "dd MMMM yyyy")

            .Update
        Else
            MsgBox "Error"
        End If
        .Close
    End With
End Sub

Private Sub ClearValues()
    txtName.Text = Empty
    cmbCollectingCenter.Text = Empty
    txtCode.Text = Empty
    cmbBank.Text = Empty
    txtAccount.Text = Empty


    txtAdditionalCommision.Text = Empty
    txtFixedCommision.Text = Empty
    txtAddress.Text = Empty
    txtEmail.Text = Empty
    txtNIC.Text = Empty
    txtNOK.Text = Empty
    txtPhone.Text = Empty
    txtAccountHolderName.Text = Empty
    txtFullName.Text = Empty
    
    chkCollector.Value = 0
    chkCommision.Value = 0
    chkFarmer.Value = 0
    chkThroughCollector.Value = 0
    
    cmbCity.Text = Empty
    cmbCollector.Text = Empty
    cmbCommision.Text = Empty
    cmbPaymentMethod.Text = Empty
    cmbSex.Text = Empty
    cmbTitle.Text = Empty
    
    cmbPaymentScheme.Text = Empty
    
    txtSporseName.Text = Empty
    txtSporseNIC.Text = Empty
    txtFirstChildName.Text = Empty
    txtSecondChildName.Text = Empty
    txtThirdChildName.Text = Empty
    txtFourthChildName.Text = Empty
    txtFifthChildName.Text = Empty
End Sub


Private Sub DisplayDetails()
    ClearValues
    With rsSupplier
        If .State = 1 Then .Close
        temSQL = "Select * from tblSupplier where SupplierID =" & Val(cmbSupplier.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!Supplier) = False Then txtName.Text = !Supplier
            If IsNull(!SupplierCode) = False Then txtCode.Text = !SupplierCode
            If IsNull(!CollectingCenterID) = False Then cmbCollectingCenter.BoundText = !CollectingCenterID
            If IsNull(!AccountNo) = False Then txtAccount.Text = !AccountNo
            If IsNull(!BankID) = False Then cmbBank.BoundText = !BankID
            If !Farmer = True Then chkFarmer.Value = 1
            If !Collector = True Then chkCollector.Value = 1
            If !Commision = True Then chkCommision.Value = 1
            If IsNull(!CommisionType) = False Then cmbCommision.BoundText = !CommisionType
            If IsNull(!AdditionalCommision) = False Then txtAdditionalCommision.Text = !AdditionalCommision
            If IsNull(!FixedCommision) = False Then txtFixedCommision.Text = !FixedCommision
            If !ThroughCollector = True Then chkThroughCollector.Value = 1
            If IsNull(!CommisionCollectorID) = False Then cmbCollector.BoundText = !CommisionCollectorID
            If IsNull(!PaymentMethodID) = False Then cmbPaymentMethod.BoundText = !PaymentMethodID
            If IsNull(!TitleID) = False Then cmbTitle.BoundText = !TitleID
            If IsNull(!SexID) = False Then cmbSex.BoundText = !SexID
            If IsNull(!DOB) = False Then dtpDOB.Value = !DOB
            If IsNull(!NICNo) = False Then txtNIC.Text = !NICNo
            If IsNull(!Address) = False Then txtAddress.Text = !Address
            If IsNull(!NOK) = False Then txtNOK.Text = !NOK
            If IsNull(!CityID) = False Then cmbCity.BoundText = !CityID
            If IsNull(!eMail) = False Then txtEmail.Text = !eMail
            If IsNull(!Phone) = False Then txtPhone.Text = !Phone
            If IsNull(!FullName) = False Then txtFullName.Text = !FullName
            If IsNull(!AccountHolder) = False Then txtAccountHolderName.Text = !AccountHolder
            If IsNull(!SporseName) = False Then txtSporseName.Text = !SporseName
            If IsNull(!SporseDOB) = False Then dtpSporseDOB.Value = !SporseDOB
            If IsNull(!SporseNIC) = False Then txtSporseNIC.Text = !SporseNIC
            If IsNull(!FirstChildNAme) = False Then txtFirstChildName.Text = !FirstChildNAme
            If IsNull(!FirstChildDOB) = False Then dtpFirstChild = !FirstChildDOB
            If IsNull(!SecondChildName) = False Then txtSecondChildName.Text = !SecondChildName
            If IsNull(!SecondChildDOB) = False Then dtpSecondChild = !SecondChildDOB
            If IsNull(!ThirdChildName) = False Then txtThirdChildName.Text = !ThirdChildName
            If IsNull(!ThirdChildDOB) = False Then dtpThirdChildDOB = !ThirdChildDOB
            If IsNull(!FourthChildName) = False Then txtFourthChildName.Text = !FourthChildName
            If IsNull(!FourthChildDOB) = False Then dtpFourthChildDOB = !FourthChildDOB
            If IsNull(!FifthChildName) = False Then txtFifthChildName.Text = !FifthChildName
            If IsNull(!FifthChildDOB) = False Then dtpFifthChildDOB = !FifthChildDOB
            If IsNull(!PaymentSchemeID) = False Then cmbPaymentScheme.BoundText = !PaymentSchemeID
        Else
            MsgBox "Error"
        End If
        .Close
    End With
End Sub

Private Sub txtAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnSave_Click
    End If
End Sub

Private Sub txtAccountHolderName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbBank.SetFocus
    End If
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbCollectingCenter.SetFocus
    End If
End Sub

Private Sub txtFixedCommision_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       chkThroughCollector.SetFocus
    End If
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtCode.SetFocus
    End If
End Sub
