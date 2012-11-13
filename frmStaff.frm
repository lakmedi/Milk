VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStaff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff Details"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12165
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
   ScaleHeight     =   9585
   ScaleWidth      =   12165
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameSearch 
      Height          =   8775
      Left            =   120
      TabIndex        =   63
      Top             =   120
      Width           =   4575
      Begin MSDataListLib.DataCombo dtcStaff 
         Height          =   7380
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   13018
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
         Text            =   ""
      End
      Begin btButtonEx.ButtonEx bttnAdd 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   8040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Appearance      =   3
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
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   495
         Left            =   2520
         TabIndex        =   2
         Top             =   8040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Appearance      =   3
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
      Begin btButtonEx.ButtonEx bttnDelete 
         Height          =   495
         Left            =   2520
         TabIndex        =   3
         Top             =   8040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Appearance      =   3
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
   End
   Begin VB.Frame frameDetails 
      Height          =   8775
      Left            =   4920
      TabIndex        =   37
      Top             =   120
      Width           =   6975
      Begin TabDlg.SSTab SSTabMain 
         Height          =   7695
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   13573
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Personal"
         TabPicture(0)   =   "frmStaff.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label21"
         Tab(0).Control(1)=   "Label20"
         Tab(0).Control(2)=   "Label1"
         Tab(0).Control(3)=   "Label11"
         Tab(0).Control(4)=   "Label12"
         Tab(0).Control(5)=   "Label13"
         Tab(0).Control(6)=   "Label18"
         Tab(0).Control(7)=   "Label26"
         Tab(0).Control(8)=   "SSTabSub"
         Tab(0).Control(9)=   "dtcSpeciality"
         Tab(0).Control(10)=   "dtcSex"
         Tab(0).Control(11)=   "dtcTitle"
         Tab(0).Control(12)=   "txtName"
         Tab(0).Control(13)=   "txtQualifications"
         Tab(0).Control(14)=   "txtRegistation"
         Tab(0).Control(15)=   "txtDesignation"
         Tab(0).Control(16)=   "txtListedName"
         Tab(0).ControlCount=   17
         TabCaption(1)   =   "Photos"
         TabPicture(1)   =   "frmStaff.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label14"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label8"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "imgSignature"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "imgPhoto"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "bttnSigDelete"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "bttnSigLoad"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "bttnPhotoDelete"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "txtPhoto"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "txtSignature"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "bttnPhotoLoad"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "Program"
         TabPicture(2)   =   "frmStaff.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label28"
         Tab(2).Control(1)=   "Label22"
         Tab(2).Control(2)=   "Label19"
         Tab(2).Control(3)=   "Label17"
         Tab(2).Control(4)=   "Label10"
         Tab(2).Control(5)=   "dtcAuthority"
         Tab(2).Control(6)=   "chkUser"
         Tab(2).Control(7)=   "txtReenterPassword"
         Tab(2).Control(8)=   "txtPassword"
         Tab(2).Control(9)=   "txtUserName"
         Tab(2).Control(10)=   "txtComments"
         Tab(2).ControlCount=   11
         Begin btButtonEx.ButtonEx bttnPhotoLoad 
            Height          =   255
            Left            =   480
            TabIndex        =   22
            Top             =   1200
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            Appearance      =   3
            Caption         =   "Load"
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
         Begin VB.TextBox txtSignature 
            Height          =   360
            Left            =   360
            TabIndex        =   66
            Top             =   6360
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtPhoto 
            Height          =   360
            Left            =   480
            TabIndex        =   65
            Top             =   3120
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtComments 
            Height          =   1320
            Left            =   -73200
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   3960
            Width           =   4815
         End
         Begin VB.TextBox txtListedName 
            Height          =   375
            Left            =   -72720
            MaxLength       =   100
            TabIndex        =   7
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox txtDesignation 
            Height          =   375
            Left            =   -72720
            MaxLength       =   100
            TabIndex        =   10
            Top             =   2880
            Width           =   3975
         End
         Begin VB.TextBox txtRegistation 
            Height          =   375
            Left            =   -72720
            MaxLength       =   100
            TabIndex        =   9
            Top             =   2400
            Width           =   3975
         End
         Begin VB.TextBox txtQualifications 
            Height          =   375
            Left            =   -72720
            MaxLength       =   100
            TabIndex        =   8
            Top             =   1920
            Width           =   3975
         End
         Begin VB.TextBox txtName 
            Height          =   375
            Left            =   -72720
            MaxLength       =   100
            TabIndex        =   6
            Top             =   960
            Width           =   3975
         End
         Begin VB.TextBox txtUserName 
            Height          =   375
            Left            =   -72120
            MaxLength       =   10
            TabIndex        =   27
            Top             =   1320
            Width           =   3615
         End
         Begin VB.TextBox txtPassword 
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   -72120
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   28
            Top             =   1920
            Width           =   3615
         End
         Begin VB.TextBox txtReenterPassword 
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   -72120
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   29
            Top             =   2520
            Width           =   3615
         End
         Begin VB.CheckBox chkUser 
            Caption         =   "User of the program"
            Height          =   495
            Left            =   -72120
            TabIndex        =   26
            Top             =   720
            Width           =   3495
         End
         Begin MSDataListLib.DataCombo dtcTitle 
            Height          =   360
            Left            =   -72720
            TabIndex        =   4
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin MSDataListLib.DataCombo dtcSex 
            Height          =   360
            Left            =   -70080
            TabIndex        =   5
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin MSDataListLib.DataCombo dtcSpeciality 
            Height          =   360
            Left            =   -72720
            TabIndex        =   11
            Top             =   3360
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin TabDlg.SSTab SSTabSub 
            Height          =   3615
            Left            =   -74760
            TabIndex        =   38
            Top             =   3840
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   6376
            _Version        =   393216
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "Private"
            TabPicture(0)   =   "frmStaff.frx":0054
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label27"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label5"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label4"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label3"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label2"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "txtPrivateMobile"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "txtPrivateEmail"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "txtPrivateFax"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "txtPrivateTel"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "txtPrivateAddress"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).ControlCount=   10
            TabCaption(1)   =   "Official"
            TabPicture(1)   =   "frmStaff.frx":0070
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lblOfficialWebsite"
            Tab(1).Control(1)=   "Label25"
            Tab(1).Control(2)=   "Label24"
            Tab(1).Control(3)=   "Label23"
            Tab(1).Control(4)=   "lblOfficialEmail"
            Tab(1).Control(5)=   "txtOfficialWebsite"
            Tab(1).Control(6)=   "txtOfficialAddress"
            Tab(1).Control(7)=   "txtOfficialTel"
            Tab(1).Control(8)=   "txtOfficialFax"
            Tab(1).Control(9)=   "txtOfficialEMail"
            Tab(1).ControlCount=   10
            Begin VB.TextBox txtOfficialEMail 
               Height          =   375
               Left            =   -73800
               TabIndex        =   15
               Top             =   2280
               Width           =   4695
            End
            Begin VB.TextBox txtOfficialFax 
               Height          =   375
               Left            =   -73800
               TabIndex        =   14
               Top             =   1800
               Width           =   4695
            End
            Begin VB.TextBox txtOfficialTel 
               Height          =   375
               Left            =   -73800
               TabIndex        =   13
               Top             =   1320
               Width           =   4695
            End
            Begin VB.TextBox txtOfficialAddress 
               Height          =   735
               Left            =   -73800
               MultiLine       =   -1  'True
               TabIndex        =   12
               Top             =   480
               Width           =   4695
            End
            Begin VB.TextBox txtOfficialWebsite 
               Height          =   375
               Left            =   -73800
               TabIndex        =   16
               Top             =   2760
               Width           =   4695
            End
            Begin VB.TextBox txtPrivateAddress 
               Height          =   735
               Left            =   1200
               MultiLine       =   -1  'True
               TabIndex        =   17
               Top             =   480
               Width           =   4695
            End
            Begin VB.TextBox txtPrivateTel 
               Height          =   375
               Left            =   1200
               MaxLength       =   100
               TabIndex        =   18
               Top             =   1320
               Width           =   4695
            End
            Begin VB.TextBox txtPrivateFax 
               Height          =   375
               Left            =   1200
               MaxLength       =   100
               TabIndex        =   20
               Top             =   2280
               Width           =   4695
            End
            Begin VB.TextBox txtPrivateEmail 
               Height          =   375
               Left            =   1200
               MaxLength       =   100
               TabIndex        =   21
               Top             =   2760
               Width           =   4695
            End
            Begin VB.TextBox txtPrivateMobile 
               Height          =   375
               Left            =   1200
               MaxLength       =   100
               TabIndex        =   19
               Top             =   1800
               Width           =   4695
            End
            Begin VB.Label lblOfficialEmail 
               BackStyle       =   0  'Transparent
               Caption         =   "E-Mail"
               Height          =   375
               Left            =   -74880
               TabIndex        =   48
               Top             =   2280
               Width           =   2175
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "Fax:"
               Height          =   375
               Left            =   -74880
               TabIndex        =   47
               Top             =   1800
               Width           =   2175
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "Telephone"
               Height          =   375
               Left            =   -74880
               TabIndex        =   46
               Top             =   1320
               Width           =   2175
            End
            Begin VB.Label Label25 
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               Height          =   375
               Left            =   -74880
               TabIndex        =   45
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label lblOfficialWebsite 
               BackStyle       =   0  'Transparent
               Caption         =   "Website"
               Height          =   375
               Left            =   -74880
               TabIndex        =   44
               Top             =   2760
               Width           =   2175
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               Height          =   375
               Left            =   120
               TabIndex        =   43
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Telephone"
               Height          =   375
               Left            =   120
               TabIndex        =   42
               Top             =   1320
               Width           =   2175
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Fax"
               Height          =   375
               Left            =   120
               TabIndex        =   41
               Top             =   2280
               Width           =   2175
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "E-Mail"
               Height          =   375
               Left            =   120
               TabIndex        =   40
               Top             =   2760
               Width           =   2175
            End
            Begin VB.Label Label27 
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile"
               Height          =   375
               Left            =   120
               TabIndex        =   39
               Top             =   1800
               Width           =   2175
            End
         End
         Begin MSDataListLib.DataCombo dtcAuthority 
            Height          =   360
            Left            =   -72120
            TabIndex        =   30
            Top             =   3120
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin btButtonEx.ButtonEx bttnPhotoDelete 
            Height          =   255
            Left            =   480
            TabIndex        =   23
            Top             =   1560
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            Appearance      =   3
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
         Begin btButtonEx.ButtonEx bttnSigLoad 
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   5520
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            Appearance      =   3
            Caption         =   "Load"
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
         Begin btButtonEx.ButtonEx bttnSigDelete 
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   5880
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            Appearance      =   3
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
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Comments"
            Height          =   375
            Left            =   -74760
            TabIndex        =   64
            Top             =   3960
            Width           =   2175
         End
         Begin VB.Image imgPhoto 
            BorderStyle     =   1  'Fixed Single
            Height          =   4575
            Left            =   1680
            Top             =   480
            Width           =   4815
         End
         Begin VB.Image imgSignature 
            BorderStyle     =   1  'Fixed Single
            Height          =   855
            Left            =   1680
            Top             =   5280
            Width           =   4815
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            Height          =   375
            Left            =   -74760
            TabIndex        =   62
            Top             =   3360
            Width           =   2175
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Listed Name"
            Height          =   375
            Left            =   -74760
            TabIndex        =   61
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Designation"
            Height          =   375
            Left            =   -74760
            TabIndex        =   60
            Top             =   2880
            Width           =   2175
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Registation"
            Height          =   375
            Left            =   -74760
            TabIndex        =   59
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Qualifications"
            Height          =   375
            Left            =   -74760
            TabIndex        =   58
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   375
            Left            =   -74760
            TabIndex        =   57
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Title"
            Height          =   375
            Left            =   -74760
            TabIndex        =   56
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label21 
            Caption         =   "Sex"
            Height          =   375
            Left            =   -70680
            TabIndex        =   55
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Signature"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   5160
            Width           =   2175
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Photo"
            Height          =   375
            Left            =   360
            TabIndex        =   53
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "User Name"
            Height          =   255
            Left            =   -74640
            TabIndex        =   52
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            Height          =   255
            Left            =   -74640
            TabIndex        =   51
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Re-enter Password"
            Height          =   255
            Left            =   -74640
            TabIndex        =   50
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Authority Level"
            Height          =   255
            Left            =   -74640
            TabIndex        =   49
            Top             =   3240
            Width           =   2055
         End
      End
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   8040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Appearance      =   3
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
      Begin btButtonEx.ButtonEx bttnCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   4800
         TabIndex        =   34
         Top             =   8040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "&Cancel"
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
         Height          =   495
         Left            =   1440
         TabIndex        =   33
         Top             =   8040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Appearance      =   3
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
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   495
      Left            =   10680
      TabIndex        =   35
      Top             =   9000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "E&xit"
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
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsStaff As New ADODB.Recordset
    Dim rsTitle As New ADODB.Recordset
    Dim rsAuthority As New ADODB.Recordset
    Dim rsSex As New ADODB.Recordset
    Dim rsSpeciality As New ADODB.Recordset
    Dim rsTemStaff As New ADODB.Recordset
    Dim temSQL As String
    Dim TemUserName As String
    Dim TemStaffID As Long

Private Sub BeforeAddEdit()
    frameSearch.Enabled = True
    frameDetails.Enabled = False
    bttnSave.Visible = False
    bttnChange.Visible = False
    bttnCancel.Visible = False
End Sub

Private Sub AfterAdd()
    frameSearch.Enabled = False
    frameDetails.Enabled = True
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnCancel.Visible = True
End Sub

Private Sub AfterEdit()
    frameSearch.Enabled = False
    frameDetails.Enabled = True
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
End Sub

Private Sub ClearValues()
    Me.txtComments.Text = Empty
    Me.txtDesignation.Text = Empty
    Me.txtListedName.Text = Empty
    Me.txtName.Text = Empty
    Me.txtOfficialAddress.Text = Empty
    Me.txtOfficialEMail.Text = Empty
    Me.txtOfficialFax.Text = Empty
    Me.txtOfficialTel.Text = Empty
    Me.txtOfficialWebsite.Text = Empty
    Me.txtPassword.Text = Empty
    Me.txtPrivateAddress.Text = Empty
    Me.txtPrivateEmail.Text = Empty
    Me.txtPrivateFax.Text = Empty
    Me.txtPrivateMobile.Text = Empty
    Me.txtPrivateTel.Text = Empty
    Me.txtQualifications.Text = Empty
    Me.txtReenterPassword.Text = Empty
    Me.txtRegistation.Text = Empty
    Me.txtUserName.Text = Empty
    Me.dtcAuthority.Text = Empty
    Me.dtcSex.Text = Empty
    Me.dtcSpeciality.Text = Empty
    Me.dtcStaff.Text = Empty
    Me.dtcTitle.Text = Empty
    Me.txtPhoto.Text = Empty
    Me.txtSignature.Text = Empty
    Me.chkUser.Value = 2
    imgPhoto.Picture = LoadPicture()
    imgSignature.Picture = LoadPicture()
End Sub

Private Function CanSave() As Boolean
    Dim tr As Integer
    CanSave = False
        If Trim(Me.txtListedName.Text) = Empty Then
            tr = MsgBox("You have not entered the Name to be listed", vbCritical, "Listed Name?")
            SSTabMain.Tab = 0
            txtListedName.SetFocus
            Exit Function
        End If
        If Trim(Me.txtName.Text) = Empty Then
            tr = MsgBox("You have not entered tha Name", vbCritical, "Name")
            SSTabMain.Tab = 0
            txtName.SetFocus
            Exit Function
        End If
        If Trim(Me.txtUserName.Text) = Empty And chkUser.Value = 1 Then
            tr = MsgBox("You have not entered a username", vbCritical, "UserName?")
            txtPassword.SetFocus
            SSTabMain.Tab = 2
            Exit Function
        End If
        If Trim(Me.txtPassword.Text) <> Trim(Me.txtReenterPassword.Text) Then
            tr = MsgBox("The passwords you entered are not matching", vbCritical, "Password?")
            txtPassword.SetFocus
            SSTabMain.Tab = 2
            Exit Function
        End If
        If IsNumeric(Me.dtcAuthority.BoundText) = False Then
            tr = MsgBox("You have not selected an authority level")
            SSTabMain.Tab = 0
            dtcAuthority.SetFocus
            Exit Function
        End If
        If chkUser.Value = 1 And TemUserName <> txtUserName.Text Then
            With rsTemStaff
                If .State = 1 Then .Close
                temSQL = "SELECT * from tblstaff where username = '" & EncreptedWord(txtUserName.Text) & "'"
                .Open temSQL, cnnStores
                If .RecordCount > 0 Then
                    tr = MsgBox("The username is already taken.", vbCritical, "Username")
                    SSTabMain.Tab = 2
                    txtUserName.SetFocus
                    SendKeys "{home}+{end}"
                    .Close
                    Exit Function
                End If
                .Close
            End With
        End If
    CanSave = True
End Function

Private Sub LocateDetails()
    With rsTemStaff
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblstaff where staffID = " & dtcStaff.BoundText
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Not IsNull(!TitleID) Then dtcTitle.BoundText = !TitleID
            If Not IsNull(!SexID) Then dtcSex.BoundText = !SexID
            If Not IsNull(!Name) Then txtName.Text = !Name
            If Not IsNull(!listedName) Then txtListedName.Text = !listedName
            If Not IsNull(!Qualifications) Then txtQualifications.Text = !Qualifications
            If Not IsNull(!Registation) Then txtRegistation.Text = !Registation
            If Not IsNull(!Designation) Then txtDesignation.Text = !Designation
            If Not IsNull(!SpecialityID) Then dtcSpeciality.BoundText = !SpecialityID
            If Not IsNull(!AuthorityID) Then dtcAuthority.BoundText = !AuthorityID
            If Not IsNull(!PrivateAddress) Then txtPrivateAddress.Text = !PrivateAddress
            If Not IsNull(!PrivatePhone) Then txtPrivateTel.Text = !PrivatePhone
            If Not IsNull(!PrivateFax) Then txtPrivateFax.Text = !PrivateFax
            If Not IsNull(!PrivateEmail) Then txtPrivateEmail.Text = !PrivateEmail
            If Not IsNull(!MobilePhone) Then txtPrivateMobile.Text = !MobilePhone
            If Not IsNull(!OfficialAddress) Then txtOfficialAddress.Text = !OfficialAddress
            If Not IsNull(!OfficialPhone) Then txtOfficialTel.Text = !OfficialPhone
            If Not IsNull(!OfficialFax) Then txtOfficialFax.Text = !OfficialFax
            If Not IsNull(!OfficialEmail) Then txtOfficialEMail.Text = !OfficialEmail
            If Not IsNull(!Website) Then txtOfficialWebsite.Text = !Website
            If Not IsNull(!Comments) Then txtComments.Text = !Comments
            If Not IsNull(!Photo) Then
                txtPhoto.Text = !Photo
                imgPhoto.Stretch = True
                On Error Resume Next
                imgPhoto.Picture = LoadPicture(txtPhoto.Text)
            End If
            If Not IsNull(!Signature) Then
                txtSignature.Text = !Signature
                imgSignature.Stretch = True
                imgSignature.Picture = LoadPicture(txtSignature.Text)
            End If
            If Not IsNull(!UserName) Then txtUserName.Text = DecreptedWord(!UserName)
            If Not IsNull(!Password) Then txtPassword.Text = DecreptedWord(!Password)
            txtReenterPassword.Text = txtPassword.Text
            If !User = True Then
                chkUser.Value = 1
            Else
                chkUser.Value = 0
            End If
        End If
        .Close
    End With
End Sub

Private Sub SaveDetails()
    With rsTemStaff
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblstaff"
        .Open temSQL, cnnStores, adOpenDynamic, adLockOptimistic
        .AddNew
        If IsNumeric(dtcTitle.BoundText) Then !TitleID = dtcTitle.BoundText
        If IsNumeric(dtcSex.BoundText) Then !SexID = dtcSex.BoundText
        !Name = txtName.Text
        !listedName = txtListedName.Text
        !Qualifications = txtQualifications.Text
        !Registation = txtRegistation.Text
        !Designation = txtDesignation.Text
        If IsNumeric(dtcSpeciality.BoundText) Then !SpecialityID = dtcSpeciality.BoundText
        If IsNumeric(dtcAuthority.BoundText) Then !AuthorityID = dtcAuthority.BoundText
        !PrivateAddress = txtPrivateAddress.Text
        !PrivatePhone = txtPrivateTel.Text
        !PrivateFax = txtPrivateFax.Text
        !PrivateEmail = txtPrivateEmail.Text
        !MobilePhone = txtPrivateMobile.Text
        !OfficialAddress = txtOfficialAddress.Text
        !OfficialPhone = txtOfficialTel.Text
        !OfficialFax = txtOfficialFax.Text
        !OfficialEmail = txtOfficialEMail.Text
        !Website = txtOfficialWebsite.Text
        !Comments = txtComments.Text
        !Photo = txtPhoto.Text
        !Signature = txtSignature.Text
        !UserName = EncreptedWord(txtUserName.Text)
        !Password = EncreptedWord(txtPassword.Text)
        If chkUser.Value = 1 Then
            !User = True
        Else
            !User = False
        End If
        .Update
    End With
End Sub

Private Sub ChangeDetails()
    With rsTemStaff
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblstaff where staffID = " & dtcStaff.BoundText
        .Open temSQL, cnnStores, adOpenDynamic, adLockOptimistic
        If IsNumeric(dtcTitle.BoundText) Then !TitleID = dtcTitle.BoundText
        If IsNumeric(dtcSex.BoundText) Then !SexID = dtcSex.BoundText
        !Name = txtName.Text
        !listedName = txtListedName.Text
        !Qualifications = txtQualifications.Text
        !Registation = txtRegistation.Text
        !Designation = txtDesignation.Text
        If IsNumeric(dtcSpeciality.BoundText) Then !SpecialityID = dtcSpeciality.BoundText
        If IsNumeric(dtcAuthority.BoundText) Then !AuthorityID = dtcAuthority.BoundText
        !PrivateAddress = txtPrivateAddress.Text
        !PrivatePhone = txtPrivateTel.Text
        !PrivateFax = txtPrivateFax.Text
        !PrivateEmail = txtPrivateEmail.Text
        !MobilePhone = txtPrivateMobile.Text
        !OfficialAddress = txtOfficialAddress.Text
        !OfficialPhone = txtOfficialTel.Text
        !OfficialFax = txtOfficialFax.Text
        !OfficialEmail = txtOfficialEMail.Text
        !Website = txtOfficialWebsite.Text
        !Comments = txtComments.Text
        !Photo = txtPhoto.Text
        !Signature = txtSignature.Text
        !UserName = EncreptedWord(txtUserName.Text)
        !Password = EncreptedWord(txtPassword.Text)
        If chkUser.Value = 1 Then
            !User = True
        Else
            !User = False
        End If
        .Update
    End With
End Sub

Private Sub LoadPhoto(StaffID As Long)

End Sub

Private Sub SavePhoto(StaffID As Long)

End Sub


Private Sub bttnAdd_Click()
    TemUserName = Empty
    Call ClearValues
    Call AfterAdd
End Sub

Private Sub bttnCancel_Click()
    Call ClearValues
    Call BeforeAddEdit
End Sub

Private Sub bttnChange_Click()
    Call ChangeDetails
    Call ClearValues
    Call FillLists
    Call BeforeAddEdit
End Sub

Private Sub bttnEdit_Click()
    Dim tr As Integer
    If Not IsNumeric(dtcStaff.BoundText) Then
        tr = MsgBox("You have not selected a staff member to edit", vbCritical, "Staff?")
        dtcStaff.SetFocus
        Exit Sub
    End If
    TemUserName = txtUserName.Text
    Call AfterEdit
End Sub

Private Sub bttnPhotoDelete_Click()
    imgPhoto.Picture = LoadPicture()
    txtPhoto.Text = Empty
End Sub

Private Sub bttnPhotoLoad_Click()
    Dim tr As Integer
    imgPhoto.Stretch = True
    CommonDialog1.Filter = "BMP|*.BMP|JPG|*.JPG;JPE;JPEG|GIF|*.GIF|All Images|*.BMP;*.JPG;*.JPE;*.JPGE;*.GIF|All Files|*.*"
    CommonDialog1.ShowOpen
    On Error GoTo PhotoError:
    imgPhoto.Picture = LoadPicture(CommonDialog1.FileName)
    txtPhoto.Text = CommonDialog1.FileName
    Exit Sub
PhotoError:
    If Err.Number = 481 Then
        tr = MsgBox("The Photo you choose is not suitable, try using a medium size BMP, JPG or GIF file", vbOKOnly, "Photo Error")
    ElseIf Err.Number = 53 Then
        tr = MsgBox("No photo exist to selected, try to select again correctly.", vbOKOnly, "Photo Error")
    Else
        tr = MsgBox("An unknown error has occured, try again," & Chr(13) & Err.Description, vbOKOnly, "Photo Error")
    End If
End Sub

Private Sub bttnSave_Click()
    If CanSave = False Then Exit Sub
    Call SaveDetails
    Call ClearValues
    Call BeforeAddEdit
    Call FillLists
End Sub

Private Sub bttnSigDelete_Click()
    imgSignature.Picture = LoadPicture()
    txtSignature.Text = Empty
End Sub

Private Sub bttnSigLoad_Click()
    Dim tr As Integer
    imgSignature.Stretch = True
    CommonDialog1.Filter = "BMP|*.BMP|JPG|*.JPG;JPE;JPEG|GIF|*.GIF|All Images|*.BMP;*.JPG;*.JPE;*.JPGE;*.GIF|All Files|*.*"
    CommonDialog1.ShowOpen
    On Error GoTo PhotoError:
    imgSignature.Picture = LoadPicture(CommonDialog1.FileName)
    txtSignature.Text = CommonDialog1.FileName
    Exit Sub
PhotoError:
    If Err.Number = 481 Then
        tr = MsgBox("The Photo you choose is not suitable, try using a medium size BMP, JPG or GIF file", vbOKOnly, "Photo Error")
    ElseIf Err.Number = 53 Then
        tr = MsgBox("No photo exist to selected, try to select again correctly.", vbOKOnly, "Photo Error")
    Else
        tr = MsgBox("An unknown error has occured, try again," & Chr(13) & Err.Description, vbOKOnly, "Photo Error")
    End If
End Sub

Private Sub ButtonEx1_Click()
    Unload Me
End Sub

Private Sub dtcStaff_Click(Area As Integer)
    If IsNumeric(dtcStaff.BoundText) Then LocateDetails
End Sub

Private Sub Form_Load()
    
    Call FillLists
    Call BeforeAddEdit
    If UserAuthority = 6 Then
        bttnAdd.Enabled = False
        bttnEdit.Enabled = False
    End If
    SSTabMain.Tab = 0
End Sub

Private Sub FillLists()
    With rsStaff
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblstaff order by name"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcStaff
        Set .RowSource = rsStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With rsTitle
        If .State = 1 Then .Close
        temSQL = "SELECT * from tbltitle order by title"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcTitle
        Set dtcTitle.RowSource = rsTitle
        .ListField = "Title"
        .BoundColumn = "TitleID"
    End With
    With rsSex
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblsex"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcSex
        Set .RowSource = rsSex
        .ListField = "Sex"
        .BoundColumn = "SexID"
    End With
    With rsSpeciality
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblSpeciality"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcSpeciality
        Set .RowSource = rsSpeciality
        .ListField = "Speciality"
        .BoundColumn = "SpecialityID"
    End With
    With rsAuthority
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblAuthority order by Authority"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcAuthority
        Set .RowSource = rsAuthority
        .ListField = "Authority"
        .BoundColumn = "AuthorityID"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsAuthority.State = 1 Then rsAuthority.Close
    If rsSex.State = 1 Then rsSex.Close
    If rsSpeciality.State = 1 Then rsSpeciality.Close
    If rsStaff.State = 1 Then rsStaff.Close
    If rsTitle.State = 1 Then rsTitle.Close
    If rsTemStaff.State = 1 Then rsTemStaff.Close
    Set rsAuthority = Nothing
    Set rsSex = Nothing
    Set rsStaff = Nothing
    Set rsTitle = Nothing
    Set rsSpeciality = Nothing
    Set rsTemStaff = Nothing
End Sub
