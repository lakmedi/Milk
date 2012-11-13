VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGRNComparisonReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GRN Comparison Report"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8340
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
   ScaleHeight     =   6525
   ScaleWidth      =   8340
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   56688643
      CurrentDate     =   39785
   End
   Begin MSDataListLib.DataCombo cmbCC 
      Height          =   360
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   56688643
      CurrentDate     =   39785
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "&Collecting Center"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "frmGRNComparisonReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
