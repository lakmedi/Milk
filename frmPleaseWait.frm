VERSION 5.00
Begin VB.Form frmPleaseWait 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please Wait"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   ScaleHeight     =   1815
   ScaleWidth      =   5385
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Please Wait ..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmPleaseWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim ProgressValue As Integer

Private Sub Form_Load()
'    ProgressBar1.Value = 0
'    ProgressValue = 0
End Sub

Private Sub Timer1_Timer()
'    ProgressValue = ProgressValue + 1
'    If ProgressValue > 100 Then ProgressValue = 0
 '   ProgressBar1.Value = ProgressValue
End Sub
