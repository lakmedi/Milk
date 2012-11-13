VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmStaffCommentDisplay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff Comments"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
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
   ScaleHeight     =   5295
   ScaleWidth      =   9120
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   7680
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&lose"
      ForeColor       =   16711680
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
      Height          =   2775
      Left            =   2520
      TabIndex        =   9
      Top             =   1680
      Width           =   6375
   End
   Begin VB.TextBox txtByStaff 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1200
      Width           =   6375
   End
   Begin VB.TextBox txtTime 
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtDate 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label5 
      Caption         =   "Comment"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Comment By"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Time"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Staff Member"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmStaffCommentDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsSTaffComment As New ADODB.Recordset
    
Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If StaffCommentIDTx = 0 Then Unload Me: Exit Sub
    With rsSTaffComment
        If .State = 1 Then .Close
        temSql = "SELECT tblStaffComment.Comment, tblStaffComment.Date, tblStaffComment.Time, tblStaff.Name, tblByStaff.Name, tblStaffComment.StaffCommentID " & _
                    "FROM (tblStaffComment LEFT JOIN tblStaff ON tblStaffComment.StaffID = tblStaff.StaffID) LEFT JOIN tblStaff AS tblByStaff ON tblStaffComment.ByStaffID = tblByStaff.StaffID " & _
                    "WHERE (((tblStaffComment.StaffCommentID)=" & StaffCommentIDTx & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtComments.Text = ![Comment]
            txtName.Text = ![tblStaff.Name]
            txtDate.Text = Format(![Date], "dd MMMM yyyy")
            txtTime.Text = ![Time]
            txtByStaff.Text = ![tblByStaff.Name]
        End If
        .Close
    End With
End Sub
