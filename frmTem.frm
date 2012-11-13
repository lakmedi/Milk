VERSION 5.00
Begin VB.Form frmTem 
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   2970
   ClientTop       =   3135
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   10650
   Begin VB.TextBox txtWord 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   3840
      Width           =   5175
   End
   Begin VB.TextBox txtNo 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "frmTem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTem As New ADODB.Recordset
    Dim temSQL As String
    Dim temCr As Double
    
    
Private Sub Command1_Click()
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblCollection"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            temCr = OwnCommisionRate(!SupplierID, !Liters)
            !CommisionRate = temCr
            !Commision = temCr * !Liters
            .Update
            .MoveNext
        Wend
    End With
End Sub

Private Sub Command2_Click()
    Dim NumToWord As New clsNumbers
    txtWord.Text = NumToWord.NumberToWord(Val(txtNo.Text))
    txtNo.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub txtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Command2_Click
    End If
End Sub
