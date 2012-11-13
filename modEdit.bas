Attribute VB_Name = "modEdit"
Option Explicit
    Public Type SaveCheck
        CanSave As Boolean
        Message As String
        Control As Control
    End Type

Public Sub EditMode(EditForm As Form)
    Dim MyCtrl As Control
    For Each MyCtrl In EditForm.Controls
        If InStr(MyCtrl.Tag, "E") > 0 Then
            MyCtrl.Enabled = True
        ElseIf InStr(MyCtrl.Tag, "S") > 0 Then
            MyCtrl.Enabled = False
        End If
    Next
End Sub

Public Sub SelectMode(EditForm As Form)
    Dim MyCtrl As Control
    For Each MyCtrl In EditForm.Controls
        If InStr(MyCtrl.Tag, "E") > 0 Then
            MyCtrl.Enabled = False
        ElseIf InStr(MyCtrl.Tag, "S") > 0 Then
            MyCtrl.Enabled = True
        End If
    Next
End Sub

Public Sub ClearEditDetails(EditForm As Form)
    Dim MyCtrl As Control
    For Each MyCtrl In EditForm.Controls
        If InStr(MyCtrl.Tag, "E") > 0 Then
            If TypeOf MyCtrl Is TextBox Then
                MyCtrl.Text = Empty
            ElseIf TypeOf MyCtrl Is DataCombo Then
                MyCtrl.Text = Empty
            ElseIf TypeOf MyCtrl Is CheckBox Then
                MyCtrl.Checked = 0
            ElseIf TypeOf MyCtrl Is OptionButton Then
                MyCtrl.Value = False
            ElseIf TypeOf MyCtrl Is ComboBox Then
                MyCtrl.Text = Empty
            End If
        End If
    Next
End Sub
