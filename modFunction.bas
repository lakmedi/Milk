Attribute VB_Name = "modFunction"
Option Explicit
    Dim FSys As New Scripting.FileSystemObject
    Dim i As Integer
    
    
        Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Public Function RepeatString(InputString As String, RepeatNo As Integer) As String
    Dim r As Integer
    For r = 1 To RepeatNo
        RepeatString = RepeatString & InputString
    Next r
End Function


Public Sub SaveCommonSettings(MyForm As Form)
    Dim MyCtrl As Control
    Dim i As Integer
    For Each MyCtrl In MyForm.Controls
        If TypeOf MyCtrl Is MSFlexGrid Then
            For i = 0 To MyCtrl.Cols - 1
                SaveSetting App.EXEName, MyForm.Name & MyCtrl.Name, i, MyCtrl.ColWidth(i)
            Next
        ElseIf TypeOf MyCtrl Is ComboBox Then
            If InStr(MyCtrl.Tag, "SS") Then
                SaveSetting App.EXEName, MyForm.Name, MyCtrl.Name, MyCtrl.Text
            End If
        End If
    Next
    SaveSetting App.EXEName, MyForm.Name, "Top", MyForm.Top
    SaveSetting App.EXEName, MyForm.Name, "Left", MyForm.Left
    If MyForm.BorderStyle = 2 Then
        SaveSetting App.EXEName, MyForm.Name, "Width", MyForm.Width
        SaveSetting App.EXEName, MyForm.Name, "Height", MyForm.Height
    End If
    SaveSetting App.EXEName, MyForm.Name, "WindowState", MyForm.WindowState
    
End Sub

Public Sub GetCommonSettings(MyForm As Form)
    Dim MyCtrl As Control
    Dim i As Integer
    
    For Each MyCtrl In MyForm.Controls
        If TypeOf MyCtrl Is MSFlexGrid Then
            For i = 0 To MyCtrl.Cols - 1
                MyCtrl.ColWidth(i) = GetSetting(App.EXEName, MyForm.Name & MyCtrl.Name, i, MyCtrl.ColWidth(i))
                MyCtrl.AllowUserResizing = flexResizeColumns
            Next
        ElseIf TypeOf MyCtrl Is ComboBox Then
            On Error Resume Next
            If InStr(MyCtrl.Tag, "SS") Then
                MyCtrl.Text = GetSetting(App.EXEName, MyForm.Name, MyCtrl.Name, "")
            End If
            On Error GoTo 0
        End If
    Next
    
    If Val(GetSetting(App.EXEName, MyForm.Name, "Width", MyForm.Top)) < MyForm.Height * 0.75 Then MyForm.Top = GetSetting(App.EXEName, MyForm.Name, "Top", MyForm.Top)
    If Val(GetSetting(App.EXEName, MyForm.Name, "Width", MyForm.Left)) < MyForm.Width * 0.75 Then MyForm.Left = GetSetting(App.EXEName, MyForm.Name, "Left", MyForm.Left)
    If MyForm.BorderStyle = 2 Then
        If Val(GetSetting(App.EXEName, MyForm.Name, "Width", MyForm.Width)) > 0 Then MyForm.Width = GetSetting(App.EXEName, MyForm.Name, "Width", MyForm.Width)
        If Val(GetSetting(App.EXEName, MyForm.Name, "Width", MyForm.Height)) > 0 Then MyForm.Height = GetSetting(App.EXEName, MyForm.Name, "Height", MyForm.Height)
    End If
    On Error Resume Next
    MyForm.WindowState = GetSetting(App.EXEName, MyForm.Name, "WindowState", MyForm.WindowState)

End Sub





Public Sub GridToExcel(ExportGrid As MSFlexGrid, Optional Topic As String, Optional SubTopic As String, Optional SaveFile As String)
    
    
    On Error Resume Next
    
    If ExportGrid.Rows <= 1 Then
        MsgBox "Noting to Export"
        Exit Sub
    End If
    
    If SaveFile = "" Then
        SaveFile = App.Path & "\" & Topic & ".xls"
    End If
    
    Dim AppExcel As Excel.Application
    Dim myworkbook As Excel.Workbook
    Dim myWorkSheet1 As Excel.Worksheet
    Dim temRow As Integer
    Dim temCol As Integer
    
    Set AppExcel = CreateObject("Excel.Application")
    Set myworkbook = AppExcel.Workbooks.Add
    Set myWorkSheet1 = AppExcel.WorkSheets(1)
    
    myWorkSheet1.Cells(1, 1) = Topic
    myWorkSheet1.Cells(2, 1) = SubTopic
    
    For temRow = 0 To ExportGrid.Rows - 1
        For temCol = 0 To ExportGrid.Cols - 1
            myWorkSheet1.Cells(temRow + 4, temCol + 1) = ExportGrid.TextMatrix(temRow, temCol)
        Next
    Next temRow
    
    myWorkSheet1.Range("A1:" & GetColumnName(CDbl(temCol)) & temRow + 2).AutoFormat Format:=xlRangeAutoFormatClassic1
    
    myWorkSheet1.Range("A" & temRow + 3 & ":" & GetColumnName(CDbl(temCol)) & temRow + 3).AutoFormat Format:=xlRangeAutoFormat3DEffects1
    
    Topic = "Day End Summery " & Format(Date, "dd MMMM yyyy")
    myworkbook.SaveAs (SaveFile)
    myworkbook.Save
   ' myworkbook.Close
    
    ShellExecute 0&, "open", SaveFile, "", "", vbMaximizedFocus
End Sub

Public Function GetColumnName(ColumnNo As Long) As String
    Dim temnum As Integer
    Dim temnum1 As Integer
    
    If ColumnNo < 27 Then
        GetColumnName = Chr(ColumnNo + 64)
    Else
        temnum = ColumnNo \ 26
        temnum1 = ColumnNo Mod 26
        GetColumnName = Chr(temnum + 64) & Chr(temnum1 + 64)
    End If
End Function


