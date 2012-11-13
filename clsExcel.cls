VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim tr As Integer
    
    Dim temSQL As String
    
    Dim AppExcel As Excel.Application
    Dim myworkbook As Excel.Workbook
    Dim myworksheet() As Excel.Worksheet
    
    Dim MainWorksheet As Excel.Worksheet
    
    
    Dim WorkSheets As Integer
    
    Dim FSys As New Scripting.FileSystemObject
        
    Dim fullExcelFileNameWithPath As String

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

    
'local variable(s) to hold property value(s)
    Private mvarTopic As String 'local copy
    Private mvarSubtopic As String 'local copy
    Private mvarTopicCol As Integer 'local copy
    Private mvarTopicRow As Integer 'local copy
    Private mvarSubtopicRow As Integer 'local copy
    Private mvarSubtopicCol As Integer 'local copy
    Private mvarPath As String 'local copy
    Private mvarFileName As String 'local copy

Public Property Let FileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileName = 5
    mvarFileName = vData
End Property


Public Property Get FileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
End Property



Public Property Let Path(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Path = 5
    mvarPath = vData
End Property


Public Property Get Path() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Path
    Path = mvarPath
End Property



Public Property Let SubtopicCol(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubtopicCol = 5
    mvarSubtopicCol = vData
End Property


Public Property Get SubtopicCol() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubtopicCol
    SubtopicCol = mvarSubtopicCol
End Property



Public Property Let SubtopicRow(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubtopicRow = 5
    mvarSubtopicRow = vData
End Property


Public Property Get SubtopicRow() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubtopicRow
    SubtopicRow = mvarSubtopicRow
End Property



Public Property Let TopicRow(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TopicRow = 5
    mvarTopicRow = vData
End Property


Public Property Get TopicRow() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TopicRow
    TopicRow = mvarTopicRow
End Property



Public Property Let TopicCol(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TopicCol = 5
    mvarTopicCol = vData
End Property


Public Property Get TopicCol() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TopicCol
    TopicCol = mvarTopicCol
End Property



Public Property Let SubTopic(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Subtopic = 5
    mvarSubtopic = vData
End Property


Public Property Get SubTopic() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Subtopic
    SubTopic = mvarSubtopic
End Property



Public Property Let Topic(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Topic = 5
    mvarTopic = vData
End Property


Public Property Get Topic() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Topic
    Topic = mvarTopic
End Property



    
Private Sub Class_Initialize()
    Set AppExcel = CreateObject("Excel.Application")
    Set myworkbook = AppExcel.Workbooks.Add
    Set MainWorksheet = myworkbook.WorkSheets(1)
    mvarTopic = ""
    mvarTopicCol = 1
    mvarTopicRow = 1
    mvarSubtopic = ""
    mvarSubtopicCol = 1
    mvarSubtopicRow = 2
    mvarFileName = "Excel " & Format(Date, "dd MMMM yyyy") & " " & Format(Time, "hh mm ss AMPM")
    mvarPath = App.Path
End Sub

Public Sub addWorkSheets(Count As Integer)
    If Count > 10 Then Exit Sub
    If Count <= WorkSheets Then Exit Sub
    ReDim Preserve myworksheet(Count)
    WorkSheets = WorkSheets + Count
    Dim i As Integer
    For i = WorkSheets - 1 To Count - 1
        Set myworksheet(i) = AppExcel.WorkSheets(i + 1)
    Next i
    
End Sub

Private Sub Class_Terminate()
    Dim i As Integer
    For i = 1 To WorkSheets
        Set Workbooks(i) = Nothing
    Next i

    Set myworkbook = Nothing
    Set AppExcel = Nothing

End Sub
    
    
Public Sub setWorkSheet(i As Integer)
    If i > WorkSheets Then Exit Sub
    myworksheet(i - 1).Activate
    Set MainWorksheet = myworksheet(i - 1)
End Sub

Public Sub setValue(row As Integer, col As Integer, CellValue As String)
    MainWorksheet.Cells(row, col) = CellValue
End Sub
    
Public Sub setTopic()
    setValue mvarTopicRow, mvarTopicCol, mvarTopic
End Sub

Public Sub setSubtopic()
    setValue mvarSubtopicRow, mvarSubtopicCol, mvarSubtopic
End Sub

Public Static Function intToCol(ByVal colNum As Long) As String
    Dim i As Long
    Dim x As Long
    For i = 6 To 0 Step -1
        x = (1 - 26 ^ (i + 1)) / (-25) - 1
        If colNum > x Then
            intToCol = intToCol & Chr(((colNum - x - 1) \ 26 ^ i) Mod 26 + 65)
        End If
    Next i
End Function
    
Public Sub setVerticalHeadings(Row1 As Integer, Col1 As Integer, Row2 As Integer, Col2 As Integer)
    Dim TemRangeAddress As String
    On Error Resume Next
    TemRangeAddress = intToCol(Col1) & Row1 & ":" & intToCol(Col2) & Row2
    MainWorksheet.Range(TemRangeAddress).Font.Bold = True
    MainWorksheet.Range(TemRangeAddress).Font.Size = 13
    MainWorksheet.Range(TemRangeAddress).BorderAround 13, xlMedium, xlColorIndexAutomatic
    MainWorksheet.Range(TemRangeAddress).Orientation = xlTickLabelOrientationUpward
    MainWorksheet.Range(TemRangeAddress).HorizontalAlignment = xlHAlignCenter
    MainWorksheet.Range(TemRangeAddress).autoFit
End Sub
    
Public Sub setHorizontalHeadings(Row1 As Integer, Col1 As Integer, Row2 As Integer, Col2 As Integer)
    On Error Resume Next
    Dim TemRangeAddress As String
    TemRangeAddress = intToCol(Col1) & Row1 & ":" & intToCol(Col2) & Row2
    MainWorksheet.Range(TemRangeAddress).Font.Bold = True
    MainWorksheet.Range(TemRangeAddress).Font.Size = 13
    MainWorksheet.Range(TemRangeAddress).BorderAround 13, xlMedium, xlColorIndexAutomatic
    MainWorksheet.Range(TemRangeAddress).Orientation = xlTickLabelOrientationHorizontal
    MainWorksheet.Range(TemRangeAddress).HorizontalAlignment = xlHAlignCenter
    MainWorksheet.Range(TemRangeAddress).autoFit
End Sub
    
Public Sub setHeadings(Row1 As Integer, Col1 As Integer, Row2 As Integer, Col2 As Integer)
    On Error Resume Next
    Dim TemRangeAddress As String
    TemRangeAddress = intToCol(Col1) & Row1 & ":" & intToCol(Col2) & Row2
    MainWorksheet.Range(TemRangeAddress).Font.Bold = True
    MainWorksheet.Range(TemRangeAddress).Font.Size = 13
    MainWorksheet.Range(TemRangeAddress).BorderAround 13, xlMedium, xlColorIndexAutomatic
    MainWorksheet.Range(TemRangeAddress).Orientation = xlTickLabelOrientationAutomatic
    MainWorksheet.Range(TemRangeAddress).HorizontalAlignment = xlHAlignCenter
    MainWorksheet.Range(TemRangeAddress).autoFit
End Sub

Public Sub finalizeExcel()
    On Error Resume Next
    Cells.Worksheet.Columns.autoFit
    MainWorksheet.Cells(mvarTopicRow, mvarTopicCol) = Topic
    MainWorksheet.Cells(mvarSubtopicRow, mvarSubtopicCol) = SubTopic
End Sub

Public Sub autofitAll()
    On Error Resume Next
    Cells.Worksheet.Columns.autoFit
End Sub


Public Sub setCurrency(Row1 As Integer, Col1 As Integer, Row2 As Integer, Col2 As Integer)
    On Error Resume Next
    Dim TemRangeAddress As String
    TemRangeAddress = intToCol(Col1) & Row1 & ":" & intToCol(Col2) & Row2
    MainWorksheet.Range(TemRangeAddress).NumberFormat = "#,##0.00"
    MainWorksheet.Range(TemRangeAddress).HorizontalAlignment = xlHAlignRight
    MainWorksheet.Range(TemRangeAddress).autoFit
End Sub

Public Sub mergeCells(Row1 As Integer, Col1 As Integer, Row2 As Integer, Col2 As Integer)
    On Error Resume Next
    Dim TemRangeAddress As String
    TemRangeAddress = intToCol(Col1) & Row1 & ":" & intToCol(Col2) & Row2
    MainWorksheet.Range(TemRangeAddress).Merge
    MainWorksheet.Range(TemRangeAddress).HorizontalAlignment = xlHAlignCenter
    MainWorksheet.Range(TemRangeAddress).autoFit
End Sub

Public Sub autoFit(Row1 As Integer, Col1 As Integer, Row2 As Integer, Col2 As Integer)
    On Error Resume Next
    Dim TemRangeAddress As String
    TemRangeAddress = intToCol(Col1) & Row1 & ":" & intToCol(Col2) & Row2
End Sub

    
Public Sub saveExcel()
On Error Resume Next
    fullExcelFileNameWithPath = mvarPath & "\" & mvarFileName
    While FSys.FileExists(fullExcelFileNameWithPath & ".xls") = True
        fullExcelFileNameWithPath = fullExcelFileNameWithPath & "1"
    Wend
    myworkbook.SaveAs (fullExcelFileNameWithPath & ".xls")

End Sub


Public Sub showExcel()
    ExcelFilePath = fullExcelFileNameWithPath & ".xls"
    frmChart.Show
End Sub
    
Public Sub openExcel()
    ShellExecute 0&, "open", fullExcelFileNameWithPath & ".xls", "", "", vbMaximizedFocus
End Sub
