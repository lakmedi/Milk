Attribute VB_Name = "ModuleFillAnyCombo"
Option Explicit
    Dim rsFill As New ADODB.Recordset
    Dim temSql As String
    
Public Sub FillAnyCombo(ComboToFill As DataCombo, table As String, Optional DoNotIncludeDeleted As Boolean)
    temSql = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then temSql = temSql & " Where Deleted = False "
    temSql = temSql & " Order by " & table
    With rsFill
        If .State = 1 Then .Close
        .Open temSql, cnnsStore, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = table
        .BoundColumn = table & "ID"
    End With
End Sub

