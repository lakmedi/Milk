Attribute VB_Name = "modMilk3Functions"
Option Explicit
    Dim temSQL As String
    Public Type ItemDeduct
        ItemID As Long
        Qty As Double
        Value As Double
        Rate As Double
    End Type

Private Function ItemToDeduct(SupplierID As Long) As ItemDeduct()
    Dim temItemDeduct() As ItemDeduct
    Dim rsTem As New ADODB.Recordset
    Dim i As Integer
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT Sum(tblItemIssue.ToPay) AS SumOfToPay, Max(tblItemIssue.Rate) AS AvgOfRate, tblItemIssue.ItemID " & _
                    "From tblItemIssue " & _
                    "WHERE (((tblItemIssue.SupplierID)=" & SupplierID & ") AND ((tblItemIssue.Deleted) = 0)) " & _
                    "GROUP BY tblItemIssue.ItemID " & _
                    "Having Sum(tblItemIssue.ToPay) > 0 "
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        ReDim temItemDeduct(.RecordCount)
        While .EOF = False
            temItemDeduct(i).ItemID = !ItemID
            temItemDeduct(i).Qty = !SumOfToPay
            temItemDeduct(i).Rate = !AvgOfRate
            temItemDeduct(i).Value = !AvgOfRate * !SumOfToPay
            .MoveNext
            i = i + 1
        Wend
    End With
    ItemToDeduct = temItemDeduct
End Function

Public Sub AutoAddDeduction(SupplierID As Long, AvailableMoney As Double, DeductDate As Date, Volume As Double)
    Dim SupItemsToDeduct() As ItemDeduct
    Dim RemainingMoney As Double
    Dim temItemMoney As Double
    Dim MaxCanPay As Double
    
    RemainingMoney = AvailableMoney - DeductVolumeDeduction(SupplierID, DeductDate, Volume)
    
    Dim i As Integer
    SupItemsToDeduct = ItemToDeduct(SupplierID)
    
    For i = 0 To UBound(SupItemsToDeduct) - 1
        If RemainingMoney <= 0 Then
            Exit Sub
        ElseIf RemainingMoney > SupItemsToDeduct(i).Value Then
            AddDeduction SupplierID, SupItemsToDeduct(i).ItemID, SupItemsToDeduct(i).Qty, SupItemsToDeduct(i).Rate, Volume, DeductDate
            RemainingMoney = RemainingMoney - SupItemsToDeduct(i).Value
        Else
            MaxCanPay = RemainingMoney / SupItemsToDeduct(i).Rate
            AddDeduction SupplierID, SupItemsToDeduct(i).ItemID, MaxCanPay, SupItemsToDeduct(i).Rate, Volume, DeductDate
            RemainingMoney = 0
        End If
    Next i
End Sub

Private Function DeductVolumeDeduction(SupplierID As Long, DeductDate As Date, Volume As Double) As Double
    Dim rsDeduction As New ADODB.Recordset
    With rsDeduction
        temSQL = "Select * from tblFarmerDailyVolumeDeduction where FarmerDailyVolumeDeductionID =0"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !DeductedDate = DeductDate
        !SupplierID = SupplierID
        !DeductedValue = VolumeDeduction(!SupplierID, Volume)
        DeductVolumeDeduction = !DeductedValue
        !DeductedVolume = Volume
        !AddedDate = Date
        !AddedTime = Time
        !AddedUserID = UserID
        !FromDate = DeductDate
        !ToDate = DeductDate
        .Update
        .Close
    End With
    
End Function

Private Sub AddDeduction(SupplierID As Long, ItemID As Long, Qty As Double, Rate As Double, Volume As Double, DeductDate As Date)
    Dim temValue As Double
    Dim temRemainingQty As Double
    
    temValue = Rate * Qty
    
    Dim rsDeduction As New ADODB.Recordset
    With rsDeduction
        temSQL = "Select * from tblDeduction where DeductionID =0"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !DeductedDate = DeductDate
        !ItemID = ItemID
        !SupplierID = SupplierID
        !Quentity = Qty
        !Rate = Rate
        !Value = temValue
        !AddedDate = Date
        !AddedTime = Time
        !AddedUserID = UserID
        .Update
        .Close
    End With
    
    Dim rsItemIssue As New ADODB.Recordset
    Dim remainingQty As Double
    Dim SettlingQty As Double
    
    remainingQty = Qty
    
    With rsItemIssue
        If .State = 1 Then .Close
        temSQL = "SELECT tblItemIssue.Quentity, tblItemIssue.Paid, tblItemIssue.ToPay, tblItemIssue.Quentity, tblItemIssue.IssueDate " & _
                    "From tblItemIssue " & _
                    "WHERE (((tblItemIssue.Quentity)>[tblItemIssue].[Paid]) AND ((tblItemIssue.ItemID)=" & ItemID & ") AND ((tblItemIssue.Deleted)= 0 )  AND ((tblItemIssue.SupplierID)=" & SupplierID & ")) " & _
                    "ORDER BY tblItemIssue.IssueDate DESC"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            Do While .EOF = False
                If !ToPay >= remainingQty Then
                    !ToPay = !ToPay - remainingQty
                    !Paid = !Paid + remainingQty
                    .Update
                    remainingQty = 0
                    Exit Do
                Else
                    SettlingQty = !ToPay
                    !ToPay = 0
                    !Paid = !Paid + SettlingQty
                    .Update
                    remainingQty = remainingQty - SettlingQty
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    
    
    
    
    
    
End Sub


Public Function VolumeDeduction(SupplierID As Long, Volume As Double) As Double
    Dim rsVoulemeDeductionType As New ADODB.Recordset
    Dim temDed As Double
    With rsVoulemeDeductionType
        If .State = 1 Then .Close
            temSQL = "Select * from tblFarmerVolumeDeduction where Deleted = 0  AND FarmerID = " & SupplierID
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                While .EOF = False
                    temDed = temDed + (Volume * !DeductionPerLiter)
                    .MoveNext
                Wend
            End If
            .Close
    End With
    VolumeDeduction = temDed
End Function

Public Function AutoAddedDeductions(AddedDate As Date, CC As Long) As Boolean
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblAutoDeduction where CollectingCenterID = " & CC & " AND AddedDate = '" & Format(AddedDate, "dd MMMM yyyy") & "'"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            AutoAddedDeductions = True
        Else
            AutoAddedDeductions = False
        End If
        .Close
    End With
End Function

Public Function FindVolumePaymentDue(FarmerID() As Double, FromDate As Date, ToDate As Date) As Double
 Dim rsDeduction As New ADODB.Recordset
    Dim i As Integer
    For i = 0 To UBound(FarmerID)
        With rsDeduction
            temSQL = "Select sum(DeductedValue) as VolumeDeduction, SupplierID from tblFarmerDailyVolumeDeduction where SupplierID = " & FarmerID(i) & " AND DeductedDate between '" & Format(FromDate, "dd MMMM yyyy") & "' AND '" & Format(ToDate, "dd MMMM yyyy") & "' AND Deleted = 0  GROUP BY SupplierID"
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                If IsNull(!VolumeDeduction) = False Then
                    If (!VolumeDeduction) > 0 Then
                        FindVolumePaymentDue = FindVolumePaymentDue + !VolumeDeduction
                    Else
                        'MsgBox !SupplierID
                    End If
                Else
                    'MsgBox !SupplierID
                End If
            End If
            .Close
        End With
    Next
    
End Function

Private Function ComBankValue(Value As Double) As String
    ComBankValue = Format(Val(Format(Value, "0.00") * 100), "000000000000")
End Function


Public Sub AddVolumeDeductionToCombankExcel(DeductionValue As Double, i As RowDividerStyleConstants, myworksheet As Excel.Worksheet, temFromDate)
            
    Dim CombankID As String
    Dim CombankBranchCode As String
    Dim ComBankAccountName As String
    Dim ComBankAccountNo As String
     
     CombankID = "7056"
     CombankBranchCode = "104"
     ComBankAccountName = "Lucky Farmers Social Security Fund"
     ComBankAccountNo = "001104014040"
            
            
            
            myworksheet.Cells(i, 1) = "0000"
            myworksheet.Cells(i, 2) = CombankID
            myworksheet.Cells(i, 3) = CombankBranchCode
            myworksheet.Cells(i, 4) = ComBankAccountNo
            
            myworksheet.Cells(i, 5) = ComBankAccountName
            myworksheet.Cells(i, 6) = "23"
            myworksheet.Cells(i, 7) = "00"
            myworksheet.Cells(i, 8) = "0"
            
            myworksheet.Cells(i, 9) = "000000"
            myworksheet.Cells(i, 10) = ComBankValue(DeductionValue)
            myworksheet.Cells(i, 11) = "SLR"
            myworksheet.Cells(i, 12) = CombankID
            myworksheet.Cells(i, 13) = CombankBranchCode
            myworksheet.Cells(i, 14) = ComBankAccountNo
            myworksheet.Cells(i, 15) = ComBankAccountName
            myworksheet.Cells(i, 16) = "0"
            myworksheet.Cells(i, 17) = "MilkTotal" & Format(temFromDate, "yy MM dd")
            myworksheet.Cells(i, 18) = Format(Date, "yyMMdd")
            myworksheet.Cells(i, 19) = ""
            myworksheet.Cells(i, 20) = "@"
    
            myworksheet.Cells(6, 16) = ""
    

End Sub
