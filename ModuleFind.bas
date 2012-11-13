Attribute VB_Name = "ModuleFind"























Option Explicit
    Dim rsPrice As New ADODB.Recordset
    Dim temSQL As String
    Public Type MilkCollection
        Supplied As Boolean
        LMR As Double
        FAT As Double
        Liters As Double
        SNF As Double
        Price As Double
        Value As Double
        OwnCommisionRate As Double
        OwnCommision As Double
    End Type
    
        
Public Function LastPaymentGeneratedDate(CollectingCenterID As Long) As Date
    Dim rsCCSummery As New ADODB.Recordset
    With rsCCSummery
        If .State = 1 Then .Close
        temSQL = "SELECT Max(tblCollectingCenterPaymentSummery.ToDate) AS MaxOfToDate " & _
                    "From tblCollectingCenterPaymentSummery " & _
                    "WHERE tblCollectingCenterPaymentSummery.CollectingCenterID =" & CollectingCenterID
               .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If IsNull(!MaxOfToDate) = False Then
            LastPaymentGeneratedDate = !MaxOfToDate
        Else
            LastPaymentGeneratedDate = #1/1/2008#
        End If
        .Close
    End With
    Set rsCCSummery = Nothing
End Function

        
Public Function LMRXLiters(LMR As Double, Liters As Double) As Double
    LMRXLiters = LMR * Liters
End Function

Public Function FATXLiters(FAT As Double, Liters As Double) As Double
    FATXLiters = FAT * Liters
End Function

Public Function SNF(LMR As Double, FAT As Double) As Double
    SNF = roundLow((LMR * 0.25) + (FAT * 0.22) + 0.72, 1)
End Function


Public Function roundLow(ToRound As Double, DigitsAfterDecimal As Integer) As Double
    Dim beforeDecimal As String
    Dim afterDecimal As String
    Dim decimalPlace As Integer
    
    decimalPlace = InStr(CStr(ToRound), ".")
    
    If decimalPlace > 0 Then
        beforeDecimal = Left(CStr(ToRound), decimalPlace - 1)
        afterDecimal = Mid(CStr(ToRound), decimalPlace, DigitsAfterDecimal + 1)
        roundLow = Val(beforeDecimal & afterDecimal)
    Else
        roundLow = ToRound
    End If
End Function


Public Function Price(FAT As Double, SNF As Double, CollectingCenterID As Long, SupplierID As Long, PriceDate As Date)
    Dim rsCC As New ADODB.Recordset
    Dim PaymentSchemeID As Long
    
    With rsCC
        If .State = 1 Then .Close
        temSQL = "Select * from tblSupplier where SupplierID = " & SupplierID
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!PaymentSchemeID) = False Then
                PaymentSchemeID = !PaymentSchemeID
            Else
                PaymentSchemeID = 0
            End If
        Else
            PaymentSchemeID = 0
        End If
        .Close
    End With
    
    If PaymentSchemeID = 0 Then
        With rsCC
            If .State = 1 Then .Close
            temSQL = "Select * from tblCollectingCenter where CollectingCenterID = " & CollectingCenterID
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                PaymentSchemeID = !PaymentSchemeID
            End If
            .Close
        End With
    End If
    
'    With rsCC
'        If .State = 1 Then .Close
'        temSql = "Select * from tblSupplier where SupplierID = " & SupplierID
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            PaymentSchemeID = !PaymentSchemeID
'        End If
'        .Close
'    End With
    
    FAT = Round(FAT, 1)
    SNF = Round(SNF, 1)
    With rsPrice
        If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblPrice WHERE FAT = " & FAT & " AND SNF = " & SNF & "AND PaymentSchemeID = " & PaymentSchemeID & " AND ((FromDate = '" & Format(PriceDate, "DD MMMM yyyy") & "') or (ToDate = '" & Format(PriceDate, "DD MMMM yyyy") & "') or (FromDate<'" & Format(PriceDate, "DD MMMM yyyy") & "' AND ToDate>'" & Format(PriceDate, "DD MMMM yyyy") & "')) ORDER By FromDate DESC"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!Price) = False Then
                Price = !Price
            Else
                Price = 0
            End If
        Else
            Price = 0
        End If
        .Close
    End With
End Function

Public Function DailyMilkSupply(SuppliedDate As Date, SupplierID As Long, SecessionID As Long, Optional TrueDate As Boolean) As MilkCollection
    Dim rsMilk As New ADODB.Recordset
    Dim MyMilkCollection As MilkCollection
    With rsMilk
        If .State = 1 Then .Close
        If TrueDate = True Then
            If SecessionID <> 0 Then
                temSQL = "SELECT  tblCollection.CollectionID,  tblCollection.CommisionRate, tblCollection.Date,  tblCollection.ProgramDate ,tblCollection.LMR, tblCollection.FAT, tblCollection.Liters, tblCollection.SNF, tblCollection.Price, tblCollection.Value, tblCollection.SecessionID, tblCollection.Date, tblCollection.SupplierID, tblCollection.Commision " & _
                            "From tblCollection " & _
                            "WHERE (((tblCollection.SecessionID)=" & SecessionID & ") AND ((tblCollection.Date) = '" & Format(SuppliedDate, "dd MMMM yyyy") & "') AND ((tblCollection.SupplierID)=" & SupplierID & ")  And ((tblCollection.Deleted) = 0 ) )"
            Else
                temSQL = "SELECT  tblCollection.CollectionID,  tblCollection.CommisionRate , tblCollection.Date,  tblCollection.ProgramDate, tblCollection.LMR, tblCollection.FAT, tblCollection.Liters, tblCollection.SNF, tblCollection.Price, tblCollection.Value, tblCollection.SecessionID, tblCollection.Date, tblCollection.SupplierID, tblCollection.Commision  " & _
                            "From tblCollection " & _
                            "WHERE (((tblCollection.Date) = '" & Format(SuppliedDate, "dd MMMM yyyy") & "') AND ((tblCollection.SupplierID)=" & SupplierID & ")  And ((tblCollection.Deleted) = 0 ) )"
            End If
        Else
            If SecessionID <> 0 Then
                temSQL = "SELECT  tblCollection.CollectionID,  tblCollection.CommisionRate , tblCollection.Date,  tblCollection.ProgramDate ,tblCollection.LMR, tblCollection.FAT, tblCollection.Liters, tblCollection.SNF, tblCollection.Price, tblCollection.Value, tblCollection.SecessionID, tblCollection.Date, tblCollection.SupplierID, tblCollection.Commision  " & _
                            "From tblCollection " & _
                            "WHERE (((tblCollection.SecessionID)=" & SecessionID & ") AND ((tblCollection.ProgramDate) = '" & Format(SuppliedDate, "dd MMMM yyyy") & "') AND ((tblCollection.SupplierID)=" & SupplierID & ")  And ((tblCollection.Deleted) = 0 ) )"
            Else
                temSQL = "SELECT  tblCollection.CollectionID, tblCollection.CommisionRate ,tblCollection.Date,  tblCollection.ProgramDate, tblCollection.LMR, tblCollection.FAT, tblCollection.Liters, tblCollection.SNF, tblCollection.Price, tblCollection.Value, tblCollection.SecessionID, tblCollection.Date, tblCollection.SupplierID, tblCollection.Commision  " & _
                            "From tblCollection " & _
                            "WHERE (((tblCollection.ProgramDate) = '" & Format(SuppliedDate, "dd MMMM yyyy") & "') AND ((tblCollection.SupplierID)=" & SupplierID & ")  And ((tblCollection.Deleted) = 0 ) )"
            End If
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            MyMilkCollection.Supplied = True
            MyMilkCollection.LMR = !LMR
            MyMilkCollection.FAT = !FAT
            MyMilkCollection.Liters = !Liters
            MyMilkCollection.SNF = !SNF
            MyMilkCollection.Price = !Price
            MyMilkCollection.Value = !Value
            MyMilkCollection.OwnCommision = !Commision
            MyMilkCollection.OwnCommisionRate = !CommisionRate
        Else
            MyMilkCollection.Supplied = False
            MyMilkCollection.LMR = 0
            MyMilkCollection.FAT = 0
            MyMilkCollection.Liters = 0
            MyMilkCollection.SNF = 0
            MyMilkCollection.Price = 0
            MyMilkCollection.Value = 0
            MyMilkCollection.OwnCommision = 0
            MyMilkCollection.OwnCommisionRate = 0
        End If
        .Close
    End With
    DailyMilkSupply = MyMilkCollection
    Set rsMilk = Nothing
End Function

Public Function PeriodMilkSupply(FromDate As Date, ToDate As Date, SupplierID As Long, SecessionID As Long) As MilkCollection
    Dim rsMilk As New ADODB.Recordset
    
    Dim MyMilkCollection As MilkCollection
    
    Dim LMRXLiters As Double
    Dim FATXLiters As Double
    Dim SuppliedCount As Long
    
    Dim totalLiters As Double
    Dim totalValue As Double
    Dim TotalOwnCommision As Double
    
    Dim rsCC As New ADODB.Recordset
    
    Dim CollectingCenterID As Long
    
    
    With rsCC
        temSQL = "Select * from tblSUpplier where SupplierID = " & SupplierID
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            CollectingCenterID = !CollectingCenterID
        Else
            CollectingCenterID = 0
        End If
        .Close
    End With
    
    With rsMilk
        If .State = 1 Then .Close
        If SecessionID <> 0 Then
            temSQL = "SELECT  tblCollection.CollectionID, tblCollection.Date, tblCollection.ProgramDate, tblCollection.LMR, tblCollection.FAT, tblCollection.Liters, tblCollection.SNF, tblCollection.Price, tblCollection.Value, tblCollection.SecessionID, tblCollection.Date, tblCollection.SupplierID, tblCollection.Commision  " & _
                        "From tblCollection " & _
                        "WHERE (((tblCollection.SecessionID)=" & SecessionID & ") AND ((tblCollection.ProgramDate) BETWEEN '" & Format(FromDate, "dd MMMM yyyy") & "' And '" & Format(ToDate, "dd MMMM yyyy") & "' ) AND ((tblCollection.SupplierID)=" & SupplierID & ")  And ((tblCollection.Deleted) = 0 ))"
        Else
            temSQL = "SELECT  tblCollection.CollectionID, tblCollection.Date, tblCollection.ProgramDate, tblCollection.LMR, tblCollection.FAT, tblCollection.Liters, tblCollection.SNF, tblCollection.Price, tblCollection.Value, tblCollection.SecessionID, tblCollection.Date, tblCollection.SupplierID, tblCollection.Commision  " & _
                        "From tblCollection " & _
                        "WHERE (((tblCollection.ProgramDate) Between '" & Format(FromDate, "dd MMMM yyyy") & "' And '" & Format(ToDate, "dd MMMM yyyy") & "') AND ((tblCollection.SupplierID)=" & SupplierID & ")  And ((tblCollection.Deleted) = 0 ))"
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            MyMilkCollection.Supplied = True
            While .EOF = False
                totalLiters = totalLiters + !Liters
                totalValue = totalValue + !Value
                TotalOwnCommision = TotalOwnCommision + !Commision
                LMRXLiters = LMRXLiters + (!LMR * !Liters)
                FATXLiters = FATXLiters + (!FAT * !Liters)
                SuppliedCount = SuppliedCount + 1
                .MoveNext
            Wend
            If totalLiters > 0 Then
                MyMilkCollection.FAT = FATXLiters / totalLiters
                MyMilkCollection.Liters = totalLiters
                MyMilkCollection.LMR = LMRXLiters / totalLiters
                MyMilkCollection.SNF = SNF(LMRXLiters / totalLiters, FATXLiters / totalLiters)
                
                'Check
                
                'MyMilkCollection.Price = Price(FATXLiters / TotalLiters, MyMilkCollection.SNF, CollectingCenterID, FromDate)
                MyMilkCollection.Price = totalValue / totalLiters
                
                
                'Check
                
                MyMilkCollection.Value = totalValue
                MyMilkCollection.OwnCommision = TotalOwnCommision
                MyMilkCollection.OwnCommisionRate = TotalOwnCommision / totalLiters
            Else
                MyMilkCollection.Supplied = False
                MyMilkCollection.LMR = 0
                MyMilkCollection.FAT = 0
                MyMilkCollection.Liters = 0
                MyMilkCollection.SNF = 0
                MyMilkCollection.Price = 0
                MyMilkCollection.OwnCommision = 0
                MyMilkCollection.OwnCommisionRate = 0
            End If
        Else
            MyMilkCollection.Supplied = False
            MyMilkCollection.LMR = 0
            MyMilkCollection.FAT = 0
            MyMilkCollection.Liters = 0
            MyMilkCollection.SNF = 0
            MyMilkCollection.Price = 0
            MyMilkCollection.Value = 0
            MyMilkCollection.OwnCommision = 0
            MyMilkCollection.OwnCommisionRate = 0
        End If
        .Close
    End With
    PeriodMilkSupply = MyMilkCollection
    Set rsMilk = Nothing
End Function

Public Function OwnCommision(SupplierID As Long, FromDate As Date, ToDate As Date, SecessionID As Long) As Double
    OwnCommision = 0
    Dim rsVolume As New ADODB.Recordset
    Dim i As Integer
    With rsVolume
         If .State = 1 Then .Close
         If SecessionID = 0 Then
             temSQL = "SELECT Sum(tblCollection.Commision) AS SumOfCommision " & _
                         "FROM tblCollection " & _
                         "WHERE (((tblCollection.ProgramDate) Between '" & Format(FromDate, "dd MMMM yyyy") & "' And '" & Format(ToDate, "dd MMMM yyyy") & "') AND ((tblCollection.SupplierID)=" & SupplierID & ")  And ((tblCollection.Deleted) = 0 ))"
         Else
             temSQL = "SELECT Sum(tblCollection.Commision) AS SumOfCommision " & _
                         "FROM tblCollection " & _
                         "WHERE (((tblCollection.ProgramDate) Between '" & Format(FromDate, "dd MMMM yyyy") & "' And '" & Format(ToDate, "dd MMMM yyyy") & "') AND ((tblCollection.SupplierID)=" & SupplierID & ")  And ((tblCollection.Deleted) = 0 ) AND ((SecessionID) = " & SecessionID & "))"
         End If
         .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
         If IsNull(!SumOfCommision) = False Then
             OwnCommision = !SumOfCommision
             .Close
         Else
             OwnCommision = 0
         End If
     End With
End Function


Public Function OthersCommision(SupplierID As Long, FromDate As Date, ToDate As Date, SecessionID As Long) As Double
    Dim i As Integer
    Dim rsS As New ADODB.Recordset
    Dim MyMilkCollection As MilkCollection
    Dim TotalMilk As Double
    Dim temDays As Integer
    Dim AvgMilk As Double
    Dim temCommsisionRate As Double
    Dim temCommsision As Double
    With rsS
        If .State = 1 Then .Close
        temSQL = "SELECT tblSupplier.SupplierID, tblSupplier.Supplier, tblSupplier.SupplierCode " & _
                    "From tblSupplier " & _
                    "Where (((tblSupplier.CommisionCollectorID) = " & SupplierID & ") AND (Deleted  = 0)) " & _
                    "ORDER BY tblSupplier.Supplier"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            TotalMilk = 0
            For i = 1 To .RecordCount
                MyMilkCollection = PeriodMilkSupply(FromDate, ToDate, !SupplierID, 0)
                TotalMilk = TotalMilk + MyMilkCollection.Liters
                .MoveNext
            Next
        End If
        .Close
    End With
    temDays = DateDiff("d", FromDate, ToDate) + 1
    If temDays <> 0 Then
        AvgMilk = TotalMilk / temDays
    Else
        AvgMilk = 0
    End If
    temCommsisionRate = OthersCommisionRate(SupplierID, AvgMilk)
    temCommsision = temCommsisionRate * TotalMilk
    OthersCommision = temCommsision
End Function

















Public Function Commision(SupplierID As Long, FromDate As Date, ToDate As Date, SecessionID As Long) As Double
    Commision = 0
    
   Dim rsVolume As New ADODB.Recordset
   Dim TotalVolume As Double
   Dim OwnVolume As Double
   
   Dim TotalDays As Integer
   Dim TemDay As Date
   TotalDays = DateDiff("d", FromDate, ToDate) + 1
   
   Dim i As Integer
   
   TemDay = FromDate
   
   For i = 1 To TotalDays
   
    
       With rsVolume
            If .State = 1 Then .Close
            If SecessionID = 0 Then
                temSQL = "SELECT Sum(tblCollection.Liters) AS SumOfLiters " & _
                            "FROM tblCollection LEFT JOIN tblSupplier ON tblCollection.SupplierID = tblSupplier.SupplierID " & _
                            "WHERE (((tblCollection.ProgramDate) Between '" & Format(TemDay, "dd MMMM yyyy") & "' And '" & Format(TemDay, "dd MMMM yyyy") & "') AND ((tblSupplier.CommisionCollectorID)=" & SupplierID & ")  And ((tblCollection.Deleted) = 0 ))"
            Else
                temSQL = "SELECT Sum(tblCollection.Liters) AS SumOfLiters " & _
                            "FROM tblCollection LEFT JOIN tblSupplier ON tblCollection.SupplierID = tblSupplier.SupplierID " & _
                            "WHERE (((tblCollection.ProgramDate) Between '" & Format(TemDay, "dd MMMM yyyy") & "' And '" & Format(TemDay, "dd MMMM yyyy") & "') AND ((tblSupplier.CommisionCollectorID)=" & SupplierID & ")  And ((tblCollection.Deleted) = 0 ) AND ((SecessionID) = " & SecessionID & "))"
            End If
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            If IsNull(!SumOfLiters) = False Then
                TotalVolume = !SumOfLiters
                .Close
            Else
                TotalVolume = 0
            End If
        End With
        
        With rsVolume
            If .State = 1 Then .Close
            If SecessionID = 0 Then
                temSQL = "SELECT Sum(tblCollection.Liters) AS SumOfLiters " & _
                            "FROM tblCollection LEFT JOIN tblSupplier ON tblCollection.SupplierID = tblSupplier.SupplierID " & _
                            "WHERE (((tblCollection.ProgramDate) Between '" & Format(TemDay, "dd MMMM yyyy") & "' And '" & Format(TemDay, "dd MMMM yyyy") & "') AND ((tblSupplier.SupplierID)=" & SupplierID & ")  And ((tblCollection.Deleted) = 0 ))"
            Else
                temSQL = "SELECT Sum(tblCollection.Liters) AS SumOfLiters " & _
                            "FROM tblCollection LEFT JOIN tblSupplier ON tblCollection.SupplierID = tblSupplier.SupplierID " & _
                            "WHERE (((tblCollection.ProgramDate) Between '" & Format(TemDay, "dd MMMM yyyy") & "' And '" & Format(TemDay, "dd MMMM yyyy") & "') AND ((tblSupplier.SupplierID)=" & SupplierID & ")  And ((tblCollection.Deleted) = 0 ) AND ((SecessionID) = " & SecessionID & "))"
            End If
                        
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            If IsNull(!SumOfLiters) = False Then
                OwnVolume = !SumOfLiters
            Else
                OwnVolume = 0
            End If
            .Close
        End With
    
        
        Dim rsSupplier As New ADODB.Recordset
        temSQL = "Select * from tblSupplier where SupplierID = " & SupplierID
        With rsSupplier
            If .State = 1 Then .Close
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                If !Commision = True Then
                    If !CommisionType = 1 Then
                        Commision = Commision + TotalVolume * CommisionRate(TotalVolume) + OwnVolume * CommisionRate(OwnVolume)
                    ElseIf !CommisionType = 2 Then
                        Commision = Commision + TotalVolume * (CommisionRate(TotalVolume) + !AdditionalCommision) + (OwnVolume * (CommisionRate(OwnVolume) + !AdditionalCommision))
                    ElseIf !CommisionType = 3 Then
                        Commision = Commision + (TotalVolume * !FixedCommision) + (OwnVolume * !FixedCommision)
                    ElseIf !CommisionType = 4 Then
                        Commision = Commision + (!FixedCommision * TotalVolume) + (OwnVolume * CommisionRate(OwnVolume)) + (!FixedCommision * OwnVolume)
                    Else
                        Commision = 0
                    End If
                Else
                    Commision = 0
                End If
            Else
                Commision = 0
            End If
            .Close
        End With
        WriteCommision SupplierID, TemDay, CommisionRate(OwnVolume), OwnVolume, SecessionID
        
        TemDay = FromDate + i
        
'        Dim rsCommision As New ADODB.Recordset
'        With rsCommision
'            If .State = 1 Then .Close
'            temSql = "Select * from tblCOllection where SupplierID = " & SupplierID & " And Date = " & TemDay
'
'
''        End If
'        End With
    Next
    
    temSQL = "Select * from tblAdditionalCommision where Deleted = 0  AND SupplierID = " & SupplierID & " AND CommisionDate between '" & Format(FromDate, "dd MMMM yyyy") & "' and '" & Format(ToDate, "dd MMMM yyyy") & "'"
    With rsSupplier
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Commision = Commision + !Value
        End If
        .Close
    End With
    
    
End Function

Public Function AdditionalCommision(SupplierID As Long, FromDate As Date, ToDate As Date) As Double
    AdditionalCommision = 0
    Dim rsCommision As New ADODB.Recordset
    temSQL = "Select sum( Value) as SumOfValue from tblAdditionalCommision where Deleted = 0  AND SupplierID = " & SupplierID & " AND CommisionDate between '" & Format(FromDate, "dd MMMM yyyy") & "' and '" & Format(ToDate, "dd MMMM yyyy") & "'"
    With rsCommision
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then AdditionalCommision = !SumOfValue
        End If
        .Close
    End With
End Function



Public Function PeriodAdditionalCommision(SupplierID As Long, FromDate As Date, ToDate As Date) As Double
    PeriodAdditionalCommision = 0
    Dim rsCommision As New ADODB.Recordset
    temSQL = "Select sum( Value) as SumOfValue from tblAdditionalCommision where Deleted = 0  AND Approved = 1 AND SupplierID = " & SupplierID & " AND CommisionDate between '" & Format(FromDate, "dd MMMM yyyy") & "' and '" & Format(ToDate, "dd MMMM yyyy") & "'"
    With rsCommision
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then PeriodAdditionalCommision = !SumOfValue
        End If
        .Close
    End With
End Function



'Public Function ActualCommisionRate(SupplierID As Long, Volume As Double)
'        Dim rsSupplier As New ADODB.Recordset
'        temSql = "Select * from tblSupplier where SupplierID = " & SupplierID
'        With rsSupplier
'            If .State = 1 Then .Close
'            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'            If .RecordCount > 0 Then
'                If !Commision = True Then
'                    If !CommisionType = 1 Then
'                        ActualCommisionRate = CommisionRate(Volume)
'                    ElseIf !CommisionType = 2 Then
'                        ActualCommisionRate = CommisionRate(Volume) + !AdditionalCommision
'                    ElseIf !CommisionType = 3 Then
'                        ActualCommisionRate = !FixedCommision
'                    ElseIf !CommisionType = 4 Then
'                        ActualCommisionRate = CommisionRate(Volume) '!FixedCommision
'                    Else
'                        ActualCommisionRate = 0
'                    End If
'                Else
'                    ActualCommisionRate = 0
'                End If
'            Else
'                ActualCommisionRate = 0
'            End If
'            .Close
'        End With
'End Function

Public Function OwnCommisionRate(SupplierID As Long, Volume As Double)

        ' ****************************************************
        '
        '   Called when entering the daily coollection
        '   frmDailyMilkCollectionReport
        '   temCr = OwnCommisionRate(cmbSupplierName.BoundText, Val(txtLiters.Text))
        '   !CommisionRate = temCr
        '   !Commision = temCr * Val(txtLiters)
        '
        ' ***************************************************

        Dim rsSupplier As New ADODB.Recordset
        temSQL = "Select * from tblSupplier where SupplierID = " & SupplierID
        With rsSupplier
            If .State = 1 Then .Close
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                If !Commision = True Then
                    If !CommisionType = 1 Then
                        OwnCommisionRate = CommisionRate(Volume)
                    ElseIf !CommisionType = 2 Then
                        OwnCommisionRate = CommisionRate(Volume) + !AdditionalCommision
                    ElseIf !CommisionType = 3 Then
                        OwnCommisionRate = !FixedCommision
                    ElseIf !CommisionType = 4 Then
                        OwnCommisionRate = CommisionRate(Volume)
                    Else
                        OwnCommisionRate = 0
                    End If
                Else
                    OwnCommisionRate = 0
                End If
            Else
                OwnCommisionRate = 0
            End If
            .Close
        End With
End Function

Public Function OthersCommisionRate(SupplierID As Long, Volume As Double)
        Dim rsSupplier As New ADODB.Recordset
        temSQL = "Select * from tblSupplier where SupplierID = " & SupplierID
        With rsSupplier
            If .State = 1 Then .Close
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                If !Commision = True Then
                    If !CommisionType = 1 Then
                        OthersCommisionRate = CommisionRate(Volume)
                    ElseIf !CommisionType = 2 Then
                        OthersCommisionRate = CommisionRate(Volume) + !AdditionalCommision
                    ElseIf !CommisionType = 3 Then
                        OthersCommisionRate = !FixedCommision
                    ElseIf !CommisionType = 4 Then
                        OthersCommisionRate = !FixedCommision
                    Else
                        OthersCommisionRate = 0
                    End If
                Else
                    OthersCommisionRate = 0
                End If
            Else
                OthersCommisionRate = 0
            End If
            .Close
        End With
End Function

Public Function CommisionRate(Volume As Double) As Double
    Dim rsCommisionRate As New ADODB.Recordset
    With rsCommisionRate
        If .State = 1 Then .Close
        temSQL = "SELECT tblCommisionRate.Commision FROM tblCommisionRate WHERE (((tblCommisionRate.Volume)=" & Round(Volume, 0) & "))"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            CommisionRate = !Commision
        Else
            CommisionRate = 0
        End If
        .Close
    End With
End Function


Public Sub WriteCommision(SupplierID As Long, SuppliedDate As Date, CommisionRate As Double, Commision As Double, SecessionID As Long)
    Dim rsCommision As New ADODB.Recordset
    If Commision <= 0 Then Exit Sub
    With rsCommision
        If .State = 1 Then .Close
        temSQL = "Select * from tblCollection where SupplierID = " & SupplierID & " And Date = '" & Format(SuppliedDate, "dd MMMM yyyy") & "' And SecessionID = " & SecessionID
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        Else
            .AddNew
            If SecessionID = 1 Then ' Morning
                !CollectionDate = SuppliedDate
                !DeliveryDate = SuppliedDate + 1
                !ProgramDate = SuppliedDate
            Else
                !CollectionDate = SuppliedDate
                !DeliveryDate = SuppliedDate + 2
                !ProgramDate = SuppliedDate + 1
            End If
        End If
        !Commision = Commision
        !CommisionRate = CommisionRate
        .Update
    End With
End Sub

Public Function PeriodDeductions(SupplierID As Long, FromDate As Date, ToDate As Date) As Double
    Dim rsDeductions As New ADODB.Recordset
    
    With rsDeductions
        temSQL = "SELECT Sum(tblDeduction.Value) AS SumOfValue " & _
                    "From tblDeduction " & _
                    "WHERE tblDeduction.Deleted = 0 AND tblDeduction.DeductedDate Between '" & Format(FromDate, "dd MMMM yyyy") & "' And '" & Format(ToDate, "dd MMMM yyyy") & "' AND tblDeduction.SupplierID=" & SupplierID & " "
                    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then
                PeriodDeductions = !SumOfValue
            Else
                PeriodDeductions = 0
            End If
        Else
            PeriodDeductions = 0
        End If
        .Close
    End With
    
    Dim AdditionalDeduction As Double
    
    AdditionalDeduction = 0
    Dim rsDeduction As New ADODB.Recordset
    temSQL = "Select sum( Value) as SumOfValue from tblAdditionalDeduction where Deleted = 0  AND SupplierID = " & SupplierID & " AND DeductionDate between '" & Format(FromDate, "dd MMMM yyyy") & "' and '" & Format(ToDate, "dd MMMM yyyy") & "'"
    With rsDeduction
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then AdditionalDeduction = !SumOfValue
        End If
        .Close
    End With
    
    Dim temPeriodVolumeDeductions As Double
    temPeriodVolumeDeductions = 0
    
    With rsDeductions
        temSQL = "SELECT Sum(tblFarmerDailyVolumeDeduction.DeductedValue) AS SumOfValue " & _
                    "From tblFarmerDailyVolumeDeduction " & _
                    "WHERE tblFarmerDailyVolumeDeduction.Deleted = 0 AND tblFarmerDailyVolumeDeduction.DeductedDate Between '" & Format(FromDate, "dd MMMM yyyy") & "' And '" & Format(ToDate, "dd MMMM yyyy") & "' AND tblFarmerDailyVolumeDeduction.SupplierID=" & SupplierID & " "
                    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then
                temPeriodVolumeDeductions = !SumOfValue
            Else
                temPeriodVolumeDeductions = 0
            End If
        Else
            temPeriodVolumeDeductions = 0
        End If
        .Close
    End With
    
    
    PeriodDeductions = PeriodDeductions + AdditionalDeduction + temPeriodVolumeDeductions
    
End Function

Public Function PeriodDeductionsComponent2(SupplierID As Long, FromDate As Date, ToDate As Date) As Double
    Dim rsDeductions As New ADODB.Recordset
    
    Dim AdditionalDeduction As Double
    
    AdditionalDeduction = 0
    Dim rsDeduction As New ADODB.Recordset
    temSQL = "Select sum( Value) as SumOfValue from tblAdditionalDeduction where Deleted = 0  AND SupplierID = " & SupplierID & " AND DeductionDate between '" & Format(FromDate, "dd MMMM yyyy") & "' and '" & Format(ToDate, "dd MMMM yyyy") & "'"
    With rsDeduction
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then AdditionalDeduction = !SumOfValue
        End If
        .Close
    End With
    
    PeriodDeductionsComponent2 = AdditionalDeduction
    
End Function

Public Function PeriodDeductionsComponent3(SupplierID As Long, FromDate As Date, ToDate As Date) As Double
    Dim rsDeductions As New ADODB.Recordset
    
    Dim temPeriodVolumeDeductions As Double
    temPeriodVolumeDeductions = 0
    
    With rsDeductions
        temSQL = "SELECT Sum(tblFarmerDailyVolumeDeduction.DeductedValue) AS SumOfValue " & _
                    "From tblFarmerDailyVolumeDeduction " & _
                    "WHERE tblFarmerDailyVolumeDeduction.Deleted = 0 AND tblFarmerDailyVolumeDeduction.DeductedDate Between '" & Format(FromDate, "dd MMMM yyyy") & "' And '" & Format(ToDate, "dd MMMM yyyy") & "' AND tblFarmerDailyVolumeDeduction.SupplierID=" & SupplierID & " "
                    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then
                temPeriodVolumeDeductions = !SumOfValue
            Else
                temPeriodVolumeDeductions = 0
            End If
        Else
            temPeriodVolumeDeductions = 0
        End If
        .Close
    End With
    
    
    PeriodDeductionsComponent3 = temPeriodVolumeDeductions
    
End Function


Public Function PeriodDeductionsComponent1(SupplierID As Long, FromDate As Date, ToDate As Date) As Double
    Dim rsDeductions As New ADODB.Recordset
    
    With rsDeductions
        temSQL = "SELECT Sum(tblDeduction.Value) AS SumOfValue " & _
                    "From tblDeduction " & _
                    "WHERE tblDeduction.Deleted = 0 AND tblDeduction.DeductedDate Between '" & Format(FromDate, "dd MMMM yyyy") & "' And '" & Format(ToDate, "dd MMMM yyyy") & "' AND tblDeduction.SupplierID=" & SupplierID & " "
                    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then
                PeriodDeductionsComponent1 = !SumOfValue
            Else
                PeriodDeductionsComponent1 = 0
            End If
        Else
            PeriodDeductionsComponent1 = 0
        End If
        .Close
    End With
    
End Function




Public Function PeriodCattleFeedDeductions(SupplierID As Long, FromDate As Date, ToDate As Date) As Double
    Dim rsDeductions As New ADODB.Recordset
    With rsDeductions
        temSQL = "SELECT Sum(tblDeduction.Value) AS SumOfValue " & _
                    "From tblDeduction " & _
                    "WHERE tblDeduction.Deleted = 0 AND tblDeduction.DeductedDate Between '" & Format(FromDate, "dd MMMM yyyy") & "' And '" & Format(ToDate, "dd MMMM yyyy") & "' AND tblDeduction.SupplierID=" & SupplierID & " "
                    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then
                PeriodCattleFeedDeductions = !SumOfValue
            Else
                PeriodCattleFeedDeductions = 0
            End If
        Else
            PeriodCattleFeedDeductions = 0
        End If
        .Close
    End With
End Function

Public Function PeriodVolumeDeductions(SupplierID As Long, FromDate As Date, ToDate As Date) As Double
    Dim rsDeductions As New ADODB.Recordset
    With rsDeductions
        temSQL = "SELECT Sum(tblFarmerDailyVolumeDeduction.DeductedValue) AS SumOfValue " & _
                    "From tblFarmerDailyVolumeDeduction " & _
                    "WHERE tblFarmerDailyVolumeDeduction.Deleted = 0 AND tblFarmerDailyVolumeDeduction.DeductedDate Between '" & Format(FromDate, "dd MMMM yyyy") & "' And '" & Format(ToDate, "dd MMMM yyyy") & "' AND tblFarmerDailyVolumeDeduction.SupplierID=" & SupplierID & " "
                    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then
                PeriodVolumeDeductions = !SumOfValue
            Else
                PeriodVolumeDeductions = 0
            End If
        Else
            PeriodVolumeDeductions = 0
        End If
        .Close
    End With
End Function

Public Function PeriodAdditionalDeductions(SupplierID As Long, FromDate As Date, ToDate As Date) As Double
    Dim AdditionalDeduction As Double
    
    PeriodAdditionalDeductions = 0
    Dim rsDeduction As New ADODB.Recordset
    temSQL = "Select sum( Value) as SumOfValue from tblAdditionalDeduction where Deleted = 0  AND SupplierID = " & SupplierID & " AND DeductionDate between '" & Format(FromDate, "dd MMMM yyyy") & "' and '" & Format(ToDate, "dd MMMM yyyy") & "'"
    With rsDeduction
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then PeriodAdditionalDeductions = !SumOfValue
        End If
        .Close
    End With
    
End Function


Public Function SupplierPaymentMethodID(SupplierID As Long)
    Dim rsSupplierPM As New ADODB.Recordset
    With rsSupplierPM
        If .State = 1 Then .Close
        temSQL = "Select * from tblSupplier where SupplierID = " & SupplierID
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!PaymentMethodID) = False Then
                SupplierPaymentMethodID = !PaymentMethodID
            Else
                SupplierPaymentMethodID = 0
            End If
        Else
            SupplierPaymentMethodID = 0
        End If
        .Close
    End With
    Set rsSupplierPM = Nothing
End Function



Public Function CCMilkSupply(SuppliedDate As Date, CCID As Long) As MilkCollection
    Dim rsMilk As New ADODB.Recordset
    Dim MyMilkCollection As MilkCollection
    With rsMilk
        If .State = 1 Then .Close
        temSQL = "SELECT sum(tblCollection.Liters) as SumOfLiters , sum(tblCollection.Value ) as SumOfValue " & _
                    "FROM tblCollection " & _
                    "WHERE (((tblCollection.Date) = '" & Format(SuppliedDate, "dd MMMM yyyy") & "') AND ((tblCollection.CollectingCenterID)=" & CCID & ")  And ((tblCollection.Deleted) = 0 ) )"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            MyMilkCollection.Supplied = True
            MyMilkCollection.LMR = 0
            MyMilkCollection.FAT = 0
            
            If IsNull(!SumOfLiters) = False Then
                MyMilkCollection.Liters = !SumOfLiters
            Else
                MyMilkCollection.Liters = 0
            End If
            
            MyMilkCollection.SNF = 0
            MyMilkCollection.Price = 0
            
            If IsNull(!SumOfValue) = False Then
                MyMilkCollection.Value = !SumOfValue
            Else
                MyMilkCollection.Value = 0
            End If
            
            If MyMilkCollection.Liters <> 0 Then
                MyMilkCollection.Price = MyMilkCollection.Value / MyMilkCollection.Liters
            Else
                MyMilkCollection.Price = 0
            End If
            
            MyMilkCollection.OwnCommision = 0
            MyMilkCollection.OwnCommisionRate = 0
        Else
            MyMilkCollection.Supplied = False
            MyMilkCollection.LMR = 0
            MyMilkCollection.FAT = 0
            MyMilkCollection.Liters = 0
            MyMilkCollection.SNF = 0
            MyMilkCollection.Price = 0
            MyMilkCollection.Value = 0
            MyMilkCollection.OwnCommision = 0
            MyMilkCollection.OwnCommisionRate = 0
        End If
        .Close
    End With
    CCMilkSupply = MyMilkCollection
    Set rsMilk = Nothing
End Function

Public Function CCGRN(SuppliedDate As Date, CCID As Long) As MilkCollection
    Dim rsMilk As New ADODB.Recordset
    Dim MyMilkCollection As MilkCollection
    With rsMilk
        If .State = 1 Then .Close
        temSQL = "SELECT Sum(tblDailyCollection.ActualValue) AS SumOfActualValue, Sum(tblDailyCollection.ActualVolume) AS SumOfActualVolume  " & _
                    "FROM tblDailyCollection " & _
                    "WHERE (((tblDailyCollection.ProgramDate)='" & Format(SuppliedDate, "dd MMMM yyyy") & "') AND ((tblDailyCollection.CollectingCenterID)=" & CCID & "))"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            MyMilkCollection.Supplied = True
            MyMilkCollection.LMR = 0
            MyMilkCollection.FAT = 0
            
            If IsNull(!SumOfActualVolume) = False Then
                MyMilkCollection.Liters = !SumOfActualVolume
            Else
                MyMilkCollection.Liters = 0
            End If
            
            If IsNull(!SumOfActualValue) = False Then
                MyMilkCollection.Value = !SumOfActualValue
            Else
                MyMilkCollection.Value = 0
            End If
            
            If MyMilkCollection.Liters <> 0 Then
                MyMilkCollection.Price = MyMilkCollection.Value / MyMilkCollection.Liters
            Else
                MyMilkCollection.Price = 0
            End If
            
            MyMilkCollection.SNF = 0
            MyMilkCollection.OwnCommision = 0
            MyMilkCollection.OwnCommisionRate = 0
        Else
            MyMilkCollection.Supplied = False
            MyMilkCollection.LMR = 0
            MyMilkCollection.FAT = 0
            MyMilkCollection.Liters = 0
            MyMilkCollection.SNF = 0
            MyMilkCollection.Price = 0
            MyMilkCollection.Value = 0
            MyMilkCollection.OwnCommision = 0
            MyMilkCollection.OwnCommisionRate = 0
        End If
        .Close
    End With
    CCGRN = MyMilkCollection
    Set rsMilk = Nothing
End Function
