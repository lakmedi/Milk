VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Lakmedipro Milk Payment System"
   ClientHeight    =   7920
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9705
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileBackup 
         Caption         =   "Backup"
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuTem 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCollectingCenters 
         Caption         =   "Collecting Centers"
      End
      Begin VB.Menu mnuSuppliers 
         Caption         =   "&Farmers"
      End
      Begin VB.Menu mnuEditPrices 
         Caption         =   "Prices"
         Begin VB.Menu mnuEditPaymentScheme 
            Caption         =   "Payment Scheme"
         End
         Begin VB.Menu mnuEditPriceCycle 
            Caption         =   "Price Cycle"
         End
         Begin VB.Menu mnuEditPricesPrices 
            Caption         =   "Prices"
         End
      End
      Begin VB.Menu mnuItemS 
         Caption         =   "Item"
      End
      Begin VB.Menu mnuEditStaff 
         Caption         =   "Staff"
      End
      Begin VB.Menu mnuEditAuthority 
         Caption         =   "Authority"
      End
      Begin VB.Menu mnuEditIncomeCategory 
         Caption         =   "Income Category"
      End
      Begin VB.Menu mnuEditExpenceCategory 
         Caption         =   "Expence Category"
      End
      Begin VB.Menu mnuItemSuppliers 
         Caption         =   "Item Suppliers"
      End
      Begin VB.Menu mnuEditVolumeDeductions 
         Caption         =   "Volume Deductions Types"
      End
      Begin VB.Menu mnuEditFarmerVolumeDeduction 
         Caption         =   "Add Volume Deductions to Farmers"
      End
      Begin VB.Menu mnuEditExit 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuMilkCollection 
      Caption         =   "&Milk Collection"
      Begin VB.Menu mnuDailyCollection 
         Caption         =   "Daily Collection"
      End
      Begin VB.Menu mnuMilkCollectionLabTestings 
         Caption         =   "Good Recieve Note"
      End
      Begin VB.Menu mnuCumulativeReport 
         Caption         =   "Cumulative Report"
      End
      Begin VB.Menu mnuMilkCollectionCollectingCenterPayAdvice 
         Caption         =   "Generate Collecting Center Pay Advice"
      End
      Begin VB.Menu mnuPrintCollectingCenterPayAdvice 
         Caption         =   "Print Collecting Center Pay Advice"
      End
      Begin VB.Menu mnuMilkPayAdvice 
         Caption         =   "Milk Pay Advice"
      End
      Begin VB.Menu mnuMilkExit 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuIssueItems 
      Caption         =   "&Issues And Payments"
      Begin VB.Menu mnuIssueItemIssue 
         Caption         =   "Item Issue"
      End
      Begin VB.Menu mnuItemPurchase 
         Caption         =   "Item Purchase"
      End
      Begin VB.Menu mnuIssuesAddAdditionalPayments 
         Caption         =   "Additional Commisions"
      End
      Begin VB.Menu mnuAddAdditioanlDeductions 
         Caption         =   "Additional Deductions"
      End
      Begin VB.Menu mnuIssuesExpences 
         Caption         =   "Expences"
      End
      Begin VB.Menu mnuIssuesIncome 
         Caption         =   "Income"
      End
      Begin VB.Menu mnuIssueExit 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuBAckOffice 
      Caption         =   "&Back Office"
      Begin VB.Menu mnuBackOfficePrintIndividualPayments 
         Caption         =   "Print Payments"
      End
      Begin VB.Menu mnuApprovals 
         Caption         =   "Approvals"
         Begin VB.Menu mnuApproveAdditionalCommisions 
            Caption         =   "Additional Commisions"
         End
         Begin VB.Menu mnuApproveAdditionalDeductions 
            Caption         =   "Additional Deductions"
         End
      End
      Begin VB.Menu mnuOutstandings 
         Caption         =   "Outsandings"
      End
      Begin VB.Menu mnuBackOfficeReports 
         Caption         =   "Reports"
         Begin VB.Menu mnuBackOfficeMIlkCollection 
            Caption         =   "Milk Collection"
            Begin VB.Menu mnuBackOfficeMilkCollectionDailyCollection 
               Caption         =   "Daily Collection"
            End
            Begin VB.Menu mnuBackOfficeMilkCollectionPayAdvice 
               Caption         =   "Collecting Center Payment Summery"
            End
            Begin VB.Menu mnuBackOfficeMIlkCollectionCumulativeReport 
               Caption         =   "Cumulative Report"
            End
            Begin VB.Menu mnuBackOfficeAllCentreCumulativeReport 
               Caption         =   "All centers cumulative report"
            End
            Begin VB.Menu mnuBAckOfficeALlCentrePaymentSummery 
               Caption         =   "All Center Payment Summery"
            End
         End
         Begin VB.Menu mnuBackOfficeReportsSuppliers 
            Caption         =   "Suppliers"
            Begin VB.Menu mnuBackOfficeReportsSuppliersMilkCollection 
               Caption         =   "Milk Collection"
            End
         End
         Begin VB.Menu mnuBackOfficeReportsCollectingCenters 
            Caption         =   "Collecting Centers"
            Begin VB.Menu mnuReportsCollectingCentersPaymentSchemeMilkSupply 
               Caption         =   "Payment Scheme Milk Supply"
            End
            Begin VB.Menu mnuBackOfficeReportsCollectingCentersMilkCollection 
               Caption         =   "Milk Collection"
            End
            Begin VB.Menu mnuReportsTotalMilkCollection 
               Caption         =   "Total Milk Collection"
            End
         End
         Begin VB.Menu mnuBackOfficeReportsMilkPaymentsCommisionsandOtherExpences 
            Caption         =   "Milk Payments, Commisions and Other Expences"
         End
         Begin VB.Menu mnuReportExpences 
            Caption         =   "Expences"
         End
         Begin VB.Menu mnuReportExpenceCategories 
            Caption         =   "Expence Categories"
         End
         Begin VB.Menu mnuAdditionalCommisionReport 
            Caption         =   "Additional Commision Report"
         End
         Begin VB.Menu mnuAdditionalDeductionReport 
            Caption         =   "Additional Deduction Report"
         End
         Begin VB.Menu mnuVitaminAndCattleFeedDeductionReport 
            Caption         =   "Vitamin & Cattle Feed Deduction Report"
         End
      End
      Begin VB.Menu mnuProfits 
         Caption         =   "Profits"
         Begin VB.Menu mnuProfitsIncome 
            Caption         =   "Income"
         End
         Begin VB.Menu mnuProfitsExpence 
            Caption         =   "Expence"
         End
         Begin VB.Menu mnuProfitsProfit 
            Caption         =   "Profit"
         End
      End
      Begin VB.Menu mnuDetectErrors 
         Caption         =   "Detect Errors"
         Begin VB.Menu mnuDeleteErrors 
            Caption         =   "Delete Errors"
         End
      End
      Begin VB.Menu mnuBackOfficeExit 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsDatabase 
         Caption         =   "Database"
      End
      Begin VB.Menu mnuOptionsPrinting 
         Caption         =   "Printing"
      End
      Begin VB.Menu mnuOptionsExit 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuPersonalDetails 
      Caption         =   "Personal Details"
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    
Private Sub mnuAdditionalCommisionReport_Click()
    frmReportAdditionalCommision.Show
    frmReportAdditionalCommision.ZOrder 0
End Sub

Private Sub mnuAdditionalDeductionReport_Click()
    frmReportAdditionalDeductions.Show
    frmReportAdditionalDeductions.ZOrder 0
End Sub

Private Sub mnuBackOfficeAllCentreCumulativeReport_Click()
    frmCumulativeDailyMilkCollectionAll.Show
    frmCumulativeDailyMilkCollectionAll.ZOrder 0
End Sub

Private Sub mnuBAckOfficeALlCentrePaymentSummery_Click()
    frmCollectingCentrePaymentSummeryAll.Show
    frmCollectingCentrePaymentSummeryAll.ZOrder 0
End Sub

Private Sub mnuBackOfficeReportsCollectingCentersMilkCollection_Click()
    frmReportCollectingCenterviceMilkCollection.Show
    frmReportCollectingCenterviceMilkCollection.ZOrder 0
End Sub



Private Sub correctError()
'    temSQL = "SELECT dbo.tblCollection.CollectionID, dbo.tblCollectingCenter.CollectingCenterID " & _
'        "                 " & _
'"FROM          dbo.tblCollectingCenter LEFT OUTER JOIN " & _
'  "                      dbo.tblSupplier ON dbo.tblCollectingCenter.CollectingCenterID = dbo.tblSupplier.CollectingCenterID RIGHT OUTER JOIN " & _
' "                       dbo.tblCollection ON dbo.tblSupplier.SupplierID = dbo.tblCollection.SupplierID " & _
'"Where (dbo.tblCollectingCenter.CollectingCenterID <> dbo.tblCollection.CollectingCenterID) " & _
'"ORDER BY dbo.tblCollection.CollectionID DESC"
'
'    Dim rsTem As New ADODB.Recordset
'    With rsTem
'        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            MsgBox "Errors in Collection. Corrected " & .RecordCount & " records."
'        End If
'        While .EOF = False
'            correctCentreId !collectionId, !CollectingCenterID
'            .MoveNext
'        Wend
'    End With
End Sub

Private Sub correctCentreId(collectionId As Long, CCID As Long)
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSQL = "select * from dbo.tblCollection where dbo.tblCollection.CollectionID =  " & collectionId
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !CollectingCenterID = CCID
            .Update
        End If
        .Close
    End With
End Sub

Private Sub MDIForm_Load()
    
    correctError
    
    
    On Error Resume Next
    
    Me.Caption = Me.Caption & "     " & UserFullName & "    -   " & " " & UserName & "      " & Time & "        " & Format(Date, "dd  dddd  mmmm yyyy")

    If UserAuthorityLevel <> Administrator Then
        mnuPersonalDetails.Visible = True
        mnuFile.Visible = False
        mnuFileBackup.Visible = False
        mnuFileRestore.Visible = False
'        mnuFileExit.Visible = False

        mnuEdit.Visible = False
        mnuEditCollectingCenters.Visible = False
        mnuSuppliers.Visible = False
        mnuEditPrices.Visible = False
        mnuItemS.Visible = False
        mnuEditStaff.Visible = False
        mnuEditAuthority.Visible = False
        mnuEditIncomeCategory.Visible = False
        mnuEditExpenceCategory.Visible = False
'        mnuItemSuppliers.Visible = False

        mnuMilkCollection.Visible = False
        mnuDailyCollection.Visible = False
        mnuMilkCollectionLabTestings.Visible = False
        mnuCumulativeReport.Visible = False
        mnuMilkCollectionCollectingCenterPayAdvice.Visible = False
        mnuPrintCollectingCenterPayAdvice.Visible = False
'        mnuMilkPayAdvice.Visible = False

        mnuIssueItems.Visible = False
        mnuIssueItemIssue.Visible = False
        mnuItemPurchase.Visible = False
'        mnuIssuesAddDeductions.Visible = False
        mnuIssuesAddAdditionalPayments.Visible = False
        mnuAddAdditioanlDeductions.Visible = False
        mnuIssuesExpences.Visible = False
'        mnuIssuesIncome.Visible = False

        mnuBAckOffice.Visible = False
'        mnuBackOfficeGenerateIndividualPayments.Visible = False
        mnuBackOfficePrintIndividualPayments.Visible = False
'        mnuConfirmPayments.Visible = False
        mnuApprovals.Visible = False
        mnuApproveAdditionalCommisions.Visible = False
'        mnuApproveAdditionalDeductions.Visible = False
        mnuOutstandings.Visible = False
        mnuBackOfficeReports.Visible = False
        mnuBackOfficeMIlkCollection.Visible = False
        mnuBackOfficeMilkCollectionDailyCollection.Visible = False
        mnuBackOfficeMilkCollectionPayAdvice.Visible = False
'        mnuBackOfficeMIlkCollectionCumulativeReport.Visible = False
        mnuBackOfficeReportsSuppliers.Visible = False
'        mnuBackOfficeReportsSuppliersMilkCollection.Visible = False
        mnuBackOfficeReportsCollectingCenters.Visible = False
        mnuBackOfficeReportsCollectingCentersMilkCollection.Visible = False
        mnuReportExpences.Visible = False

        mnuProfits.Visible = False
        mnuProfitsExpence.Visible = False
        mnuProfitsIncome.Visible = False
'        mnuProfitsProfit.Visible = False

        mnuOptions.Visible = False
        mnuOptionsDatabase.Visible = False
'        mnuOptionsPrinting.Visible = False
        mnuPersonalDetails.Visible = False

        mnuWindow.Visible = False
'        mnuBackOfficeGenerateIndividualPayments.Visible = False
        mnuBackOfficePrintIndividualPayments.Visible = False
'        mnuConfirmPayments.Visible = False
'        mnuDetectErrors.Visible = False
    End If



Select Case UserAuthorityLevel

        Case Authority.Viewer '1
        
            mnuEdit.Visible = True
            mnuFile.Visible = True
            mnuIssueItems.Visible = True
            mnuBAckOffice.Visible = True
            mnuWindow.Visible = True
            mnuOptions.Visible = True
            mnuMilkCollection.Visible = True
            
            mnuCumulativeReport.Visible = True
            mnuMilkPayAdvice.Visible = True

            mnuBackOfficeReports.Visible = True
            mnuReportExpences.Visible = False
            mnuDetectErrors.Visible = False

            mnuBackOfficeMIlkCollection.Visible = True
            mnuBackOfficeReportsSuppliers.Visible = True
            mnuBackOfficeReportsCollectingCenters.Visible = True
            mnuAdditionalCommisionReport.Visible = True
            mnuAdditionalDeductionReport.Visible = True
            mnuVitaminAndCattleFeedDeductionReport.Visible = True
            mnuBackOfficeMilkCollectionDailyCollection.Visible = True
            mnuBackOfficeMilkCollectionPayAdvice.Visible = True
            mnuBackOfficeMIlkCollectionCumulativeReport.Visible = True
            mnuBackOfficeReportsSuppliersMilkCollection.Visible = True
            mnuBackOfficeReportsCollectingCentersMilkCollection.Visible = True
            mnuReportsTotalMilkCollection.Visible = True

            mnuReportExpences.Visible = False
            mnuDetectErrors.Visible = False
            
            mnuPersonalDetails.Visible = True

        Case Authority.Analyzer '2
            
            mnuDailyCollection.Visible = True
            mnuCumulativeReport.Visible = True
            mnuMilkPayAdvice.Visible = True
            
            mnuFileBackup.Visible = True
            mnuFileRestore.Visible = True
            
            mnuBAckOffice.Visible = True
            mnuOutstandings.Visible = False
            mnuBackOfficeReports.Visible = True
            
            mnuEdit.Visible = True
            mnuFile.Visible = True
            mnuIssueItems.Visible = True
            mnuBAckOffice.Visible = True
            mnuWindow.Visible = True
            mnuOptions.Visible = True
            mnuMilkCollection.Visible = True
            
            mnuBackOfficeReports.Visible = True
            mnuReportExpences.Visible = False
            mnuDetectErrors.Visible = False
            mnuBackOfficeMIlkCollection.Visible = True
            mnuBackOfficeReportsSuppliers.Visible = True
            mnuBackOfficeReportsCollectingCenters.Visible = True
            mnuAdditionalCommisionReport.Visible = True
            mnuAdditionalDeductionReport.Visible = True
            mnuVitaminAndCattleFeedDeductionReport.Visible = True
            mnuBackOfficeMilkCollectionDailyCollection.Visible = True
            mnuBackOfficeMilkCollectionPayAdvice.Visible = True
            mnuBackOfficeMIlkCollectionCumulativeReport.Visible = True
            mnuBackOfficeReportsSuppliersMilkCollection.Visible = True
            mnuBackOfficeReportsCollectingCentersMilkCollection.Visible = True
            mnuReportsTotalMilkCollection.Visible = True
            
            mnuPersonalDetails.Visible = True

        Case Authority.OrdinaryUser '3

            mnuEdit.Visible = True
            mnuFile.Visible = True
            mnuIssueItems.Visible = True
            mnuBAckOffice.Visible = True
            mnuWindow.Visible = True
            mnuOptions.Visible = True
            mnuMilkCollection.Visible = True

            mnuItemSuppliers.Visible = True

            mnuDailyCollection.Visible = True
            mnuMilkCollectionLabTestings.Visible = True
            mnuMilkCollectionCollectingCenterPayAdvice.Visible = True

            mnuCumulativeReport.Visible = True
            mnuPrintCollectingCenterPayAdvice.Visible = True
            mnuMilkPayAdvice.Visible = True

            mnuIssueItemIssue.Visible = True
            mnuItemPurchase.Visible = True
'            mnuIssuesAddDeductions.Visible = True
            mnuIssuesAddAdditionalPayments.Visible = True
            mnuAddAdditioanlDeductions.Visible = True
            mnuIssuesExpences.Visible = True
            mnuIssuesIncome.Visible = True

            mnuOutstandings.Visible = True
            mnuBackOfficeReports.Visible = True
            mnuProfits.Visible = True
            mnuProfitsIncome.Visible = True
            mnuProfitsExpence.Visible = True
            mnuProfitsProfit.Visible = True

            mnuItemSuppliers.Visible = True

            mnuDetectErrors.Visible = False
            mnuBAckOffice.Visible = True
            mnuBackOfficeReports.Visible = True
            mnuReportExpences.Visible = False
            mnuDetectErrors.Visible = False
            mnuBackOfficeMIlkCollection.Visible = True
            mnuBackOfficeReportsSuppliers.Visible = True
            mnuBackOfficeReportsCollectingCenters.Visible = True
            mnuAdditionalCommisionReport.Visible = True
            mnuAdditionalDeductionReport.Visible = True
            mnuVitaminAndCattleFeedDeductionReport.Visible = True
            mnuBackOfficeMilkCollectionDailyCollection.Visible = True
            mnuBackOfficeMilkCollectionPayAdvice.Visible = True
            mnuBackOfficeMIlkCollectionCumulativeReport.Visible = True
            mnuBackOfficeReportsSuppliersMilkCollection.Visible = True
            mnuBackOfficeReportsCollectingCentersMilkCollection.Visible = True
            mnuReportsTotalMilkCollection.Visible = True
            
            mnuPersonalDetails.Visible = True

        Case Authority.PowerUser '4
        
        
            mnuEdit.Visible = True
            mnuFile.Visible = True
            mnuIssueItems.Visible = True
            mnuBAckOffice.Visible = True
            mnuWindow.Visible = True
            mnuOptions.Visible = True
            mnuMilkCollection.Visible = True
            
            mnuSuppliers.Visible = True
            mnuItemS.Visible = True
            mnuEditExpenceCategory.Visible = True
            mnuEditIncomeCategory.Visible = True
            mnuItemSuppliers.Visible = False

            mnuDailyCollection.Visible = True
            mnuMilkCollectionLabTestings.Visible = True
            mnuCumulativeReport.Visible = True
            mnuMilkCollectionCollectingCenterPayAdvice.Visible = True
            mnuPrintCollectingCenterPayAdvice.Visible = True
            mnuMilkPayAdvice.Visible = True

            mnuIssueItemIssue.Visible = True
            mnuItemPurchase.Visible = True
'            mnuIssuesAddDeductions.Visible = True
            mnuIssuesAddAdditionalPayments.Visible = True
            mnuAddAdditioanlDeductions.Visible = True
            mnuIssuesExpences.Visible = True
            mnuIssuesIncome.Visible = True

'            mnuBackOfficeGenerateIndividualPayments.Visible = True
            mnuBackOfficePrintIndividualPayments.Visible = True
            mnuOutstandings.Visible = True
            mnuBackOfficeReports.Visible = True
            mnuProfits.Visible = True

            mnuDetectErrors.Visible = False
            mnuBAckOffice.Visible = True
            mnuBackOfficeReports.Visible = True
            mnuReportExpences.Visible = True
            mnuDetectErrors.Visible = False
            mnuBackOfficeMIlkCollection.Visible = True
            mnuBackOfficeReportsSuppliers.Visible = True
            mnuBackOfficeReportsCollectingCenters.Visible = True
            mnuAdditionalCommisionReport.Visible = True
            mnuAdditionalDeductionReport.Visible = True
            mnuVitaminAndCattleFeedDeductionReport.Visible = True
            mnuBackOfficeMilkCollectionDailyCollection.Visible = True
            mnuBackOfficeMilkCollectionPayAdvice.Visible = True
            mnuBackOfficeMIlkCollectionCumulativeReport.Visible = True
            mnuBackOfficeReportsSuppliersMilkCollection.Visible = True
            mnuBackOfficeReportsCollectingCentersMilkCollection.Visible = True
            mnuReportsTotalMilkCollection.Visible = True
'            mnuBackOfficeGenerateIndividualPayments.Visible = True
            mnuBackOfficePrintIndividualPayments.Visible = True
            mnuOutstandings.Visible = True
            mnuProfits.Visible = True
            mnuProfitsIncome.Visible = True
            mnuProfitsExpence.Visible = True
            mnuProfitsProfit.Visible = True

'            mnuBackOfficeGenerateIndividualPayments.Visible = True
            mnuBackOfficePrintIndividualPayments.Visible = True
'            mnuConfirmPayments.Visible = True
            mnuDetectErrors.Visible = True
            mnuDeleteErrors.Visible = True
            
            mnuPersonalDetails.Visible = True

        Case Authority.SuperUser '5

            mnuEdit.Visible = True
            mnuFile.Visible = True
            mnuIssueItems.Visible = True
            mnuBAckOffice.Visible = True
            mnuWindow.Visible = True
            mnuOptions.Visible = True
            mnuMilkCollection.Visible = True

            mnuFileBackup.Visible = True
            mnuFileRestore.Visible = True

            mnuEditCollectingCenters.Visible = True
            mnuSuppliers.Visible = True
            mnuEditPrices.Visible = True
            mnuItemS.Visible = True
            mnuEditStaff.Visible = False
            mnuEditAuthority.Visible = False
            mnuEditIncomeCategory.Visible = True
            mnuEditExpenceCategory.Visible = True
            mnuItemSuppliers.Visible = True

            mnuDailyCollection.Visible = True
            mnuMilkCollectionLabTestings.Visible = True
            mnuCumulativeReport.Visible = True
            mnuMilkCollectionCollectingCenterPayAdvice.Visible = True
            mnuPrintCollectingCenterPayAdvice.Visible = True
            mnuMilkPayAdvice.Visible = True

            mnuIssueItemIssue.Visible = True
            mnuItemPurchase.Visible = True
'            mnuIssuesAddDeductions.Visible = True
            mnuIssuesAddAdditionalPayments.Visible = True
            mnuAddAdditioanlDeductions.Visible = True
            mnuIssuesExpences.Visible = True
            mnuIssuesIncome.Visible = True

'            mnuBackOfficeGenerateIndividualPayments.Visible = True
            mnuBackOfficePrintIndividualPayments.Visible = True
'            mnuConfirmPayments.Visible = True
            mnuApprovals.Visible = True
            mnuOutstandings.Visible = True
            mnuBackOfficeReports.Visible = True
            mnuProfits.Visible = True

            mnuOptionsDatabase.Visible = False
            mnuOptionsPrinting.Visible = True

            mnuPersonalDetails.Visible = True

            mnuDetectErrors.Visible = True
            mnuDeleteErrors.Visible = True
            mnuBAckOffice.Visible = True
            mnuBackOfficeReports.Visible = True
            mnuReportExpences.Visible = True
            mnuBackOfficeMIlkCollection.Visible = True
            mnuBackOfficeReportsSuppliers.Visible = True
            mnuBackOfficeReportsCollectingCenters.Visible = True
            mnuAdditionalCommisionReport.Visible = True
            mnuAdditionalDeductionReport.Visible = True
            mnuVitaminAndCattleFeedDeductionReport.Visible = True
            mnuBackOfficeMilkCollectionDailyCollection.Visible = True
            mnuBackOfficeMilkCollectionPayAdvice.Visible = True
            mnuBackOfficeMIlkCollectionCumulativeReport.Visible = True
            mnuBackOfficeReportsSuppliersMilkCollection.Visible = True
            mnuBackOfficeReportsCollectingCentersMilkCollection.Visible = True
            mnuReportsTotalMilkCollection.Visible = True
            mnuProfitsIncome.Visible = True
            mnuProfitsExpence.Visible = True
            mnuProfitsProfit.Visible = True

            mnuApproveAdditionalDeductions.Visible = True
            mnuApproveAdditionalCommisions.Visible = True

'            mnuBackOfficeGenerateIndividualPayments.Visible = True
            mnuBackOfficePrintIndividualPayments.Visible = True
'            mnuConfirmPayments.Visible = True
            mnuDetectErrors.Visible = True
            mnuDeleteErrors.Visible = True

        Case Authority.Administrator '6
        
            mnuEdit.Visible = True
            mnuFile.Visible = True
            mnuIssueItems.Visible = True
            mnuBAckOffice.Visible = True
            mnuWindow.Visible = True
            mnuOptions.Visible = True
            mnuMilkCollection.Visible = True
        
            mnuFileBackup.Visible = True
            mnuFileRestore.Visible = True

            mnuEditCollectingCenters.Visible = True
            mnuSuppliers.Visible = True
            mnuEditPrices.Visible = True
            mnuItemS.Visible = True
            mnuEditStaff.Visible = True
            mnuEditAuthority.Visible = True
            mnuEditIncomeCategory.Visible = True
            mnuEditExpenceCategory.Visible = True
            mnuItemSuppliers.Visible = True

            mnuDailyCollection.Visible = True
            mnuMilkCollectionLabTestings.Visible = True
            mnuCumulativeReport.Visible = True
            mnuMilkCollectionCollectingCenterPayAdvice.Visible = True
            mnuPrintCollectingCenterPayAdvice.Visible = True
            mnuMilkPayAdvice.Visible = True

            mnuIssueItemIssue.Visible = True
            mnuItemPurchase.Visible = True
'            mnuIssuesAddDeductions.Visible = True
            mnuIssuesAddAdditionalPayments.Visible = True
            mnuAddAdditioanlDeductions.Visible = True
            mnuIssuesExpences.Visible = True
            mnuIssuesIncome.Visible = True

'            mnuBackOfficeGenerateIndividualPayments.Visible = True
            mnuBackOfficePrintIndividualPayments.Visible = True
'            mnuConfirmPayments.Visible = True
            mnuApprovals.Visible = True
            mnuOutstandings.Visible = True
            mnuBackOfficeReports.Visible = True
            mnuProfits.Visible = True

            mnuPersonalDetails.Visible = True

            mnuDetectErrors.Visible = True
            mnuBAckOffice.Visible = True
            mnuBackOfficeReports.Visible = True
            mnuReportExpences.Visible = True
            mnuBackOfficeMIlkCollection.Visible = True
            mnuBackOfficeReportsSuppliers.Visible = True
            mnuBackOfficeReportsCollectingCenters.Visible = True
            mnuAdditionalCommisionReport.Visible = True
            mnuAdditionalDeductionReport.Visible = True
            mnuVitaminAndCattleFeedDeductionReport.Visible = True
            mnuBackOfficeMilkCollectionDailyCollection.Visible = True
            mnuBackOfficeMilkCollectionPayAdvice.Visible = True
            mnuBackOfficeMIlkCollectionCumulativeReport.Visible = True
            mnuBackOfficeReportsSuppliersMilkCollection.Visible = True
            mnuBackOfficeReportsCollectingCentersMilkCollection.Visible = True
            mnuReportsTotalMilkCollection.Visible = True
            mnuProfitsIncome.Visible = True
            mnuProfitsExpence.Visible = True
            mnuProfitsProfit.Visible = True

            mnuApproveAdditionalDeductions.Visible = True
            mnuApproveAdditionalCommisions.Visible = True

'            mnuBackOfficeGenerateIndividualPayments.Visible = True
            mnuBackOfficePrintIndividualPayments.Visible = True
'            mnuConfirmPayments.Visible = True
            mnuDetectErrors.Visible = True
            mnuDeleteErrors.Visible = True

            mnuDeleteErrors.Visible = True
            mnuReportExpences.Visible = True
            mnuBackOfficeReports.Visible = True
        Case Else
'
    End Select
    
'-----------------------------------------------------
'Authorities
'-----------------------------------------------------

'File

    mnuFileBackup.Visible = BackupAllowed
    mnuFileRestore.Visible = RestoreAllowed
    
'Edit

    mnuEditCollectingCenters.Visible = CollectingCenterAllowed
    mnuSuppliers.Visible = FarmersAllowed
    mnuEditPrices.Visible = PricesAllowed
    mnuItemS.Visible = ItemAllowed
    mnuEditStaff.Visible = StaffsAllowed
    mnuEditAuthority.Visible = AuthorityAllowed
    mnuEditIncomeCategory.Visible = IncomeCategoryAllowed
    mnuEditExpenceCategory.Visible = ExpenceCategoryAllowed
    mnuItemSuppliers.Visible = ItemSuppiersAllowed
    
'Milk Collection

    mnuDailyCollection.Visible = DailyCollectionAllowed
    mnuMilkCollectionLabTestings.Visible = GoodRecieveNoteAllowed
    mnuCumulativeReport.Visible = CumulativeReportAllowed
    mnuMilkCollectionCollectingCenterPayAdvice.Visible = GenarateCollectingCenterPayAdviceAllowed
    mnuPrintCollectingCenterPayAdvice.Visible = PrintCollectingCenterPayAdviceAllowed
    mnuMilkPayAdvice.Visible = MilkPayAdviceAllowed
    
'Issue & Payments
    mnuIssueItemIssue.Visible = ItemIssueAllowed
    mnuItemPurchase.Visible = ItemPurchaseAllowed
'    mnuIssuesAddDeductions.Visible = AddDeductionsAllowed
    mnuIssuesAddAdditionalPayments.Visible = AdditionalCommisionsAllowed
    mnuAddAdditioanlDeductions.Visible = AdditionalDeductionsAllowed
    mnuIssuesExpences.Visible = ExpencesAllowed
    mnuIssuesIncome.Visible = IncomeAllowed
    
'Back Office

'    mnuBackOfficeGenerateIndividualPayments.Visible = GenerateIndividualPaymentsAllowed
    mnuBackOfficePrintIndividualPayments.Visible = PrintIndividualPaymentsAllowed
'    mnuConfirmPayments.Visible = ConfirmPaymentsAllowed
    mnuApprovals.Visible = ApprovalsAllowed
    mnuOutstandings.Visible = OutstandingAllowed
    mnuBackOfficeReports.Visible = ReportsAllowed
    mnuProfits.Visible = ProfitsAllowed
    mnuDetectErrors.Visible = DetectErrorsAllowed
    
'Options

    mnuOptionsDatabase.Visible = DatabaseAllowed
    mnuOptionsPrinting.Visible = PrintingAllowed
   
'-----------------------------------------------------
'End Authorities
'-----------------------------------------------------
    
        
''Dear Dr.Buddhika,
''Me and Mr.Sampath Discussed about authority levels of Milk Paying system.
''We divided authority levels in to 6 levels.(Mentioned Below)
''
''1.Viewer
''2. Analyzer.
''3. User.
''4. Power user.
''5. Supper User.
''6. Administrator.
''
''Authority types mentioned below.
'
''1.File (Menu)>>>>Backup  >>6
''                   >>>>Restore  >>6
'
''2.Edit (Menu)>>>>Collecting center  >>5,6
''                    >>>>Farmers  >>4,5,6
''                    >>>>Prices  >>5,6
''                    >>>>Item  >>4,5,6
''                    >>>>Staff  >>5,6      ;This form must update with out authority functions to level 5.
''                                                     in level 5 they can only update employees details but they cant assign authority.
''                                                     in level 6 they can assign authority password resetting, user name settings,user accounts deleting modifying, disabling etc..
''
''                    >>>>Position  >>6
''                    >>>>Authority  >>6
''                    >>>>Income Category  >>4,5,6
''                    >>>>Expense Category  >>4,5,6
'
''                    >>>>Item Supplier>>>>Edit  >>4,5,6
''                                               >>>>Add  >>3,4,5,6

''3.Milk Collection (Menu)>>>>Daily Collection  >>2,3,4,5,6
''                                   >>>>Lab Testing  >>3,4,5,6
''                                   >>>>Cumulative Report  >>1,2,3,4,5,6
''                                                                     >>>>print  >>3,4,5,6
''                                   >>>>Generate Collecting Center Pay Advice  >>4,5,6
''                                   >>>>Print Collecting Center Pay Advice  >>3,4,5,6
''                                   >>>>Supplier Milk Pay Advice  >>1,2,3,4,5,6

''4.Issue and payment>>>>3,4,5,6 >>Delete button  >>4,5,6

''5.Back Office>>>>Approve Collecting center Payments >>4,5,6   save Button  >>5,6
''                   >>>>Print Individual Payments  >>4,5,6
''                   >>>>Confirm Payments  >>5,6
''                   >>>>Approvals  >>5,6
''                   >>>>Outstanding  >>2,3,4,5,6
''                   >>>>Reports  >>2,3,4,5,6
''                   >>>>Profits  >>3,4,5,6
''                   >>>>Option  >>6
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer
    i = MsgBox("Are you sure you want to exit?", vbYesNo, "EXIT?")
    If i = vbNo Then Cancel = True
End Sub

Private Sub mnuAddAdditioanlDeductions_Click()
    frmAddAdditionalDeductions.Show
    frmAddAdditionalDeductions.ZOrder
End Sub

Private Sub mnuApproveAdditionalCommisions_Click()
    frmApproveAdditionalPayments.Show
    frmApproveAdditionalPayments.ZOrder 0
End Sub

Private Sub mnuApproveAdditionalDeductions_Click()
    frmApproveAdditionalDeductions.Show
    frmApproveAdditionalDeductions.ZOrder 0
End Sub

Private Sub mnuBackOfficeGenerateIndividualPayments_Click()
    frmGenerateIndividualPayments.Show
    frmGenerateIndividualPayments.ZOrder
End Sub

Private Sub mnuBAckOfficeMIlkCollectionCumulativeReport_Click()
    frmViewCumulativeDailyMilkCollection.Show
    frmViewCumulativeDailyMilkCollection.ZOrder 0
End Sub

Private Sub mnuBackOfficeMilkCollectionDailyCollection_Click()
    frmViewDailyMilkCollectionReport.Show
    frmViewDailyMilkCollectionReport.ZOrder 0
End Sub

Private Sub mnuBackOfficeMilkCollectionPayAdvice_Click()
    frmPrintCollectingCentrePaymentSummeryDisplay.Show
    frmPrintCollectingCentrePaymentSummeryDisplay.ZOrder 0
End Sub

Private Sub mnuBackOfficePrintBankSlips_Click()
    frmPrintingSlips.Show
    frmPrintingSlips.ZOrder 0
End Sub

Private Sub mnuBackOfficePrintPayments_Click()
    frmBanking.Show
    frmBanking.ZOrder 0
End Sub

Private Sub mnuBackOfficePrintIndividualPayments_Click()
    frmPrintIndividualPayments.Show
    frmPrintIndividualPayments.ZOrder 0
End Sub

Private Sub mnuBackOfficeReportsMilkPaymentsCommisionsandOtherExpences_Click()
    frmReportsExpence.Show
    frmReportsExpence.ZOrder 0
End Sub

Private Sub mnuBackOfficeReportsSuppliersMilkCollection_Click()
    frmReportSupplierviceMilkCollection.Show
    frmReportSupplierviceMilkCollection.ZOrder 0
End Sub

Private Sub mnuConfirmPayments_Click()
    frmPaymentConfirmation.Show
    frmPaymentConfirmation.ZOrder 0
End Sub

Private Sub mnuCumulativeReport_Click()
    frmCumulativeDailyMilkCollection.Show
    frmCumulativeDailyMilkCollection.ZOrder 0
End Sub

Private Sub mnuDailyCollection_Click()
    frmDailyMilkCollectionReport.Show
    frmDailyMilkCollectionReport.ZOrder 0
End Sub

Private Sub mnuDeleteErrors_Click()
    frmErrorDetection.Show
    frmErrorDetection.ZOrder 0
End Sub

Private Sub mnuEditAuthority_Click()
    frmAuthority.Show
    frmAuthority.ZOrder 0
End Sub

Private Sub mnuEditCollectingCenters_Click()
    frmCollectingCenters.Show
    frmCollectingCenters.ZOrder 0
End Sub

Private Sub mnuEditExpenceCategory_Click()
    frmExpenceCategory.Show
    frmExpenceCategory.ZOrder 0
End Sub

Private Sub mnuEditFarmerVolumeDeduction_Click()
    frmFarmerVolumeDeduction.Show
    frmFarmerVolumeDeduction.ZOrder 0
End Sub

Private Sub mnuEditIncomeCategory_Click()
    frmIncomeCategory.Show
    frmIncomeCategory.ZOrder 0
End Sub

Private Sub mnuEditPosition_Click()

End Sub


Private Sub mnuEditPaymentScheme_Click()
    frmPaymentScheme.Show
    frmPaymentScheme.ZOrder 0
End Sub

Private Sub mnuEditPriceCycle_Click()
    frmPriceCycle.Show
    frmPriceCycle.ZOrder 0
End Sub

Private Sub mnuEditPricesPrices_Click()
    frmPriceAdjustment.Show
    frmPriceAdjustment.ZOrder 0
'    frmNewPriceAdjustment.Show
'    frmNewPriceAdjustment.ZOrder 0
End Sub

Private Sub mnuEditStaff_Click()
    frmStaffDetails1.Show
    frmStaffDetails1.ZOrder 0
End Sub

Private Sub mnuEditVolumeDeductions_Click()
frmVolumeDeduction.Show
frmVolumeDeduction.ZOrder 0
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuIssueItemIssue_Click()
    frmIssueItems.Show
    frmIssueItems.ZOrder 0
End Sub

Private Sub mnuIssuesAddAdditionalPayments_Click()
    frmAddAdditionalPayments.Show
    frmAddAdditionalPayments.ZOrder 0
End Sub

Private Sub mnuIssuesAddDeductions_Click()
    frmAddDeductions.Show
    frmAddDeductions.ZOrder 0
End Sub


Private Sub mnuIssuesExpences_Click()
    frmExpence.Show
    frmExpence.ZOrder
End Sub

Private Sub mnuIssuesIncome_Click()
    frmIncome.Show
    frmIncome.ZOrder 0
End Sub

Private Sub mnuItemPurchase_Click()
    frmItemPurchase.Show
    frmItemPurchase.ZOrder 0
End Sub

Private Sub mnuItems_Click()
    frmItems.Show
    frmItems.ZOrder 0
End Sub

Private Sub mnuItemSuppliers_Click()
    frmItemSuppliers.Show
    frmItemSuppliers.ZOrder 0
End Sub

Private Sub mnuMilkCollectionCollectingCenterPayAdvice_Click()
    frmCollectingCentrePaymentSummery.Show
    frmCollectingCentrePaymentSummery.ZOrder 0
End Sub

Private Sub mnuMilkCollectionLabTestings_Click()
    frmGRN.Show
    frmGRN.ZOrder 0
End Sub

Private Sub mnuMilkPayAdvice_Click()
    frmMilkPayAdviceDisplay1.Show
    frmMilkPayAdviceDisplay1.ZOrder 0
End Sub

Private Sub mnuOptionInstitution_Click()

End Sub

Private Sub mnuOptionsExit_Click()
    Unload Me
End Sub

Private Sub mnuOptionsPrinting_Click()
    frmPrintingPreferances.Show
    frmPrintingPreferances.ZOrder 0
End Sub

Private Sub mnuOutstandings_Click()
    frmOutstandingBalance.Show
    frmOutstandingBalance.ZOrder 0
End Sub

Private Sub mnuPersonalDetails_Click()
    frmStaffPersonalDetails.Show
    frmStaffPersonalDetails.ZOrder 0
End Sub

Private Sub mnuPrintCollectingCenterPayAdvice_Click()
    frmPrintCollectingCentrePaymentSummery.Show
    frmPrintCollectingCentrePaymentSummery.ZOrder 0
End Sub

Private Sub mnuProfitsExpence_Click()
    frmViewExpence.Show
    frmViewExpence.ZOrder 0
End Sub

Private Sub mnuProfitsIncome_Click()
    frmViewIncome.Show
    frmViewIncome.ZOrder 0
End Sub

Private Sub mnuProfitsProfit_Click()
    frmProfit.Show
    frmProfit.ZOrder 0
End Sub

Private Sub mnuReportExpenceCategories_Click()
    frmReportsExpencesCategories.Show
    frmReportsExpencesCategories.ZOrder 0
End Sub

Private Sub mnuReportExpences_Click()
    frmReportsExpences.Show
    frmReportsExpences.ZOrder 0
End Sub

Private Sub mnuReportsCollectingCentersPaymentSchemeMilkSupply_Click()
    frmPaymentSchemeMilkCollection.Show
    frmPaymentSchemeMilkCollection.ZOrder 0
End Sub

Private Sub mnuReportsTotalMilkCollection_Click()
    frmAllCenterMilkPayments.Show
    frmAllCenterMilkPayments.ZOrder 0
End Sub

Private Sub mnuSuppliers_Click()
    frmSuppliers.Show
    frmSuppliers.ZOrder 0
End Sub

Private Sub mnuTem_Click()
frmTem.Show
'    Dim rsTem As New ADODB.Recordset
'    Dim temSQL As String
'    With rsTem
'        If .State = 1 Then .Close
'        temSQL = "Select * from tblSupplier where SupplierID = CommisionCollectorID"
'        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
'        If .RecordCount > 0 Then
'            While .EOF = False
'                !ThroughCollector = False
'                !CommisionCollectorID = 0
'                .Update
'                .MoveNext
'            Wend
'        End If
'        .Close
'    End With
        
    
End Sub

Private Sub mnuVitaminAndCattleFeedDeductionReport_Click()
    frmReportVitaminAndCattleFeedDeduction.Show
    frmReportVitaminAndCattleFeedDeduction.ZOrder 0
End Sub

