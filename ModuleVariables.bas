Attribute VB_Name = "ModuleVariables"
Option Explicit


' database Variables
Public Database As String
Public cnnStores As New ADODB.Connection


' Server Variables

Public Server As String
Public SQLServer As String
Public ServerDatabase As String
Public ServerUserName As String
Public ServerPassword As String


' User Variables
Public UserName As String
Public UserID As Long
Public UserAuthority As Long
Public UserFullName As String

'---------------------------------------------------------------------
'Authorities
'---------------------------------------------------------------------
'File
Public BackupAllowed As Boolean
Public RestoreAllowed As Boolean

'Edit
Public CollectingCenterAllowed As Boolean
Public FarmersAllowed As Boolean
Public PricesAllowed As Boolean
Public ItemAllowed As Boolean
Public StaffsAllowed As Boolean
Public AuthorityAllowed As Boolean
Public IncomeCategoryAllowed As Boolean
Public ExpenceCategoryAllowed As Boolean
Public ItemSuppiersAllowed As Boolean

Public ItemSuppiersEditAllowed As Boolean

'Milk Collection
Public DailyCollectionAllowed As Boolean
Public GoodRecieveNoteAllowed As Boolean
Public CumulativeReportAllowed As Boolean
Public GenarateCollectingCenterPayAdviceAllowed As Boolean
Public PrintCollectingCenterPayAdviceAllowed As Boolean
Public MilkPayAdviceAllowed As Boolean

Public CumulativeReportPrintAllowed As Boolean

'Issue & Payments
Public ItemIssueAllowed As Boolean
Public ItemPurchaseAllowed As Boolean
Public AddDeductionsAllowed As Boolean
Public AdditionalCommisionsAllowed As Boolean
Public AdditionalDeductionsAllowed As Boolean
Public ExpencesAllowed As Boolean
Public IncomeAllowed As Boolean

Public IssuePaymentDeleteAllowed As Boolean

'Back Office
Public GenerateIndividualPaymentsAllowed As Boolean
Public PrintIndividualPaymentsAllowed As Boolean
Public ConfirmPaymentsAllowed As Boolean
Public ApprovalsAllowed As Boolean
Public OutstandingAllowed As Boolean
Public ReportsAllowed As Boolean
Public ProfitsAllowed As Boolean
Public DetectErrorsAllowed As Boolean

Public ApprovePaymentsSaveAllowed As Boolean

'Option
Public DatabaseAllowed As Boolean
Public PrintingAllowed As Boolean

'---------------------------------------------------------------------
'End Authorities
'---------------------------------------------------------------------

Public InstitutionName As String
Public InstitutionAddressLine1 As String
Public InstitutionAddressLine2 As String
Public InstitutionAddressLine3 As String


' Printing Preferances
Public PrintingOnBlankPaper As Boolean
Public PrintingOnPrintedPaper As Boolean
Public BillPrinterName As String
Public BillPaperName As String
Public ReportPrinterName As String
Public ReportPaperName As String
Public BillPaperHeight As Long
Public BillPaperWidth As Long
Public ReportPaperWidth As Long
Public ReportPaperHeight As Long

Public LongDateFormat As String
Public ShortDateFormat As String

' Program Variable
Public DemoCopy As Boolean
Public ExpiaryDate As Date
Public DemoCount As Long


' Data transfer variables
Public StaffCommentIDTx As Long
Public ExcelFilePath As Long

'Authorizing Variables
Public FileOK As Boolean
Public EditOK As Boolean
Public StaffOK         As Boolean
Public MilkCollectionOK           As Boolean
Public PayAdviceOK As Boolean
Public CashBankingOK                     As Boolean
Public DeductionsOK                As Boolean
Public IncomeExpencesOK           As Boolean
Public PreferancesOK      As Boolean

Public Enum Authority
    Viewer
    Analyzer
    OrdinaryUser
    PowerUser
    SuperUser
    Administrator
    NotIdentified
End Enum

Public UserAuthorityLevel As Authority
