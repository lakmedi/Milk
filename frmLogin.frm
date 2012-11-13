VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTemUsername 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   4005
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Cancel"
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
   Begin btButtonEx.ButtonEx cmdOK 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Login"
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
   Begin btButtonEx.ButtonEx btnServer 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Server"
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
   Begin VB.Label Label1 
      Caption         =   "&User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "&Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FSys As New Scripting.FileSystemObject
    Dim SuppliedWord As String
    
    Dim rsStaff As New ADODB.Recordset
    Dim rsHospital As New ADODB.Recordset
    Dim temSQL As String
    Dim constr As String
    Dim TemUserPassward As String
    
Private Sub btnServer_Click()
    frmServer.Show 1
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
    
    Dim TemResponce As Byte
    Dim UserNameFound As Boolean
    UserNameFound = False
    
    TemUserPassward = txtPassword.Text
    If Trim(txtUserName.Text) = "" Then
        TemResponce = MsgBox("You have not entered a username", vbCritical, "Username")
        txtUserName.SetFocus
        Exit Sub
    End If
    If Trim(txtPassword.Text) = "" Then
        TemResponce = MsgBox("You have not entered a password", vbCritical, "Password")
        txtPassword.SetFocus
        Exit Sub
    End If
    With rsStaff
        If .State = 1 Then .Close
        temSQL = "Select * from tblstaff Where Deleted = 0 "
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount < 1 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            txtTemUsername.Text = DecreptedWord(!UserName)
            If UCase(txtUserName.Text) = UCase(txtTemUsername.Text) Then
                UserNameFound = True
                If txtPassword.Text = DecreptedWord(!Password) Then
                    UserName = UCase(txtUserName.Text)
                    UserID = !StaffID
                    UserFullName = !Staff
                    
                    Open App.Path & "\Login.TXT" For Append As #1
                    Print #1, "Date : " & Format(Date, "dd MMMM yyyy")
                    Print #1, "Time :" & Time
                    Print #1, "User Name : " & UserName
                    Print #1, "User : " & UserFullName
                    Print #1, "Login : " & "Success"
                    Print #1,
                    Close #1
                    
                    
                    
                    If !NeedPasswordReset = True Then
                        frmResetPassword.Show 1
                    End If
                    
                    If Weekday(Date) = vbMonday Then
                        If !ResetDate = Date Then
                    
                        Else
                            frmResetPassword.Show 1
                        End If
                    End If
                    
                    If Not IsNull(!AuthorityID) Then
                        UserAuthority = !AuthorityID
                    Else
                        UserAuthority = 0
                    End If
                    Select Case UserAuthority
                        Case 1: UserAuthorityLevel = Viewer
                        Case 2: UserAuthorityLevel = Analyzer
                        Case 3: UserAuthorityLevel = OrdinaryUser
                        Case 4: UserAuthorityLevel = PowerUser
                        Case 5: UserAuthorityLevel = SuperUser
                        Case 6: UserAuthorityLevel = Administrator
                        Case Else: UserAuthorityLevel = NotIdentified
                    End Select
                    
                    '----------------------------------------------------------
                    'Authorities
                    '----------------------------------------------------------
                    
                    'File
                    BackupAllowed = !BackupAllowed
                    RestoreAllowed = !RestoreAllowed
                    
                    'File
                    CollectingCenterAllowed = !CollectingCenterAllowed
                    FarmersAllowed = !FarmersAllowed
                    PricesAllowed = !PricesAllowed
                    ItemAllowed = !ItemAllowed
                    StaffsAllowed = !StaffsAllowed
                    AuthorityAllowed = !AuthorityAllowed
                    IncomeCategoryAllowed = !IncomeCategoryAllowed
                    ExpenceCategoryAllowed = !ExpenceCategoryAllowed
                    ItemSuppiersAllowed = !ItemSuppiersAllowed
                    
                    ItemSuppiersEditAllowed = !ItemSuppiersEditAllowed
                    
                    'Milk Collection
                    DailyCollectionAllowed = !DailyCollectionAllowed
                    GoodRecieveNoteAllowed = !GoodRecieveNoteAllowed
                    CumulativeReportAllowed = !CumulativeReportAllowed
                    GenarateCollectingCenterPayAdviceAllowed = !GenarateCollectingCenterPayAdviceAllowed
                    PrintCollectingCenterPayAdviceAllowed = !PrintCollectingCenterPayAdviceAllowed
                    MilkPayAdviceAllowed = !MilkPayAdviceAllowed
                    
                    CumulativeReportPrintAllowed = !CumulativeReportPrintAllowed
                   
                    'Issue & Payments
                    ItemIssueAllowed = !ItemIssueAllowed
                    ItemPurchaseAllowed = !ItemPurchaseAllowed
                    AddDeductionsAllowed = !AddDeductionsAllowed
                    AdditionalCommisionsAllowed = !AdditionalCommisionsAllowed
                    AdditionalDeductionsAllowed = !AdditionalDeductionsAllowed
                    ExpencesAllowed = !ExpencesAllowed
                    IncomeAllowed = !IncomeAllowed
                    
                    IssuePaymentDeleteAllowed = !IssuePaymentDeleteAllowed
                    
                    'Back Office
                    GenerateIndividualPaymentsAllowed = !GenerateIndividualPaymentsAllowed
                    PrintIndividualPaymentsAllowed = !PrintIndividualPaymentsAllowed
                    ConfirmPaymentsAllowed = !ConfirmPaymentsAllowed
                    ApprovalsAllowed = !ApprovalsAllowed
                    OutstandingAllowed = !OutstandingAllowed
                    ReportsAllowed = !ReportsAllowed
                    ProfitsAllowed = !ProfitsAllowed
                    DetectErrorsAllowed = !DetectErrorsAllowed
                   
                    ApprovePaymentsSaveAllowed = !ApprovePaymentSaveAllowed
                    
                    'Option
                    DatabaseAllowed = !DatabaseAllowed
                    PrintingAllowed = !PrintingAllowed
                    
                    '----------------------------------------------------------
                    'End Authorities
                    '----------------------------------------------------------
                    
                    
                    
                    Exit Do
                Else
                    TemResponce = MsgBox("The username and password you entered are not matching. Please try again", vbCritical, "Wrong Username and Password")
                    
                    
                    Open App.Path & "\Login.TXT" For Append As #1
                    Print #1, "Date : " & Format(Date, "dd MMMM yyyy")
                    Print #1, "Time :" & Time
                    Print #1, "User Name : " & txtUserName.Text
                    Print #1, "Password : " & txtPassword.Text
                    Print #1, "Login : " & "Failed"
                    Print #1,
                    Close #1
                    
                    txtUserName.SetFocus
                    On Error Resume Next
                    SendKeys "{home}+{end}"
                    Exit Sub
                End If
            Else
            End If
            .MoveNext
        Loop
        .Close
        If UserNameFound = False Then
            
            Open App.Path & "\Login.TXT" For Append As #1
            Print #1, "Date : " & Format(Date, "dd MMMM yyyy")
            Print #1, "Time :" & Time
            Print #1, "User Name : " & txtUserName.Text
            Print #1, "Password : " & txtPassword.Text
            Print #1, "Login : " & "Failed"
            Print #1,
            Close #1
            
            
            TemResponce = MsgBox("There is no such  a username, Please try again", vbCritical, "Username")
            txtUserName.SetFocus
            On Error Resume Next: SendKeys "{home}+{end}"
            Exit Sub
        End If
        End With
        
'        If DemoCopy = True Then
'            If Date > ExpiaryDate Or DemoCount > 50 Then
'                MsgBox "Demo Copy EXPIRED"
'                End
'            Else
'                MsgBox "Demo Copy Only. Will Expire Soon"
'                SaveSetting App.EXEName, "Options", "DemoCount", DemoCount + 1
'            End If
'        End If
        
        MDIMain.Show
        Unload Me
End Sub

Private Sub Form_Load()
    DemoCopy = False
    ExpiaryDate = #4/17/2010#
    Dim TemResponce As Byte
    Dim tempath As String
    Call LoadPreferances
    
    
    If DemoCopy = True Then
        If Date > ExpiaryDate Then
            MsgBox "Demo Copy Expired. Please contact Lakmedipro 071 5812399, 091 2241603"
            End
        Else
            MsgBox "Demo Copy. Will Expire soon."
        End If
    End If
    
    While connectToDatabase = False
        MsgBox "Please select the correct settings for the SQL Server 2005"
        frmServer.Show 1
    Wend

End Sub


Private Function connectToDatabase() As Boolean
    On Error GoTo eh:
    connectToDatabase = False
    constr = "Provider=MSDataShape.1;Persist Security Info=True;Data Source=" & Server & "\" & SQLServer & ";User ID=" & ServerUserName & ";Password=" & ServerPassword & ";Initial Catalog=" & ServerDatabase & ";Data Provider=SQLOLEDB.1"
    If cnnStores.State = 1 Then cnnStores.Close
    cnnStores.Open constr
    DataEnvironment1.deCnnStores.ConnectionString = constr
    connectToDatabase = True
    Exit Function
eh:
    connectToDatabase = False
    
End Function


Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtUserName.Text <> "" Then cmdOK_Click: Exit Sub
    If KeyAscii = 13 And txtUserName.Text = "" Then txtUserName.SetFocus: Exit Sub
End Sub


Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPassword.SetFocus
End Sub

Private Sub LoadPreferances()
    Database = GetSetting(App.EXEName, "Options", "Database", "")




    Server = DecreptedWord(GetSetting(App.EXEName, Me.Name, "Server", ""))
    SQLServer = DecreptedWord(GetSetting(App.EXEName, Me.Name, "SQLServer", ""))
    ServerDatabase = DecreptedWord(GetSetting(App.EXEName, Me.Name, "ServerDatabase", ""))
    ServerUserName = DecreptedWord(GetSetting(App.EXEName, Me.Name, "ServerUserName", ""))
    ServerPassword = DecreptedWord(GetSetting(App.EXEName, Me.Name, "ServerPassword", ""))
    



    LongDateFormat = "yyyy mmmm dd" ' GetSetting(App.EXEName, "Options", "LongDateFormat", "dddd, dd MMMM yyyy")
    ShortDateFormat = "yy mm dd"  ' GetSetting(App.EXEName, "Options", "ShortDateFormat", "dd MM yy")
    
    BillPrinterName = GetSetting(App.EXEName, "Options", "BillPrinterName", "")
    BillPaperName = GetSetting(App.EXEName, "Options", "BillPaperName", "")
    ReportPrinterName = GetSetting(App.EXEName, "Options", "ReportPrinterName", "")
    ReportPaperName = GetSetting(App.EXEName, "Options", "ReportPaperName", "")
    PrintingOnBlankPaper = GetSetting(App.EXEName, "Options", "PrintingOnBlankPaper", True)
    PrintingOnPrintedPaper = GetSetting(App.EXEName, "Options", "PrintingOnPrintedPaper", False)
    DemoCount = GetSetting(App.EXEName, "Options", "DemoCount", 1)
    
    
    InstitutionName = "Lucky Lanka Dairies (Pvt) Ltd."
    InstitutionAddressLine1 = "Bibulewela, Karagoda, Uyangoda."
    InstitutionAddressLine2 = "Tel. 041 2292652, 041 2293032, 041 2292031"
    InstitutionAddressLine3 = "Fax. 041 2292831"
    
End Sub


