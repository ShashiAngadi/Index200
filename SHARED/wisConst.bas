Attribute VB_Name = "wisConst"
Option Explicit
Public Const SE_PRIVILEGE_ENABLED As Long = 2
Public Const EWX_REBOOT = 2

Public Const wisGray = &H80000000
'Shashi 4/12/2000
'Public Const vbWhite = &H80000005  '&H80000005&
Public Const wisWhite = &H80000005    '&H80000005&
Public Const gDelim = ";"

' Status variable constants...
Public Const wis_CANCEL = 0
Public Const wis_FAILURE = 0
Public Const wis_OK = 1
Public Const wis_SUCCESS = 2
Public Const wis_COMPLETE = 3
Public Const wis_EVENT_SUCCESS = 4
Public Const wis_SHOW_FIRST = 5
Public Const wis_SHOW_PREVIOUS = 6
Public Const wis_SHOW_NEXT = 7
Public Const wis_SHOW_LAST = 8
Public Const wis_PRINT_CURRENT = 9
Public Const wis_PRINT_ALL = 10
Public Const wis_PRINT_ALL_PAUSE = 11
Public Const wis_PRINT_CURRENT_PAUSE = 12
Public Const wis_Print_Excel = 13

' Database updation mode.
Public Const wis_INSERT = 1
Public Const wis_UPDATE = 2

' Query mode constants...
Public Const wis_QUERY_BY_CUSTOMERID = 1
Public Const wis_QUERY_BY_CUSTOMERNAME = 2

' The key name for this application in Registry.
Public Const wis_INDEX2000_KEY = "Software\Waves Information Systems\Index 2000V3"

' Password for database.
Public Const wis_PWD = "wis"

' Title for Message boxes.
Public Const wis_MESSAGE_TITLE = "Index-2000 Info..."


' Report Constants
Enum wisReports
    wisTradingAccount = 1
    wisDebitCreditStatement = 2
    wisProfitLossStatement = 3
    wisBalanceSheet = 4
    wisDailyRegister = 5
    wisBankBalance = 6
    wisDailyDebitCredit = 7
    wisTrialBalance
    wisRepCreditTrans
    wisRepDebitTrans
    wisDetailCashBook
    wisBalancing
    wisDailyCashBook
End Enum


Enum wisModules
    wis_None = 0
    wis_CustReg = 1
    wis_SBAcc = 2000 '2
    wis_CAAcc = 3000 '3
    wis_RDAcc = 4000 '4
    wis_PDAcc = 5000 '5
    wis_BKCC = 6
    wis_BKCCLoan = 7
    wis_Members = 8000 ' 8
    wis_Users = 9
    wis_MatAcc = 10
    wis_SuspAcc = 11
    wis_Deposits = 100
    wis_DepositLoans = 200
    wis_Loans = 300
    
    'The MAximum Module Id
    'If we create new deposit
    'we will Use Max Module Id
    wis_MaxModId = 300
    wis_BankAccounts = 10000
End Enum


Public Enum wis_DepositType
    wisDeposit_SB = 20  'Changed on Feb 2016 '1
    wisDeposit_CA = 30  'Changed on Feb 2016 '2
    wisDeposit_RD = 40 'Changed on Feb 2016 '4
    wisDeposit_PD = 50 'Changed on Feb 2016 '8
    wisDeposit_FD = 10 '100 'Changed on Feb 2016 '10
End Enum


'enumeration to identy the amoun ttyep
'tha amountmay be Principle ,interest,
'penalinterest,payble

Public Enum wis_AmountType
    wisPrincipal = 1
    wisRegularInt = 2
    wisPenalInt = 3
    wisPayable = 4
    wisMisc = 5
End Enum

'Enumerated error values...
Enum errors
    wis_DATABASE_NOT_OPEN
    wis_INVALID_DATABASE
    wis_DUPLICATE_ACCNO
    wis_INVALID_MODULEID
    wis_INVALID_ACCNO
    wis_ACCNO_NOT_SET
    wis_MODULEID_NOT_SET
    wis_INIT_FAIL
    wis_FILENOTFOUND
End Enum

'The Report order definned here
Public Enum wis_ReportOrder
    wisByName = 1
    wisByAccountNo = 2
End Enum

'The Gender order definned here
Public Enum wis_Gender
    wisNoGender = 0
    wisMale = 1
    wisFemale = 2
End Enum


'New TransCtion Types defined below
Public Enum wisTransactionTypes
    ' wInterest & wCharges Are w.r.t Bank/Society
'     wInterest = 2
'     wCharges = -2
    
    ' wDeposit & wWithdraw are w.r.t to any Accounts
    wDeposit = 1        'Customers money into account
    'Redefnation Money Comes Into Accountirrespsctive of the account type
    wWithdraw = 2      'Customers money out of account
     
    wContraDeposit = 3        'Customers money into account
    'Redefnation Money Comes Into Accountirrespsctive of the account type
    wContraWithdraw = 4      'Customers money out of account
End Enum


Enum wisLoanCategories
    wisAgriculural = 1
    wisNonAgriculural = 2
End Enum

Enum wisLoanTerm
    wisShortTerm = 1
    wisMidTerm = 2
    wisLongTerm = 3
End Enum

Enum wis_ChequeStatus
   wisPending = 1
   wisCleared = 2
   wisBounced = 4
   wisDiscount = 8
End Enum

Public Enum wis_ChequeTrans
    wischqIssue = 1
    wischqPay = 2
    wischqStop = 3
    wischqLoss = 4
End Enum
 
Public Const wis_ROWS_PER_PAGE2 = 70
Public Const wis_ROWS_PER_PAGE1 = 11
Public Const wis_ROWS_PER_PAGE_A4 = 60
Public Const wis_ROWS_PER_PAGE = 38  ' No.of max rows to print in passbook page

 
Function ErrMsg(errNum As Integer, Optional ByVal errParam As String) As String
' Returns the user-defined error description.
Select Case errNum
    Case wis_ACCNO_NOT_SET
        ErrMsg = "The member variable AccNo is not set."
    Case wis_MODULEID_NOT_SET
        ErrMsg = "The member variable 'ModuleID' is not set."
    Case wis_INVALID_MODULEID
        'ErrMsg = "No module installed having id number " & errParam & "."
        ErrMsg = "Module license information not found.  Please contact the vendor for licensed version of software modules."
    Case wis_INIT_FAIL
        ErrMsg = "Error in initializing the module."
    Case wis_INVALID_ACCNO
        ErrMsg = "Account number should be a valid number."
    Case wis_FILENOTFOUND
        ErrMsg = "Could not find the file  - " & errParam & "."
    Case wis_DATABASE_NOT_OPEN
        ErrMsg = "A database must be open, for a query to be executed."
    Case wis_DUPLICATE_ACCNO
        ErrMsg = "The account number " & errParam & " is in use."
    Case wis_INVALID_DATABASE
        ErrMsg = "The database is not proper or is corrupt."
End Select
End Function


