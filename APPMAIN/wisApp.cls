VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "wisApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public gAppPath As String
'Public gDbTrans As clsTransact
'Public gDbTrans As Object
Private ClassObj(4) As Object

' Declare object variables for different modules.
'Private CustObj As clsCustReg
'Private WithEvents SbObj As clsSBAcc
'Private WithEvents CAObj As clsCAAcc
'Private WithEvents LoanObj As clsLoan
'Private FDObj As clsFDAcc
'Private PDObj As clsPDAcc
'Private RDObj As clsRDAcc
'Private MMObj As clsMMAcc
'Private RTObj As clsReports
'Private BankObj As clsAccTrans
'Private MatObj As clsMatAcc
'Private UserObj As clsUsers
'Private UtilsObj As clsUtils
'Private ClearingObj As clsClearing
'Private ClassObj As Object
' User defined events...
Public Event UpdateStatus(strMsg As String)
'This function returns the availiable class object to load the
'if this class already loaded it returns the exiting class
Private Sub ReturnClassID(ByVal strClassName As String, _
            ByRef ClassExists As Boolean, ByRef ClassNo As Integer)

Dim count As Byte
Dim MaxCount As Byte
Dim NotExistNo As Integer
Dim obj As Object
Dim Repeat As Boolean
strClassName = UCase(strClassName)

FirstLine:

NotExistNo = 10

MaxCount = 4

Repeat = True
While Repeat
    Repeat = False

    For count = 0 To MaxCount
        If ClassObj(count) Is Nothing Then
            NotExistNo = count
        Else
            If count > NotExistNo Then
                Repeat = True
                Set ClassObj(NotExistNo) = ClassObj(count)
                Set ClassObj(count) = Nothing
                Exit For
            End If
            If ClassObj(count).ObjectHandle = 0 Then Set ClassObj(count) = Nothing: Repeat = True
        End If
    Next
Wend


For count = 0 To MaxCount
    If ClassObj(count) Is Nothing Then Exit For
    If UCase(ClassObj(count).ObjectName) = strClassName Then
        ClassExists = True
        ClassNo = count
        Exit For
    End If
Next

If ClassExists Then Exit Sub

If count > MaxCount Then
    Debug.Assert count = 10
    
    'GoTo FirstLine
End If
    
ClassNo = count

End Sub

Public Sub MakeWindowsActive()
Dim count As Integer
For count = 0 To 4
    If Not ClassObj(count) Is Nothing Then
        If ClassObj(count).ObjectHandle Then _
            SetActiveWindow (ClassObj(count).ObjectHandle)
    End If
Next
End Sub

Public Sub ShowCADialog()
'Set ClassObj = New clsCAAcc

'ClassObj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("CAClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New ClsCAAcc
ClassObj(ClsNo).Show

End Sub
Public Sub ShowSHGDialog()
'Set Classobj = New clsShg
'Classobj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("SHGClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsShg
ClassObj(ClsNo).Show

End Sub


Public Sub ShowBankDialog()
'Set ClassObj = New clsAccTrans
'ClassObj.ShowAccTrans
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("AccTrans", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsAccTrans
ClassObj(ClsNo).ShowAccTrans

End Sub


Public Sub ShowCompanyDetails()
Debug.Print IIf(gDbTrans Is Nothing, "NOthing", "object")
Load frmCompanyDetails
gWindowHandle = frmCompanyDetails.hwnd
frmCompanyDetails.Show 1
gWindowHandle = 0

End Sub

Public Sub ShowCompanyCreation()

Load frmCreateCompany
gWindowHandle = frmCreateCompany.hwnd
frmCreateCompany.Show 1
gWindowHandle = 0

End Sub

Public Sub ShowSuspence()
'Set ClassObj = New clsSuspAcc
'ClassObj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("SuspClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsSuspAcc
ClassObj(ClsNo).Show


End Sub


Public Sub ShowPassing()
frmPassing.Show 1
'Set ClassObj = New clsSuspAcc
'ClassObj.Show
'Dim ClsNo As Integer
'Dim blExist As Boolean
'Call ReturnClassID("Passing", blExist, ClsNo)
'If ClsNo > 4 Then Exit Sub
'If Not blExist Then Set ClassObj(ClsNo) = New clsSuspAcc
'ClassObj(ClsNo).Show


End Sub


Public Sub ShowContra()
'Set ClassObj = New clsContra
'ClassObj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("ContraClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsContra
ClassObj(ClsNo).Show

End Sub

Public Sub ShowCustomerSearch()
 gWindowHandle = frmCustSearch.hwnd
 frmCustSearch.Show
 gWindowHandle = 0
End Sub

Public Sub ShowDepositLoan()
'    Set ClassObj = New clsDepLoan
'    ClassObj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("DepLoanClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsDepLoan
ClassObj(ClsNo).Show

End Sub

Public Sub ShowKCCDialog()
'Set Classobj = New clsBkcc
'Classobj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("KCCClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsBkcc
ClassObj(ClsNo).Show

End Sub


Public Sub ShowMaterialReport()
    frmRepSelect.Show 1
End Sub

Public Sub ShowOwnCompany()
    
    frmCompanyDetails.Show vbModal
    
End Sub

Public Sub ShowCustInfo()
'Set ClassObj = New clsCustReg
'ClassObj.ShowCustInfo

Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("CustClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsCustReg
ClassObj(ClsNo).ShowCustInfo

End Sub

Public Sub ShowMaterialInvoiceDetails()
'Set ClassObj = New clsInvoiceDet
'ClassObj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("InvoiceClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsInvoiceDet
ClassObj(ClsNo).Show

End Sub

Public Sub ShowMaterialPurchase()
Dim tmpClassObj As New clsMaterial
Dim headID As Long
'Set ClassObj = New clsMaterial

headID = tmpClassObj.GetHeadID
'If HeadID = 0 Then HeadID = parPurchase + 1
If headID = 0 Then Exit Sub
Set tmpClassObj = Nothing


'Set ClassObj = New clsPurchase
'ClassObj.VendorID = HeadID
'ClassObj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("PurchaseClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsPurchase
ClassObj(ClsNo).VendorID = headID
ClassObj(ClsNo).Show


End Sub

Public Sub ShowMaterialSales()

'Dim HeadID As Long
'Dim tmpClassObj As New clsMaterial
'
'HeadID = tmpClassObj.GetHeadIDFromHeadsList(Enum_Customers)
'
'If HeadID = 0 Then Exit Sub
'Set tmpClassObj = Nothing

'Set ClassObj = New clsSales
'ClassObj.VendorID = HeadID
'ClassObj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("SalesClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsSales
'ClassObj(ClsNo).VendorID = HeadID
ClassObj(ClsNo).Show

End Sub


Public Sub ShowMaterialTransfer()
'Set ClassObj = New clsTransferNew
'ClassObj.Show

Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("TransferClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsTransferNew
ClassObj(ClsNo).Show

'Set TfrClass = Nothing

End Sub


Public Sub ShowUtils()
'If UtilsObj Is Nothing Then Set UtilsObj = New clsUtils
'UtilsObj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("UtilClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsUtils
ClassObj(ClsNo).Show
End Sub

Public Sub ShowLoanCreateDialog()
'Set Classobj = New clsLoan
'Classobj.ShowCreateLoanAccount
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("LOanClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsLoan
ClassObj(ClsNo).ShowCreateLoanAccount

End Sub

Public Sub ShowLoanSchemeCreateDialog()
'Set ClassObj = New clsLoan
'ClassObj.ShowLoanSchemes

Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("LoanClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsLoan
ClassObj(ClsNo).ShowLoanSchemes

End Sub

Public Sub ShowLoanReportDialog()
'Set Classobj = New clsLoan
'Classobj.ShowLoanReport
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("LoanClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsLoan
ClassObj(ClsNo).ShowLoanReport

End Sub

Public Sub ShowLoanTransDialog()
'Set ClassObj = New clsLoan
'ClassObj.ShowLoanAccountDetail

Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("LoanClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsLoan
ClassObj(ClsNo).ShowLoanAccountDetail

End Sub

Public Sub ShowFDDialog()
 
'Set ClassObj = New clsFDAcc
'ClassObj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("FDClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsFDAcc
ClassObj(ClsNo).Show

End Sub

Public Sub ShowMemberDialog()
'If MMObj Is Nothing Then Set MMObj = New clsMMAcc
'MMObj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("MemberClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsMMAcc
ClassObj(ClsNo).Show

End Sub

Public Sub ShowMaterialDialog()

'dim tmpClassObj = New clsMaterial

Dim headID As Long
'HeadID = tmpClassobj.GetHeadIDFromHeadsList(Enum_Stockist)

'If HeadId = 0 Then Exit Sub
With frmProductPropertyNew
 '   .lblCompanyName = Classobj.GetCompanyName(HeadID)
    .Show vbModal
End With
'Set tmpClassObj = Nothing

End Sub

Public Sub ShowClearingDialog()

'Set ClassObj = New clsClearing
'ClassObj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("ClearingClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsClearing
ClassObj(ClsNo).Show


End Sub

Public Sub ShowPDDialog()
'If PDObj Is Nothing Then Set PDObj = New clsPDAcc
'PDObj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("PDClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsPDAcc
ClassObj(ClsNo).Show

End Sub

Public Sub ShowRDDialog()
'If RDObj Is Nothing Then Set RDObj = New clsRDAcc
'RDObj.Show
Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("RDClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsRDAcc
ClassObj(ClsNo).Show

End Sub

Public Sub ShowReportDialog(Optional ReportType As wisReports)
    
Dim SetupClass As clsSetup

'Dim ClsNo As Integer
'Dim blExist As Boolean
'Call ReturnClassID("ReportClass", blExist, ClsNo)
'If ClsNo > 4 Then Exit Sub
'If Not blExist Then Set ClassObj(ClsNo) = New clsReports
'ClassObj(ClsNo).Show

Dim newReport As Boolean

newReport = CBool(GetConfigValue("NewReport"))

Dim RTobj As clsReports
Set RTobj = New clsReports

With RTobj

    If ReportType Then
        If ReportType = wisRepCreditTrans Or ReportType = wisRepDebitTrans Then
            'Do Nothing
        Else
            Dim ShowFromDate As Boolean
            'ShowFromDate = (ReportType = wisBankBalance) Or (ReportType = wisDetailCashBook) Or (ReportType = wisDailyCashBook) Or (ReportType = wisBalancing)
            ShowFromDate = (ReportType = wisDetailCashBook) Or (ReportType = wisDailyCashBook) Or (ReportType = wisBalancing)
            ShowFromDate = Not ShowFromDate
            .ShowReportDate (ShowFromDate)
            If .FromRpDate = "" Then Exit Sub
            
        End If
    End If
    Screen.MousePointer = vbHourglass
    'If ReportType = wisBalanceSheet Then .ShowBalanceSheet
    If ReportType = wisBalanceSheet Then
        Set SetupClass = New clsSetup
        newReport = CBool(SetupClass.ReadSetupValue("General", "ShowNewBalSheet", "True"))
        Set SetupClass = Nothing
        If newReport Then .ShowConditionBalanceSheet_New Else .ShowConditionBalanceSheet
    End If
    If ReportType = wisDailyDebitCredit Then .ShowRPReport
    'If ReportType = wisDebitCreditStatement Then .ShowRPReport
    'If ReportType = wisProfitLossStatement Then .ShowPandLAccount
    'If ReportType = wisDebitCreditStatement Or _
        ReportType = wisProfitLossStatement Then .ShowReport1 (ReportType)
    If ReportType = wisDebitCreditStatement Or _
                    ReportType = wisProfitLossStatement Then
            Set SetupClass = New clsSetup
            
            If UCase(SetupClass.ReadSetupValue("General", "RPorPLBothHeads", "True")) = "TRUE" Then
                .ShowReport2 (ReportType)
            Else
                .ShowReport1 (ReportType)
            End If
            Set SetupClass = Nothing
    End If
    If ReportType = wisTradingAccount Then .ShowTradingAccount
    If ReportType = wisTrialBalance Then .ShowTrialBalance
    If ReportType = wisDetailCashBook Then .ShowDailyCashBook (newReport)
    If ReportType = wisDailyCashBook Then .ShowDailyCashBook (newReport)
    If ReportType = wisBalancing Then .ShowBalancing
    If ReportType = wisBankBalance Then .ShowGeneralLedger (newReport)
End With
    
    'Screen.MousePointer = vbDefault
    If ReportType = wisRepCreditTrans Or ReportType = wisRepDebitTrans Then
        
        Screen.MousePointer = vbDefault
        RTobj.ShowReportDate (False)
        If RTobj.ToRpDate = "" Then Exit Sub
        Screen.MousePointer = vbHourglass
        With frmReportview
            .ReportType = ReportType
            .FromIndianDate = RTobj.FromRpDate
            .ToIndianDate = RTobj.ToRpDate
        End With
        Unload frmRptDt
        Load frmReportview
        Screen.MousePointer = vbDefault
        frmReportview.Show 1
        
    ElseIf ReportType Then
        Screen.MousePointer = vbDefault
        RTobj.ShowCurrentReport
    Else
        Screen.MousePointer = vbDefault
        RTobj.ShowReportForm
    End If
    
    Set RTobj = Nothing

End Sub

Public Sub ShowSBDialog()

'If SbObj Is Nothing Then Set SbObj = New clsSBAcc
'SbObj.Show
'Set Classobj(0) = New clsSBAcc
'Debug.Print Classobj(0).ObjectName
'Debug.Print Classobj(0).Name
'Classobj(0).Show

Dim ClsNo As Integer
Dim blExist As Boolean
Call ReturnClassID("SBClass", blExist, ClsNo)
If ClsNo > 4 Then Exit Sub
If Not blExist Then Set ClassObj(ClsNo) = New clsSBAcc
ClassObj(ClsNo).Show

End Sub


Private Sub Class_Terminate()

' Release SB Object.
'If Not SbObj Is Nothing Then Set SbObj = Nothing

' Release CA Object.
'If Not CAObj Is Nothing Then Set CAObj = Nothing

' Release Loan Object.
'If Not LoanObj Is Nothing Then Set LoanObj = Nothing

' Release pigmy Deposit Object.
'If Not PDObj Is Nothing Then Set PDObj = Nothing

' Release RD Object.
'If Not RDObj Is Nothing Then Set RDObj = Nothing

' Release FD Object.
'If Not FDObj Is Nothing Then Set FDObj = Nothing

' Release member Object.
'If Not MMObj Is Nothing Then Set MMObj = Nothing

Dim count As Integer
For count = 0 To 4
    Set ClassObj(count) = Nothing
Next


End Sub

Public Sub ShowDataEntry()
    frmDataMain.Show vbModal
End Sub
