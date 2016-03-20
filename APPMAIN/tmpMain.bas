Attribute VB_Name = "bastmpMain"
Option Explicit

Public gStrDate As String
Public gAppPath As String
Public Const gAppName = "INDEX 2000"
Public gDbTrans As clsTransact
'Public gDbTrans As Object
Public gCurrUser As clsUsers
Public gUserID As Long
Public gCompanyName As String
Public gCancel As Byte
Public gWindowHandle As Long
Public gHandles(4) As Long
Public gCashier As Boolean
Public gPassing As Boolean

''Start 10 Sep 2013
Public Const INDEX2000_IMAGE_PATH = "C:\IndexPhotos"
Public Const INDEX2000_PHOTO = "photo"
Public Const INDEX2000_SIGN = "sign"
Public Const INDEX2000_IMG_FILE_TYPES = "jpg|jpeg|png|bmp|tif"
Public gImagePath As String
''start 10 Sep 2013
 
'By shashi on 9/9/2001
'Structure for holding the fields info for table.
'This is used by CreateDB function.
Public Enum wis_DBOperation
    Insert = 1
    Update = 2
    DeleteRec = 3
End Enum

Private Sub BeginDayTrans()
Dim rst As Recordset

BeginLine:

'check whether any body has logged in and the UserID
'they has start the Day
gDbTrans.SqlStmt = "SELECT * FROM Install " & _
                    " Where KeyData = 'BeginDate'"
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = "Insert Into Install " & _
                "(KeyData,ValueData ) Values " & _
                "('BeginDate','');"
    If Not gDbTrans.SQLExecute Then GoTo EndLine
    gDbTrans.SqlStmt = "Insert Into Install " & _
                "(KeyData,ValueData ) Values " & _
                "('EndDate','');"
    If Not gDbTrans.SQLExecute Then GoTo EndLine
    gDbTrans.CommitTrans
    GoTo BeginLine
End If
If Len(FormatField(rst("ValueData"))) > 0 Then _
        DayBeginDate = FormatField(rst("ValueData")): GoTo ExitLine

'Now Check the user permissions
Dim Perm As wis_Permissions

Perm = gCurrUser.UserPermissions

'if user not a bank administrator 'then do not allow him to begin the day
If Not ((Perm And perBankAdmin) > 0 Or (Perm And perOnlyWaves) > 0) Then
    MsgBox "You have not permitted to begin the day operation"
    GoTo EndLine
End If

Load frmDayBegin
frmDayBegin.ShowDayBegin

ExitLine:

    Exit Sub


EndLine:
    gDbTrans.CloseDB
    MsgBox "Unable to Begin the day transaction", vbInformation, wis_MESSAGE_TITLE
    End


End Sub

Public Sub CenterMe(frmForm As Form)

With frmForm
    .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
End With
End Sub


Public Sub ExitApplication(Confirm As Boolean, Cancel As Integer)

If Confirm Then
    ' Ask for user confirmation.
    Dim nRet As Integer
    'nRet = MsgBox("Do you want to exit this application?", _
            vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
    nRet = MsgBox(GetResourceString(750), _
            vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
    If nRet = vbNo Then Cancel = True: Exit Sub
End If

If gWindowHandle Then CloseWindow (gWindowHandle)

On Error Resume Next
If gLangOffSet Then
    Call NudiResetAllFlags: Call NudiStopKeyboardEngine
    
    If gLangShree Then
        UNLOADTRANSLITERATION
        'To set keyboard back to English language
        Call SAMHITA.SHREE_SETSCRIPT(Pass1, Pass2, ENG)
        'Before closing application,Deactivate samhita by calling Close_Shree API call
        SAMHITA.CLOSE_SHREE
    End If
End If
Debug.Print IIf(NudiGetLastError = 0, 2, 1)

Unload wisMain
Set wisMain = Nothing
Set gCurrUser = Nothing
'Set wisAppObj = Nothing

If Not gDbTrans Is Nothing Then gDbTrans.CloseDB
Set gDbTrans = Nothing

'1 For Exit
'2 For Reboot
'Call ExitWindowsEx(2, 0)
   
   End


End Sub

Public Function GetAccRecordSet(AccHeadID As Long, _
                    Optional AccNum As String = "", _
                    Optional SearchName As String = "") As Recordset

On Error Resume Next

Dim rstReturn As Recordset
Dim SqlStr As String
Dim pos As Long
Dim sqlClause As String

Set rstReturn = Nothing
On Error GoTo Exit_Line

Dim ModuleId As wisModules

ModuleId = GetModuleIDFromHeadID(AccHeadID)

Set rstReturn = Nothing

'Members
If ModuleId >= wis_Members And ModuleId < wis_Members + 100 Then
        SqlStr = "Select AccNum, A.AccID as ID,A.CustomerID,MemberType From MemMaster A"

    If ModuleId >= wis_Members And ModuleId < wis_Members + 100 Then sqlClause = " ANd A.MemberType= " & ModuleId - wis_Members
    
'BKcc Account
ElseIf ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then
        SqlStr = "Select AccNum, A.LoanID as ID From BKCCMaster A"
'Current Account
ElseIf ModuleId = wis_CAAcc Then
    SqlStr = "Select AccNum, AccId as Id,A.CustomerID From CAMaster A"

'DepositLoans
ElseIf ModuleId >= wis_DepositLoans And ModuleId < wis_DepositLoans + 100 Then
    SqlStr = "Select AccNum, LoanId as ID,A.CustomerID From DepositLoanMaster A"
    If ModuleId > wis_DepositLoans Then _
        sqlClause = " ANd A.DepositType = " & ModuleId - wis_DepositLoans

'Deposit Accounts like Fd
ElseIf ModuleId >= wis_Deposits And ModuleId < wis_Deposits + 100 Then
    SqlStr = "Select AccNum, AccId as ID,A.CustomerID From FDMaster A"
    If ModuleId > wis_Deposits Then _
        sqlClause = " ANd A.DepositType = " & ModuleId - wis_Deposits
'Loan Accounts
ElseIf ModuleId >= wis_Loans And ModuleId < wis_Loans + 100 Then
    SqlStr = "Select AccNum, LoanId as ID,A.CustomerID From LoanMaster A"
    If ModuleId > wis_Loans Then _
        sqlClause = " AND A.SchemeID = " & ModuleId - wis_Loans

'Pigmy Accounts
ElseIf ModuleId = wis_PDAcc Then
    SqlStr = "Select AccNum, AccId as ID,A.CustomerID From PDMaster A"

'Recurring Accounts
ElseIf ModuleId = wis_RDAcc Then
    SqlStr = "Select AccNum, AccId as ID,A.CustomerID From RDMaster A"

'Savings account
ElseIf ModuleId >= wis_SBAcc And ModuleId < wis_SBAcc + 100 Then
    SqlStr = "Select AccNum, AccId as ID,A.CustomerID From SBMaster A"

'Suspencaccount
ElseIf ModuleId = wis_SuspAcc Then
    SqlStr = "Select TransId as AccNum, AccId as ID,A.CustomerID " & _
            " From SuspAccount "
    
Else
    MsgBox "Plese select the account type", vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If
    
'If Moduleid <> wis_SuspAcc Then

    SqlStr = Trim(SqlStr)
    pos = InStr(1, SqlStr, "FROM", vbTextCompare)
    
    If pos Then
        SqlStr = Left(SqlStr, pos - 1) & _
          ", Title +' '+FirstName+' '+MiddleName+' '+LastName as CustName" & _
          " " & Mid(SqlStr, pos)
    End If
    
    SqlStr = SqlStr & ", NameTab B WHERE B.CustomerID = A.CustomerID"
        
    If ModuleId = wis_SuspAcc Then
        SqlStr = SqlStr & " ANd Cleared = 0 And CustomerID > 0" & _
            " ANd (TransType = " & wDeposit & " OR TransType = " & wContraDeposit & ")"
        If AccNum <> "" Then _
            SqlStr = SqlStr & " AND A.TransID = " & Val(AccNum)
    
    Else
        If AccNum <> "" Then _
            SqlStr = SqlStr & " AND A.AccNum = " & AddQuotes(AccNum, True)
        
        If Trim(SearchName) <> "" Then _
            SqlStr = SqlStr & " AND (FirstName like '" & SearchName & "%' " & _
                " Or MiddleName like '" & SearchName & "%' " & _
                " Or LastName like '" & SearchName & "%')"
    End If
    
    gDbTrans.SqlStmt = SqlStr & " " & sqlClause & " ORDER By FirstName"

    If gDbTrans.Fetch(rstReturn, adOpenStatic) < 1 Then
        MsgBox "There are no customers in the " & _
            GetHeadName(AccHeadID), vbInformation, wis_MESSAGE_TITLE
        Exit Function
    End If

'End If

Exit_Line:

Set GetAccRecordSet = rstReturn

End Function

Public Function GetAccountNumber(ByVal AccHeadID As Long, _
    Optional ByVal AccountId As Long = 0, Optional ByVal CustomerID As Long = 0, _
    Optional returnCustomerName As String) As String

On Error Resume Next

returnCustomerName = ""
If AccountId = 0 And CustomerID = 0 Then Exit Function

Dim SqlStr As String

On Error GoTo Exit_Line

Dim ModuleId As wisModules

ModuleId = GetModuleIDFromHeadID(AccHeadID)

If ModuleId = wis_None Then Exit Function

'Members
If ModuleId >= wis_Members And ModuleId < wis_Members + 100 Then
    SqlStr = "Select AccNum, A.AccID as ID," & _
        " Title +' '+ FIrstName +' '+ MiddleName +' ' + LastName as CustName" & _
        " From MemMaster A, NameTab B" & _
        " Where A.CustomerID = B.CustomerID "
        
'Savings Account
ElseIf ModuleId >= wis_SBAcc And ModuleId < wis_SBAcc + 100 Then
    SqlStr = "Select AccNum, A.AccID as ID," & _
        " Title +' '+ FIrstName +' '+ MiddleName +' ' + LastName as CustName" & _
        " From SBMaster A, Nametab B" & _
        " Where A.CustomerID = B.CustomerID "
        
'Current Account
ElseIf ModuleId = wis_CAAcc Then
    SqlStr = "Select AccNum, A.AccID as ID," & _
        " Title +' '+ FIrstName +' '+ MiddleName +' ' + LastName as CustName" & _
        " From CAMaster A,Nametab B" & _
        " Where A.CustomerID = B.CustomerID "
        
'Pigmy Accounts
ElseIf ModuleId = wis_PDAcc Then
    SqlStr = "Select AccNum, A.AccID as ID, " & _
        " Title +' '+ FIrstName +' '+ MiddleName +' ' + LastName as CustName" & _
        " From PDMaster A, Nametab B" & _
        " Where A.CustomerID = B.CustomerID "
        
'Recurring Accounts
ElseIf ModuleId = wis_RDAcc Then
    SqlStr = "Select AccNum, A.AccID as ID," & _
        " Title +' '+ FIrstName +' '+ MiddleName +' ' + LastName as CustName" & _
        " From RDMaster A,Nametab B" & _
        " Where A.CustomerID = B.CustomerID "
        

'Deposit Accounts like Fd
ElseIf ModuleId >= wis_Deposits And ModuleId < wis_Deposits + 100 Then
    SqlStr = "Select AccNum, A.AccID as ID, " & _
        " Title +' '+ FIrstName +' '+ MiddleName +' ' + LastName as CustName" & _
        " From FDMaster A,Nametab B" & _
        " Where A.CustomerID = B.CustomerID "
        
'BKCC Account
ElseIf ModuleId = wis_BKCC Or ModuleId = wis_BKCCLoan Then
    SqlStr = "Select AccNum, A.LoanID as ID," & _
        " Title +' '+ FIrstName +' '+ MiddleName +' ' + LastName as CustName" & _
        " From BKCCMaster A,Nametab B" & _
        " Where A.CustomerID = B.CustomerID "
        
'DepositLoans
ElseIf ModuleId >= wis_DepositLoans And ModuleId < wis_DepositLoans + 100 Then
    SqlStr = "Select AccNum, A.LoanID as ID," & _
        " Title +' '+ FIrstName +' '+ MiddleName +' ' + LastName as CustName" & _
        " From DepositLoanMaster A,Nametab B" & _
        " Where A.CustomerID = B.CustomerID "
        
'Loan Accounts
ElseIf ModuleId >= wis_Loans And ModuleId < wis_Loans + 100 Then
    SqlStr = "Select AccNum, A.LoanID as ID," & _
        " Title +' '+ FIrstName +' '+ MiddleName +' ' + LastName as CustName" & _
        " From LoanMaster A,Nametab B" & _
        " Where A.CustomerID = B.CustomerID "
Else
    Exit Function
End If
    
If Trim(SqlStr) = "" Then Exit Function
    
If AccountId <= 0 Then
    SqlStr = SqlStr & " AND B.CustomerID = " & CustomerID
Else
    If InStr(10, SqlStr, "AccID", vbTextCompare) Then
        SqlStr = SqlStr & " AND A.AccID = " & AccountId
    Else
        SqlStr = SqlStr & " AND A.LoanID = " & AccountId
    End If
End If
    
    SqlStr = Trim(SqlStr)
    gDbTrans.SqlStmt = SqlStr

Dim rst As Recordset
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Function

GetAccountNumber = FormatField(rst("AccNum"))

If Not IsMissing(returnCustomerName) Then _
        returnCustomerName = Trim$(FormatField(rst("CustName")))


Exit_Line:


End Function

Public Function GetDataBaseName() As String

GetDataBaseName = gDbTrans.DataBaseName

Exit Function

Dim DBPath As String
Dim DbName As String

DBPath = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "server")
If DBPath = "" Then
    'Give the local path of the MDB FILE
    'DBPath = App.Path
    DBPath = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "DBPath")
    If DBPath = "" Then DBPath = App.Path
    
Else
    If InStr(1, DBPath, "\\", vbTextCompare) Then DBPath = Right(DBPath, Len(DBPath) - 2)
    DBPath = "\\" & DBPath & "\Index 2000"
    ''
    'MsgBox "No Database Functions vailiable for Net work DataBase"
End If


'Get the database Name
DbName = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "DBName")
If Len(GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "DSN")) Then
    GetDataBaseName = DbName
    Exit Function
End If
If DbName = "" Then
    DbName = "Index 2000.MDB"
Else
    DbName = IIf(StrComp(Right(DbName, 3), "mdb", vbTextCompare), "", ".mdb")
End If


GetDataBaseName = DBPath & "\" & DbName

End Function

Public Function GetFromDateString(strFromDate As String, Optional strToDate As String) As String

If strFromDate = "" Then strFromDate = strToDate
If strFromDate = strToDate Then strToDate = ""

If Len(strToDate) Then
    If gLangOffSet Then
        GetFromDateString = strFromDate & " " & _
            GetResourceString(342) & " " & _
            strToDate & " " & GetResourceString(108)
    Else
        GetFromDateString = GetResourceString(107) & " " & _
            strFromDate & " " & GetResourceString(108) & _
            " " & strToDate
    End If
Else
    If gLangOffSet Then
        GetFromDateString = strFromDate & " " & GetResourceString(362)
    Else
        GetFromDateString = GetResourceString(362) & " " & strFromDate
    End If
End If

End Function

'This function Returns the Last Transaction Id of then
'Contra transaction
Public Function GetMaxContraTransID() As Long

Dim ContraID As Long
Dim rst As Recordset

'Get the Contra ID
'put withdrawal transction details int to contra Table
gDbTrans.SqlStmt = "Select max(Contraid) as MaxContraid from ContraTrans "
ContraID = 0
If gDbTrans.Fetch(rst, adOpenForwardOnly) Then _
                        GetMaxContraTransID = FormatField(rst(0))

End Function

Public Function GetTransactionType(cmbTrans As ComboBox) As wisTransactionTypes
    On Error GoTo ErrLine
    Dim Trans As wisTransactionTypes
    Select Case cmbTrans.Text
        Case GetResourceString(271)
            Trans = wDeposit
        Case GetResourceString(272)
            Trans = wWithdraw
        Case GetResourceString(273)
            Trans = wContraWithdraw
        Case GetResourceString(47)
            Trans = wContraDeposit
    End Select
    GetTransactionType = Trans
    Exit Function
ErrLine:
    MsgBox ("Error in GettingTransaction Type")
End Function

Public Function GetUptoString(strTillDate As String) As String

If gLangOffSet Then
    GetUptoString = strTillDate & " " & GetResourceString(108)
Else
    'GetUptoString = GetResourceString(362) & " " & strFromDate
    GetUptoString = "Up to " & strTillDate
End If

End Function


Public Function GetChangeString(StrFirst As String, strSecond As String) As String

If gLangOffSet Then
    GetChangeString = strSecond & " " & StrFirst
Else
    GetChangeString = StrFirst & " " & strSecond
End If

End Function



Public Function GetMemberNumber(ByVal CustomerID As Long) As String

GetMemberNumber = ""
Dim rst As Recordset
gDbTrans.SqlStmt = "SELECT AccNum From MemMaster " & _
    " WHERE customerID = " & CustomerID

If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
        GetMemberNumber = FormatField(rst("AccNum"))

Set rst = Nothing

End Function

'This Function Returns The Name of the Account Type
'Input given to the function is HeadiD of transcation Ledger head
'
Public Function GetModuleName(AccountHeadID As Long) As String
Dim rstTemp As Recordset

Dim AccModule As wisModules

GetModuleName = ""

gDbTrans.SqlStmt = "Select AccType,HeadName " & _
            " From BankHeadId's " & _
            " Where HeadID = " & AccountHeadID
If gDbTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then Exit Function

AccModule = FormatField(rstTemp("AccType"))
GetModuleName = FormatField(rstTemp("HeadName"))
Set rstTemp = Nothing
Exit Function

If AccModule = wis_BKCC Then _
    GetModuleName = GetResourceString(229, 43)
If AccModule = wis_BKCCLoan Then _
    GetModuleName = GetResourceString(229)
If AccModule = wis_CAAcc Then _
    GetModuleName = GetResourceString(422)
If AccModule = wis_SBAcc Then _
    GetModuleName = GetResourceString(421)
If AccModule = wis_Deposits Then _
    GetModuleName = GetResourceString(43)
If AccModule = wis_DepositLoans Then _
    GetModuleName = GetResourceString(43, 58)
If AccModule = wis_Loans Then _
    GetModuleName = GetResourceString(58)
If AccModule = wis_Members Then _
    GetModuleName = GetResourceString(53, 36)
If AccModule = wis_RDAcc Then _
    GetModuleName = GetResourceString(424)
If AccModule = wis_PDAcc Then _
    GetModuleName = GetResourceString(425)
If AccModule = wis_SuspAcc Then _
    GetModuleName = GetResourceString(365)

End Function

Public Sub Initialize()

'''Check for the demo validity
Dim ProductID As String
Dim ProductVal As String
Dim NoOfEntries As Long
Dim Days As Long
Dim strTemp As String
Dim strToday As String
Dim strwisRegKey As String
Dim count As Integer
Dim MaxCount As Integer

Dim InstallDate As String
Dim ExpiryDate As String
Dim USInstallDate As Date
Dim USExpiryDate As Date

strwisRegKey = "Software\Waves Information Systems"
strToday = Format(Now, "dd/mm/yyyy")

ProductID = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "ProductID")
If ProductID = "" Then
    'Now Register the Product Key & Reutrn it to user
    Randomize
    ProductID = Format(Rnd(1) * 1000, "0000") & "-" & Format(Rnd(1) * 10000, "0000") & _
                "-" & Format(Rnd(1) * 10000, "0000") & "-" & Format(Rnd(1) * 10000, "0000")
    
    Call SetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "ProductId", ProductID)
End If

'Now Get the Product Value
MaxCount = Len(ProductID)
For count = 1 To MaxCount
    strTemp = Mid(ProductID, count, 1)
    If count Mod 5 Then strTemp = Right(CStr((Val(strTemp) + 1) * 3), 1)
    ProductVal = ProductVal & strTemp
Next

strwisRegKey = "Software\Waves Information Systems\" & ProductID
If ProductVal <> GetRegistryValue(HKEY_CURRENT_USER, strwisRegKey, "License") Then
    MsgBox "Product Key " & ProductID & vbCrLf & _
        "Your Validity period expired Please CONTACT" & vbCrLf & _
        "Waves Information Systems., Gadag", vbInformation, wis_MESSAGE_TITLE
    'End  'Exit Sub
End If

count = Val(GetRegistryValue(HKEY_CURRENT_USER, strwisRegKey, "Validity"))
InstallDate = GetRegistryValue(HKEY_CURRENT_USER, strwisRegKey, "InstallDate")

If count > 0 Then
    NoOfEntries = Val(GetRegistryValue(HKEY_CURRENT_USER, strwisRegKey, "Count"))
    If DateValidate(InstallDate, "/", True) Then
        InstallDate = GetSysFormatDate(InstallDate)
        ExpiryDate = DateAdd("d", CInt(count), InstallDate)
        If DateDiff("d", CDate(ExpiryDate), Now) > 0 Then
            Call SetRegistryValue(HKEY_CURRENT_USER, strwisRegKey, "License", "1111-2222-3333-4444")
            MsgBox "Your Demo Period Has Been Expired, Contact The Vendor", vbInformation, wis_MESSAGE_TITLE & " - Demo"
            End
        End If
    End If
    
    'In case he changes his system date
    'we will restrict the user by counting the no of logins
    'We consider he logs 3 time per day
    If Val(NoOfEntries) > (count * 3) Then
        Call SetRegistryValue(HKEY_CURRENT_USER, strwisRegKey, ProductID, "1111-2222-3333-4444")
        MsgBox "Your Demo Period Has Been Expired, Contact The Vendor", vbInformation, wis_MESSAGE_TITLE & " - Demo"
        End
    End If
    NoOfEntries = CStr(Val(NoOfEntries) + 1)
    Call SetRegistryValue(HKEY_CURRENT_USER, strwisRegKey, "Count", NoOfEntries)
End If

If DateValidate(InstallDate, "/", True) Then
    USInstallDate = GetSysFormatDate(InstallDate)
    
    'Now Check the Date of today And CHeck the Expiry date
    'If Expiry date earlir then the login date i.e. today
    'Then it means that he has no paid the AMC Charge
    'Then warn him for the First time
    'then allow a the user to use soft ware for 1 week
    'After a week or two do not allow him to user the
    'Dim ExpiryDate As String
    Dim Warned As Boolean
    ExpiryDate = "24/08/2071"
    If DateDiff("d", USInstallDate, Now) > 365 Then
        ExpiryDate = GetRegistryValue(HKEY_CURRENT_USER, strwisRegKey, "ExpiryDate")
        If Len(ExpiryDate) = 0 Then
            ExpiryDate = GetSysFormatDate(ExpiryDate)
            ExpiryDate = DateAdd("m", 1, DateAdd("YYYY", 1, ExpiryDate))
            ExpiryDate = GetIndianDate(CDate(ExpiryDate))
            Call SetRegistryValue(HKEY_CURRENT_USER, strwisRegKey, "ExpiryDate", ExpiryDate)
        End If
    End If
    
    USExpiryDate = GetSysFormatDate(ExpiryDate)
    Days = DateDiff("d", USExpiryDate, Now)
    If Days > 0 Then
        
        Warned = IIf(UCase(GetRegistryValue(HKEY_CURRENT_USER, strwisRegKey, "Warned")) = "TRUE", True, False)
        If Not Warned Then
            MsgBox "You have not paid the the Amc charges of the the current year", vbInformation, wis_MESSAGE_TITLE
            Call SetRegistryValue(HKEY_CURRENT_USER, strwisRegKey, "Warned", "True")
        Else
            MsgBox "You have not paid the the Amc charges " & _
                " of the the current year" & vbCrLf & _
                "This product will not work after " & ExpiryDate _
                , vbInformation, wis_MESSAGE_TITLE
        End If
        If Days > 15 Then End  'Appliaction terminates
        
    End If
End If

'Initialize the global variables
If gDbTrans Is Nothing Then Set gDbTrans = New clsTransact
'If gDbTrans Is Nothing Then Set gDbTrans = CreateObject("Transaction.Transact")
gAppPath = App.Path

Dim ret As Long
Dim strRet As String

gImagePath = App.Path & "\"
If Len(gImagePath) = 0 Then
    strRet = String$(512, 0)
    ret = GetPrivateProfileString("system", "imagePath", "", strRet, Len(strRet), App.Path & "\\index_images.ini")
    If (ret < 0) Then
        MsgBox "Unable to read configuration file 'index_images.ini'.  Ensure that this file exists."
    Else
        gImagePath = Mid$(strRet, 1, ret)
        ''Check for the existsnce of path
        If Len(Dir$(gImagePath)) < 1 Then
            MsgBox "Please Create the folder " & gImagePath & " for images", vbOKOnly, "Index 2000"
            gImagePath = App.Path & "\"
        End If
    End If
End If

Exit Sub

End Sub

Public Sub LoadAccountGroups(cmbObject As ComboBox, Optional SetDefault As Boolean)

Dim rst As Recordset
gDbTrans.SqlStmt = "SELECT * From AccountGroup"
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    With cmbObject
        .Clear
        .Locked = True
        .Visible = False
        Dim CountGroup As Byte
        While Not rst.EOF
            .AddItem FormatField(rst("GroupName"))
            .ItemData(.newIndex) = rst("AccGroupId")
            rst.MoveNext
            CountGroup = CountGroup + 1
        Wend
        If CountGroup > 1 Then
            .AddItem "", 0
            .Locked = False
            .Visible = True
            .ZOrder 0
        End If
        
        .ListIndex = 0
        
    End With
Else

    gDbTrans.SqlStmt = "Insert Into AccountGroup " & _
            "(AccGroupID,GroupName) " & _
            " Values (1, " & AddQuotes(GetResourceString(339), True) & ")"
    gDbTrans.BeginTrans
    Call gDbTrans.SQLExecute
    gDbTrans.CommitTrans
    With cmbObject
        .Clear
        '.AddItem ""
        .AddItem GetResourceString(339)
        .ItemData(.newIndex) = 1
        .ListIndex = 0
        .Locked = True
        .Visible = False
    End With
End If

End Sub

Public Sub LoadCustomerTypes(cmbObject As ComboBox)
Dim rst As Recordset

'Load Customer Type
'Dim Rst As Recordset
gDbTrans.SqlStmt = "SELECT * From CustomerType "

If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    While Not rst.EOF
        With cmbObject
            .AddItem FormatField(rst("custTypeName"))
            .ItemData(.newIndex) = FormatField(rst("custType"))
        End With
        rst.MoveNext
    Wend
    Set rst = Nothing
Else
    Dim InTrans As Boolean
    InTrans = gDbTrans.BeginTrans
   gDbTrans.SqlStmt = "Insert Into CustomerType " & _
            " (CustType, CustTypeName,UIType ) " & _
            " Values (" & 1 & ", " & _
            AddQuotes(GetResourceString(339)) & ", " & _
            0 & ")"
    If Not gDbTrans.SQLExecute Then Exit Sub
     If InTrans Then gDbTrans.CommitTrans
End If

End Sub

Public Sub LoadDepositTypes(cmbObject As ComboBox)
 'add the Loan deposit types here
 Dim Deptype As wis_DepositType
      
      With cmbObject
        .Clear
         .AddItem ""
         .ItemData(.newIndex) = 0
         
        Dim rstDeposit As Recordset
        gDbTrans.SqlStmt = "SELECT * From DepositName"
        If gDbTrans.Fetch(rstDeposit, adOpenDynamic) > 0 Then
            While Not rstDeposit.EOF
                .AddItem FormatField(rstDeposit("DepositName"))
                '.ItemData(.NewIndex) = wis_Deposits + RstDeposit("DepositID")
                .ItemData(.newIndex) = FormatField(rstDeposit("DepositID"))
                rstDeposit.MoveNext
            Wend
        End If

         Deptype = wisDeposit_RD
        .AddItem GetResourceString(424) '"Recuring Deposit"
        .ItemData(.newIndex) = Deptype
        
        Deptype = wisDeposit_PD
        .AddItem GetResourceString(425) '"Pigmy Depoisit"
        .ItemData(.newIndex) = Deptype
        
      End With
End Sub

Public Sub LoadPlaces(cmbObject As ComboBox)
Dim rst As Recordset
gDbTrans.SqlStmt = "Select * From PlaceTab"
cmbObject.AddItem ""
If gDbTrans.Fetch(rst, adOpenDynamic) Then
    With cmbObject
        While Not rst.EOF
            .AddItem FormatField(rst("Place"))
            rst.MoveNext
        Wend
    End With
Else
    cmbObject.AddItem "Home Town"
End If
End Sub
Public Sub LoadFarmerTypes(cmbObject As ComboBox)
Dim rst As Recordset
cmbObject.AddItem ""
gDbTrans.SqlStmt = "Select * From FarmerTypeTab"
If gDbTrans.Fetch(rst, adOpenDynamic) Then
    With cmbObject
        While Not rst.EOF
            .AddItem FormatField(rst("TypeName"))
            .ItemData(.newIndex) = FormatField(rst("FarmerTypeID"))
           rst.MoveNext
        Wend
    End With
End If
End Sub

Public Sub LoadCastes(cmbObject As ComboBox)
Dim rst As Recordset
cmbObject.AddItem ""
gDbTrans.SqlStmt = "Select * From CasteTab"
If gDbTrans.Fetch(rst, adOpenDynamic) Then
    With cmbObject
        While Not rst.EOF
            .AddItem FormatField(rst("Caste"))
            rst.MoveNext
        Wend
    End With
Else
    cmbObject.AddItem "Indian"
End If
End Sub

Public Sub LoadGender(cmbObject As ComboBox)
'LoadAccountGroup(cmb(1))
Dim Gender As wis_Gender
With cmbObject
    Gender = wisNoGender
    .AddItem GetResourceString(338) ''All
    .ItemData(.newIndex) = Gender
    
    Gender = wisMale
    .AddItem GetResourceString(385) ''mALE
    .ItemData(.newIndex) = Gender
    
    Gender = wisFemale
    .AddItem GetResourceString(386) ''Female
    .ItemData(.newIndex) = Gender

End With
End Sub


Public Sub Main()

frmSplash.Show 'vbModal
Call PauseApplication(1)

'Get Date FormglangOffset =
'gStrDate = Format(Now) ', "mm/dd/yyyy")
'gStrDate = Left(gStrDate, InStr(1, gStrDate, " ", vbTextCompare) - 1)
gStrDate = Format(Now, "DD/MM/YYYY")
DateFormat = "dd/mm/yyyy"
Call Initialize

Call KannadaInitialize

Load wisMain
Unload frmSplash
Set frmSplash = Nothing

Set gCurrUser = New clsUsers

wisMain.Show

With gCurrUser
    .MaxRetries = 3
    .ShowLoginDialog
    If Not .LoginStatus Then
        Do
            Call ExitApplication(False, 1)
            'Unload wisMain
        Loop
        End
    End If
End With


'Temprary code for a year
'Now Crete the Tab Main
Call KannadaInitialize
Call PostLoginInitialize

'If It is online then Show theBegin date
If gOnLine Then Call BeginDayTrans
''Check for any Tempray execution
Call TempSub

Exit Sub

End Sub
Private Sub TempSub()
    'Check For Memei Id o
    gDbTrans.SqlStmt = "UPDATE LoanMaster INNER JOIN MemMaster ON " & _
        " LoanMaster.CustomerID = MemMaster.CustomerID " & _
        " SET LoanMaster.MemId = MemMaster.AccID where LoanMaster.MemId=0"
    gDbTrans.BeginTrans
    If gDbTrans.SQLExecute Then
        gDbTrans.CommitTrans
    Else
        gDbTrans.RollBack
    End If
    
    
     'Check For Memei Id o
    gDbTrans.SqlStmt = "UPDATE ShareTrans A INNER JOIN MemTrans B" & _
        " ON A.AccID = B.AccID AND (A.SaleTransid = B.TransID or A.ReturnTransID=b.TransID)" & _
        " SET A.ReturnTransID = Null " & _
        " Where Not exists (Select TransID From MemTrans C Where (C.TransType =2 or C.TransType =4) " & _
        " AND C.AccId=A.AccID and C.TransID = A.ReturnTransID )"
    gDbTrans.BeginTrans
    If gDbTrans.SQLExecute Then
        gDbTrans.CommitTrans
    Else
        gDbTrans.RollBack
    End If
End Sub

Private Sub PostLoginInitialize()

gUserID = gCurrUser.UserID

Dim SetUp As clsSetup
Set SetUp = New clsSetup

gImagePath = SetUp.ReadSetupValue("General", "ImagePath", "")

If Len(Trim(gImagePath)) > 0 Then _
    If Dir(gImagePath, vbDirectory) = "" Then MakeDirectories (gImagePath)

Set SetUp = Nothing
gDbTrans.SqlStmt = "SELECT CustomerID,Title + ' ' + FirstName + ' ' + MiddleName +' '+ " & _
        " LastName as NAME,Place,Caste,Gender,IsciName From NameTab"
gDbTrans.CreateView ("QryName")

gDbTrans.SqlStmt = "SELECT CustomerID,Title + ' ' + FirstName + ' ' + MiddleName +' '+ " & _
        " LastName as NAME,IsciName From NameTab"
gDbTrans.CreateView ("QryOnlyName")

gDbTrans.SqlStmt = "SELECT X.CustomerID,Y.AccNum as MemberNum,Y.ACCID as MemID, Title + ' ' + FirstName + ' ' + MiddleName +' '+ " & _
        " LastName as NAME,Place,Caste,Gender,IsciName,MemberType From NameTab X " & _
        " Inner Join MemMaster Y on Y.CustomerID = X.CustomerID"
gDbTrans.CreateView ("QryMemName")

'Free Cust Id
gDbTrans.BeginTrans

gDbTrans.SqlStmt = "Delete * FROM FreeCustId Where FreeID in (Select distinct customerID from NameTab)"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack

gDbTrans.SqlStmt = "Update FreeCustId set Selected = False Where FreeID in (Select distinct customerID from NameTab)"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack

'TransID
gDbTrans.SqlStmt = "Delete * FROM IdFromInventory Where TransID not in (Select distinct TransID from AccTrans)" & _
    " and TransID not in (Select distinct TransID from TransParticulars)"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack

'TEMPORARY CODE REMOVE IN 2016
'UPdate the New ModuleIds OD SB and Memebre
gDbTrans.SqlStmt = "Update BankHeadIDs Set AccType = 2000 where AccType = 2"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack
gDbTrans.SqlStmt = "Update InterestTab Set ModuleID = 2000 where ModuleID = 2"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack
gDbTrans.SqlStmt = "Update NoteTab Set ModuleID = 2000 where ModuleID = 2"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack


gDbTrans.SqlStmt = "Update BankHeadIDs Set AccType = 3000 where AccType = 3"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack
gDbTrans.SqlStmt = "Update InterestTab Set ModuleID = 3000 where ModuleID = 3"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack
gDbTrans.SqlStmt = "Update NoteTab Set ModuleID = 3000 where ModuleID = 3"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack

gDbTrans.SqlStmt = "Update BankHeadIDs Set AccType = 4000 where AccType = 4"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack
gDbTrans.SqlStmt = "Update InterestTab Set ModuleID = 4000 where ModuleID = 4"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack
gDbTrans.SqlStmt = "Update NoteTab Set ModuleID = 4000 where ModuleID = 4"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack

gDbTrans.SqlStmt = "Update BankHeadIDs Set AccType = 5000 where AccType = 5"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack
gDbTrans.SqlStmt = "Update InterestTab Set ModuleID = 5000 where ModuleID = 5"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack
gDbTrans.SqlStmt = "Update NoteTab Set ModuleID = 5000 where ModuleID = 5"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack

gDbTrans.SqlStmt = "Update BankHeadIDs Set AccType = 8000 where AccType = 8"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack
gDbTrans.SqlStmt = "Update InterestTab Set ModuleID = 8000 where ModuleID = 8"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack
gDbTrans.SqlStmt = "Update NoteTab Set ModuleID = 8000 where ModuleID = 8"
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack


gDbTrans.SqlStmt = "Update BankHeadIDs Set AccType = AccType + " & 10 & _
        " where AccType < " & wis_Deposits + 10 & " and AccType > " & wis_Deposits

If Not gDbTrans.SQLExecute Then gDbTrans.RollBack

gDbTrans.SqlStmt = "Update BankHeadIDs Set AccType = AccType + " & 10 & _
        " where AccType < " & wis_DepositLoans + 10 & " and AccType > " & wis_DepositLoans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBack
'TEMPORARY CODE REMOVE IN 2016

gDbTrans.CommitTrans

''Correct the Deposit ID
Dim RstDep As Recordset
gDbTrans.SqlStmt = "SELECT * from DepositName where DepositID > 10 "
If gDbTrans.Fetch(RstDep, adOpenDynamic) > 0 Then Exit Sub 'Already converted

'DepositLoanMaster

gDbTrans.BeginTrans
'Select All the Deposit WHich are More than 4
Dim rstDepLoan As Recordset
Set RstDep = Nothing
gDbTrans.SqlStmt = "SELECT LoanId,CustomerID,DepositType,PledgeDescription from DepositLoanMaster where DepositType = 4 or DepositType = 8 "

If gDbTrans.Fetch(rstDepLoan, adOpenDynamic) > 0 Then
    ''414/1;414;417;422;425
    Dim depNum As String
    Dim DepositId As Integer
    
    Do
        'Get the Deposit
        depNum = FormatField(rstDepLoan("PledgeDescription"))
        DepositId = FormatField(rstDepLoan("DepositType"))
        
        gDbTrans.SqlStmt = "SELECT * from " & IIf(DepositId = wisDeposit_RD, "RD", "PD") & "Master" & _
            " Where AccNum = " & AddQuotes(depNum, True)
        If gDbTrans.Fetch(RstDep, adOpenDynamic) < 1 Then
            gDbTrans.SqlStmt = "Update DepositLoanMaster Set DepositType = DepositType + 10 Where LoanId = " & rstDepLoan("LoanID")
            If Not gDbTrans.SQLExecute Then
                gDbTrans.RollBack
                Exit Sub
            End If
        End If
NextDepLoan:
        rstDepLoan.MoveNext
        If rstDepLoan.EOF = True Then Exit Do
        
    Loop

End If
gDbTrans.SqlStmt = "Update DepositName Set DepositID = DepositID +10"
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Sub
End If
'FDMaster
gDbTrans.SqlStmt = "Update FDMAster Set DepositType = DepositType + 10"
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Sub
End If
'DepositLoanMaster
gDbTrans.SqlStmt = "Update DepositLoanMaster Set DepositType = DepositType + 10 where DepositType < 10 and DepositType <> 4 and DepositType <> 8"
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Sub
End If

'DepositLoanMaster
gDbTrans.SqlStmt = "Update BankHeadIDs Set AccType = AccType + " & wis_Deposits + 10 & _
        " where AccType < " & wis_Deposits + 10 & " and AccType > " & wis_Deposits
        
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Sub
End If


gDbTrans.CommitTrans


End Sub


Public Function SetTransactionCombo(cmbTrans As ComboBox, transType As wisTransactionTypes) As Boolean
    On Error GoTo ErrLine
    SetTransactionCombo = False
    Dim strTrans As String
        If transType = wDeposit Then strTrans = GetResourceString(271)
        If transType = wWithdraw Then strTrans = GetResourceString(272)
        If transType = wContraDeposit Then strTrans = GetResourceString(273)
        If transType = wContraWithdraw Then strTrans = GetResourceString(47)
    
    With cmbTrans
        Dim I As Integer
        For I = 0 To .ListCount - 1
            If .List(I) = strTrans Then '
                .ListIndex = I
                Exit For
            End If
        Next
    End With
    SetTransactionCombo = True
    Exit Function
ErrLine:
End Function
Public Function GetResourcseString(ResourceID As Long)
    GetResourcseString = GetResourceString(ResourceID)
End Function
'-----------------------------------------------------------
' SUB: UpdateStatus
'
' "Fill" (by percentage) inside the PictureBox and also
' display the percentage filled
'
' IN: [pic] - PictureBox used to bound "fill" region
'     [sngPercent] - Percentage of the shape to fill
'     [fBorderCase] - Indicates whether the percentage
'        specified is a "border case", i.e. exactly 0%
'        or exactly 100%.  Unless fBorderCase is True,
'        the values 0% and 100% will be assumed to be
'        "close" to these values, and 1% and 99% will
'        be used instead.
'
' Notes: Set AutoRedraw property of the PictureBox to True
'        so that the status bar and percentage can be auto-
'        matically repainted if necessary
'-----------------------------------------------------------
'
Public Sub UpdateStatus(pic As PictureBox, ByVal sngPercent As Single, Optional ByVal fBorderCase As Boolean = False)
    Dim strPercent As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer

    'For this to work well, we need a white background and any color foreground (blue)
    Const colBackground = &HFFFFFF ' white
    Const colForeground = &H800000 ' dark blue

    pic.ForeColor = colForeground
    pic.BackColor = colBackground
    
    '
    'Format percentage and get attributes of text
    '
    Dim intPercent
    intPercent = Int(100 * sngPercent + 0.5)
    
    'Never allow the percentage to be 0 or 100 unless it is exactly that value.  This
    'prevents, for instance, the status bar from reaching 100% until we are entirely done.
    If intPercent <= 0 Then
        intPercent = 1
        If Not fBorderCase Then
            intPercent = 1
        End If
    ElseIf intPercent >= 100 Then
        intPercent = 99
        If Not fBorderCase Then
            intPercent = 99
        End If
    End If
    
    strPercent = Format$(intPercent) & "%"
    intWidth = pic.TextWidth(strPercent)
    intHeight = pic.TextHeight(strPercent)

    '
    'Now set intX and intY to the starting location for printing the percentage
    '
    intX = pic.Width / 2 - intWidth / 2
    intY = pic.Height / 2 - intHeight / 2

    '
    'Need to draw a filled box with the pics background color to wipe out previous
    'percentage display (if any)
    '
    pic.DrawMode = 13 ' Copy Pen
    pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF

    '
    'Back to the center print position and print the text
    '
    pic.CurrentX = intX
    pic.CurrentY = intY
    pic.Print strPercent

    '
    'Now fill in the box with the ribbon color to the desired percentage
    'If percentage is 0, fill the whole box with the background color to clear it
    'Use the "Not XOR" pen so that we change the color of the text to white
    'wherever we touch it, and change the color of the background to blue
    'wherever we touch it.
    '
    pic.DrawMode = 10 ' Not XOR Pen
    If sngPercent > 0 Then
        pic.Line (0, 0)-(pic.Width * sngPercent, pic.Height), pic.ForeColor, BF
    Else
        pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF
    End If

    pic.Refresh
End Sub

