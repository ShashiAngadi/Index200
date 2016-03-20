Attribute VB_Name = "basMain"
Option Explicit

Public gStrDate As String
Public gAppPath As String
Public Const gAppName = "INDEX 2000"
Public gDBTrans As clsTransact
Public gCurrUser As clsUsers
Public gTranstype As wisTransactionTypes
Public gUser As clsUsers
Public gUserID As Long
Public gCompanyName As String
'Added On 23/9/2000

Public gCancel As Boolean
Public gWindowHandle As Long
Public gBank As Boolean


'By shashi on 9/9/2001
' Structure for holding the fields info for table.
' This is used by CreateDB function.
Private Type TabStruct
    Field As String
    Type As String
    Length As Integer
    'Index As Boolean
    Required As Boolean
    'Primary As Boolean
    AutoIncrement As Boolean
End Type



Private Function ChangePassWord(DBPath As String) As Boolean
Dim TmpDB As Database
Set TmpDB = Workspaces(0).OpenDatabase(DBPath, _
        True, False, ";pwd=WIS!@#")
TmpDB.NewPassword "WIS!@#", "PRAGMANS"
TmpDB.Close
Set TmpDB = Nothing
ChangePassWord = True
End Function

Private Function CheckPassWord(DBPath As String) As Boolean
On Error GoTo ErrLine
Dim TmpDB As Database
Set TmpDB = Workspaces(0).OpenDatabase(DBPath, _
        True, False, ";pwd=PRAGMANS")
TmpDB.Close
Set TmpDB = Nothing
CheckPassWord = True
Exit Function
ErrLine:
    CheckPassWord = False
    Err.Clear
End Function


'
'This Function Will carry Out Every Time
'INdex Runs 'if there is any change in Data base
' it will rectify the Database
' udsually it has to create the table in database
' The code in this shoul not be altered at least one year
'
Private Sub ChangesInDataBase(DBName As String, StrPAssWord As String)
Dim Count As Integer
'Count will Keep the No of changes has to be done
'If it runs More than the required time


' then it is sure that theres is problem in altering database
'Whenver code changes the count has to change

' when you write code please speicify the date
'Once changed Code has to run atleas one year
'to update the database every where

'Here Creating Table to keep the opening balance of the bank every day
'
Dim TheTable As TableDef
Dim TheFields() As Field, TheIndex As Index
Dim i As Integer
Dim PrimaryKeySet As Boolean
'On Error GoTo err_line

If Not CheckPassWord(DBName) Then ChangePassWord (DBName)
Count = 4

Dim TmpDB As Database

If StrPAssWord = "" Then
    Set TmpDB = Workspaces(0).OpenDatabase(DBName, StrPAssWord)
Else
    Set TmpDB = OpenDatabase(DBName, False, False, ";pwd=" & StrPAssWord)
End If
On Error Resume Next
Dim TblCount As Integer
Dim TableName() As String
ReDim TableName(4)
TableName(1) = "OBTab"
TableName(2) = "Material"
TableName(3) = "BKCCTrans"
TableName(4) = "TEMP"

    For TblCount = 1 To Count
Recheck:
'''        If TblCount = 1 Then
            Set TheTable = TmpDB.CreateTableDef(TableName(TblCount))
            If Err.Number = 3010 Then 'table already exists
                Err.Clear
                GoTo NextTable
            ElseIf Err.Number Then
                GoTo Err_Line
            End If
            ' Create and add the fields.
            Dim TheField As Field
            With TheTable
                Select Case TblCount
                    Case 1
                        'Create First field
                        Set TheField = .CreateField("OBDate", dbDate)
                        TheField.Required = True
                        .fields.Append TheField
                        'Create 2nd field
                        Set TheField = .CreateField("OBAmount", dbCurrency)
                        TheField.Required = True
                        .fields.Append TheField
                        '3rd field
                        Set TheField = .CreateField("Module", dbInteger)
                        TheField.Required = False
                        .fields.Append TheField
                    Case 2
                        'Create First field
                        Set TheField = .CreateField("TransDate", dbDate)
                        TheField.Required = True
                        .fields.Append TheField
                        'Create 2nd field
                        Set TheField = .CreateField("Amount", dbCurrency)
                        TheField.Required = True
                        .fields.Append TheField
                        '3rd field
                        Set TheField = .CreateField("Module", dbInteger)
                        TheField.Required = True
                        .fields.Append TheField
                        '4th field
                        Set TheField = .CreateField("TransType", dbInteger)
                        TheField.Required = True
                        .fields.Append TheField
                    Case 3
                        'Create First field
                        Set TheField = .CreateField("LoanID", dbLong)
                        TheField.Required = True
                        .fields.Append TheField
                        Set TheField = .CreateField("TransID", dbLong)
                        TheField.Required = True
                        .fields.Append TheField
                        Set TheField = .CreateField("TransType", dbInteger)
                        TheField.Required = True
                        .fields.Append TheField
                        Set TheField = .CreateField("TransDate", dbDate)
                        TheField.Required = True
                        .fields.Append TheField
                        'Create 2nd field
                        Set TheField = .CreateField("Amount", dbCurrency)
                        TheField.Required = True
                        .fields.Append TheField
                        Set TheField = .CreateField("Balance", dbCurrency)
                        TheField.Required = True
                        .fields.Append TheField
                        '3rd field
                        Set TheField = .CreateField("Particulars", dbText, 40)
                        TheField.Required = True
                        TheField.AllowZeroLength = True
                        .fields.Append TheField
                    Case 4
                        'Create First field
                        Set TheField = .CreateField("AccID", dbLong)
                        .fields.Append TheField
                        Set TheField = .CreateField("MaxTransID", dbLong)
                        .fields.Append TheField
                        Set TheField = .CreateField("UserID", dbLong)
                        .fields.Append TheField
               End Select
            End With
            TmpDB.TableDefs.Append TheTable
            If Err.Number = 3010 Then
                Err.Clear
            End If
NextTable:
        Next TblCount

Set TheTable = Nothing
    
    TmpDB.Close
    Set TmpDB = Nothing

Err_Line:
    If Err.Number = 3010 Then
        Err.Clear
        GoTo Recheck
        Exit Sub
    ElseIf Err.Number > 0 Then
        Debug.Print Err.Number
        MsgBox Err.Number & vbCrLf & Err.Description
    End If
    On Error Resume Next
'''    TmPDB.Close
    Set TmpDB = Nothing
End Sub





'
Public Function MakeBackUp(DataFileName As String) As Boolean
Dim BaseFolder As String
Dim PresentDate As String
Dim FolderName As String
Dim PrevFolder As String

PresentDate = Format(Now, "DD/YY/MM")
BaseFolder = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\waves information systems\index 2000\settings", "BaseFolder")
PrevFolder = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\waves information systems\index 2000\settings", "dbfoldername")

'if date = 22/2/01 then
'foldername becomes wis2201

If Trim(BaseFolder) = "" Then BaseFolder = App.Path
If Right(BaseFolder, 1) = "\" Then BaseFolder = Left(BaseFolder, Len(BaseFolder) - 1)
FolderName = BaseFolder + "\" + "wis" + Left(PresentDate, 2) + Right(PresentDate, 2)

'If todays back up already taken then exit

If PrevFolder = FolderName Then GoTo ExitLine
    On Error Resume Next
    'Before deleteing the daata base check the root
    If App.Path <> PrevFolder Then
        Kill PrevFolder & "\index 2000.mdb"
        RmDir PrevFolder
    End If
    On Error GoTo ErrLine
    ' check for the existence of path
    If Not DoesPathExist(FolderName) Then
        MakeDirectories (FolderName)
    End If
    gDBTrans.CloseDB
    FileCopy App.Path + "\" + "index 2000.mdb", FolderName + "\index 2000.mdb"
    Call SetRegistryValue(HKEY_LOCAL_MACHINE, "software\waves information systems\index 2000\settings", "dbfoldername", FolderName)
'''    Call gDBTrans.OpenDB(DataFileName, "WIS!@#")
    
ExitLine:

MakeBackUp = True

'Exit Function
 
ErrLine:
    If Err Then
        MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & "Unable to take back up", vbInformation
    End If
    Call gDBTrans.OpenDB(DataFileName, "WIS!@#")
End Function
'This Function Checks Whether The Customer Has Already an Active account Operating In
'Requested Module, if yes It Returns True ---
'30th Aug 2000 Vinay
Public Function CustomerExisting(ModuleID As wisModules, CustomerId As Long) As Boolean
Dim TableName As String

On Error GoTo ErrLine

'Fire The Query
If ModuleID = wis_CAAcc Then
TableName = "CAMaster"
ElseIf ModuleID = wis_dlAcc Then
TableName = "DLMaster"
ElseIf ModuleID = wis_FDAcc Then
TableName = "FDMaster"
ElseIf ModuleID = wis_Members Then
TableName = "MMMaster"
ElseIf ModuleID = wis_PDAcc Then
TableName = "PDMaster"
ElseIf ModuleID = wis_RDAcc Then
TableName = "RDMaster"
ElseIf ModuleID = wis_sbacc Then
TableName = "SBMaster"
Else
    GoTo ExitLine
End If

gDBTrans.SQLStmt = "Select * From " & TableName & " Where CustomerID = " & CustomerId
    
    If gDBTrans.SQLFetch > 0 Then
        CustomerExisting = True
    Else
         CustomerExisting = False
    End If
 
 GoTo ExitLine

ErrLine:
    MsgBox "Error In Function CustomerExisting"

ExitLine:
End Function

#If Handle Then
Public Sub SetLoadedWindows()
Static Repeat As Boolean
On Error GoTo ExitLine
Dim Count As Byte
If Repeat Then
    Exit Sub
End If
Repeat = True
For Count = LBound(gWindowHandles) To UBound(gWindowHandles)
   If gWindowHandles(Count) <> 0 Then
      Call SetActiveWindow(gWindowHandles(Count))
   End If
Next Count
Repeat = False
'Call SetActiveWindow(gWindowHandle)
ExitLine:

End Sub

Public Sub SetWindowHandle(hWnd As Long, Show As Boolean, Optional AbortClose As Integer)
   
   'Static gWindowHandles() As Long
   Static PresentHandle As Long
   Static RefNo As Byte
   Dim Count As Byte
         
   On Error GoTo ErrLine
   If Show Then ' If Form is sHOWINF THEN
      For Count = LBound(gWindowHandles) To UBound(gWindowHandles)
         If gWindowHandles(Count) = hWnd Then
            Exit Sub
         End If
      Next Count
      RefNo = RefNo + 1
      ReDim Preserve gWindowHandles(RefNo)
      gWindowHandles(RefNo - 1) = PresentHandle
      gWindowHandle = hWnd
      PresentHandle = hWnd
   
   Else  ' If form is unloading
      If gWindowHandle = hWnd Then
         RefNo = RefNo - 1
         PresentHandle = gWindowHandles(RefNo)
         gWindowHandle = gWindowHandles(RefNo)
         ReDim Preserve gWindowHandles(RefNo)
      Else 'Check Whether This window has shown previously
         For Count = LBound(gWindowHandles) To UBound(gWindowHandles)
            If gWindowHandles(Count) = hWnd Then
               Dim LoadCount As Byte
               For LoadCount = Count + 1 To UBound(gWindowHandles)
                Call CloseWindow(gWindowHandles(LoadCount))
               Next
               Exit Sub
            End If
         Next Count
      End If
   End If
   Exit Sub
ErrLine:
If Err.Number = 6 Then
   RefNo = 1
   Resume
   
ElseIf Err.Number = 9 Then
   ReDim gWindowHandles(RefNo)
   Resume
End If

End Sub

#End If

Public Sub Initialize()

'''Check for the demo validity
Dim ProductID As String
Dim InstallDate As String
Dim ExpiryDate As String
Dim NoOfEntries As String
Dim Days As String

'''ProductID = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\waves information systems\Index 2000\ProductID", "ProductID")
'''If ProductID = "" Then
'''    MsgBox "Your application has corrupted Please contact vendor" & vbCrLf & _
'''       "Waves Information Systtems., Bangalore (or Gadag)", vbInformation, wis_MESSAGE_TITLE
'''    End  'Exit Sub
'''End If
'''
'''
'''InstallDate = GetRegistryValue(HKEY_CURRENT_USER, "Software\waves information systems\" & ProductID & "\InstallDate", "Date")
'''Days = GetRegistryValue(HKEY_CURRENT_USER, "Software\waves information systems\" & ProductID & "\Validity", "Days")
'''NoOfEntries = GetRegistryValue(HKEY_CURRENT_USER, "Software\waves information systems\" & ProductID & "\Validity", "Count")
'''
'''If Days = Null Then Days = ""
'''If NoOfEntries = Null Then NoOfEntries = ""
'''If Val(Days) > 0 Then
'''    ExpiryDate = FormatDate(DateAdd("d", CInt(Days), FormatDate(InstallDate)))
'''    If WisDateDiff(ExpiryDate, Format(Now, "dd/mm/yyyy")) > 0 Then
'''        MsgBox "Your Demo Period Has Been Expired, Contact The Vendor", vbInformation, wis_MESSAGE_TITLE & " - Demo"
'''        End
'''    End If
'''    'In case he changes his system date
'''    'we will restrict the user by caounnig the no of entries
'''    If Val(NoOfEntries) > CInt(Days) * 10 Then
'''        MsgBox "Your Demo Period Has Been Expired, Contact The Vendor", vbInformation, wis_MESSAGE_TITLE & " - Demo"
'''        End
'''    End If
'''    NoOfEntries = CStr(Val(NoOfEntries) + 1)
'''    Call SetRegistryValue(HKEY_CURRENT_USER, "software\Waves information Systems\" & ProductID & "\Validity", "Count", NoOfEntries)
'''End If

'Initialize the global variables
    gAppPath = App.Path
    If gDBTrans Is Nothing Then
        Set gDBTrans = New clsTransact
    End If

'Get the database name
Dim DBPath As String
Dim DBFileName As String

DBPath = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\waves information systems\index 2000\settings", "server")

 If DBPath = "" Then
    'Give the local path of the MDB FILE
    DBPath = App.Path
Else
    If InStr(1, DBPath, "\\", vbTextCompare) Then DBPath = Right(DBPath, Len(DBPath) - 2)
    DBPath = "\\" & DBPath & "\Index 2000"
End If

DBFileName = DBPath & "\" & gAppName & ".MDB"


'Check for the cahnges in databse
Call ChangesInDataBase(DBFileName, "PRAGMANS")


'Open the data base
'''    Debug.Assert DBFileName = ""
    If Not gDBTrans.OpenDB(DBFileName, "PRAGMANS") Then
        If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
            End
        End If
        If Not gDBTrans.CreateDB(gAppPath & "\Index 2000.TAB", "WIS!@#") Then
            MsgBox "Unable to create new database !", vbCritical, gAppName & " - Error"
            On Error Resume Next
             Kill gAppPath & "\Index 2000.MDB"
             End
        End If
    Else
'''        gDBTrans.CompactDataBase
        'Make a routine check to the data base
        gDBTrans.SQLStmt = "Select Count(*) as TOTUsers from UserTab"
        If gDBTrans.SQLFetch <= 0 Then
            MsgBox "Initialization Error", vbCritical, gAppName & " - Error"
            End
        End If
        If Val(gDBTrans.Rst("TotUsers")) = 0 Then  'Insert users into the databse
            gDBTrans.BeginTrans
            gDBTrans.SQLStmt = "Insert into NameTab (CustomerID, FirstName, Gender, Reference) values (1,'Administrator',0,0)"
            If Not gDBTrans.SQLExecute Then
                MsgBox "Initialization Error", vbCritical, gAppName & " - Error"
                gDBTrans.RollBack
                End
            End If
            gDBTrans.SQLStmt = "Insert into UserTab(UserID,CustomerID,LoginName,Password,Permissions) values (1,1,'admin','admin',14)"
            If Not gDBTrans.SQLExecute Then
                MsgBox "Initialization Error", vbCritical, gAppName & " - Error"
                gDBTrans.RollBack
                End
            End If
            gDBTrans.CommitTrans
        End If
        
        'Now get whether Firm is bank or Society
        gBank = False
        gDBTrans.SQLStmt = "Select * from Install where KeyData = " & AddQuotes("Firm", True)
        If gDBTrans.SQLFetch > 0 Then
            If StrComp(FormatField(gDBTrans.Rst("ValueData")), "BANK", vbTextCompare) Then
                gBank = True
            End If
        End If
        
        'Now get the Name of the Bank /Society from DataBase
        gDBTrans.SQLStmt = "Select * from Install where KeyData = " & AddQuotes("CompanyName", True)
        If gDBTrans.SQLFetch > 0 Then
            gCompanyName = FormatField(gDBTrans.Rst("ValueData"))
        End If
    End If

'Get the User Up
    If gUser Is Nothing Then
        Set gUser = New clsUsers
    End If
'Call GetServerIndianDate
gStrDate = GetSereverDate
End Sub


Public Function GetOpeningBalanceNew(AsOnIndiandate As String, Optional ReportType As wisReports) As Currency
Dim FromIndianDate As String
Dim ToIndianDate As String
Dim DaysDiff As Long
Dim PrevOPBalance As Currency
Dim Balance As Currency

gDBTrans.SQLStmt = "SELECT Top 1 * FROM OBTab Where " & _
    " OBDate < #" & FormatDate(AsOnIndiandate) & "# AND Module = 0 " & _
    " Order By OBDate Desc"
    
If gDBTrans.SQLFetch > 0 Then
    If WisDateDiff(AsOnIndiandate, FormatField(gDBTrans.Rst("OBDate"))) = -1 Then
        GetOpeningBalanceNew = FormatField(gDBTrans.Rst("OBAmount"))
        Exit Function
    Else
        'If We have not get the OB On requred day we have to get the
        'Ob By adding Receipt & Paymnet Fron the LAst OB date
        PrevOPBalance = FormatField(gDBTrans.Rst("OBAmount"))
        FromIndianDate = FormatDate(DateAdd("d", 1, FormatField(gDBTrans.Rst("OBDate"))))
    End If
Else
    Exit Function
End If

'Get The ToDate
If ReportType = wisDebitCreditStatement Then
    ToIndianDate = FormatDate(DateAdd("d", -1, CDate(FormatDate(AsOnIndiandate))))
Else
    ToIndianDate = AsOnIndiandate
End If
    
Dim Count As Integer
Dim LoopCount As Integer
Dim Debits As Currency
Dim Credits As Currency
Count = 1
'MemAcc
Dim Class As Object
Dim clsMMAcc As New clsMMAcc
Set Class = New clsMMAcc
With Class
    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate)
    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate)
    Debits = Debits + .ShareFee(FromIndianDate, ToIndianDate)
    Debits = Debits + .MembershipFee(FromIndianDate, ToIndianDate)
    Credits = Credits + .Loss(FromIndianDate, ToIndianDate)
End With
Set Class = Nothing
     
'SBAccount
Set Class = New clsSBAcc
With Class
    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate)
    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate)
    Debits = Debits + .Profit(FromIndianDate, ToIndianDate)
End With
Set Class = Nothing
  'Frmcancel
DoEvents
If gCancel Then Exit Function

'CAAccount
Set Class = New clsCAAcc
With Class
    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate)
    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate)
    Debits = Debits + .Profit(FromIndianDate, ToIndianDate)
    Credits = Credits + .Loss(FromIndianDate, ToIndianDate)
End With
Set Class = Nothing
  
'FD Account
Set Class = New clsFDAcc
With Class
    Debits = Debits + .DepositDeposits(FromIndianDate, ToIndianDate)
    Credits = Credits + .WithDrawlDeposits(FromIndianDate, ToIndianDate)
    Debits = Debits + .ProfitDeposits(FromIndianDate, ToIndianDate)
    Credits = Credits + .LossDeposits(FromIndianDate, ToIndianDate)
    Credits = Credits + .ContraRPLossDeposits(FromIndianDate, ToIndianDate)
    Debits = Debits + .DepositLoans(FromIndianDate, ToIndianDate)
    Credits = Credits + .WithDrawlLoans(FromIndianDate, ToIndianDate)
    Debits = Debits + .ProfitLoans(FromIndianDate, ToIndianDate)
    Credits = Credits + .LossLoans(FromIndianDate, ToIndianDate)
End With
Set Class = Nothing
  
'RD Account
Set Class = New clsRDAcc
With Class
    Debits = Debits + .DepositDeposits(FromIndianDate, ToIndianDate)
    Credits = Credits + .WithDrawlDeposits(FromIndianDate, ToIndianDate)
    Debits = Debits + .ProfitDeposits(FromIndianDate, ToIndianDate)
    Credits = Credits + .LossDeposits(FromIndianDate, ToIndianDate)
    Debits = Debits + .DepositLoans(FromIndianDate, ToIndianDate)
    Credits = Credits + .WithDrawlLoans(FromIndianDate, ToIndianDate)
    Debits = Debits + .ProfitLoans(FromIndianDate, ToIndianDate)
    Credits = Credits + .LossLoans(FromIndianDate, ToIndianDate)
End With
Set Class = Nothing
  'Frmcancel
DoEvents
If gCancel Then Exit Function

'PD Account
Dim UtilClass As New clsUtils
Set UtilClass = New clsUtils
Debits = Debits + UtilClass.Deposits(FormatDate(FromIndianDate), FormatDate(ToIndianDate), wis_PDLoan)
Credits = Credits + UtilClass.WithDrawals(FormatDate(FromIndianDate), FormatDate(ToIndianDate), wis_PDLoan)
Debits = Debits + UtilClass.Deposits(FormatDate(FromIndianDate), FormatDate(ToIndianDate), wis_PD)
Credits = Credits + UtilClass.WithDrawals(FormatDate(FromIndianDate), FormatDate(ToIndianDate), wis_PD)
Debits = Debits + UtilClass.Profit(FromIndianDate, ToIndianDate, wis_PD)
Credits = Credits + UtilClass.Loss(FromIndianDate, ToIndianDate, wis_PD)
Debits = Debits + UtilClass.Profit(FromIndianDate, ToIndianDate, wis_PDLoan)
Credits = Credits + UtilClass.Loss(FromIndianDate, ToIndianDate, wis_PDLoan)
Set UtilClass = Nothing

'Dl Account
Set Class = New clsDLAcc
With Class
    Debits = Debits + .DepositDeposits(FromIndianDate, ToIndianDate)
    Credits = Credits + .WithDrawlDeposits(FromIndianDate, ToIndianDate)
    Debits = Debits + .ProfitDeposits(FromIndianDate, ToIndianDate)
    Credits = Credits + .LossDeposits(FromIndianDate, ToIndianDate)
    Debits = Debits + .DepositLoans(FromIndianDate, ToIndianDate)
    Credits = Credits + .WithDrawlLoans(FromIndianDate, ToIndianDate)
    Debits = Debits + .ProfitLoans(FromIndianDate, ToIndianDate)
    Credits = Credits + .LossLoans(FromIndianDate, ToIndianDate)
End With
Set Class = Nothing
  
' Loan Account
Dim NameStr() As String
Dim SchemeId() As Long
Dim IdNo() As Long
Set Class = New clsLoan
With Class
    Debits = Debits + .BKCCLoanRepayments(FromIndianDate, ToIndianDate)
    Credits = Credits + .BKCCLoanWithDrawals(FromIndianDate, ToIndianDate)
    Debits = Debits + .BKCCLoanProfit(FromIndianDate, ToIndianDate)
    Debits = Debits + .BKCCLoanProfit(FromIndianDate, ToIndianDate, True)
    Credits = Credits + .BKCCLoanLoss(FromIndianDate, ToIndianDate)
    Debits = Debits + .BKCCDepositRepayments(FromIndianDate, ToIndianDate)
    Credits = Credits + .BKCCDepositWithDrawals(FromIndianDate, ToIndianDate)
    Debits = Debits + .BKCCDepositProfit(FromIndianDate, ToIndianDate)
    Credits = Credits + .BKCCDepositLoss(FromIndianDate, ToIndianDate)
    LoopCount = .LoanList(NameStr, SchemeId, wisAgriculural)
    If LoopCount > 0 Then
          For LoopCount = LBound(SchemeId) To UBound(SchemeId)
                    Debits = Debits + .LoanRepayments(FromIndianDate, ToIndianDate, SchemeId(LoopCount))
                    Credits = Credits + .LoanWithDrawls(FromIndianDate, ToIndianDate, SchemeId(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, SchemeId(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, SchemeId(LoopCount), True)
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, SchemeId(LoopCount))
           Next LoopCount
    End If
End With
Set Class = Nothing
'Fire the Loanname Array
ReDim NameStr(0)
ReDim IdNo(0)

Set Class = New clsLoan
With Class
    LoopCount = .LoanList(NameStr, SchemeId, wisNonAgriculural)
    If LoopCount > 0 Then
          For LoopCount = LBound(SchemeId) To UBound(SchemeId)
                    Debits = Debits + .LoanRepayments(FromIndianDate, ToIndianDate, SchemeId(LoopCount))
                    Credits = Credits + .LoanWithDrawls(FromIndianDate, ToIndianDate, SchemeId(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, SchemeId(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, SchemeId(LoopCount), True)
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, SchemeId(LoopCount))
           Next LoopCount
    End If
End With
Set Class = Nothing
'Fire the Loanname Array
ReDim NameStr(0)
ReDim IdNo(0)
    
'Now Looad the transaCtion with Other Banks
' And Bank's Income & expenses
'BankkAcc
    Set Class = New clsBankAcc
    With Class
        LoopCount = .Heads_ShareCapital(NameStr, IdNo)    'Share Capital
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
    
        LoopCount = .Heads_BankAccounts(NameStr, IdNo)    'BANK Accounts
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
    
        LoopCount = .Heads_Asset(NameStr, IdNo)    'BANK Assets
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
    
        LoopCount = .Heads_Investments(NameStr, IdNo)    'INvestmnets
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
        
        LoopCount = .Heads_ReserveFund(NameStr, IdNo)    'INvestmnets
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
    
        LoopCount = .Heads_Payments(NameStr, IdNo)    'Payments '(kodaathakkavugaLu)
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
            
        LoopCount = .Heads_Repayments(NameStr, IdNo)    'Payments '(kodaathakkavugaLu)
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
            
        LoopCount = .Heads_BankLoan(NameStr, IdNo)    'Bank Loan Accounts
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
        
        LoopCount = .Heads_GovtLoanSubsidy(NameStr, IdNo)    'Govt Loan Subsidery
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
        
        LoopCount = .Heads_MemberDeposits(NameStr, IdNo)    'Member Deposits
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Credits = Credits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
    
        LoopCount = .Heads_Advance(NameStr, IdNo)    'Advances
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
            
        '"NOw Trading Eapenses & INcome
        LoopCount = .Heads_TradingExpense(NameStr, IdNo)    'Trading Expenses
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
        
        LoopCount = .Heads_TradingIncome(NameStr, IdNo)    'Trading Income
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
    
        LoopCount = .Heads_Expense(NameStr, IdNo)    'Trading Expenses
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
        
        LoopCount = .Heads_Income(NameStr, IdNo)    'Trading Income
            For LoopCount = LBound(IdNo) To UBound(IdNo)
                    Debits = Debits + .Deposits(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .WithDrawls(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Debits = Debits + .Profit(FromIndianDate, ToIndianDate, IdNo(LoopCount))
                    Credits = Credits + .Loss(FromIndianDate, ToIndianDate, IdNo(LoopCount))
            Next LoopCount
    
    End With
    
    Set Class = Nothing
    'Free the Loanname Array
    ReDim NameStr(0)
    ReDim IdNo(0)
 'Frmcancel
      DoEvents
      If gCancel Then Exit Function
   
'Material Account Details
    Dim MatID() As Long
    Set Class = New clsUtils
    With Class
            Debits = Debits + .Deposits(FormatDate(FromIndianDate), FormatDate(ToIndianDate), wis_Stock)
            Credits = Credits + .WithDrawals(FormatDate(FromIndianDate), FormatDate(ToIndianDate), wis_Stock)
    End With
'''    Class.Close
    Set Class = Nothing
    'Frmcancel
      DoEvents
      If gCancel Then Exit Function

    ReDim NameStr(0)
    ReDim IdNo(0)
    
    '***
    GetOpeningBalanceNew = PrevOPBalance + Debits - Credits

End Function

Public Function GetTradingOpeningBalance(AsOnIndiandate As String) As Currency

Dim PrevOb As Currency
Dim FromIndianDate As String
Dim ToIndianDate As String
Dim BalanceIndianDate As String
Dim DaysDiff As Long
Dim OPBalance As Currency

    'Get TheRecent Ob w.r.t. date
Dim Module As wisModules
Module = wis_MatAcc

gDBTrans.SQLStmt = "SELECT Top 1 * FROM OBTab WHERE " & _
            " OBDate <= #" & FormatDate(AsOnIndiandate) & "# " & _
            " AND Module = " & Module & " ORDER BY TransDate Desc"

If gDBTrans.SQLFetch > 0 Then
    BalanceIndianDate = FormatField(gDBTrans.Rst("OBDate"))
    PrevOb = FormatField(gDBTrans.Rst("OBAmount"))
    If WisDateDiff(AsOnIndiandate, BalanceIndianDate) = 0 Then
        GetTradingOpeningBalance = PrevOb
        Exit Function
    End If
    GoTo NextAction
End If


gDBTrans.SQLStmt = "Select * From Install where KeyData = 'BalanceSheetDate'"

If gDBTrans.SQLFetch < 1 Then Exit Function
    BalanceIndianDate = FormatField(gDBTrans.Rst("ValueData"))
    If WisDateDiff(BalanceIndianDate, AsOnIndiandate) < 0 Then Exit Function
    
    'Frmcancel
      DoEvents
      If gCancel Then Exit Function
    
    gDBTrans.SQLStmt = "Select * from Install Where KeyData = 'TradingOpeningBalance'"
    If gDBTrans.SQLFetch < 1 Then
        PrevOb = 0
    Else
        PrevOb = CCur(FormatField(gDBTrans.Rst("ValueData")))
    End If

NextAction:

    Dim Class As Object
    Dim strName() As String
    Dim MatID() As Long
    Dim Debits As Currency
    Dim Credits As Currency
    Dim Count As Integer
    
'Get The Date
    ToIndianDate = FormatDate(DateAdd("d", -1, CDate(AsOnIndiandate)))
    Set Class = New clsMatAcc
    Call Class.MaterialList(strName, MatID)
    For Count = LBound(MatID) To UBound(MatID)
        Debits = Debits + Class.Sales_Cash(FromIndianDate, ToIndianDate, MatID(Count))
        Credits = Credits + Class.Purchase_Cash(FromIndianDate, ToIndianDate, MatID(Count))
    Next Count
    Set Class = Nothing
    
    Set Class = New clsBankAcc
    Dim Bk As clsBankAcc
    
    Call Class.Heads_TradingExpense(strName, MatID)
    For Count = LBound(MatID) To UBound(MatID)
        Debits = Debits + Class.Profit(FromIndianDate, ToIndianDate, MatID(Count))
        Credits = Credits + Class.Loss(FromIndianDate, ToIndianDate, MatID(Count))
    Next Count
        
    GetTradingOpeningBalance = PrevOb + Debits - Credits

End Function

Public Sub Main()
'frmSplash.Show vbModal
'Call PauseApplication(2)
Call Initialize
Call KannadaInitialize
Load wisMain
'Now unload the splash form
'Unload frmSplash

wisMain.Show

'Ask Login
    On Error GoTo LOGIN_ERROR
    Set gCurrUser = New clsUsers
    gCurrUser.MaxRetries = 3
    gCurrUser.CancelError = True
    gCurrUser.ShowLoginDialog
    If Not gCurrUser.LoginStatus Then
        GoTo LOGIN_ERROR
    End If
    gUserID = gCurrUser.UserID
    Dim Perm As wis_Permissions
    Perm = wisFullPermissions
'    If gUser.UserPermissions = Perm Then

    Call TransferBKCC

        wisMain.Refresh
        With frmCancel
            gCancel = False
            .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
            .lblMessage = "Confirming the password"
            .cmdCancel.Visible = False
            .Show
            .Refresh
        End With
'''        Load frmStartOFDay
        Unload frmCancel
'''        If frmStartOFDay.fShowTheForm Then
'''            frmStartOFDay.Show vbModal
'''        End If
        
'    End If
    
Exit Sub

LOGIN_ERROR:
    'MsgBox gAppName & " could not log you on !", vbExclamation, gAppName & " - Error"
    MsgBox gAppName & " " & LoadResString(gLangOffSet + 751), vbExclamation, gAppName & " - Error"
    'Unload wisMain
    On Error Resume Next
    gDBTrans.CloseDB
Set gDBTrans = Nothing
Set gUser = Nothing
'code added
If gLangOffSet = wis_KannadaOffset Then
    AppActivate "Akruti Engine 1998", False
    SendKeys "%{F4}Y", True ' i & "{+}", false
End If

    Set wisMain = Nothing
    End
    

End Sub

Sub PauseApplication(Secs As Integer)
Dim PauseTime, Start, Finish, TotalTime
        PauseTime = Secs   ' Set duration.
        Start = Timer   ' Set start time.
        Do While Timer < Start + PauseTime
            DoEvents    ' Yield to other processes.
        Loop
        Finish = Timer  ' Set end time.
        TotalTime = Finish - Start  ' Calculate total time.
End Sub

Public Function StoreBalances(Module As Integer, BalanceDate As String, Amount As Currency, Optional ReportType As wisReports) As Boolean
If ReportType = wisBalanceSheet Then
    gDBTrans.SQLStmt = "SELECT * FROM OBTab WHERE Module = " & Module & _
        " AND OBDate = #" & BalanceDate & "#"
Else
    gDBTrans.SQLStmt = "SELECT * FROM OBTab WHERE Module = " & Module & _
        " AND OBDate = #" & BalanceDate & "#"
End If
If gDBTrans.SQLFetch < 1 Then
    gDBTrans.SQLStmt = "INSERT INTO OBTab (OBDate, OBAmount, Module) VALUES " & _
        "(#" & BalanceDate & "#, " & Amount & ", " & Module & ");"
Else
    gDBTrans.SQLStmt = "UPDATE OBTab SET OBAmount = " & Amount & _
        " WHERE OBDate = #" & BalanceDate & "# AND Module = " & Module
End If
gDBTrans.BeginTrans
If Not gDBTrans.SQLExecute Then
    gDBTrans.RollBack
    Exit Function
Else
    gDBTrans.CommitTrans
End If
StoreBalances = True
End Function



Public Function OBOfAccount(Module As Integer, OBDate As String, Optional ReportType As wisReports) As Currency
If ReportType = wisBalanceSheet Then
    gDBTrans.SQLStmt = "SELECT TOP 1 * FROM OBTab WHERE OBDate < #" & _
        OBDate & "# AND Module = " & Module & _
        " ORDER BY OBDate DESC;"
Else
    gDBTrans.SQLStmt = "SELECT TOP 1 * FROM OBTab WHERE OBDate <= #" & _
        OBDate & "# AND Module = " & Module & _
        " ORDER BY OBDate DESC;"
End If
If gDBTrans.SQLFetch < 1 Then
    If Module = wis_FromProfit Then
        OBOfAccount = -1
    End If
    Exit Function
End If
OBOfAccount = FormatField(gDBTrans.Rst("OBAmount"))

End Function



Private Function TransferBKCC() As Boolean
Dim SchemeId As Long
Dim NewIndex As Index
gDBTrans.SQLStmt = "SELECT * FROM Install WHERE KeyData = 'BKCCID'"
If gDBTrans.SQLFetch > 0 Then
    SchemeId = FormatField(gDBTrans.Rst("ValueData"))
    gDBTrans.BeginTrans
    gDBTrans.SQLStmt = "CREATE UNIQUE INDEX LoanIDTransID " _
        & "ON BKCCTrans (LoanID,TransID) WITH IGNORE NULL;"
    If Not gDBTrans.SQLExecute Then
        MsgBox "Error in Creating Index In BKCC Transaction Table", vbInformation, "Creating Index..."
        gDBTrans.RollBack
        Exit Function
    End If
    gDBTrans.SQLStmt = "INSERT INTO BKCCTrans SELECT A.* FROM LoanTrans A, " & _
        "LoanMaster B WHERE B.SchemeID = " & SchemeId & " AND A.LoanID = B.LoanID"
    If Not gDBTrans.SQLExecute Then
        MsgBox "Error in Transferring BKCC Accounts", vbInformation, "BKCC Transfer"
        gDBTrans.RollBack
        Exit Function
    End If
    gDBTrans.SQLStmt = "DELETE A.* FROM LoanTrans A, " & _
        "LoanMaster B WHERE B.SchemeID = " & SchemeId & " AND A.LoanID = B.LoanID"
    If Not gDBTrans.SQLExecute Then
        MsgBox "Error in deleting BKCC Accounts in OLD Loan Table", vbInformation, "Deleting BKCC Accounts ..."
        gDBTrans.RollBack
        Exit Function
    End If
    gDBTrans.CommitTrans
    
End If

End Function


