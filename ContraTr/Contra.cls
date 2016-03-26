VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_frmTrans As frmContra
Attribute m_frmTrans.VB_VarHelpID = -1

Private m_FromAccType As Long
Private m_FromAccID As Long
Private m_fromAccNum As String
Private m_InitLoad As Boolean

Private m_NoOfAccounts
Private m_GridRow As Integer
Private m_FromAmount As Currency
Private m_ToAmount As Currency 'new

Private m_Amount() As Currency
'Private m_AccType() As wisModules
Private m_AccHeadId() As Long
Private m_AccID() As Long
'Private m_AccNum() As String
Private m_AmountType() As wis_AmountType
Private m_TransDate As Date
Private m_VoucherNo As String

Private m_Status As wis_FunctionReturned


'
Public Sub Show()
Dim StartTfr As Boolean

If gCurrUser Is Nothing Then
    Set gCurrUser = New clsUsers
Else
    'Now Check the permission
    'whether logged inperson is having permission of sb account or not
    Dim Perms As wis_Permissions
    Perms = gCurrUser.UserPermissions
    'If Perms And perBankAdmin Then StartTfr = True
    If (Perms And perClerk) Or gCurrUser.IsAdmin Then StartTfr = True
    If StartTfr Then
        If m_frmTrans Is Nothing Then Set m_frmTrans = New frmContra
        Load m_frmTrans
        gWindowHandle = m_frmTrans.hwnd
        If m_FromAmount Then
            m_InitLoad = True
            Dim cancel As Integer
            'Call m_frmTrans.InitialiseValue(m_TransDate, m_VoucherNo, m_FromAccType, m_fromAccNum, m_FromAmount, Cancel)
        End If
        m_frmTrans.Show vbModal
        gWindowHandle = 0
    Else
        MsgBox "You have no permission to operate this account", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
End If

Dim Amount As Currency
'    Amount = Me.ContraDeposits("2/4/2002", "6/4/2002")
End Sub

Public Property Get ObjectName()
    ObjectName = "ContraClass"
End Property


Public Function SaveDetails() As Boolean

On Error GoTo Exit_Line

Dim PrincAmount As Currency
Dim Amount As Currency
Dim IntAmount As Currency
Dim PenalAmount As Currency
Dim BalanceAmount As Currency

Dim ModuleID As wisModules
Dim FromModuleID As wisModules
Dim ParentID As Long
Dim ContraID As Long

Dim InTrans As Boolean
Dim count As Integer
Dim MaxCount As Integer
Dim VoucherNo As String
Dim TransDate As Date

Dim NoOfWithDraws As Integer
Dim WithDrawAccNo As Byte
Dim DepositAccNo As Byte
Dim ClsObject As Object
Dim FromHeadID As Long
Dim ToHeadID As Long
Dim FromParentID As Long
Dim ToParentID As Long

Dim exteranlAccount As Boolean
Dim internalAccount As Boolean


'First withdraw the amount from the account
BalanceAmount = m_FromAmount

TransDate = m_TransDate 'formatdate(m_c
TransDate = GetSysFormatDate(m_frmTrans.txtTransDate)
If Not m_InitLoad Then
    gDbTrans.BeginTrans
    InTrans = True
End If

Dim rstTemp As ADODB.Recordset
Dim bankClass As clsBankAcc
Set bankClass = New clsBankAcc

MaxCount = UBound(m_AccID)

ModuleID = GetModuleIDFromHeadID(m_AccHeadId(0))
ParentID = GetParentID(m_AccHeadId(0))
If ModuleID > 100 Then ModuleID = ModuleID - (ModuleID Mod 100)
'Assign from head values
FromModuleID = ModuleID
FromHeadID = m_AccHeadId(0)
FromParentID = ParentID


If Not m_InitLoad Then
    'Get the Transcation Id for the contra table
    ContraID = GetMaxContraTransID + 1

    m_NoOfAccounts = 1: NoOfWithDraws = 1
    If ModuleID = 0 And (ParentID = parBankAccount Or ParentID = parBankLoanAccount) Then
        exteranlAccount = True
        'Here we need not do the transction for the sub heads of account
        'Insert the Details in the Contra Table
        
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
                    "(ContraID,TransID, " & _
                    " AccHeadID,AccId," & _
                    "TransType,Amount,UserID) " & _
                    " VALUES (" & ContraID & "," & 0 & "," & _
                    m_AccHeadId(0) & "," & 0 & "," & _
                    wContraWithdraw & "," & m_Amount(0) & "," & gCurrUser.UserID & ")"
        
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
        
    ElseIf ModuleID = wis_SBAcc Or ModuleID = wis_CAAcc Or ModuleID = wis_SuspAcc Then
        internalAccount = True
        If ModuleID = wis_CAAcc Then Set ClsObject = New ClsCAAcc
        If ModuleID = wis_SBAcc Then Set ClsObject = New clsSBAcc
        If ModuleID = wis_SuspAcc Then
            Set ClsObject = New clsSuspAcc
            'Dim Cls As clsSuspAcc
            Call ClsObject.WithdrawAmount(m_AccHeadId(0), m_AccID(0), 0, "", _
                TransDate, m_Amount(0), , 0, VoucherNo, False, ContraID)
                'TransDate, m_Amount(0), , m_AccID(0), VoucherNo, False, ContraID)
        Else
            If ClsObject.WithdrawAmount(CInt(m_AccID(0)), m_Amount(0), _
                "Tfr to ", TransDate, VoucherNo, False, ContraID) = 0 Then GoTo Exit_Line
            Set ClsObject = Nothing
        End If
    End If
Else
    'First check whether all withdrawls and deposits are
    'in sorting order or not
    Call SortAccountHeads

    NoOfWithDraws = m_NoOfAccounts
    For count = 0 To m_NoOfAccounts - 1
        If m_Amount(count) < 0 Then
            NoOfWithDraws = NoOfWithDraws - 1
            Debug.Assert 100 < 99
        End If
    Next

    If m_NoOfAccounts > 1 Then
    'If NoOfWithDraws > 1 Then
        'IF there is more than one account where the amount withdrwan
        'then it is difficult to mantian the contra transaction
        'In such case withdraw all the withdrawal amount from first account
        'then later compansate the first of withdrawl and deposited account
        For count = 1 To m_NoOfAccounts - 1
        'For Count = 1 To NoOfWithDraws - 1
            If m_Amount(count) > 0 Then
                If Not bankClass.UpdateContraTrans(m_AccHeadId(count), _
                        m_AccHeadId(m_NoOfAccounts), m_Amount(count), _
                                            TransDate) Then GoTo Exit_Line
                Amount = Amount + m_Amount(count)
            Else
                If Not bankClass.UpdateContraTrans(m_AccHeadId(0), _
                        m_AccHeadId(count), Abs(m_Amount(count)), _
                                            TransDate) Then GoTo Exit_Line
            End If
            
        Next
    End If
End If

Dim I As Integer
Dim tmpModuleId As wisModules
Dim TransferAmount As Currency

'Assign the Module ID if We have transferred from any module ID
FromModuleID = ModuleID
'For Count = NoOfWithDraws To MaxCount 'm_NoOfAccounts To MaxCount
For count = m_NoOfAccounts To MaxCount
    If m_Amount(count) <= 0 Then GoTo NextAmount
    'If m_Amount(COunt) < 0 Then TransType = wContraWithdraw
    
    IntAmount = 0: PenalAmount = 0
    PrincAmount = 0
    
    'Now search whether he is Paying any Interest Amount
    'on this loan account
    'or He is making the two transcation to any account
    I = count
    ModuleID = GetModuleIDFromHeadID(m_AccHeadId(count))
    ParentID = GetParentID(m_AccHeadId(count))
    Do
        If I > MaxCount Then Exit Do
        tmpModuleId = GetModuleIDFromHeadID(m_AccHeadId(I))
        If m_AccID(count) = m_AccID(I) And ModuleID = tmpModuleId Then
            ToHeadID = m_AccHeadId(I)
            TransferAmount = m_Amount(I)
            If m_AmountType(I) = wisPrincipal And PrincAmount = 0 Then
                If PrincAmount Then GoTo Exit_Line
                PrincAmount = m_Amount(I)
                'Call BankClass.UpdateContraTrans(m_AccHeadId(0), m_AccHeadId(I), PrincAmount, TransDate)
                m_Amount(I) = 0
            ElseIf m_AmountType(I) = wisRegularInt And IntAmount = 0 Then
                If IntAmount Then GoTo Exit_Line
                IntAmount = m_Amount(I)
                'Call BankClass.UpdateContraTrans(m_AccHeadId(0), m_AccHeadId(I), IntAmount, TransDate)
                m_Amount(I) = 0
            ElseIf m_AmountType(I) = wisPenalInt And PenalAmount = 0 Then
                If PenalAmount Then GoTo Exit_Line
                PenalAmount = m_Amount(I)
                'Call BankClass.UpdateContraTrans(m_AccHeadId(0), m_AccHeadId(I), PenalAmount, TransDate)
                m_Amount(I) = 0
            Else
                TransferAmount = 0  '' new Line of code just not execut UpdateContraTrans
            End If
            'Call BankClass.UpdateContraTrans(m_AccHeadId(0), m_AccHeadId(I), TransferAmount, TransDate)
            If FromParentID = parBankAccount Or FromParentID = parBankLoanAccount Then
                ''IN case of Banks account in Other banks the Cash in hand should increase,
                ' So we reversing the accounts head id
                bankClass.BankAccountToCustomerAccount = True
                Dim strDesc As String
                strDesc = GetResourceString(307) '& " " & GetResourceString(342) & " " & _
                    ToHeadID & " " & GetResourceString(108) & " " & FromHeadID
                
                    ''Set Up the Transaction class Directly
                    ' So that We can have saperate TransID/Record for Contra Transaction
                    'between External bank/loan account and Internal/loan account
                    Dim contraAccTransClass As clsAccTrans
                    Set contraAccTransClass = New clsAccTrans
                    Call contraAccTransClass.AllTransHeadsAdd(ToHeadID, TransferAmount, 0)
                    Call contraAccTransClass.AllTransHeadsAdd(FromHeadID, 0, TransferAmount)
                    If contraAccTransClass.SaveVouchers(Journal, TransDate, strDesc) <> Success Then GoTo Exit_Line
                    'If Not BankClass.UpdateContraTrans(ToHeadID, FromHeadID, TransferAmount, TransDate) Then GoTo Exit_line
                
            Else
                If Not bankClass.UpdateContraTrans(FromHeadID, ToHeadID, TransferAmount, TransDate) Then GoTo Exit_Line
            End If
        End If
        I = I + 1
    Loop

    
    If ModuleID > 100 Then ModuleID = ModuleID - (ModuleID Mod 100)
    
    If ModuleID = 0 And (ParentID = parBankAccount Or ParentID = parBankLoanAccount) Then
        'Here we need not do the transction for the sub heads of account
        'Insert the Details in the Contra Table
        If ContraID Then
            ''Delete the Contra Transaction as both are similar Outside the bank accounts
            If FromParentID = ParentID Then
                MsgBox GetResourceString(535), vbInformation, wis_MESSAGE_TITLE
                GoTo Exit_Line
            End If
        Else
            ContraID = GetMaxContraTransID + 1
        End If
        'ContraID = GetMaxContraTransID + 1
        gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
                    "(ContraID,TransID, " & _
                    " AccHeadID,AccId," & _
                    "TransType,Amount,UserID) " & _
                    " VALUES (" & ContraID & "," & 0 & "," & _
                    m_AccHeadId(count) & "," & 0 & "," & _
                    wContraDeposit & "," & PrincAmount & "," & gCurrUser.UserID & ")"
        
        If Not gDbTrans.SQLExecute Then GoTo Exit_Line
        ContraID = 0
        
        BalanceAmount = BalanceAmount - PrincAmount
        
    ElseIf ModuleID = wis_CAAcc Or ModuleID = wis_SBAcc Or _
                    ModuleID = wis_PDAcc Or ModuleID = wis_BKCC Or _
                    ModuleID = wis_Deposits Or ModuleID = wis_RDAcc Then
        internalAccount = True
        If ModuleID = wis_BKCC Then Set ClsObject = New clsBkcc
        If ModuleID = wis_SBAcc Then Set ClsObject = New clsSBAcc
        If ModuleID = wis_Deposits Then Set ClsObject = New clsFDAcc
        If ModuleID = wis_PDAcc Then Set ClsObject = New clsPDAcc
        If ModuleID = wis_CAAcc Then Set ClsObject = New ClsCAAcc
        If ModuleID = wis_RDAcc Then Set ClsObject = New clsRDAcc
        
        BalanceAmount = BalanceAmount - PrincAmount
        
        If IntAmount Or PenalAmount Then GoTo Exit_Line
        
        If FromParentID = parBankAccount Or FromParentID = parBankLoanAccount Then
                If ClsObject.WithdrawAmount(m_AccID(count), PrincAmount, _
                         "Tfr From ", TransDate, m_VoucherNo, False, ContraID) = 0 Then GoTo Exit_Line
        Else
            If ModuleID = wis_BKCC Then
                If ClsObject.DepositAmount(m_AccID(count), PrincAmount, 0, 0, _
                         "Tfr From ", TransDate, m_VoucherNo, False, ContraID) = 0 Then GoTo Exit_Line
            Else
                If ClsObject.DepositAmount(m_AccID(count), PrincAmount, _
                         "Tfr From ", TransDate, m_VoucherNo, False, ContraID) = 0 Then GoTo Exit_Line
            End If
        End If
    ElseIf ModuleID = wis_DepositLoans Or _
            ModuleID = wis_BKCC Or ModuleID = wis_BKCCLoan Or ModuleID = wis_Loans Or _
            ModuleID = wis_Members Then
        
        If ModuleID > 100 Then ModuleID = ModuleID - (ModuleID Mod 100)
        
        internalAccount = True
        
        If ModuleID = wis_DepositLoans Then
            If PenalAmount Then GoTo Exit_Line
            Set ClsObject = New clsDepLoan
            If ClsObject.DepositAmount(m_AccID(count), PrincAmount, _
                IntAmount, "Tfr From ", TransDate, m_VoucherNo) = 0 Then GoTo Exit_Line
            BalanceAmount = BalanceAmount - PrincAmount - IntAmount
        
        ElseIf (ModuleID = wis_Loans And ModuleID < wis_Loans + 100) Or ModuleID = wis_BKCCLoan Or ModuleID = wis_BKCC Then
        
            If ModuleID = wis_Loans Then Set ClsObject = New clsLoan
            If ModuleID = wis_BKCC Or ModuleID = wis_BKCCLoan Then _
                                    Set ClsObject = New clsBkcc
            
            
            If FromParentID = parBankAccount Or FromParentID = parBankLoanAccount Then
                If ClsObject.WithdrawAmount(m_AccID(count), PrincAmount, _
                         "Tfr From ", TransDate, m_VoucherNo, False, ContraID) = 0 Then GoTo Exit_Line
            Else
                If ClsObject.DepositAmount(CLng(m_AccID(count)), PrincAmount, _
                    IntAmount, PenalAmount, "Tfr From ", TransDate, m_VoucherNo, False, ContraID) = 0 Then GoTo Exit_Line
            End If
            
            BalanceAmount = BalanceAmount - PrincAmount - IntAmount - PenalAmount
        
        ElseIf ModuleID = wis_Members Then
            Set ClsObject = New clsMMAcc
            If ClsObject.DepositAmount(CLng(m_AccID(count)), PrincAmount, _
                        IntAmount, "Tfr From ", TransDate, m_VoucherNo, False, ContraID) = 0 Then
                GoTo Exit_Line
            End If
            BalanceAmount = BalanceAmount - PrincAmount - IntAmount - PenalAmount
        End If
    
    ElseIf ModuleID = wis_SuspAcc Then
        Set ClsObject = New clsSuspAcc
        
        If ClsObject.DepositAmount(m_AccHeadId(0), m_AccID(I), 0, "", m_TransDate, _
                    PrincAmount, 0, VoucherNo, False, ContraID) Then GoTo Exit_Line
        
    End If
    
    Set ClsObject = Nothing
    'Make Transction in Acc Trans
    Debug.Assert MaxCount = 0
    
NextAmount:

Next

If m_NoOfAccounts > 1 Then
    'IF theres is more than one account where the amount withdrwan
    'then it is difficult to Manatian the Contra transaction
    'In such case withdrwa all the withdraw amount from first account
    'then later compansate the first of withdrawl and deposited account
    If Amount > 0 Then
    If Not bankClass.UndoContraTrans(m_AccHeadId(0), _
            m_AccHeadId(m_NoOfAccounts), Amount, TransDate) Then GoTo Exit_Line
    End If
End If

Debug.Assert BalanceAmount = 0
If BalanceAmount = 0 Then
    If Not m_InitLoad Then gDbTrans.CommitTrans: InTrans = False
    SaveDetails = True
End If
'Exit Function

Exit_Line:

If Not bankClass Is Nothing Then bankClass.BankAccountToCustomerAccount = False
If InTrans Then
    gDbTrans.RollBack
    MsgBox "Unable to save the Transaction", vbInformation, wis_MESSAGE_TITLE
'    Err.Clear
End If

If Err Then
    MsgBox "Error In Contra save Details", vbInformation, wis_MESSAGE_TITLE
    'Resume
    Err.Clear
End If

End Function

'Moduel Created : Shashidhar Angadi
'Date           : 28th November 2003
'Argumentds     : Nil
''Descr:        : This Function sorts the account heads
'   where transaction takes place'
'There are head which or with draw side & does the Deposit transaction
'and some accounts are on he deposit side & does the with draw transaction
'in both side such accounts are idenified by the negtive amounts
'so in theis sub
'we will put the negtive amount heads in the last part of the same type
'for if there are 2 with drawing heads & in that ine does the deposit transaction
'in such case we will put that head second in the whle order
Private Sub SortAccountHeads()

If m_NoOfAccounts = 1 Then Exit Sub

Dim count As Integer
Dim MaxCount As Integer
Dim FoundNegative As Boolean
Dim NeedtoRectify As Boolean

MaxCount = m_NoOfAccounts - 1

KeepChecking:

''NeedtoRectify = False
'FoundNegative = False
'
'For Count = 0 To MaxCount
'    If m_Amount(Count) < 0 Then FoundNegative = True
'    If FoundNegative And m_Amount(Count) > 0 Then NeedtoRectify = True: Exit For
'Next
'
'If Not NeedtoRectify Then Exit Sub


'If Count < 1 Then Exit Sub
Dim tmpAmount As Double
Dim tmpHeadId As Long
Dim tmpAccid As Long
Dim tmpAmountType As wis_AmountType

'First Sort the Deposit Accounts
MaxCount = m_NoOfAccounts - 1

Do
    FoundNegative = False
    For count = 1 To MaxCount
        Debug.Assert m_Amount(0) > 0
        If m_Amount(count) > m_Amount(count - 1) Then
            FoundNegative = True
            
            tmpAmount = m_Amount(count)
            tmpHeadId = m_AccHeadId(count)
            tmpAccid = m_AccID(count)
            tmpAmountType = m_AmountType(count)
            
            m_Amount(count) = m_Amount(count - 1)
            m_AccHeadId(count) = m_AccHeadId(count - 1)
            m_AccID(count) = m_AccID(count - 1)
            m_AmountType(count) = m_AmountType(count - 1)
            
            m_Amount(count - 1) = tmpAmount
            m_AccHeadId(count - 1) = tmpHeadId
            m_AccID(count - 1) = tmpAccid
            m_AmountType(count - 1) = tmpAmountType
        End If
        
    Next
    
Loop While FoundNegative

'Now sort the Witht Depsit Accounts
MaxCount = UBound(m_AccID)
Do
    FoundNegative = False
    For count = m_NoOfAccounts + 1 To MaxCount
        Debug.Assert m_Amount(0) > 0
        If m_Amount(count) > m_Amount(count - 1) Then
            FoundNegative = True
            
            tmpAmount = m_Amount(count)
            tmpHeadId = m_AccHeadId(count)
            tmpAccid = m_AccID(count)
            tmpAmountType = m_AmountType(count)
            
            m_Amount(count) = m_Amount(count - 1)
            m_AccHeadId(count) = m_AccHeadId(count - 1)
            m_AccID(count) = m_AccID(count - 1)
            m_AmountType(count) = m_AmountType(count - 1)
            
            m_Amount(count - 1) = tmpAmount
            m_AccHeadId(count - 1) = tmpHeadId
            m_AccID(count - 1) = tmpAccid
            m_AmountType(count - 1) = tmpAmountType
        End If
        
    Next
    
Loop While FoundNegative


End Sub

Public Function UndoTransaction(ByVal ContraTransID As Long, _
                    ByVal TransDate As Date) As wis_FunctionReturned

On Error GoTo Exit_Line

Dim Amount As Currency
Dim BalanceAmount As Currency
Dim InTrans As Boolean
Dim MaxCount As Integer
Dim ParentID As Long

Dim rstTrans As Recordset
 
gDbTrans.SqlStmt = "Select * from ContraTrans " & _
                " Where ContraID = " & ContraTransID & _
                " And Amount <> 0 "

If gDbTrans.Fetch(rstTrans, adOpenDynamic) < 0 Then GoTo Exit_Line

Dim count As Integer
Dim transType As wisTransactionTypes
Dim TotalToAmount As Currency
Dim ModuleID As wisModules

Dim TransID() As Long
ReDim TransID(0)

'Now get the account details where amount is withdrawn
count = 0: MaxCount = 0
m_FromAmount = 0
While Not rstTrans.EOF
    transType = rstTrans("TransType")
    If transType = wContraWithdraw Then
        ReDim Preserve m_Amount(MaxCount)
        ReDim Preserve m_AccID(MaxCount)
        ReDim Preserve m_AccHeadId(MaxCount)
        ReDim Preserve TransID(MaxCount)
        ReDim Preserve m_AmountType(MaxCount)
        
        m_FromAmount = m_FromAmount + rstTrans("Amount")
        m_Amount(MaxCount) = rstTrans("Amount")
        m_AccID(MaxCount) = rstTrans("AccID")
        m_AccHeadId(MaxCount) = rstTrans("AccHeadID")
        TransID(MaxCount) = rstTrans("TransID")
         
'        ParentId = GetParentID(m_AccHeadId(m_AccHeadId(Count)))
'        Select Case ParentId
'            Case parMemDepIntPaid, parMemLoanPenalInt
'
'        End Select
        MaxCount = MaxCount + 1
    End If
    rstTrans.MoveNext
Wend
m_NoOfAccounts = MaxCount

rstTrans.MoveFirst

'Now get the details amount is deposited
count = 0
While Not rstTrans.EOF
    transType = rstTrans("TransType")
    If transType = wContraDeposit Then
        ReDim Preserve m_Amount(MaxCount)
        ReDim Preserve m_AccID(MaxCount)
        ReDim Preserve m_AccHeadId(MaxCount)
        ReDim Preserve TransID(MaxCount)
        ReDim Preserve m_AmountType(MaxCount)
        
        TotalToAmount = TotalToAmount + rstTrans("Amount")
        m_Amount(MaxCount) = rstTrans("Amount")
        m_AccID(MaxCount) = rstTrans("AccID")
        m_AccHeadId(MaxCount) = rstTrans("AccHeadID")
        TransID(MaxCount) = rstTrans("TransID")
        
        MaxCount = MaxCount + 1
    End If
    rstTrans.MoveNext
Wend

'Sort the Amount where it has with drawn
Call SortAccountHeads

'First withdraw the amount from the account
BalanceAmount = m_FromAmount
Dim ClsObject As Object

'TransDate = m_TransDate 'formatdate(m_c
gDbTrans.BeginTrans: InTrans = True

Dim rstTemp As ADODB.Recordset
Dim bankClass As clsBankAcc
Dim tmpModuleId As wisModules

Set bankClass = New clsBankAcc

ModuleID = GetModuleIDFromHeadID(m_AccHeadId(0))
If ModuleID > 100 Then ModuleID = ModuleID - (ModuleID Mod 100)
    
    'IF there is more than one accountHead where the amount has been withdrwan
    'then it is difficult to undo the Contra transaction
    'In such case Undo all the withdrawn amount from first account
    'then later compansate the first of exta amount
    For count = 1 To m_NoOfAccounts - 1
        If m_Amount(count) > 0 Then
            If Not bankClass.UndoContraTrans(m_AccHeadId(count), _
                    m_AccHeadId(m_NoOfAccounts), m_Amount(count), TransDate) Then GoTo Exit_Line
            Amount = Amount + m_Amount(count)
        Else
            'If Not BankClass.UndoContraTrans(m_AccHeadId(0), _
                    m_AccHeadId(Count), Abs(m_Amount(Count)), TransDate) Then GoTo Exit_line
            If Not bankClass.UndoContraTrans(m_AccHeadId(0), _
                    m_AccHeadId(count), Abs(m_Amount(count)), TransDate) Then GoTo Exit_Line
                    
        End If
    Next

Dim I As Integer


'First Undo The Accounts where amount has with drawn
'In Such Case do Not Withdraw any amount from the ledger heads
Dim StartCount As Integer
Dim EndCount As Integer
Dim PrinAmount As Currency
        
Dim DepositTrans  As Boolean
StartCount = 0: EndCount = m_NoOfAccounts - 1
DepositTrans = False
BalanceAmount = 0

Restart:

For count = StartCount To EndCount 'm_NoOfAccounts To MaxCount
    If m_Amount(count) = 0 Then GoTo NextAmount
    PrinAmount = m_Amount(count)
    
    If DepositTrans Then
        'If undo the TransCtion with Both withdraw & Deposit
        'then the amount will be doubled 'so undo the tranction
        'while delteing deposit transaction
        If m_AccHeadId(count) Mod HEAD_OFFSET = 0 Then m_AccHeadId(count) = m_AccID(count)
        If Not bankClass.UndoContraTrans(m_AccHeadId(0), m_AccHeadId(count), _
                m_Amount(count), TransDate) Then GoTo Exit_Line
        BalanceAmount = BalanceAmount - PrinAmount
    Else
        BalanceAmount = BalanceAmount + PrinAmount
    End If
    
    ModuleID = GetModuleIDFromHeadID(m_AccHeadId(count))

     'Now search whether he is Paying any Interest Amount
     'on this loan account
     I = count + 1 'Count
     PrinAmount = 0
     Do
         If I > EndCount Then Exit Do
         
         'If the AccountNo and ModuleId both are same then
         'the both transaction related to same module like (
         ' One transaction is of interest then and otheris of principal
         'in such case  only on undo will take place to that module
         tmpModuleId = GetModuleIDFromHeadID(m_AccHeadId(I))
         If m_AccID(count) = m_AccID(I) And ModuleID = tmpModuleId Then
             PrinAmount = m_Amount(I)
             If DepositTrans Then
                 If Not bankClass.UndoContraTrans(m_AccHeadId(0), _
                     m_AccHeadId(I), PrinAmount, TransDate) Then GoTo Exit_Line
                 BalanceAmount = BalanceAmount - PrinAmount
             Else
                 Debug.Assert m_Amount(I)
                 BalanceAmount = BalanceAmount + PrinAmount
             End If
             m_Amount(I) = 0
         End If
         I = I + 1
     Loop
     
     If ModuleID > 100 Then ModuleID = ModuleID - (ModuleID Mod 100)

     If ModuleID = wis_BKCC Then Set ClsObject = New clsBkcc
     If ModuleID = wis_SBAcc Then Set ClsObject = New clsSBAcc
     If ModuleID = wis_Deposits Then Set ClsObject = New clsFDAcc
     If ModuleID = wis_PDAcc Then Set ClsObject = New clsPDAcc
     If ModuleID = wis_CAAcc Then Set ClsObject = New ClsCAAcc
     If ModuleID = wis_RDAcc Then Set ClsObject = New clsRDAcc
     If ModuleID = wis_SuspAcc Then Set ClsObject = New clsSuspAcc
     If (ModuleID >= wis_BankAccounts And ModuleID < wis_BankAccounts + 100) Then ClsObject = New clsClearing
    
     If ModuleID = wis_SuspAcc Then Set ClsObject = New clsSuspAcc
     If ModuleID = wis_DepositLoans Then Set ClsObject = New clsDepLoan
     If ModuleID = wis_Loans Then Set ClsObject = New clsLoan
     If ModuleID = wis_BKCC Or ModuleID = wis_BKCCLoan Then Set ClsObject = New clsBkcc
     If ModuleID = wis_Members Then Set ClsObject = New clsMMAcc
     
    'If THe transaction is of Bank Account or Bank loan Account
    'then need not do any transaction deletion
    'In case of Bank a/c transction the moduleid will be 0
     If ModuleID Then
        If Not ClsObject.UndoTransaction(m_AccID(count), TransID(count)) Then GoTo Exit_Line
     End If
     Set ClsObject = Nothing
    
    'Make Transction in Acc Trans
    
NextAmount:

Next

If Not DepositTrans Then
    'Now undo the Transaction where amount has deposited
    DepositTrans = True
    StartCount = m_NoOfAccounts: EndCount = MaxCount - 1
    GoTo Restart
End If

If m_NoOfAccounts > 1 Then
    'IF theres is more than one account where the amount withdrwan
    'then it is difficult to Manatian the Contra transaction
    'In such case withdrwa all the withdraw amount from first account
    'then later compansate the first of withdrawa and deposited account
    Debug.Assert Amount = 0
    If Not bankClass.UpdateContraTrans(m_AccHeadId(0), _
                m_AccHeadId(m_NoOfAccounts), Amount, TransDate) Then GoTo Exit_Line
End If

Debug.Assert BalanceAmount = 0
'Now delete the Contra ranns details
gDbTrans.SqlStmt = "Delete * From ContraTrans Where ContraID = " & ContraTransID
If Not gDbTrans.SQLExecute Then GoTo Exit_Line
If BalanceAmount = 0 Then
    gDbTrans.CommitTrans: InTrans = False
    UndoTransaction = Success
Else
    Debug.Assert ContraTransID <> 0
End If
'Exit Function

Exit_Line:

If InTrans Then gDbTrans.RollBack

If Err Then
    MsgBox "Error In Contra undo Transaction ", vbInformation, wis_MESSAGE_TITLE
    'Resume
    Err.Clear
End If

End Function
Property Get Status() As wis_FunctionReturned
    Status = m_Status
End Property

Property Let TransDate(NewValue As Date)
    m_TransDate = NewValue
End Property


Property Let VoucherNo(NewValue As String)
    m_VoucherNo = NewValue
End Property

Private Sub Class_Initialize()

m_TransDate = CDate(gStrDate)
Call Clear
m_VoucherNo = ""

End Sub

Private Sub m_frmTrans_AddClicked()

On Error Resume Next

Dim curRow As Integer
Dim ParentID As Long

'If m_AccHeadId(0) = 0 Or m_Id(0) = 0 Then
If m_AccHeadId(0) = 0 Or m_AccID(0) = 0 Then
curRow = 1
With m_frmTrans
    .grd.Row = 1
    ' check whether he has selected the account type or not
    If .cmbFrom.ListIndex < 0 Then
        'MsgBox "Please select the Account type", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(36, 35, 296), _
            vbInformation, wis_MESSAGE_TITLE 'Accoun type not specified
        .cmbFrom.SetFocus
        Exit Sub
    End If
    
    'Now Get the Parent Id of the Head
    'We enterd this loop if m_AccHeadId(0) = 0
    ' So how the below code gives a value for ParentID
    'So next line gives the error * exists sub
    ParentID = GetParentID(m_AccHeadId(0))
    
    '''So ge the ParentID of the selected Account type
    ParentID = GetParentID(.cmbFrom.ItemData(.cmbFrom.ListIndex))
    'parBankLoanAccount = 40000  parMemberShare = 10100
    'parMemberLoan = 130100        parMemDepLoan = 130200
    
    'If the Account head is Sb or Ca then Only check for the
    'Account Number
    'If Not (ParentID = parBankAccount Or ParentID = parBankLoanAccount) Then
    If ParentID = parMemberShare Or ParentID = parMemberLoan Or _
        ParentID = parMemDepLoan Or ParentID = parSuspAcc Or ParentID = parMemberDeposit Then
        If .txtFromAccNo = "" Then
            MsgBox GetResourceString(36, 60, 296), _
                vbInformation, wis_MESSAGE_TITLE 'Account type not specified
            .txtFromAccNo.SetFocus
            Exit Sub
        End If
        .optRegInt.Enabled = True
        .optPenalInt.Enabled = True
    Else
     .optPrincipal.Value = True
     .optRegInt.Enabled = False
     .optPenalInt.Enabled = False
    End If
    If .txtFromAmount.Value = 0 Then
        MsgBox GetResourceString(40, 296), _
            vbInformation, wis_MESSAGE_TITLE 'Accoun type not specified
        .txtFromAmount.SetFocus
        Exit Sub
    End If
    
On Error GoTo ErrLine:

    'If He has not selected the account from
    'where he is withdrawing the amount
    With .cmbFrom
        m_AccHeadId(0) = .ItemData(.ListIndex)
    End With
    
    If AddToGrid Then
        .grd.Rows = .grd.Rows + 1
        .fraTo.Enabled = True
        .cmbFrom.Enabled = False
        .txtFromAccNo.Enabled = False
        .txtFromAmount.Enabled = False
        .cmdFrom.Enabled = False
    End If

End With
    
    
Exit Sub

End If


'Here he is selecting the Account where He is transferring the amount
'check whether he has selected the account type or not

With m_frmTrans
    If .cmbTo.ListIndex < 0 Then
        'MsgBox "Please select the Account tyep", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(36, 35, 296), _
            vbInformation, wis_MESSAGE_TITLE 'Accoun type not specified
        '.cmbTo.SetFocus
        Exit Sub
    End If
   
    'Now Get the PArent Id of the Head
    ParentID = GetParentID(.cmbTo.ItemData(.cmbTo.ListIndex))
    'If the Account head is Sb or Ca thenOnly check for the
    'Account Number
    If Not (ParentID = parBankAccount Or ParentID = parBankLoanAccount) Then
        If .txtToAccNo = "" Then
            MsgBox GetResourceString(36, 60, 296), _
                vbInformation, wis_MESSAGE_TITLE 'Account type not specified
            .txtToAccNo.SetFocus
            Exit Sub
        End If
    End If
    If .txtToAmount.Value = 0 Then
        MsgBox GetResourceString(40, 296), _
            vbInformation, wis_MESSAGE_TITLE 'Accoun type not specified
        .txtToAmount.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrLine:
    'If He has not selected the account from
    'where he is withdrawing the amount
    If AddToGrid Then
        .grd.Rows = .grd.Rows + 1
        .optPrincipal.Value = True
    End If
End With

Exit Sub

ErrLine:
    MsgBox "Error In AddClicked", vbInformation, wis_MESSAGE_TITLE

End Sub

Private Sub m_frmTrans_ClearClicked()

If Not m_InitLoad Then
    Call Clear
    With m_frmTrans
        'If m_InitLoad Then Call InitialiseValue
        'grd.Clear
        .fraTo.Enabled = False
        .cmbFrom.Enabled = True
        .txtFromAccNo.Enabled = True
        .txtFromAmount.Enabled = True
        .cmdFrom.Enabled = True
        .optRegInt.Enabled = True
        .optPenalInt.Enabled = True
    End With
Else
    Dim I As Integer
    I = IIf(m_NoOfAccounts > 0, m_NoOfAccounts - 1, 0)
    ReDim Preserve m_AccHeadId(I)
    'ReDim Preserve m_Id(0)
    ReDim Preserve m_AccID(I)
    ReDim Preserve m_AmountType(I)
    ReDim Preserve m_Amount(I)
    m_frmTrans.grd.Rows = I + 2
End If


End Sub


Private Function AddToGrid() As Boolean

On Error GoTo ErrLine

AddToGrid = False
Dim GrdRow As Integer
Dim ArrNo As Integer

Dim AccHeadID As Long
Dim ModuleID As wisModules
Dim ParentID As Long
Dim IntHeadID As Long
Dim headName As String

Dim CustName As String
Dim AccNum As String
Dim StrAcctype As String

If m_GridRow = 0 Then m_GridRow = 1
If m_InitLoad Then m_GridRow = m_GridRow + 1

With m_frmTrans
    If m_GridRow = 1 Then
        .grd.Row = 1
        headName = .cmbFrom.Text
        AccHeadID = .cmbFrom.ItemData(.cmbFrom.ListIndex)
        AccNum = Trim(.txtFromAccNo)
        StrAcctype = .cmbFrom.Text
        m_Amount(0) = .txtFromAmount.Value
    Else
        'Get the grid row
        With .grd
            .Col = 0
            If .Rows <= m_NoOfAccounts + .FixedRows And m_InitLoad Then _
                                    .Rows = m_NoOfAccounts + .FixedRows + 1
            For ArrNo = 1 To .Rows - 1
                .Row = ArrNo
                If .Text = "" Then Exit For
            Next
            
        End With
        headName = .cmbTo.Text
        AccHeadID = .cmbTo.ItemData(.cmbTo.ListIndex)
        AccNum = Trim(.txtToAccNo)
        StrAcctype = .cmbTo.Text
    End If
    
    GrdRow = .grd.Row
    ArrNo = .grd.Row - 1
End With

'Now check the whether entered Details are correct are not
'Now Check the account No w.r.t Account tYPE

'Now get the Account if Not specified
If UBound(m_AccID) < ArrNo Then
    ReDim Preserve m_AccID(ArrNo)
    ReDim Preserve m_AccHeadId(ArrNo)
    ReDim Preserve m_AmountType(ArrNo)
    ReDim Preserve m_Amount(ArrNo)
End If

ModuleID = GetModuleIDFromHeadID(AccHeadID)
ParentID = GetParentID(AccHeadID)

'Id we do not have the unique Id of the account type then
'Get the unique id of the account type
Dim RstAcc As Recordset
'First Check for the Bank & bank Loan Account
'If ParentID = parBankAccount Or ParentID = parBankLoanAccount Then
If ModuleID = wis_None Then
    'Here need not to get the account number
    'the head name itself is the account number
ElseIf ModuleID = wis_SuspAcc Then
    With m_frmTrans
        Set RstAcc = GetAccRecordSet(.cmbFrom.ItemData(.cmbFrom.ListIndex), .txtFromAccNo)
    End With
    
    If ArrNo Then
        m_AccID(ArrNo) = m_AccID(0)
    Else
        m_AccID(0) = m_frmTrans.txtFromAccNo
    End If
    
    If RstAcc Is Nothing Then
        AccNum = m_frmTrans.txtFromAccNo
        CustName = GetResourceString(365)
    Else
        AccNum = RstAcc("AccNum")
        CustName = RstAcc("CustName")
    End If
Else
    Set RstAcc = GetAccRecordSet(AccHeadID, AccNum)
    
    If RstAcc Is Nothing Then
        With m_frmTrans
            Call ActivateTextBox(IIf(m_GridRow <= 1, .txtFromAccNo, .txtToAccNo))
        End With
        Exit Function
    End If
    m_AccID(ArrNo) = RstAcc("Id")
    AccNum = RstAcc("AccNum")
    CustName = RstAcc("CustName")
End If

m_AccHeadId(ArrNo) = AccHeadID
m_AmountType(ArrNo) = wisPrincipal

If m_frmTrans.optRegInt Then
    m_AmountType(ArrNo) = wisRegularInt
    ParentID = GetParentID(AccHeadID)
    If ParentID = parBankAccount Or ParentID = parBankLoanAccount Then
        'Head NAme already slected in the top
        'while seleting the grid row
    ElseIf ParentID = parMemberShare Then
        headName = ""
        If ModuleID Mod 100 > 0 Then headName = GetMemberTypeName(ModuleID Mod 100)
        headName = headName & " " & GetResourceString(53, 191)
        ParentID = parBankIncome
    ElseIf ParentID = parMemberDeposit Then
        headName = GetHeadName(AccHeadID) & " " & GetResourceString(487)
        ParentID = parMemDepIntPaid
    ElseIf ParentID = parMemDepLoan Then
        headName = GetHeadName(AccHeadID) & " " & GetResourceString(483)
        ParentID = parMemDepLoanIntReceived
    ElseIf ParentID = parMemberLoan Then
        headName = GetHeadName(AccHeadID) & " " & GetResourceString(344)
        ParentID = parMemLoanIntReceived
    End If
    
    IntHeadID = GetIndexHeadID(headName)
    If IntHeadID = 0 Then IntHeadID = GetHeadID(headName, ParentID)
    If IntHeadID = 0 Then GoTo ErrLine
    
    m_AccHeadId(ArrNo) = IIf(IntHeadID, IntHeadID, AccHeadID)
End If
If m_frmTrans.optPenalInt Then
    m_AmountType(ArrNo) = wisPenalInt
    ParentID = parMemLoanPenalInt
    headName = GetHeadName(AccHeadID) & " " & GetResourceString(345)
    
    IntHeadID = GetIndexHeadID(headName)
    If IntHeadID = 0 Then IntHeadID = GetHeadID(headName, ParentID)
    If IntHeadID = 0 Then GoTo ErrLine
    
    m_AccHeadId(ArrNo) = IIf(IntHeadID, IntHeadID, AccHeadID)
    
End If


'Now add the details to the grid
With m_frmTrans
    With .grd
        .Col = 0: .Text = .Row
        .Col = 1: .Text = StrAcctype
        .Col = 3: .Text = AccNum
        .Col = 4: .Text = CustName
        .Col = 5
    End With
    If Not .fraTo.Enabled Then
        .grd.Text = .txtFromAmount.Value
        .grd.CellAlignment = 1
        If Not m_InitLoad Then _
            m_FromAmount = m_FromAmount + .txtFromAmount
    Else
       .grd.Text = .txtToAmount
       m_ToAmount = m_ToAmount + .txtToAmount
       m_Amount(ArrNo) = .txtToAmount.Value
    End If
    .grd.Col = 2
    .grd.Text = GetResourceString(IIf(.optPrincipal, 310, _
            IIf(.optRegInt, 344, 345)))
    
    .txtToAmount.Value = m_FromAmount - m_ToAmount
    .cmdSave.Enabled = IIf(.txtToAmount.Value, False, True)
End With

ExitLine:
AddToGrid = True

m_GridRow = m_GridRow + 1

Exit Function

ErrLine:
If Err Then
    MsgBox "Error in AddtoGrid"
    'Resume
    Err.Clear
End If
    
AddToGrid = False

End Function

'
Private Sub Clear()
'first Intitielaise the Varible
ReDim m_AccHeadId(0)
ReDim m_AccID(0)
ReDim m_AmountType(0)
ReDim m_Amount(0)
m_GridRow = 0
'After That initialize the grid
If m_frmTrans Is Nothing Then Set m_frmTrans = New frmContra
Call InitGrid
With m_frmTrans
    .cmbFrom.ListIndex = -1
    .cmbTo.ListIndex = -1
    
    .fraTo.Enabled = False
    .cmbFrom.Enabled = True
    .txtFromAccNo.Enabled = True
    .txtFromAmount.Enabled = True
    .cmdFrom.Enabled = True
    .optPrincipal = True
    
    m_FromAmount = 0
    m_ToAmount = 0
    .cmdUndo.Enabled = False
    .cmdSave.Enabled = False
End With

End Sub

Private Sub InitGrid()
With m_frmTrans.grd
    .Clear
    .AllowUserResizing = flexResizeBoth
    .SelectionMode = flexSelectionByRow
    .Rows = 2
    .Cols = 6
    .FixedRows = 1
    .FixedCols = 1
    .Row = 0
    .Col = 0: .Text = GetResourceString(33): .ColWidth(0) = 450 'SlnO
    .Col = 1: .Text = GetResourceString(36, 35): .ColWidth(1) = 1650 'Accoun Name
    .Col = 2: .Text = GetResourceString(36): .ColWidth(2) = 950 'Name
    .Col = 3: .Text = GetResourceString(36, 60): .ColWidth(3) = 850 'Account No
    .Col = 4: .Text = GetResourceString(35): .ColWidth(4) = 1850 'Name
    .Col = 5: .Text = GetResourceString(40): .ColWidth(5) = 1050 'Amount
    
End With
    
End Sub


Private Sub InitialiseValue(ByVal TransDate As Date, ByVal VoucherNo As String, ByVal AccHeadID As Long, ByVal AccNum As String, ByVal Amount As Currency, cancel As Integer, Optional AmountType As wis_AmountType = wisPrincipal)
'Public Sub InitialiseValue(ByVal TransDate As Date, ByVal VoucherNo As String, ByVal Module As wisModules, ByVal AccNum As String, ByVal Amount As Currency, Cancel As Integer, Optional AmountType As Wis_AmountType = wisPrincipal)
Dim StrAcctype As String

If Amount = 0 Then GoTo Exit_Line
If AccHeadID = 0 Then GoTo Exit_Line

'If transferring amount already loaded
'then exit the sub
cancel = True

If m_frmTrans.fraTo.Enabled Then GoTo Exit_Line
m_InitLoad = True
Call Clear

With m_frmTrans
    .txtTransDate.Text = GetIndianDate(CStr(TransDate)): .txtTransDate.Locked = True
    .txtvoucher = VoucherNo: .txtvoucher.Locked = True
    
    .txtFromAccNo = AccNum
    StrAcctype = GetResourceString(307)
    .cmbFrom.Text = GetResourceString(307)
    .txtFromAmount = Amount
    m_Amount(0) = Amount
    
    'Id we do not have the unique Id of the account type then
    'Get the unique id of the account type
    Dim RstAcc As Recordset
    Set RstAcc = GetAccRecordSet(AccHeadID, AccNum)
    If RstAcc Is Nothing Then
        'MsgBox "Invalid account no specified", vbInformation, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(500), vbInformation, wis_MESSAGE_TITLE
        Call ActivateTextBox(IIf(.grd.Row = 1, .txtFromAccNo, .txtToAccNo))
        Exit Sub
    End If
End With


m_AccID(0) = RstAcc("Id")
m_AccHeadId(0) = AccHeadID
'If Module > 100 Then m_AccType(0) = Module - Module Mod 100
m_AmountType(0) = wisPrincipal

'If optRegInt Then m_AmountType(ArrNo) = wisRegularInt
'If optPenalInt Then m_AmountType(ArrNo) = wisPenalInt

'Now add the details to the grid
With m_frmTrans.grd
    .Rows = 3
    .Row = 1
    .Col = 0: .Text = .Row
    .Col = 1: .Text = StrAcctype
    .Col = 3: .Text = RstAcc("AccNum")
    .Col = 4: .Text = RstAcc("CustName")
    .Col = 5: .Text = m_frmTrans.txtFromAmount
    .CellAlignment = 1
    m_FromAmount = m_FromAmount + m_frmTrans.txtFromAmount
    .Col = 2
End With

With m_frmTrans
    .grd.Col = 2
    .grd.Text = GetResourceString(IIf(.optPrincipal, 310, _
                IIf(.optRegInt, 344, 345)))
    .txtToAmount = m_FromAmount - m_ToAmount

    .fraTo.Enabled = True
    .cmbFrom.Enabled = False
    .txtFromAccNo.Enabled = False
    .txtFromAmount.Enabled = False
    .cmdFrom.Enabled = False
End With

cancel = False

Exit_Line:

End Sub


Public Function Transfer(ByVal TransDate As Date, ByVal VoucherNo As String, _
            ByVal AccHeadID As Long, ByVal AccNum As String, _
            ByVal Amount As Currency, _
            Optional AmountType As wis_AmountType = wisPrincipal) As wis_FunctionReturned

Dim StrAcctype As String

Transfer = Failure

If Amount = 0 Then GoTo Exit_Line
If AccHeadID = 0 Then GoTo Exit_Line

'If m_frmTrans.fraTo.Enabled Then GoTo Exit_Line
'Call Clear
If Not m_InitLoad Then
    Call Clear
    With m_frmTrans
        .txtTransDate.Text = GetIndianDate(CStr(TransDate))
        .txtTransDate.Locked = True
        .txtvoucher = VoucherNo: .txtvoucher.Locked = True
    End With
End If
m_InitLoad = True

On Error GoTo Exit_Line

With m_frmTrans
    .txtFromAccNo = AccNum
    StrAcctype = GetResourceString(307)
    .cmbFrom.Text = GetResourceString(307)
    .txtFromAmount = Amount
    
    'Id we do not have the unique Id of the account type then
    'Get the unique id of the account type
    Dim RstAcc As Recordset
    Set RstAcc = GetAccRecordSet(AccHeadID, AccNum)
    If RstAcc Is Nothing Then
        'MsgBox "Invalid account no specified", vbInformation, wis_MESSAGE_TITLE
        'MsgBox GetResourceString(500), vbInformation, wis_MESSAGE_TITLE
        Call ActivateTextBox(IIf(.grd.Row = 1, .txtFromAccNo, .txtToAccNo))
        Exit Function
    End If
End With

Dim I As Integer
'If m_GridRow = 0 Then m_GridRow = 1
I = m_GridRow
ReDim Preserve m_AccID(I)
ReDim Preserve m_AccHeadId(I)
ReDim Preserve m_Amount(I)
ReDim Preserve m_AmountType(I)
m_AccID(I) = RstAcc("Id")
m_AccHeadId(I) = AccHeadID
'If Module > 100 Then m_AccType(0) = Module - Module Mod 100
m_AmountType(I) = AmountType
m_Amount(I) = Amount

'Now add the details to the grid
With m_frmTrans.grd
    .Rows = 3 + I
    .Row = I + 1
    .Col = 0: .Text = .Row
    .Col = 1: .Text = StrAcctype
    .Col = 3: .Text = RstAcc("AccNum")
    .Col = 4: .Text = RstAcc("CustName")
    .Col = 5: .Text = m_frmTrans.txtFromAmount
    .CellAlignment = 1
    m_FromAmount = m_FromAmount + m_frmTrans.txtFromAmount
    .Col = 2
End With

With m_frmTrans
    .grd.Col = 2
    .grd.Text = GetResourceString(IIf(.optPrincipal, 310, _
                                                    IIf(.optRegInt, 344, 345)))
    .txtToAmount = m_FromAmount - m_ToAmount

    .fraTo.Enabled = True
    .cmbFrom.Enabled = False
    .txtFromAccNo.Enabled = False
    .txtFromAmount.Enabled = False
    .cmdFrom.Enabled = False
End With

m_GridRow = m_GridRow + 1
m_NoOfAccounts = m_NoOfAccounts + 1

Transfer = Success

Exit_Line:

End Function

Private Sub LoadAccountBalance()
If m_frmTrans.cmbFrom.ListIndex < 0 Then Exit Sub

Dim ret As Integer
Dim BalanceAmount As Currency
Dim AccId As Long
Dim rst As ADODB.Recordset
Dim AccType As wisModules

'Check for the Selected AccountType if Account type is Sb then
'get the SB balance and if it is CA then Get the CA Balance-siddu
Dim LstInd As Long
With m_frmTrans.cmbFrom
    AccType = .ItemData(.ListIndex)
    LstInd = .ListIndex
    If .ListIndex = 0 Then AccType = wis_SBAcc
    If .ListIndex = 1 Then AccType = wis_CAAcc
End With



'Check for the Account tyeps
'AccType = Accnum
'If it is Current Account then get the Balance
If AccType = wis_CAAcc Then
    gDbTrans.SqlStmt = "SELECT AccID FROM CAMaster " & _
        "WHERE AccNum = " & AddQuotes(Trim$(m_frmTrans.txtFromAccNo.Text), True)
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        'MsgBox "Account number does not exists !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        m_frmTrans.txtFromAmount.Text = ""
        Exit Sub
    End If
   
    AccId = FormatField(rst("AccID"))
    
    'get teh TotalBalance Form the Particular Account
    gDbTrans.SqlStmt = "Select TOP 1 Balance from CATrans where AccID = " & _
        AccId & " order by TransID DESC"
    ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If ret <= 0 Then
        MsgBox "No Records"
    Else
        BalanceAmount = FormatField(rst(0))
        m_frmTrans.txtFromAmount.Value = BalanceAmount
        Set rst = Nothing
    End If
End If

'if the Account is SB Get the Sb Account balance Given from the AccountId
If AccType = wis_SBAcc Then
    'AccId=
    gDbTrans.SqlStmt = "Select * from SBMaster" & _
                " where AccNum = " & AddQuotes(m_frmTrans.txtFromAccNo, True)
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        With m_frmTrans
            .txtFromAmount.Text = ""
            .txtFromAccNo.SetFocus
        End With
        Exit Sub
    End If
    
    AccId = FormatField(rst("AccId"))
    gDbTrans.SqlStmt = "Select TOP 1 Balance from SBTrans where AccID = " & _
        AccId & " order by TransID DESC"
    ret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If ret <= 0 Then
        MsgBox GetResourceString(525), vbExclamation, gAppName & " - Error"
        Exit Sub
    Else
        BalanceAmount = FormatField(rst(0))
        m_frmTrans.txtFromAmount.Value = BalanceAmount
        Set rst = Nothing
    End If
End If

End Sub

Private Sub m_frmTrans_SaveClicked()

'First Check the amount withdrawing and
'depositing are same are not
Dim WithdrwaAmount As Currency
Dim DepositAmount As Currency
Dim count As Integer

WithdrwaAmount = m_FromAmount
DepositAmount = 0
    
With m_frmTrans.grd
    .Col = .Cols - 1
    If Not m_InitLoad Then m_NoOfAccounts = 1
    For count = m_NoOfAccounts + 1 To .Rows - 1
        .Row = count
        DepositAmount = DepositAmount + Val(.Text)
    Next
    If DepositAmount <> WithdrwaAmount Then
        'MsgBox "Invalid amount specified", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(506), vbInformation, wis_MESSAGE_TITLE
        m_frmTrans.txtFromAmount.SetFocus
        Exit Sub
    End If
    m_Status = Success
    
End With

Dim cancel As Integer
Dim TransDate As Date

TransDate = GetSysFormatDate(m_frmTrans.txtTransDate)
'RaiseEvent TransferAmount(m_AccHeadId, m_Id, m_Amount, m_AmountType, TransDate, txtvoucher, Cancel)
'If Not Cancel Then Call Clear
If m_InitLoad Then
    m_frmTrans.Hide
    'Unload m_frmTrans
Else
    SaveDetails
    Call Clear
End If

End Sub

Private Sub m_frmTrans_TransferAmount(AccHeadID() As Long, AccId() As Long, _
                                Amount() As Currency, AmountType() As wis_AmountType, _
                                TransDate As Date, VoucherNo As String, frmUnload As Integer)

ReDim m_Amount(0)
ReDim m_AccHeadId(0)
ReDim m_AccID(0)
ReDim m_AccNum(0)

'Assing the varibales to the module level varibles
m_Amount = Amount
m_TransDate = TransDate
m_AccHeadId = AccHeadID
m_AccID = AccId
m_AmountType = AmountType

m_Status = Success


frmUnload = True
If m_InitLoad Then
    frmUnload = False
Else
    If SaveDetails Then frmUnload = False
End If

End Sub


Private Sub m_frmTrans_UndoClicked()

With m_frmTrans.grd
    
    If .Row < 2 Then
        If .Rows = 3 Then Call Clear
        Exit Sub
    End If
    
    'if the user selected more than one row
    'then need not to delete that row
    If .RowSel <> .Row Then Exit Sub
    If .RowSel = .Rows - 1 Then Exit Sub
    'if selected row is Trasnferring row then exit sub
    If .RowSel <= m_NoOfAccounts Then Exit Sub
    
    Dim RemRow As Integer
    Dim MaxCount As Integer
    Dim count As Integer


'If The deleting row is not tHe last row then
    RemRow = .Row
    MaxCount = UBound(m_AccID)
    If .Rows - 2 <> .Rows Then
        .Col = 0
        For count = RemRow To MaxCount
            .Row = count
            .Text = Val(.Text) - 1
            m_AccID(count - 1) = m_AccID(count)
            m_AccHeadId(count - 1) = m_AccHeadId(count)
            m_AmountType(count - 1) = m_AmountType(count)
            m_Amount(count - 1) = m_Amount(count)
        Next
    End If
    
    .Row = RemRow
    .Col = .Cols - 1
    m_ToAmount = m_ToAmount - Val(.Text)
    'remove the row
    .RemoveItem RemRow
    'Redimension the variables
    ReDim Preserve m_AccID(MaxCount - 1)
    ReDim Preserve m_AccHeadId(MaxCount - 1)
    ReDim Preserve m_AmountType(MaxCount - 1)
    ReDim Preserve m_Amount(MaxCount - 1)
    
End With

With m_frmTrans
    .cmdSave.Enabled = IIf(m_ToAmount - m_FromAmount = 0, True, False)
    .optPrincipal = True
End With


End Sub


Private Sub m_frmTrans_WindowClosed()
    Set m_frmTrans = Nothing
End Sub


Public Property Get ObjectHandle() As Long
    If Not m_frmTrans Is Nothing Then _
                ObjectHandle = m_frmTrans.hwnd
End Property

