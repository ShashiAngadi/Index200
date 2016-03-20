VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSuspAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_frmSusp As frmSuspAcc
Attribute m_frmSusp.VB_VarHelpID = -1
Private m_ParentHandle As Long

'This Function will Depist the Amount to The Speicfied Account No
'And It put the particulars  as deposited notes
'This function Can Be Used in Standing Instuctions Presently this is Using by Clearing Class
Public Function DepositAmount(ByVal AccHeadID As Long, ByVal AccId As Long, _
                    ByVal CustomerID As Long, ByVal CustName As String, _
                    ByVal TransDate As Date, ByVal Amount As Currency, _
                    ByVal AccountTransID As Integer, Optional VoucherNo As String = "", _
                    Optional CashTransaction As Boolean = False, Optional ContraTransID As Long = 0) As Integer

Dim Rst As ADODB.Recordset


'Now Get the Max TransID On specified  Date Then compare it With Max(TransID)
gDbTrans.SqlStmt = "SELECT MAX(TransID) From SuspAccount"

Dim TransID As Long
Dim transType As wisTransactionTypes
Dim Balance As Currency

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 0 Then
    'Assign The Necessary Values
    TransID = Rst("TransID")
    Balance = FormatField(Rst("Balance"))
End If

'Query To Put The Acquired Values into Corresponding Table
TransID = TransID + 1
Balance = Amount

transType = IIf(CashTransaction, wDeposit, wContraDeposit)

Dim UserID As Long
UserID = gCurrUser.UserID

gDbTrans.SqlStmt = "INSERT INTO SuspAccount (TransId," & _
         " CustomerID,AccHeadID,Accid,CustName," & _
         " AccTransID,TransDate,Amount, TransType, " & _
         " Balance,VoucherNo,UserId) " & _
         " VALUES (" & TransID & "," & _
         CustomerID & "," & AccHeadID & "," & _
         AccId & "," & AddQuotes(CustName, True) & "," & _
         AccountTransID & ", #" & TransDate & "#," & _
         Amount & "," & transType & "," & Balance & "," & _
         AddQuotes(VoucherNo, True) & "," & UserID & " )"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

'update it to te Contra Trans table
'Dim AccHeadID As Long
AccHeadID = GetIndexHeadID(GetResourceString(365))
If Not CashTransaction Then
    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = ContraTransID
    
    If ContraID = 0 Then ContraID = GetMaxContraTransID
    If ContraID = 0 Then ContraID = 1
    
    gDbTrans.SqlStmt = "Insert INTO ContraTrans" & _
                "(ContraID,AccHeadID,AccId," & _
                "TransType, TransID,Amount,VoucherNo,UserId)" & _
                " Values (" & ContraID & "," & AccHeadID & "," & _
                AccHeadID & ", " & transType & ", " & _
                TransID & "," & Amount & ", " & _
                AddQuotes(VoucherNo, True) & "," & gUserID & " )"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
End If

DepositAmount = TransID

ErrLine:

End Function

Public Sub Show()
If m_frmSusp Is Nothing Then Set m_frmSusp = New frmSuspAcc

If gWindowHandle Then m_ParentHandle = 0
gWindowHandle = m_frmSusp.hwnd
m_frmSusp.Show 1

End Sub

'This Function will withdraw the Amount given From The Speicfied Account No
'And It put the particulars  as deposited notes
Public Function WithdrawAmount(ByVal AccountHeadID As Long, ByVal AccId As Long, _
                        ByVal CustomerID As Long, ByVal CustName As String, _
                        ByVal TransDate As Date, Amount As Currency, _
                        Optional AccountTransID As Integer, _
                        Optional PrevTransID As Long, _
                        Optional VoucherNo As String, _
                        Optional CashTransaction As Boolean = False, Optional ContraTransID As Long = 0) As Integer

If VoucherNo = "" Then VoucherNo = " "

Dim Rst As ADODB.Recordset
Dim SqlStmt As String

'Now Get the Max TransID On specified  Date Then compare it With Max(TransID)
gDbTrans.SqlStmt = "SELECT MAX(TransID) From SuspAccount "

Dim TransID As Long
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then TransID = FormatField(Rst(0))

TransID = TransID + 1
If Not CashTransaction Then
    'Get the Contra Transaction ID
    Dim ContraID As Long
    ContraID = ContraTransID
    If ContraID = 0 Then ContraID = GetMaxContraTransID + 1
End If

Dim Balance As Currency

'Assign The Necessary Values
Balance = Amount

Dim transType As wisTransactionTypes
Dim UserID As Long

UserID = gCurrUser.UserID
transType = IIf(CashTransaction, wWithdraw, wContraWithdraw)

'Query To Put The Acquired Values into Corresponding Table
gDbTrans.SqlStmt = "INSERT INTO SuspAccount (TransId," & _
         " CustomerID,AccHeadID,Accid,CustName," & _
         " AccTransID,TransDate,Amount, TransType, " & _
         " Balance,VoucherNo,UserId) " & _
         " VALUES (" & TransID & "," & _
         CustomerID & "," & AccountHeadID & "," & _
         AccId & "," & AddQuotes(CustName, True) & "," & _
         AccountTransID & ", #" & TransDate & "#," & _
         Amount & "," & transType & "," & Balance & "," & _
         AddQuotes(VoucherNo, True) & "," & UserID & " )"

If Not gDbTrans.SQLExecute Then GoTo ErrLine

'Get the Headid
Dim ClsBank As clsBankAcc
Dim HeadID As Long
    
If Not CashTransaction Then
    HeadID = GetIndexHeadID(GetResourceString(365))
    'update it to te Contra Trans table
    gDbTrans.SqlStmt = "Insert INTO ContraTrans " & _
            "(ContraID,AccHeadID,AccID," & _
            "TransType, TransID,Amount,UserId,VoucherNo)" & _
            " Values (" & ContraID & "," & HeadID & "," & _
            AccId & ", " & transType & ", " & TransID & "," & Amount & ", " & _
            UserID & "," & AddQuotes(VoucherNo, True) & ")"
    ' Execute the updation.
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
End If

If PrevTransID Then
    'if already related transaction (i.e. deposit from the customer made)
    'has made then mark the transaction as Cleared
    gDbTrans.SqlStmt = "Update SuspAccount " & _
            " set Cleared = 1 Where TransId = " & PrevTransID
    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    'mark the new record also has cleared
    gDbTrans.SqlStmt = "Update SuspAccount " & _
            " set Cleared = 1 Where TransId = " & TransID
    If Not gDbTrans.SQLExecute Then GoTo ErrLine

End If
WithdrawAmount = TransID

Exit Function

ErrLine:

End Function

Public Property Get ObjectName()
    ObjectName = "SuspClass"
End Property


Public Function UndoTransaction(AccId As Long, TransID As Long) As Boolean
    
    'Get last transaction record
    Dim Amount As Currency
    Dim lastTransID As Long
    Dim TransDate As Date
    Dim transType As wisTransactionTypes
    Dim DiffAmount As Currency
    Dim Rst As ADODB.Recordset
     
    gDbTrans.SqlStmt = "Select TOP 1 * from SuspAccount " & _
                    " order by TransID desc"
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Function
    Amount = FormatField(Rst("Amount"))
    lastTransID = FormatField(Rst("TransID"))
    TransDate = Rst("TransDate")
    transType = FormatField(Rst("TransType"))
    
    If TransID <> lastTransID Then
        gDbTrans.SqlStmt = "Select TOP 1 * from SuspAccount " & _
                    " Where TransID = " & TransID
        TransDate = Rst("TransDate")
        transType = FormatField(Rst("TransType"))
        DiffAmount = Amount * IIf(transType = wContraDeposit Or transType = wDeposit, 1, -1)
        'Exit Function
    End If
    
'Delete record from Data base
    'detete the transaction
    gDbTrans.SqlStmt = "Delete from SuspAccount where TransID = " & lastTransID
    If Not gDbTrans.SQLExecute Then Exit Function
    
    gDbTrans.SqlStmt = "UpDate SuspAccount Set Balance = Balance - " & DiffAmount & _
                        " where TransID > " & lastTransID & _
                        " AND TransDate >= #" & TransDate & "#"
    If Not gDbTrans.SQLExecute Then Exit Function
    
    UndoTransaction = True

End Function


Private Sub m_frmSusp_AddAmount(ByVal AccHeadID As Long, _
                        ByVal AccountId As Long, ByVal CustomerID As Long, _
                        ByVal Name As String, ByVal TransDate As Date, _
                        ByVal Amount As Currency)
    

On Error GoTo ErrLine

Dim BankClass As clsBankAcc
Set BankClass = New clsBankAcc

gDbTrans.BeginTrans

If DepositAmount(AccHeadID, AccountId, CustomerID, _
        Name, TransDate, Amount, 0, , True) < 0 Then gDbTrans.RollBacknRaiseError
    

If Not BankClass.AddCashDeposits(AccHeadID, Amount, TransDate) Then _
                        gDbTrans.RollBacknRaiseError
Set BankClass = Nothing

gDbTrans.CommitTrans

Exit Sub

ErrLine:

    Set BankClass = Nothing
    MsgBox GetResourceString(535), , wis_MESSAGE_TITLE

End Sub

Private Sub m_frmSusp_RemoveAmount(ByVal AccHeadID As Long, ByVal AccountId As Long, _
                    ByVal Name As String, ByVal TransDate As Date, _
                    ByVal Amount As Currency, ByVal PrevTrans As Long)


On Error GoTo ErrLine

Dim BankClass As clsBankAcc
Set BankClass = New clsBankAcc
Dim custId As Long

gDbTrans.BeginTrans

If WithdrawAmount(AccHeadID, AccountId, custId, _
        Name, TransDate, Amount, 0, PrevTrans, , True) < 0 Then gDbTrans.RollBacknRaiseError
    

If Not BankClass.UpdateCashWithDrawls(AccHeadID, Amount, TransDate) Then _
                        gDbTrans.RollBacknRaiseError
Set BankClass = Nothing

gDbTrans.CommitTrans

Exit Sub

ErrLine:

    Set BankClass = Nothing
    MsgBox GetResourceString(535), , wis_MESSAGE_TITLE

End Sub


Private Sub m_frmSusp_UndoTrans()

Dim Rst As Recordset
Dim lastTransID As Long
Dim transType As wisTransactionTypes
Dim TransDate As Date
Dim Amount As Currency

gDbTrans.SqlStmt = "Select Top 1 * From SuspAccount" & _
            " Order By TransId Desc"
If gDbTrans.Fetch(Rst, adOpenDynamic) < 1 Then Exit Sub

transType = Rst("TransType")
TransDate = Rst("TransDate")
lastTransID = Rst("TransID")
Amount = Rst("Amount")

If transType = wContraDeposit Then
    MsgBox GetResourceString(775), , wis_MESSAGE_TITLE
    Exit Sub
End If


Dim SuspHeadID As Long

SuspHeadID = GetIndexHeadID(GetResourceString(365))

If transType = wContraWithdraw Then
    'In case of contra transaction
    'Get the headname of the counter part
    gDbTrans.SqlStmt = "SELECT * From ContraTrans " & _
            " WHERE AccHeadID = " & SuspHeadID & _
            " And  TransID = " & lastTransID
    If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then
        Dim ContraClass As clsContra
        Set ContraClass = New clsContra
        Call ContraClass.UndoTransaction(Rst("ContraID"), TransDate)  '= Success Then _

        Set ContraClass = Nothing
        Exit Sub
    End If
End If
 


gDbTrans.BeginTrans


gDbTrans.SqlStmt = "Delete From SuspAccount Where " & _
            " TransID = " & lastTransID
If Not gDbTrans.SQLExecute Then GoTo ErrLine

Dim BankClass As New clsBankAcc
If transType = wDeposit Then _
    If Not BankClass.UndoCashDeposits(SuspHeadID, Amount, TransDate) Then GoTo ErrLine
If transType = wWithdraw Then _
    If Not BankClass.UndoCashWithdrawls(SuspHeadID, Amount, TransDate) Then GoTo ErrLine

gDbTrans.CommitTrans

Set BankClass = Nothing
Exit Sub

ErrLine:
    
    gDbTrans.RollBack
    Set BankClass = Nothing
    MsgBox GetResourceString(535), , wis_MESSAGE_TITLE
        
End Sub

Private Sub m_frmSusp_WindowClosed(wHandle As Long)

If gWindowHandle = wHandle Then gWindowHandle = 0
If m_ParentHandle Then gWindowHandle = m_ParentHandle

Set m_frmSusp = Nothing
End Sub


Public Property Get ObjectHandle() As Long
    If Not m_frmSusp Is Nothing Then _
            ObjectHandle = m_frmSusp.hwnd
End Property

