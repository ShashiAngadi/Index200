VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Public Property Get ObjectName()

End Property

Public Property Let AccountId(NewValue As Long)
    
End Property


Public Property Let CustomerID(NewValue As Long)
End Property


Public Function UndoTransaction(AccID As Long, TransID As Long) As Boolean
End Function

Public Function ContraWithdrawls(FromDate As Date, ToDate As Date) As Currency
End Function

Public Function ContraDeposits(FromDate As Date, ToDate As Date) As Currency
End Function

Public Function CashDeposits(FromDate As Date, ToDate As Date) As Currency

End Function


Public Function CashWithdrawls(FromDate As Date, ToDate As Date) As Currency

End Function

Public Function CustomerBalance(CustomerID As Long, retAccNum As String, AsOnDate As Date) As Currency

End Function
'Public f_frmSBAcc As frmSBAcc
Public Sub Show()
Dim StartCA As Boolean

If gCurrUser Is Nothing Then
    Set gCurrUser = New clsUsers
Else
    'Now Check the permission
    'whether logged inperson is having permission of sb account or not
    Dim Perms As wis_Permissions
    Perms = gCurrUser.UserPermissions
    If Perms And perBankAdmin Then StartCA = True
    If Perms And perCreateAccount Then StartCA = True
    If Perms And perClerk Then StartCA = True
    If Perms And perCashier Then StartCA = True
    If Perms And perReadOnly Then StartCA = True
    If StartCA Then
        If m_frmCAAcc Is Nothing Then Set m_frmCAAcc = New frmCAAcc
        Load m_frmCAAcc
        If gWindowHandle Then m_ParentHandle = gWindowHandle
        gWindowHandle = m_frmCAAcc.hwnd
        m_frmCAAcc.Show
        If m_AccID < 0 Then m_frmCAAcc.AccountLoad (m_AccID)
    Else
        MsgBox "You have no permission to operate this account", vbInformation, wis_MESSAGE_TITLE
        Exit Sub
    End If
    
End If

End Sub
Public Function Deposits(FromDate As Date, ToDate As Date) As Currency
End Function

Public Function Loss(FromDate As Date, ToDate As Date) As Currency
End Function

'THis Function Retuns the Next Account Blance of the customewr
'Before calling this function, we should call the CustomerBlance function
Public Function NextBalance(retAccNum As String) As Currency
End Function


'This Function will withdraw the Amount given From The Speicfied Account No
'And It put the particulars  as deposited notes
'This function Can Be Used in Standing Instuctions Presently this is Using by Clearing Class
Public Function WithdrawAmount(AccID As Integer, Amount As Currency, _
    Description As String, TransDate As Date, Optional VoucherNo As String = "", _
    Optional CashTransaction As Boolean = False) As Integer

End Function

'This Function will Deposit the Amount given From The Speicfied Account No
'And It put the particulars  as deposited notes
'This function Can Be Used in Standing Instuctions Presently this is Using by Clearing Class
Public Function DepositAmount(ByVal AccID As Long, ByVal Amount As Currency, _
                ByVal Description As String, ByVal TransDate As Date, _
                Optional VoucherNo As String = "", _
                Optional CashTransaction As Boolean = False) As Integer

End Function


Public Function WithDrawls(FromDate As Date, ToDate As Date) As Currency
End Function

Public Function Profit(FromDate As Date, ToDate As Date) As Currency
Dim Amount As Currency
Dim TransType As wisTransactionTypes

Amount = 0

'first get the profit direct cash
TransType = wWithdraw
Amount = PLTransactionDetails(FromDate, ToDate, TransType)

';Now calculate the contra transaction
TransType = wContraWithdraw
Amount = Amount + PLTransactionDetails(FromDate, ToDate, TransType)

Profit = Amount
End Function

Public Function GetAccountID(AccNum As String) As Integer
End Function


Public Function Balance(AsOnDate As Date) As Currency

End Function

