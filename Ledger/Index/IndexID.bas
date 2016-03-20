Attribute VB_Name = "basIndexID"
Option Explicit

Public Enum WIS_IndexIDs
    indMember = 1

    indDepositSB
    indDepositCA
    indDepositPigmy
    indDepositRD
    indDepositFD
    indDepositBKCC
    
    indLoansMember
    indLoansDeposit
    indLoansFD
    indLoansRD
    indLoansPigmy
    indLoansBKCC
    
    indProfitMember
    indProfitDepositSB
    indProfitDepositCA
    indProfitDepositFD
    indProfitDepositRD
    indProfitDepositPigmy
    indProfitDepositBKCC
    indProfitLoansDeposit
    indProfitLoansRD
    indProfitLoansPigmy
    indProfitLoansMembers
    indProfitLoansBKCC
    
    indLossDepositSB
    indLossDepositCA
    indLossDepositPigmy
    indLossDepositRD
    indLossDepositBKCC
    indLossLoansDeposit
    indLossLoansRD
    indLossLoansPigmy
    indLossLoansMembers
    indLossLoansBKCC
    
    indPayAbleDepositPigmy
    indPayAbleDepositRD
    indPayAbleDepositFD
    
    indPayAbleLoans

    
End Enum
' This Function Will Read the IndexIDs and Return the Respective Material ID
' ID will be kept the IndexIDs Table in the Database
Public Function GetIDForIndexEnum(ByVal IndexIds As WIS_IndexIDs) As Long

On Error GoTo Hell:

Dim rstID As ADODB.Recordset

GetIDForIndexEnum = 0

gDbTrans.SqlStmt = " SELECT MaterialID From IndexIDs WHERE IndexID=" & IndexIds

If gDbTrans.Fetch(rstID, adOpenForwardOnly) < 1 Then Exit Function

GetIDForIndexEnum = FormatField(rstID.Fields("MaterialID"))

Set rstID = Nothing

Exit Function

Hell:
        
End Function

Public Function GetModuleIDFromHeadID(headID As Long) As wisModules

gDbTrans.SqlStmt = "Select * From BankHeadIDs Where HeadID = " & headID

Dim rstTemp As Recordset
If gDbTrans.Fetch(rstTemp, adOpenDynamic) < 1 Then Exit Function

GetModuleIDFromHeadID = FormatField(rstTemp("AccType"))

Set rstTemp = Nothing
    
End Function


