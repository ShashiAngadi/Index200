Attribute VB_Name = "basGeneral"
Option Explicit


Public Function ExecuteSql(SQLStmt As String) As Boolean
    ExecuteSql = False
    gDbTrans.BeginTrans
    gDbTrans.SQLStmt = SQLStmt '"Update SBMaster set ClosedDate = NULL where AccID = " & AccNo
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Function
    End If
    gDbTrans.CommitTrans
    ExecuteSql = True
End Function


Public Function RemoveFromAmountReceivable(AccHeadID As Long, AccID As Long, _
                        AccTransID As Long, TransDate As Date, _
                        Amount As Currency, DueHeadID As Long) As Boolean

Dim rstTemp As Recordset
Dim lngTransID As Long
Dim Balance As Currency

gDbTrans.SQLStmt = "Select Balance From AmountReceivAble" & _
        " WHere AccHeadID = " & AccHeadID & _
        " ANd AccId = " & AccID & _
        " Order By TransID Desc"
Balance = 0
If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
                    Balance = FormatField(rstTemp("Balance"))

'Now Get the TransID
gDbTrans.SQLStmt = "Select Max(TransID) as MaxTransID From AmountReceivAble"
If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
                    lngTransID = FormatField(rstTemp("MaxTransID"))

lngTransID = lngTransID + 1
Balance = Balance - Amount
If Balance < 0 Then Balance = 0
'Now insert this details Into amount receivable table
gDbTrans.SQLStmt = "Insert Into AmountReceivAble" & _
                "( TransID,AccHeadId,AccID," & _
                " TransType,TransDate,AccTransID,Amount," & _
                " Balance,UserID,DueHeadID) VALUES (" & _
                lngTransID & "," & AccHeadID & "," & AccID & "," & _
                wDeposit & ",#" & TransDate & "#," & AccTransID & "," & _
                Amount & "," & Balance & "," & gCurrUser.UserID & "," & _
                DueHeadID & ")"

'Now Execute the query
RemoveFromAmountReceivable = gDbTrans.SQLExecute

End Function

Public Function AddToAmountReceivable(AccHeadID As Long, AccID As Long, _
                        AccTransID As Long, TransDate As Date, _
                        Amount As Currency, DueHeadID As Long) As Boolean

Dim rstTemp As Recordset
Dim lngTransID As Long
Dim Balance As Currency

gDbTrans.SQLStmt = "Select Balance From AmountReceivAble" & _
        " WHere AccHeadID = " & AccHeadID & _
        " ANd AccId = " & AccID & _
        " Order By TransID Desc"
Balance = 0
If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
                    Balance = FormatField(rstTemp("Balance"))

'Now Get the TransID
gDbTrans.SQLStmt = "Select Max(TransID) as MaxTransID From AmountReceivAble"
If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
                    lngTransID = FormatField(rstTemp("MaxTransID"))

lngTransID = lngTransID + 1
Balance = Balance + Amount

'Now insert this details Into amount receivable table
gDbTrans.SQLStmt = "Insert Into AmountReceivAble" & _
                "(TransID,AccHeadId,AccID," & _
                " TransType,TransDate,AccTransID,Amount," & _
                " Balance,UserID,DueHeadID) VALUES (" & _
                lngTransID & "," & AccHeadID & "," & AccID & "," & _
                wWithdraw & ",#" & TransDate & "#," & AccTransID & "," & _
                Amount & "," & Balance & "," & gCurrUser.UserID & "," & _
                DueHeadID & ")"

'Now Execute the query
AddToAmountReceivable = gDbTrans.SQLExecute

End Function

Public Function UndoAmountReceivable(AccHeadID As Long, _
                    AccID As Long, AccTransID As Long) As Boolean

'Now insert this details Into amount receivable table
gDbTrans.SQLStmt = "Delete from AmountReceivAble" & _
                " Where AccheadID = " & AccHeadID & _
                " And AccID = " & AccID & _
                " And AccTransID = " & AccTransID
                
'Now Execute the query
UndoAmountReceivable = gDbTrans.SQLExecute

End Function


Function SaveInterest(ByVal ModuleID As Integer, ByVal SchemeName As String, _
        ByVal IntRate As Single, Optional ByVal EmpIntRate As Single, _
        Optional ByVal SeniorIntRate As Single, Optional ByVal AsOnDate As Date) As Boolean
        

If IsMissing(AsOnDate) Then AsOnDate = DayBeginUSDate


If IntRate <= 0 Then Exit Function

If EmpIntRate = 0 Then EmpIntRate = IntRate
If SeniorIntRate = 0 Then SeniorIntRate = IntRate

If IntRate < 1 Then IntRate = IntRate * 100
If EmpIntRate < 1 Then EmpIntRate = EmpIntRate * 100
If SeniorIntRate < 1 Then SeniorIntRate = SeniorIntRate * 100
 
Dim strIntRate As String
 
strIntRate = IntRate & ";" & EmpIntRate & ";" & SeniorIntRate

Dim IntClass As clsInterest
Set IntClass = New clsInterest
If Not IntClass.SaveInterest(ModuleID, SchemeName, IntRate, _
                EmpIntRate, SeniorIntRate, AsOnDate) Then GoTo ExitLine

Set IntClass = Nothing

SaveInterest = True
        
ExitLine:
End Function


Function GetInterestRate(ByVal ModuleID As Integer, ByVal SchemeName As String, _
            Optional ByVal IsEmployee As Boolean, _
            Optional ByVal IsSeniorCitizen As Single, _
            Optional ByVal AsOnDate As Date) As Single
        
GetInterestRate = 0


If IsMissing(AsOnDate) Then AsOnDate = DayBeginUSDate

'Now Get the last interest update on this module
'First Get The Rate Of Interest As On From indianDate
gDbTrans.SQLStmt = "Select Top 1 * from InterestTab " & _
            " Where StartDate <= #" & AsOnDate & "#" & _
            " And ModuleID = " & ModuleID & _
            " And SchemeName = " & AddQuotes(SchemeName, True) & _
            " Order by StartDate Desc "

Dim Rst As ADODB.Recordset
If gDbTrans.Fetch(Rst, adOpenForwardOnly) <= 0 Then
    gDbTrans.SQLStmt = "Select Top 1 * from InterestTab " & _
            " Where ModuleID = " & ModuleID & _
            " And SchemeName = " & AddQuotes(SchemeName, True) & _
            " Order by StartDate Desc "
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 0 Then Exit Function
End If

Dim strIntRate As String
Dim pos As Integer

strIntRate = FormatField(Rst("InterestRate"))
Set Rst = Nothing

pos = InStr(1, strIntRate, ";")
If pos = 0 Then GoTo ExitLine

If IsEmployee Then
    strIntRate = Mid(strIntRate, pos + 1)
ElseIf IsSeniorCitizen Then
    pos = InStr(pos + 1, strIntRate, ";")
    If pos = 0 Then GoTo ExitLine
    strIntRate = Mid(strIntRate, pos + 1)
End If

ExitLine:

GetInterestRate = Val(strIntRate)
        

End Function

Function GetInterestRateOnDate(ByVal ModuleID As Integer, _
            ByVal SchemeName As String, ByVal AsOnDate As Date) As String
        
GetInterestRateOnDate = ""

'Now Get the last interest update on this module
'First Get The Rate Of Interest As On From indianDate
gDbTrans.SQLStmt = "Select Top 1 * from InterestTab " & _
            " Where StartDate = #" & AsOnDate & "#" & _
            " And ModuleID = " & ModuleID & _
            " And SchemeName = " & AddQuotes(SchemeName, True)

Dim strIntRate As String
Dim Rst As ADODB.Recordset

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 0 Then Exit Function

strIntRate = FormatField(Rst("InterestRate"))
Set Rst = Nothing

ExitLine:

GetInterestRateOnDate = strIntRate
        

End Function



