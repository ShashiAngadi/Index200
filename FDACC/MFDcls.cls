VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatFd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function InterestOnDeposits(AsOnIndiandate As String) As Currency
   On Error GoTo ErrLine
  'declare the variables necessary
  
Dim Retval As Long
Dim rstMain As Recordset
Dim AccId As Long
Dim DepositID As Long
Dim InterestRate As Currency
Dim LastIntDate As String
Dim CreateDate As String
Dim MatDate As String
Dim Duration As Long
Dim IntAmount As Currency
Dim Balance As Currency
Dim Loan As Boolean

Dim TotalAmount As Currency

'initialize some variables
Loan = False

'Build The Querry
'''gDBTrans.SQLStmt = " SELECT Balance, A.AccId, A.DepositId, CreateDate, MaturityDate, TransDate, RateOfInterest" & _
                                 " From FDMaster A, FDTrans B Where Loan = " & Loan & _
                            "  And A.AccId = B.AccId And A.DepositId = B.DepositID " & _
                            "  And (ClosedDate = NULL or ClosedDate >= #" & FormatDate(AsOnIndiandate) & "#)" & _
                            "  And CreateDate < #" & FormatDate(AsOnIndiandate) & "# And TransID =  " & _
                                    " ( Select max(TransID) From FDTrans C Where A.AccId = C.AccId  " & _
                                    " and transDate <= #" & FormatDate(AsOnIndiandate) & "# And A.DepositId = C.DepositID" & _
                                    " And Loan = " & Loan & " ) " & _
                            " Order By A.AccId, A.DepositId "
 
 
gDBTrans.SQLStmt = " SELECT Balance, A.AccId, A.DepositId, CreateDate, MaturityDate, TransDate, RateOfInterest" & _
                " From FDMaster A, FDTrans B Where Loan = " & Loan & _
                "  And A.AccId = B.AccId And A.DepositId = B.DepositID " & _
                "  And ClosedDate = NULL " & _
                "  And CreateDate < #" & FormatDate(AsOnIndiandate) & "# And TransID =  " & _
                " ( Select max(TransID) From FDTrans C Where A.AccId = C.AccId  " & _
                " and transDate <= #" & FormatDate(AsOnIndiandate) & "# And A.DepositId = C.DepositID" & _
                " And Loan = " & Loan & " ) " & _
                " Order By A.AccId, A.DepositId "
 
 
Retval = gDBTrans.SQLFetch
If Retval < 1 Then GoTo ExitLine
Set rstMain = gDBTrans.Rst.Clone

While Not rstMain.EOF
    AccId = Val(FormatField(rstMain("AccId")))
    DepositID = Val(FormatField(rstMain("DepositID")))
    Balance = CCur(FormatField(rstMain("Balance")))

    LastIntDate = FormatField(rstMain("TransDate"))
    CreateDate = FormatField(rstMain("CreateDate"))
    InterestRate = CCur(FormatField(rstMain("RateofInterest")))
    MatDate = FormatField(rstMain("MaturityDate"))
    
    ' Necessary Changes
' if interest on deposits has to be considered
' take lastintdate or BalanceSheetDate whichevere is later
' BalanaceSheetDate is availble at Install Table




    'Now Get The Date Difference
    ' On that Difference Get the InterestRate
    Duration = WisDateDiff(LastIntDate, AsOnIndiandate)
    If Duration <= 0 Then GoTo NextDeposit
    
    If InterestRate = 0 Then
        InterestRate = CCur(GetFDDepositInterest(CInt(Duration), AsOnIndiandate))
    End If
    
    Duration = WisDateDiff(LastIntDate, AsOnIndiandate)
    If Duration <= 0 Then GoTo NextDeposit
    IntAmount = ((InterestRate / 100) * Balance * Duration) / 365
    TotalAmount = CCur(TotalAmount + IntAmount)
    
NextDeposit:
    rstMain.MoveNext
Wend
InterestOnDeposits = TotalAmount \ 1
GoTo ExitLine


ErrLine:
MsgBox "Error In FDClass--Interest On Deposits", vbCritical, wis_MESSAGE_TITLE

ExitLine:
End Function

'This Function wil Depist the Amount given From The Speicfied Account No
'And It put the particulars  as deposited notes
'This function Can Be Used in Standing Instuctions Presently this is Using by Clearing Class

Public Function DepositLoanAmount(AccId As Long, DepositID As Long, Amount As Currency, Description As String, IndianTransDate As String) As Boolean

Dim PrevTransdate As String
Dim Loan As Boolean

'Check Whether Any Loan Exists Or not on this account
Loan = False
gDBTrans.SQLStmt = "Select TransID, TransDate, Balance From FDTrans Where Balance > 0 And Loan = " & Loan & _
                        " And AccId = " & AccId & " And DepositId = " & DepositID & _
                        " And TransId = (Select max(transId) from FDTrans Where AccId = " & AccId & _
                        " And DepositId = " & DepositID & " And Loan = " & Loan & ")"

If gDBTrans.SQLFetch < 1 Then GoTo ErrLine  ' If  there is no loan

' If Transaction Has done on later date the exit the function
PrevTransdate = FormatField(gDBTrans.Rst("TransDate"))
If WisDateDiff(IndianTransDate, PrevTransdate) > 0 Then GoTo ErrLine

Dim TransID As Long
Dim Balance As Currency
Dim Particulars As String
Dim TransType As wisTransactionTypes

'Assign The Necessary Values
TransID = Val(FormatField(gDBTrans.Rst("TransID")))
Balance = Val(FormatField(gDBTrans.Rst("Balance")))
Particulars = Description
TransType = wDeposit
' Now calculate the interest on this a\loan amount till date

Dim InterestRate As Single
Dim InterestAmount As Currency
Dim Days As Integer

gDBTrans.SQLStmt = "Select CreateDate,RateOfInterest From FDMaster Where AccId = " & AccId & _
                        " And DepositId = " & DepositID
If gDBTrans.SQLFetch < 1 Then GoTo ErrLine
Days = WisDateDiff(FormatField(gDBTrans.Rst("CreateDate")), IndianTransDate)
InterestRate = FormatField(gDBTrans.Rst("RateOfInterest")) + 2
'InterestRate = GetfdLoanInterest(Days)
InterestAmount = FormatCurrency(ComputeFDInterest(Balance, FormatField(gDBTrans.Rst("CreateDate")), IndianTransDate, CDbl(InterestRate), True))
If InterestAmount > Amount Then GoTo ErrLine
'Now Calculate the Principle Amount

Amount = Amount - InterestAmount


'Query To Put The Acquired Values into Corresponding Table

gDBTrans.BeginTrans
Loan = True
' first insert the Interest
     TransType = wCharges
     Loan = True
     TransID = TransID + 1
      gDBTrans.SQLStmt = "INSERT INTO FDTrans (Accid ,DepositID,TransId ,Amount ,Balance" & _
               ", TransType, Loan, Particulars,TransDate)  VALUES(" & AccId & "," & DepositID & "," & _
               TransID & "," & InterestAmount & "," & Balance & "," & TransType & ", " & Loan & ", " & _
               " '" & Particulars & " ', #" & FormatDate(IndianTransDate) & "# )"

        If Not gDBTrans.SQLExecute Then
           gDBTrans.RollBack
           GoTo ErrLine
        End If

' nEXT insert the PRINCIPLE
     TransType = wDeposit
     TransID = TransID + 1
     Balance = Balance - Amount
     gDBTrans.SQLStmt = "INSERT INTO FDTrans (Accid ,DepositID,TransId ,Amount ,Balance" & _
               ", TransType, Loan, Particulars,TransDate)  VALUES(" & AccId & "," & DepositID & "," & _
               TransID & "," & InterestAmount & "," & Balance & "," & TransType & ", " & Loan & ", " & _
               " '" & Particulars & "', #" & FormatDate(IndianTransDate) & "# )"

        If Not gDBTrans.SQLExecute Then
           gDBTrans.RollBack
           GoTo ErrLine
        End If

gDBTrans.CommitTrans

DepositLoanAmount = True

ErrLine:

End Function


Public Function CustomerLoanBalance(CustomerId As Long, Optional AccId As Long, Optional AsOnIndiandate As String) As Currency
'To trap an error
On Error GoTo ErrLine

CustomerLoanBalance = 0
'Sql stmt to get Accid
gDBTrans.SQLStmt = "Select Accid from FDMaster where Customerid=" & CustomerId
If gDBTrans.SQLFetch < 1 Then
    AccId = 0
    Exit Function
End If
AccId = FormatField(gDBTrans.Rst(0))
If Trim$(AsOnIndiandate) = "" Then
    AsOnIndiandate = FormatDate(gStrDate)
End If
Dim Loan As Boolean
Loan = True
gDBTrans.SQLStmt = "Select  top 1 Balance from FDTrans where Accid=" & AccId & " And Loan = " & Loan & _
                " And Transid = (Select max(transID) from FDtrans Where Accid=" & AccId & " And TransDate <= #" & FormatDate(AsOnIndiandate) & "# " & _
                " And Loan = " & Loan & " )"

If gDBTrans.SQLFetch > 0 Then
    CustomerLoanBalance = FormatField(gDBTrans.Rst("Balance"))
End If
 
ExitLine:
    Exit Function
ErrLine:
    If Err Then
        MsgBox "AccountBalance: " & vbCrLf & Err.Description, vbCritical
        GoTo ExitLine
    End If

End Function


Public Function CustomerBalance(CustomerId As Long, Optional AccId As Long, Optional AsOnIndiandate As String) As Currency
'To trap an error
On Error GoTo ErrLine

CustomerBalance = 0
'Sql stmt to get Accid
gDBTrans.SQLStmt = "Select Accid from FDMaster where Customerid=" & CustomerId
If gDBTrans.SQLFetch < 1 Then
    AccId = 0
    Exit Function
End If
AccId = FormatField(gDBTrans.Rst(0))
gDBTrans.SQLStmt = "Select  top 1 Balance from FDTrans where Accid=" & AccId & " ORDER BY Transid  Desc"
If Trim$(AsOnIndiandate) = "" Then
    AsOnIndiandate = FormatDate(gStrDate)
End If
Dim Loan As Boolean
Loan = False
gDBTrans.SQLStmt = "Select  Balance from FDTrans where Accid=" & AccId & " And Loan = " & Loan & _
                " And Transid = (Select max(transID) from FDtrans Where Accid=" & AccId & " And TransDate <= #" & FormatDate(AsOnIndiandate) & "# " & _
                " And Loan = " & Loan & " )"


If gDBTrans.SQLFetch > 0 Then
    CustomerBalance = FormatField(gDBTrans.Rst("Balance"))
End If
 
ExitLine:
    Exit Function
ErrLine:
    If Err Then
        MsgBox "AccountBalance: " & vbCrLf & Err.Description, vbCritical
        GoTo ExitLine
    End If
Resume
End Function



Public Function InterestOnLoans(AsOnIndiandate As String) As Currency
   On Error GoTo ErrLine
  'declare the variables necessary
  
Dim Retval As Long
Dim rstMain As Recordset
Dim AccId As Long
Dim DepositID As Long
Dim InterestRate As Currency
Dim LastIntDate As String
Dim CreateDate As String
Dim MatDate As String
Dim Duration As Long
Dim IntAmount As Currency
Dim Balance As Currency
Dim Loan As Boolean

Dim TotalAmount As Currency

'initialize some variables
Loan = True

'Build The Querry
gDBTrans.SQLStmt = " SELECT Balance, A.AccId, A.DepositId, CreateDate, MaturityDate, TransDate, RateOfInterest" & _
                                 " From FDMaster A, FDTrans B Where Loan = " & Loan & _
                            "  And A.AccId = B.AccId And A.DepositId = B.DepositID " & _
                            "  And (ClosedDate = NULL or ClosedDate >= #" & FormatDate(AsOnIndiandate) & "#)" & _
                            "  And CreateDate < #" & FormatDate(AsOnIndiandate) & "# And TransID =  " & _
                                    " ( Select max(TransID) From FDTrans C Where A.AccId = C.AccId  " & _
                                    " and transDate <= #" & FormatDate(AsOnIndiandate) & "# And A.DepositId = C.DepositID" & _
                                    " And Loan = " & Loan & " ) " & _
                            " Order By A.AccId, A.DepositId "
 
Retval = gDBTrans.SQLFetch
If Retval < 1 Then GoTo ExitLine
Set rstMain = gDBTrans.Rst.Clone

While Not rstMain.EOF
    AccId = Val(FormatField(rstMain("AccId")))
    DepositID = Val(FormatField(rstMain("DepositID")))
    Balance = CCur(FormatField(rstMain("Balance")))
    LastIntDate = FormatField(rstMain("TransDate"))
    CreateDate = FormatField(rstMain("CreateDate"))
    InterestRate = CCur(FormatField(rstMain("RateofInterest")))
    MatDate = FormatField(rstMain("MaturityDate"))
    
    'Now Get The Date Difference
    ' On that Difference Get the InterestRate
    Duration = WisDateDiff(LastIntDate, AsOnIndiandate)
    If Duration <= 0 Then GoTo NextDeposit
    
    If InterestRate = 0 Then
        InterestRate = CCur(GetFDLoanInterest(CInt(Duration), CreateDate))
    End If
    
    Duration = WisDateDiff(LastIntDate, AsOnIndiandate)
    If Duration <= 0 Then GoTo NextDeposit
    IntAmount = ((InterestRate / 100) * Balance * Duration) / 365
    TotalAmount = CCur(TotalAmount + IntAmount)
    
NextDeposit:
    rstMain.MoveNext
Wend
InterestOnLoans = TotalAmount
GoTo ExitLine


ErrLine:
MsgBox "Error In FDClass--Interest On Loans", vbCritical, wis_MESSAGE_TITLE
'''Debug.Print "Kannada"

ExitLine:

End Function

Public Function Liability(AsOnIndiandate As String) As Currency
' Computes the Liability (Deposit plus interest on Deposit)

Dim Setup As New clsSetup
Dim Loan As Boolean
Dim LiabBeforeInterest As Currency
Dim InterestOnDeposit As Double

Loan = False
Liability = 0

If Not DateValidate(AsOnIndiandate, "/", True) Then Exit Function

LiabBeforeInterest = ComputeTotalFDLiability(AsOnIndiandate, Loan)

'InterestOnDeposit = CDbl(Setup.ReadSetupValue("FDacc", "Interest On FDDeposit", 12#))

'Liability = ComputeFDInterest(LiabBeforeInterest, InterestOnDeposit)
' The  above function CompuntFDInterest needs period which needs to be clarified
End Function


Private Function TransactionDetails(FromIndianDate As String, ToIndianDate As String, Loan As Boolean, TransType As wisTransactionTypes) As Currency
TransactionDetails = -1
If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
If Not DateValidate(ToIndianDate, "/", True) Then Exit Function

Dim SqlStr As String

SqlStr = " Select Sum(Amount) as Deposits from FDTrans where " & _
    " TransType = " & TransType & _
    " And TransDate >= #" & FormatDate(FromIndianDate) & "# " & _
    " And TransDate <= #" & FormatDate(ToIndianDate) & "#" & _
    " And Loan = " & Loan

gDBTrans.SQLStmt = SqlStr
If gDBTrans.SQLFetch < 1 Then Exit Function

Dim Rst As Recordset
Set Rst = gDBTrans.Rst.Clone
If Not Rst Is Nothing Then _
    TransactionDetails = Val(FormatField(Rst(0)))

End Function

 

Public Sub Show()
    gWindowHandle = frmFDAcc.hWnd
    frmFDAcc.Show 'vbModal
End Sub

Public Function DepositDeposits(FromIndianDate As String, ToIndianDate As String) As Currency
' Total Fixed Deposit flown in the bank between 2 given dates

Dim TransType As wisTransactionTypes
Dim Loan As Boolean
DepositDeposits = -1

If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
If Not DateValidate(ToIndianDate, "/", True) Then Exit Function
Loan = False
TransType = wDeposit

DepositDeposits = TransactionDetails(FromIndianDate, ToIndianDate, Loan, TransType)

End Function

Public Function LossDeposits(FromIndianDate As String, ToIndianDate As String) As Currency
        LossDeposits = 0
        
        Dim Loan As Boolean
        Dim TransType As wisTransactionTypes
        
        If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
        If Not DateValidate(ToIndianDate, "/", True) Then Exit Function
        
        Loan = False
        TransType = wInterest
        LossDeposits = TransactionDetails(FromIndianDate, ToIndianDate, Loan, TransType)

End Function


Public Function ContraRPLossDeposits(FromIndianDate As String, ToIndianDate As String) As Currency
        ContraRPLossDeposits = 0
        
        Dim Loan As Boolean
        Dim TransType As wisTransactionTypes
        
        If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
        If Not DateValidate(ToIndianDate, "/", True) Then Exit Function
        
        Loan = False
        TransType = wRPInterest
        ContraRPLossDeposits = TransactionDetails(FromIndianDate, ToIndianDate, Loan, TransType)

End Function


Public Function ContraLossDeposits(FromIndianDate As String, ToIndianDate As String) As Currency
        ContraLossDeposits = 0
        
        Dim Loan As Boolean
        Dim TransType As wisTransactionTypes
        
        If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
        If Not DateValidate(ToIndianDate, "/", True) Then Exit Function
        
        Loan = False
        TransType = wContraInterest
        ContraLossDeposits = TransactionDetails(FromIndianDate, ToIndianDate, Loan, TransType)

End Function

Public Function WithDrawlDeposits(FromIndianDate As String, ToIndianDate As String) As Currency
'Fixed Deposit Given back to the customer
WithDrawlDeposits = -1

Dim TransType As wisTransactionTypes
Dim Loan As Boolean

TransType = wWithDraw
Loan = False

If FromIndianDate <> "" And Not DateValidate(FromIndianDate, "/", True) Then Exit Function

If ToIndianDate <> "" And Not DateValidate(ToIndianDate, "/", True) Then Exit Function

WithDrawlDeposits = TransactionDetails(FromIndianDate, ToIndianDate, Loan, TransType)

End Function
Public Function ProfitDeposits(FromIndianDate As String, ToIndianDate As String) As Currency

' Profit to the bank through deposits

        ProfitDeposits = 0
        
        Dim Loan As Boolean
        Dim TransType As wisTransactionTypes
        
        
        If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
        If Not DateValidate(ToIndianDate, "/", True) Then Exit Function
        
        Loan = False
        TransType = wCharges
        ProfitDeposits = TransactionDetails(FromIndianDate, ToIndianDate, Loan, TransType)
        
End Function


Public Function BalanceLoans(AsOnIndiandate As String) As Currency

End Function
Public Function WithDrawlLoans(FromIndianDate As String, ToIndianDate As String) As Currency

Dim TransType As wisTransactionTypes
Dim Loan As Boolean

WithDrawlLoans = -1
Loan = True
TransType = wWithDraw

If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
If Not DateValidate(ToIndianDate, "/", True) Then Exit Function
    
WithDrawlLoans = TransactionDetails(FromIndianDate, ToIndianDate, Loan, TransType)

End Function


Public Function LossLoans(FromIndianDate As String, ToIndianDate As String) As Currency
        LossLoans = 0
        
        Dim Loan As Boolean
        Dim TransType As wisTransactionTypes
        
        
        If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
        If Not DateValidate(ToIndianDate, "/", True) Then Exit Function
        
        Loan = True
        TransType = wInterest
        LossLoans = TransactionDetails(FromIndianDate, ToIndianDate, Loan, TransType)

End Function

Public Function ProfitLoans(FromIndianDate As String, ToIndianDate As String) As Currency
        ProfitLoans = 0
        
        Dim Loan As Boolean
        Dim TransType As wisTransactionTypes
        
        If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
        If Not DateValidate(ToIndianDate, "/", True) Then Exit Function
        
        Loan = True
        TransType = wCharges
        ProfitLoans = TransactionDetails(FromIndianDate, ToIndianDate, Loan, TransType)
        
End Function





Public Function DepositLoans(FromIndianDate As String, ToIndianDate As String) As Currency

Dim TransType As wisTransactionTypes
Dim Loan As Boolean
Loan = True
DepositLoans = -1

If Not DateValidate(FromIndianDate, "/", True) Then Exit Function
If Not DateValidate(ToIndianDate, "/", True) Then Exit Function

TransType = wDeposit

DepositLoans = TransactionDetails(FromIndianDate, ToIndianDate, Loan, TransType)

End Function





Public Function BalanceDeposits(AsOnIndiandate As String) As Currency
' fixed deposits in the bank as on the given date
BalanceDeposits = 0
On Error GoTo ErrLine
    

gDBTrans.SQLStmt = "Select SUM(Balance) From MFDTrans A, FDMaster B " & _
    " Where " & _
    " A.Accid = B.Accid And A.DepositId = B.DepositID " & _
    " AND A.Loan = " & Loan & " And TransID = " & _
    "(Select Max(TransID) from MFDTrans C where A.Accid = C.Accid " & _
    " AND A.DepositId = C.DepositID AND Loan = " & Loan & " and TransDate <= " & _
    " #" & FormatDate(AsOnIndiandate) & "# )"

If gDBTrans.SQLFetch < 1 Then Exit Function

BalanceDeposits = FormatField(gDBTrans.Rst(0))
    
ErrLine:
DoEvents

End Function



Public Function BalanceMaturedFD(AsOnIndiandate As String) As Currency

' fixed deposits in the bank as on the given date
BalanceMaturedFD = -1
gDBTrans.SQLStmt = "Select SUM(Balance) FROM FDTrans A, FDMaster B" & _
    " Where (MaturedOn <> NULL AND ((ClosedDate >= #" & AsOnIndiandate & "#" & _
    " Or ClosedDate Is Null) AND MaturedOn <= #" & AsOnIndiandate & "#))" & _
    " And A.Accid = B.Accid And A.DepositId = B.DepositID And Loan = " & False & _
    " And TransID = (Select Max(TransID) from FDTrans C where A.Accid = C.Accid " & _
    " AND A.DepositId = C.DepositID AND Loan = " & False & " AND TransDate <= #" & _
    AsOnIndiandate & "# )"
If gDBTrans.SQLFetch > 0 Then
   BalanceMaturedFD = FormatField(gDBTrans.Rst(0))
End If
End Function
