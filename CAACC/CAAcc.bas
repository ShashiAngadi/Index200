Attribute VB_Name = "basCAAcc"
Option Explicit
Public Enum wis_CAReports
    repCABalance = 1
    repCADayBook = 2
    repCALedger = 3
    repCAAccOpen = 4
    repCAAccClose = 5
    repCAJoint = 6
    repCAProduct = 7
    repCACheque = 8
    repCAMonthlyBalance = 9
    repCACashBook
End Enum



Public Function CAAccountExists(ByVal AccId As Long, Optional ClosedON As String) As Boolean
Dim ret As Integer
Dim Rst As ADODB.Recordset

'Query Database
    gDbTrans.SqlStmt = "Select * from CAmaster where " & _
                        "AccID = " & AccId
    ret = gDbTrans.Fetch(Rst, adOpenForwardOnly)
    If ret <= 0 Then Exit Function
    
    If ret > 1 Then  'Screwed case
        'MsgBox "Data base curruption !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(601), vbExclamation, gAppName & " - Error"
        Set Rst = Nothing
        Exit Function
    End If
    
'Check the closed status
    If Not IsMissing(ClosedON) Then _
        ClosedON = FormatField(Rst("ClosedDate"))
    Set Rst = Nothing

    CAAccountExists = True

End Function


Public Function ComputeCAInterest(ByVal Product As Currency, Rate As Double) As Currency
    ComputeCAInterest = (Product * 1 * Rate) / (100 * 12)
End Function
'
Public Function ComputeCAProducts(AccIDArr() As Long, Mon As Integer, Yr As Integer, ByRef Products() As Currency) As Currency

Dim newfrmCancel As New frmCancel
Dim I As Long
Dim rst_Main As ADODB.Recordset
Dim rst_PM As ADODB.Recordset
Dim rst_1_10 As ADODB.Recordset
Dim rst_11_30 As ADODB.Recordset
Dim TransDate1 As String
Dim TransDate2 As String

'Validate month and year
If Mon < 1 Or Mon > 12 Then Exit Function

On Error Resume Next
'newfrmCancel.Show
On Error GoTo 0
newfrmCancel.PicStatus.Visible = True
newfrmCancel.Move Screen.Width * 3 / 4, Screen.Height * 1 / 4

'First get rec set of all the accounts in cA
    newfrmCancel.lblMessage = " Calulating no of accounts"
    gDbTrans.SqlStmt = "Select AccID, ClosedDate from CAMaster order by AccID"
    If gDbTrans.Fetch(rst_Main, adOpenStatic) < 1 Then Set rst_Main = Nothing

    DoEvents
    If gCancel Then Exit Function

'Get Balance upto Last Day of PrevMonth
'This query gets the maximum transaction performed on or before the
'last day of the previous month specified for every account
    
    TransDate1 = GetSysFormatDate("1/" & Mon & "/" & Yr)
    TransDate1 = DateAdd("d", -1, TransDate1)
    gDbTrans.SqlStmt = "Select * from CATrans as A where TransID = " & _
                       " (Select MAX(TransID) from CATrans B where " & _
                        " A.AccID = B.AccID and TransDate <= #" & TransDate1 & _
                        "#) order by AccID"
    If gDbTrans.Fetch(rst_PM, adOpenStatic) < 1 Then Set rst_PM = Nothing

    DoEvents
    If gCancel Then Exit Function

'Get max Balance between 1 to 10
    TransDate1 = GetSysFormatDate("1/" & Mon & "/" & Yr)
    TransDate2 = GetSysFormatDate("10/" & Mon & "/" & Yr)

'Build the Query for max trans between 1 and 10
    gDbTrans.SqlStmt = "Select * from CATrans A where TransDate = " & _
                        "(Select MAX (TransDate) from CATrans B " & _
                        " where A.AccID = B.AccID and TransDate between #" & _
                        TransDate1 & "# and #" & TransDate2 & "#) order by AccID"
    
    If gDbTrans.Fetch(rst_1_10, adOpenStatic) < 1 Then Set rst_1_10 = Nothing
    
    DoEvents
    If gCancel Then Exit Function

'Get balance upto LAST DAY of current month
'Set the TransCction dates according to month
    TransDate1 = GetSysFormatDate("11/" & Mon & "/" & Yr)
    TransDate2 = GetSysLastDate(TransDate1)          'Last Day day the Month)

    gDbTrans.SqlStmt = "Select * from CATrans A where Balance = " & _
                        "(Select MIN(Balance) from CATrans B " & _
                        " where A.AccID = B.AccID and TransDate between #" & _
                        TransDate1 & "# and #" & TransDate2 & "#) order by AccID"
    
    If gDbTrans.Fetch(rst_11_30, adOpenStatic) < 1 Then Set rst_11_30 = Nothing

    DoEvents
    If gCancel Then Exit Function

'Get 1 and Last Day of current month
    TransDate1 = GetSysFormatDate("1/" & Mon & "/" & Yr)
    'TransDate2 =  :ast Day of the Current month

    Dim transType As wisTransactionTypes
    Dim rst_Trans As ADODB.Recordset
    Dim rst_Trans_Count As ADODB.Recordset

    transType = wWithdraw

'Get transactions for all the accounts made during that month
    gDbTrans.SqlStmt = "Select AccID, TransDate from CATrans where " & _
                        "TransDate >= #" & TransDate1 & "# and " & _
                        "TransDate <= #" & TransDate2 & "# and " & _
                        "TransType = " & transType & _
                        " order by AccID, TransDate, TransID"
    If gDbTrans.Fetch(rst_Trans, adOpenStatic) < 1 Then Set rst_Trans = Nothing

    DoEvents
    If gCancel Then Exit Function

'Get Count of transactions for all the accounts made during that month
    gDbTrans.SqlStmt = "Select Count(*) as TotalTrans, AccID from CATrans where " & _
                        "TransDate >= #" & TransDate1 & "# and " & _
                        "TransDate <= #" & TransDate2 & "# and " & _
                        " TransType = " & transType & " group by AccID"
    If gDbTrans.Fetch(rst_Trans_Count, adOpenStatic) < 1 Then _
        Set rst_Trans_Count = Nothing

    DoEvents
    If gCancel Then Exit Function
 
Dim Balance As Currency
Dim AccId As Long
Dim ClosedDate As String
Dim Day7 As Date, Day14 As Date, Day21 As Date, Day30 As Date
Day7 = GetSysFormatDate("7/" & Mon & "/" & Yr)
Day14 = GetSysFormatDate("14/" & Mon & "/" & Yr)
Day21 = GetSysFormatDate("21/" & Mon & "/" & Yr)
Day30 = GetSysLastDate(Day21)  'Previous day of (1st of Next Month)
Dim Count7 As Integer, Count14 As Integer, Count21 As Integer, Count30 As Integer

'Loop through all the accounts to calculate the products
If rst_Main Is Nothing Then Exit Function
    For I = 1 To rst_Main.RecordCount
        Balance = 0
        AccId = FormatField(rst_Main("AccID"))
        If Not rst_PM Is Nothing Then
            If Not rst_PM.EOF Then
                If AccId = FormatField(rst_PM("AccID")) Then
                    Balance = FormatField(rst_PM("Balance"))
                    rst_PM.MoveNext
                End If
            End If
        End If
        If Not rst_1_10 Is Nothing Then
            If Not rst_1_10.EOF Then
                If AccId = FormatField(rst_1_10("AccID")) Then
                    Balance = FormatField(rst_1_10("Balance"))
                    rst_1_10.MoveNext
                End If
            End If
        End If
        
        If Not rst_11_30 Is Nothing Then
            If Not rst_11_30.EOF Then
                If AccId = FormatField(rst_11_30("AccID")) Then
                    Balance = IIf(Balance < FormatField(rst_11_30("Balance")), Balance, FormatField(rst_11_30("Balance")))
                    rst_11_30.MoveNext
                End If
            End If
        End If
        
        Dim TotalTrans As Integer
        Dim CheckOut As Boolean
        Dim MaxTrans As Integer
        CheckOut = False
        TotalTrans = 0
        MaxTrans = 2
        Count7 = 0: Count14 = 0: Count21 = 0: Count30 = 0
        If Not rst_Trans_Count Is Nothing Then
            If Not rst_Trans_Count.EOF Then
                If AccId = FormatField(rst_Trans_Count("AccID")) Then
                    TotalTrans = FormatField(rst_Trans_Count("TotalTrans"))
                    If TotalTrans > MaxTrans * 4 Then
                        'Set balance = 0
                        Balance = 0
                    ElseIf TotalTrans <= MaxTrans Then
                        'Do nothing
                    Else
                        CheckOut = True
                    End If
                    rst_Trans_Count.MoveNext
                End If
            End If
        End If
        
        If CheckOut Then 'Traverse thro the rec set
            If Not rst_Trans Is Nothing Then
                If Not rst_Trans.EOF Then
                    Do
                        If rst_Trans.EOF Then
                            Exit Do
                        End If
                        If FormatField(rst_Trans("AccID")) <> AccId Then
                            Exit Do
                        End If
                        
                        If Not rst_Trans.EOF Then
                            If DateDiff("d", Day7, rst_Trans("TransDate")) <= 0 Then
                                Count7 = Count7 + 1
                            ElseIf DateDiff("d", Day14, rst_Trans("TransDate")) <= 0 Then
                                Count14 = Count14 + 1
                            ElseIf DateDiff("d", Day21, rst_Trans("TransDate")) <= 0 Then
                                Count21 = Count21 + 1
                            Else
                                Count30 = Count30 + 1
                            End If
                            rst_Trans.MoveNext
                        End If
                    Loop
                End If
            End If
        Else
            'Reach the current account number
            If Not rst_Trans Is Nothing Then
                If Not rst_Trans.EOF Then
                    If AccId = FormatField(rst_Trans("AccID")) Then
                        rst_Trans.Move TotalTrans
                    End If
                End If
            End If
        End If
        
        If Count7 > MaxTrans Or Count14 > MaxTrans Or Count21 > MaxTrans Or Count30 > MaxTrans Then
            Balance = 0
        End If
        
        'Check for closure
        If Not IsNull(rst_Main("ClosedDate")) Then  'Account has been closed
            ClosedDate = rst_Main("ClosedDate")
            If (Yr = Year(ClosedDate) And Mon >= Month(ClosedDate)) Or Yr > Year(ClosedDate) Then
                Balance = -1
            End If
        End If
        
        AccIDArr(UBound(AccIDArr)) = AccId
        Products(UBound(Products)) = Balance
        ReDim Preserve AccIDArr(UBound(AccIDArr) + 1)
        ReDim Preserve Products(UBound(Products) + 1)
     
        rst_Main.MoveNext
    
    Next I

End Function
Public Function ComputeTotalCALiability(AsOnDate As Date) As Currency

Dim ret As Long
Dim Rst As ADODB.Recordset

ComputeTotalCALiability = 0

'''''Changed By Shashi on 23/11/2001****************************

gDbTrans.SqlStmt = "Select Sum(Balance) FROm CATrans A,CAMaster B WHERE " & _
            " B.AccID = A.AccID AND TransID = " & _
            "(Select MAX(TransId) from CATrans C " & _
            " where C.AccID = A.AccID and TransDate <= " & _
            "#" & AsOnDate & "#)"

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Function

ComputeTotalCALiability = FormatField(Rst(0))
Set Rst = Nothing

End Function

