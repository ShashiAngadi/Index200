Attribute VB_Name = "CaTransfer"
'This BAs file is used to Transfer
'SbMaster & Sb TranscTION dETAILS
'FROM oLD DATABASE TO NEW DATA BASE
Option Explicit

'just calling this function we can transafer the sbmaster Old to new
'Arguments for this function are OldSbTrans & new SbTrans
'Old sb Trans is assigned to Old database
'and NewSBTrans has assigned to new database
Public Function TransferCA(OldDBName As String, NewDBName As String) As Boolean
Debug.Print Now
Dim OldTrans As New clsDBUtils
Dim NewTrans As New clsDBUtils
If Not OldTrans.OpenDB(OldDBName, "PRAGMANS") Then Exit Function
If Not NewTrans.OpenDB(NewDBName, "WIS!@#") Then
    OldTrans.CloseDB
    Exit Function
End If

    If Not TransferCAMaster(OldTrans, NewTrans) Then Exit Function
    If Not TransferCATrans(OldTrans, NewTrans) Then Exit Function
    
    TransferCA = True
    
End Function


'this function is used to transfer the
'SB MAster details form OLdb to new one
'and NewSBTrans has assigned to new database
Private Function TransferCAMaster(OldSBTrans As clsDBUtils, NewSBTrans As clsDBUtils) As Boolean
Dim SqlStr As String

On Error GoTo Err_Line

'Fetch the detials of Sb Account
Dim ACCID As Long, IntroId As Long
Dim rst As ADODB.Recordset
Dim rstTemp As ADODB.Recordset

SqlStr = "SELECT * FROM CAMASTER ORDER BY AccID"
OldSBTrans.SQLStmt = SqlStr
If OldSBTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

While Not rst.EOF
    If ACCID = FormatField(rst("AccID")) Then GoTo NextAccount
    
    IntroId = FormatField(rst("Introduced"))
    'Get the Introducer ID
    If IntroId > 0 Then
        OldSBTrans.SQLStmt = "SELECT CustomerID FROM CAMASTER " & _
            "WHERE AccID = " & IntroId
        If OldSBTrans.Fetch(rstTemp, adOpenForwardOnly) > 0 Then _
            IntroId = FormatField(rstTemp("CustomerID"))
    End If
    
    ACCID = rst("AccID")
    'First insert into Sb joint table
    SqlStr = "Insert INTO CAJOINT (" & _
        "AccID,CustomerID,CustomerNum)" & _
        "VALUES (" & _
        rst("AccID") & "," & rst("CustomerID") & "," & _
        "1 )"
        
    NewSBTrans.BeginTrans
    NewSBTrans.SQLStmt = SqlStr
'    If Not NewSBTrans.SQLExecute Then
'        NewSBTrans.RollBack
'        MsgBox "Unable to transafer the CA MAster data"
'        NewSBTrans.RollBack
'        Exit Function
'    End If
    
'NOW insert into
    SqlStr = "Insert INTO CAMASTER (" & _
        "AccID,CustomerID,AccNUM,CreateDate,ModifiedDate,ClosedDate," & _
        "JointHolder,Nominee,Introduced,LedgerNo,FolioNo ," & _
        "AccGroup, NomineeID ,InOperative,LastPrintId )"
    
    SqlStr = SqlStr & " VALUES (" & _
        rst("AccID") & "," & rst("CustomerID") & "," & _
        AddQuotes(rst("AccID"), True) & "," & _
        FormatDateField(rst("CreateDate")) & "," & FormatDateField(rst("Modifieddate")) & "," & _
        FormatDateField(rst("ClosedDate")) & " ," & _
        AddQuotes(FormatField(rst("JointHolder")), True) & ", " & _
        AddQuotes(FormatField(rst("Nominee")), True) & "," & _
        rst("Introduced") & "," & Val(rst("LedgerNo")) & "," & Val(rst("FolioNo")) & " ," & _
        "'GEN' ,0 ," & _
        False & "," & _
        "0 )"
        
    NewSBTrans.SQLStmt = SqlStr
    If Not NewSBTrans.SQLExecute Then
        NewSBTrans.RollBack
        MsgBox "Unable to transafer the SB MAster data"
        Exit Function
    End If
    NewSBTrans.CommitTrans
    
NextAccount:
    rst.MoveNext
Wend

TransferCAMaster = True

Exit Function

Err_Line:
    If Err Then MsgBox "eror In CAMaster " & vbCrLf & Err.Description
    
End Function


'this function is used to transfer the
'SB transaction details form OLd Db to new one
'and NewSBTrans has assigned to new database
Private Function TransferCATrans(OldSBTrans As clsDBUtils, NewSBTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim IsIntTrans As Boolean

On Error GoTo Err_Line

Dim OldTrans As wisTransactionTypes, NewTransType As Integer
Dim Balance As Currency
Dim TransID As Long
Dim rst As ADODB.Recordset
Dim ACCID As Long
Dim Amount As Currency
Dim TransDate As Date
    'Fetch the detials of Sb Account

SqlStr = "SELECT * FROM CATrans ORDER BY AccID,TransId"
OldSBTrans.SQLStmt = SqlStr
If OldSBTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function

While Not rst.EOF
    IsIntTrans = False
    OldTrans = FormatField(rst("TransType"))
    TransID = FormatField(rst("TransID"))
'    If OldTrans = wContraInterest Or OldTrans = wInterest Then IsIntTrans = True
'    If OldTrans = wContraCharges Or OldTrans = wCharges Then IsIntTrans = True
    If OldTrans = 4 Or OldTrans = 2 Then IsIntTrans = True
    If OldTrans = -2 Or OldTrans = -4 Then IsIntTrans = True
    
    If OldTrans = 1 Then NewTransType = wDeposit
    If OldTrans = -1 Then NewTransType = wWithDraw
    If OldTrans = 3 Then NewTransType = wContraDeposit
    If OldTrans = -3 Then NewTransType = wContraWithDraw
    
    
    If IsIntTrans Then
'        If OldTrans = wInterest Then NewTransType = -1
'        If OldTrans = wContraInterest Then NewTransType = -3
'        If OldTrans = wCharges Then NewTransType = 3
'        If OldTrans = wContraCharges Then NewTransType = 1
        If OldTrans = 2 Or OldTrans = 4 Then NewTransType = wContraWithDraw
        If OldTrans = -2 Or OldTrans = -4 Then NewTransType = wContraDeposit
        
        TransDate = rst("TransDate")
        
        SqlStr = "Insert INTO CAPLTrans ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance,Particulars," & _
            "TransType )"
        SqlStr = SqlStr & "VALUES (" & _
            rst("AccID") & "," & _
            TransID & "," & _
            FormatDateField(rst("TransDate")) & "," & _
            rst("Amount") & "," & _
            "0  ," & _
            AddQuotes(FormatField(rst("Particulars")), True) & "," & _
            NewTransType & " )"
        NewSBTrans.BeginTrans
        NewSBTrans.SQLStmt = SqlStr
        
        If Not NewSBTrans.SQLExecute Then
            MsgBox "Unable to transafer the SB MAster data"
            NewSBTrans.RollBack
            Exit Function
        End If
        NewSBTrans.CommitTrans
        Amount = rst("Amount")
        ACCID = rst("AccID")
        'if balance of the Last transaction and this is same then
        'it has two transaction one for Profit and other for Receipt
        
'If Balance = FormatField(Rst("balance")) Then Rst.MoveNext
        
        'Insted of the above code the beleow code is written for
        'shiggaon case onmly
        'Except shiggaon in all banks the aboce code works 100%
        Do
            If TransDate <> rst("Transdate") Or ACCID <> rst("AccId") Then
                rst.MovePrevious
                Exit Do
            End If
            If OldTrans > 0 And Balance + Amount = rst("Balance") Then Exit Do
            If OldTrans < 0 And Balance - Amount = rst("Balance") Then Exit Do
            rst.MoveNext
        Loop
        OldTrans = rst("TransType")
        'After this transaction the transaction in the sb Table is contra
        'Therefore
        If OldTrans = 1 Then NewTransType = wDeposit
        If OldTrans = -1 Then NewTransType = wWithDraw
        If OldTrans = 3 Then NewTransType = wContraDeposit
        If OldTrans = -3 Then NewTransType = wContraWithDraw
    End If
    
    SqlStr = "Insert INTO CATrans ( " & _
        "AccID,TransID,TransDate," & _
        "Amount,Balance,Particulars," & _
        "TransType,ChequeNo)"
    
    SqlStr = SqlStr & "VALUES (" & _
        rst("AccID") & "," & _
        TransID & "," & _
        FormatDateField(rst("TransDate")) & "," & _
        rst("Amount") & "," & _
        rst("Balance") & "," & _
        AddQuotes(FormatField(rst("Particulars")), True) & "," & _
        NewTransType & "," & rst("ChequeNo") & " )"
    
    NewSBTrans.BeginTrans
    NewSBTrans.SQLStmt = SqlStr
    If Not NewSBTrans.SQLExecute Then
        MsgBox "Unable to transafer the SB Trans data"
        NewSBTrans.RollBack
        Exit Function
    End If
    'If Rst.AbsolutePosition Mod 5000 = 0 Then Debug.Print Now
    NewSBTrans.CommitTrans
    Balance = FormatField(rst("Balance"))
NextAccount:
    rst.MoveNext
Wend

TransferCATrans = True

Exit Function

Err_Line:
    If Err Then MsgBox "Error in SBTrans" & Err.Description
    
End Function
