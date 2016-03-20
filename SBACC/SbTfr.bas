Attribute VB_Name = "SbTransfer"
'This BAs file is used to Transfer
'SbMaster & Sb TranscTION dETAILS
'FROM oLD DATABASE TO NEW DATA BASE
Option Explicit

'just calling this function we can transafer the sbmaster Old to new
'Arguments for this function are OldSbTrans & new SbTrans
'Old sb Trans is assigned to Old database
'and NewSBTrans has assigned to new database
Public Function TransferSB(OldDBName As String, NewDBName As String) As Boolean
Screen.MousePointer = vbHourglass
Dim OldTrans As clsDBUtils
Dim NewTrans As clsDBUtils

Set OldTrans = New clsDBUtils
Set NewTrans = New clsDBUtils

If Not OldTrans.OpenDB(OldDBName, "PRAGMANS") Then Exit Function
If Not NewTrans.OpenDB(NewDBName, "WIS!@#") Then
    OldTrans.CloseDB
    Exit Function
End If

    If Not TransferSBMaster(OldTrans, NewTrans) Then GoTo ErrLine
    gDBTrans.CloseDB
    Call NewTrans.WISCompactDB(NewDBName, "WIS!@#", "WIS!@#")
    If Not TransferSBTrans(OldTrans, NewTrans) Then GoTo ErrLine
    TransferSB = True

ErrLine:

OldTrans.CloseDB
NewTrans.CloseDB
Set OldTrans = Nothing
Set NewTrans = Nothing
Call gDBTrans.OpenDB(NewDBName, "WIS!@#")
Screen.MousePointer = vbNormal

End Function

'this function is used to transfer the
'SB MAster details form OLdb to new one
'and NewSBTrans has assigned to new database
Private Function TransferSBMaster(OldSBTrans As clsDBUtils, NewSBTrans As clsDBUtils) As Boolean
Dim SqlStr As String
Dim SngSpace As String
Dim ACCID As Long, IntroId As Long
Dim rstMain As ADODB.Recordset
Dim rst As ADODB.Recordset

On Error GoTo Err_Line

'Fetch the detials of Sb Account
SqlStr = "SELECT * FROM SBMASTER ORDER BY AccID"

OldSBTrans.SQLStmt = SqlStr
If OldSBTrans.Fetch(rstMain, adOpenForwardOnly) < 1 Then Exit Function
While Not rstMain.EOF
    If ACCID = FormatField(rstMain("AccID")) Then GoTo NextAccount
    
    IntroId = FormatField(rstMain("Introduced"))
    'Get the Introducer ID
    If IntroId > 0 Then
        SqlStr = "SELECT CustomerID FROM SBMASTER " & _
            " WHERE AccID = " & IntroId
        OldSBTrans.SQLStmt = SqlStr
        If OldSBTrans.Fetch(rst, adOpenForwardOnly) > 0 Then IntroId = FormatField(rst("CustomerID"))
    End If
    
    ACCID = rstMain("AccID")
    'First insert into Sb joint table
'    SqlStr = "Insert INTO SBJOINT (" & _
'        "AccID,CustomerID,CustomerNum)" & _
'        "VALUES (" & _
'        rstMain("AccID") & "," & rstMain("CustomerID") & "," & _
'        "1 )"
        
    NewSBTrans.BeginTrans
'    NewSBTrans.SQLStmt = SqlStr
'    If Not NewSBTrans.SQLExecute Then
'        NewSBTrans.RollBack
'        MsgBox "Unable to transafer the SB MAster data"
'        NewSBTrans.RollBack
'        Exit Function
'    End If
    
'NOW insert into
    SqlStr = "Insert INTO SBMASTER (" & _
        "CustomerID,AccNUM,CreateDate,ModifiedDate,ClosedDate," & _
        "JointHolder,Nominee,Introduced,LedgerNo,FolioNo ," & _
        "AccGroup, NomineeID ,InOperative,LastPrintId )"
    
    SqlStr = SqlStr & " VALUES (" & _
        rstMain("CustomerID") & "," & _
        AddQuotes(rstMain("AccID"), True) & "," & _
        FormatDateField(rstMain("CreateDate")) & "," & FormatDateField(rstMain("ModifiedDate")) & "," & _
        FormatDateField(rstMain("ClosedDate")) & "," & _
        AddQuotes(FormatField(rstMain("JointHolder")), True) & ", " & _
        AddQuotes(Left(FormatField(rstMain("Nominee")), 25), True) & "," & _
        rstMain("Introduced") & ",'" & Val(rstMain("LedgerNo")) & "','" & Val(rstMain("FolioNo")) & "' ," & _
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
    rstMain.MoveNext
Wend

TransferSBMaster = True
Exit Function

Err_Line:
    If Err Then MsgBox "eror In SBMaster " & Err.Description
    
End Function

'this function is used to transfer the
'SB transaction details form OLd Db to new one
'and NewSBTrans has assigned to new database
Private Function TransferSBTrans(OldSBTrans As clsDBUtils, NewSBTrans As clsDBUtils) As Boolean

On Error GoTo Err_Line

Dim SqlStr As String
Dim IsIntTrans As Boolean
Dim OldTrans As wisTransactionTypes, NewTransType As Integer
Dim Balance As Currency
Dim TransID As Long
Dim rstMain As ADODB.Recordset
Dim ACCID As Long
Dim Amount As Currency
Dim TransDate As Date
'Fetch the detials of Sb Account

SqlStr = "SELECT * FROM SBTrans ORDER BY AccID,TransId"
OldSBTrans.SQLStmt = SqlStr
If OldSBTrans.Fetch(rstMain, adOpenStatic) < 1 Then Exit Function

While Not rstMain.EOF
    IsIntTrans = False
    OldTrans = FormatField(rstMain("TransType"))
    TransID = FormatField(rstMain("TransID"))
    If OldTrans = 4 Or OldTrans = 2 Then IsIntTrans = True
    If OldTrans = -2 Or OldTrans = -4 Then IsIntTrans = True
    
    NewTransType = OldTrans
    If OldTrans = 1 Then NewTransType = wDeposit
    If OldTrans = -1 Then NewTransType = wWithDraw
    If OldTrans = 3 Then NewTransType = wContraDeposit
    If OldTrans = -3 Then NewTransType = wContraWithDraw
    
    If IsIntTrans Then
        If OldTrans = 2 Or OldTrans = 4 Then NewTransType = wContraDeposit
        If OldTrans = -2 Or OldTrans = -4 Then NewTransType = wContraWithDraw
        
        TransDate = rstMain("TransDate")
        
        SqlStr = "Insert INTO SBPLTrans ( " & _
            "AccID,TransID,TransDate," & _
            "Amount,Balance,Particulars," & _
            "TransType )"
        SqlStr = SqlStr & "VALUES (" & _
            rstMain("AccID") & "," & _
            TransID & "," & _
            "#" & rstMain("TransDate") & "#," & _
            rstMain("Amount") & "," & _
            "0  ," & _
            AddQuotes(FormatField(rstMain("Particulars")), True) & "," & _
            NewTransType & " )"
        NewSBTrans.BeginTrans
        NewSBTrans.SQLStmt = SqlStr
        
        If Not NewSBTrans.SQLExecute Then
            MsgBox "Unable to transafer the SB MAster data"
            NewSBTrans.RollBack
            Exit Function
        End If
        NewSBTrans.CommitTrans
        Amount = rstMain("Amount")
        ACCID = rstMain("AccID")
        'if balance of the Last transaction and this is same then
        'it has two transaction one for Profit and other for Receipt
        
        'If Balance = FormatField(Rst("balance")) Then Rst.MoveNext
        
        'Insted of the above code the beleow code is written for
        'shiggaon case onmly
        'Except shiggaon in all banks the aboce code works 100%
        Do
            If TransDate <> rstMain("transdate") Or ACCID <> rstMain("AccId") Then
                rstMain.MovePrevious
                Exit Do
            End If
            If OldTrans > 0 And Balance + Amount = rstMain("Balance") Then Exit Do
            If OldTrans < 0 And Balance - Amount = rstMain("Balance") Then Exit Do
            rstMain.MoveNext
        Loop
        OldTrans = rstMain("TransType")
        'After this transaction the transaction in the sb Table is contra
        'Therefore
        NewTransType = (OldTrans / Abs(OldTrans)) * 3
        If OldTrans = 1 Then NewTransType = wDeposit
        If OldTrans = -1 Then NewTransType = wWithDraw
        If OldTrans = 3 Then NewTransType = wContraDeposit
        If OldTrans = -3 Then NewTransType = wContraWithDraw
    End If
    
    SqlStr = "Insert INTO SBTrans ( " & _
        "AccID,TransID,TransDate," & _
        "Amount,Balance,Particulars," & _
        "TransType,ChequeNo)"
    
    SqlStr = SqlStr & "VALUES (" & _
        rstMain("AccID") & "," & _
        TransID & "," & _
        "#" & rstMain("TransDate") & "#," & _
        rstMain("Amount") & "," & _
        rstMain("Balance") & "," & _
        AddQuotes(FormatField(rstMain("Particulars")), True) & "," & _
        NewTransType & "," & FormatField(rstMain("ChequeNo")) & " )"
    
    NewSBTrans.BeginTrans
    NewSBTrans.SQLStmt = SqlStr
    If Not NewSBTrans.SQLExecute Then
        MsgBox "Unable to transafer the SB Trans data"
        NewSBTrans.RollBack
        Exit Function
    End If
    'If Rst.AbsolutePosition Mod 5000 = 0 Then Debug.Print Now
    NewSBTrans.CommitTrans
    Balance = FormatField(rstMain("Balance"))
NextAccount:
    rstMain.MoveNext
Wend
TransferSBTrans = True
Exit Function
Err_Line:
    If Err Then
        MsgBox "Error in SBTrans" & Err.Description
    End If
    
End Function


