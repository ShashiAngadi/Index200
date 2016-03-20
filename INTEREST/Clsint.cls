VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Dim m_ModuleID As wisModules
'Dim m_SchemeName As String

Dim m_InterestRate() As Single
Dim m_NoOfInterest As Integer

Dim M_Date() As String
Dim M_NoOfDate As Integer

Dim m_PresentInterestNum As Integer
Dim m_PresentDateNum As Integer

Public Function UndoLastInterest(ModuleID As Integer, _
                        SchemeName As String, InterestRate As Single, _
                        StartDate As Date) As Boolean

On Error GoTo ErrLine
'undo the last transaction
Dim TransID As Long
Dim Rst As ADODB.Recordset
'get the max transid for the particular Moduleid
'fire SQL

gDbTrans.SqlStmt = "SELECT Top 1 * from InterestTab" & _
            " where ModuleID =" & ModuleID & _
            " AND SchemeName = " & AddQuotes(SchemeName, True)
        
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then GoTo ErrLine
TransID = Val(FormatField(Rst(0)))

gDbTrans.SqlStmt = "Delete * from InterestTab" & _
            " where ModuleID =" & ModuleID & _
            " AND SchemeName = " & AddQuotes(SchemeName, True) & _
            " AND EndDate is Null"

If Not gDbTrans.SQLExecute Then GoTo ErrLine
 
gDbTrans.SqlStmt = "SELECT MAX (TransID) from InterestTab  where ModuleID =" & ModuleID & _
        " AND SchemeName = " & AddQuotes(SchemeName, True)
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then GoTo ExitLine

TransID = Val(FormatField(Rst(0)))

If TransID > 100 Then
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = "Update InterestTab Set EndDate = NULL " & _
        " Where ModuleID =" & ModuleID & _
        " AND SchemeName = " & AddQuotes(SchemeName, True) & _
        " And StartDate = (Select Max(StartDate) From InterestTab" & _
                " Where ModuleID =" & ModuleID & _
                " AND SchemeName = " & AddQuotes(SchemeName, True) & ")"

    If Not gDbTrans.SQLExecute Then GoTo ErrLine
    
End If

ExitLine:
UndoLastInterest = True
Exit Function

ErrLine:
    MsgBox "Error In Undo Interest Class", vbExclamation, gAppName & " - Error"
    'MsgBox GetResourceString(533), vbExclamation, gAppName & " - Error"
End Function

Public Sub ClearInterest()
    m_NoOfInterest = 0
    M_NoOfDate = 0
    m_PresentInterestNum = 0
    m_PresentDateNum = 0
    ReDim m_InterestRate(0)
    ReDim M_Date(0)
    
End Sub

'This Function Will give the Rate of interest existed int the Period specified
'It Also set the the interest Rate changed Period and Number of Changes
'It Also sets the Next Interest Date
Public Function InterestRate(ByVal ModuleID As Integer, _
                ByVal SchemeName As String, ByVal fromDate As Date, _
                Optional ByVal toDate As Date, Optional IsEmployee As Boolean, _
                Optional IsSeniorCitizen As Boolean) As Single
      
Err.Clear
On Error GoTo ErrLine
Dim Rst As ADODB.Recordset


If Trim$(SchemeName) = "" Then _
      Err.Raise 5003, "Interest Class", "SchemeName not set"

m_NoOfInterest = 0
m_PresentInterestNum = 0
M_NoOfDate = 0
If IsDate(toDate) Then toDate = CDate(gStrDate)

Dim Retval As Integer
Dim strIntRate As String
Dim pos As Integer
ReDim m_InterestRate(0)
ReDim M_Date(0)

'First Get The Rate Of Interest As On From indianDate
gDbTrans.SqlStmt = "Select Top 1 * from InterestTab " & _
            " Where StartDate <= #" & fromDate & "#" & _
            " And ModuleID = " & ModuleID & _
            " And SchemeName = " & AddQuotes(SchemeName, True) & _
            " Order by StartDate Desc "
            
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then
    'if the no interest spcified between nt period specified then
    'get the exitin interest
     gDbTrans.SqlStmt = "Select Top 1 * from InterestTab" & _
            " Where ModuleID = " & ModuleID & _
            " And SchemeName = " & AddQuotes(SchemeName, True) & _
            " Order by StartDate "
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then GoTo ExitLine
End If

m_NoOfInterest = 1
M_NoOfDate = 1
strIntRate = FormatField(Rst("InterestRate"))
pos = InStr(1, strIntRate, ";")
m_InterestRate(0) = Val(strIntRate)
If IsEmployee Then m_InterestRate(0) = Val(Mid(strIntRate, pos + 1))
If IsSeniorCitizen Then m_InterestRate(0) = Val(Mid(strIntRate, InStr(pos + 1, strIntRate, ";")))

'Check Whether Interest between Given Period
gDbTrans.SqlStmt = "Select * from InterestTab " & _
            " Where ModuleID = " & ModuleID & _
            " And SchemeName = " & AddQuotes(SchemeName, True) & _
            " And StartDate > #" & fromDate & "# " & _
            " And StartDate < #" & DateAdd("d", 1, toDate) & "#" & _
            " Order by StartDate Asc"
      
Retval = gDbTrans.Fetch(Rst, adOpenDynamic)
m_NoOfInterest = m_NoOfInterest + Retval
M_NoOfDate = M_NoOfDate + Retval
m_PresentInterestNum = 1
 
Retval = 1
ReDim Preserve m_InterestRate(Retval)
ReDim Preserve M_Date(Retval)

If Not Rst Is Nothing Then
    While Not Rst.EOF   'formatfield(Rst("
        
        strIntRate = FormatField(Rst("Interestrate"))
        m_InterestRate(Retval) = Val(strIntRate)
        pos = InStr(1, strIntRate, ";")
        If IsEmployee Then m_InterestRate(Retval) = Val(Mid(strIntRate, pos + 1))
        If IsSeniorCitizen Then m_InterestRate(Retval) = Val(Mid(strIntRate, InStr(pos + 1, strIntRate, ";")))
        
        M_Date(Retval) = Rst("StartDate")
        Retval = Retval + 1
        Rst.MoveNext
    Wend
End If
InterestRate = m_InterestRate(0)

m_PresentDateNum = LBound(M_Date)
M_NoOfDate = UBound(M_Date)

m_PresentInterestNum = LBound(m_InterestRate)
m_NoOfInterest = UBound(m_InterestRate)

ExitLine:
      Exit Function

ErrLine:

    If Err Then _
         MsgBox Err.Number & vbCrLf & "Error : " & Err.Description, vbExclamation, wis_MESSAGE_TITLE
    
   'Resume

End Function

Public Function LastInterestRate() As Integer
'Gives The Last Interest Added On Given Two Periods
'use InterestRate function Before Using This

If m_NoOfInterest > 0 Then
   LastInterestRate = m_NoOfInterest
   Exit Function
End If

LastInterestRate = 0
End Function

'This Function returnts the NextInterestDate in system Format between specified date
'Before Calling This Function You must call InterestRate functin
'which Sets the Array of INterest rates
'And Array of Dates when InterestRate Has been Changed
'If there is no next date then it returns the NULL date
Public Function NextInterestDate() As Date
   
   On Error GoTo ExitLine
    
    NextInterestDate = vbNull
    'Check Whether before calling this Function
    'The InterestRate Function Has called or not
    If m_NoOfInterest < 0 Then _
            Err.Raise 5002, "Interest Class", "Invalid procedure Call"
    
    If m_PresentDateNum = M_NoOfDate Then GoTo ExitLine
    
    'Returnt the Next Interest Rate from  Interestrate array
    m_PresentDateNum = m_PresentDateNum + 1
    NextInterestDate = m_InterestRate(m_PresentDateNum)
    NextInterestDate = M_Date(m_PresentDateNum)

ExitLine:

End Function

'This Function returnts the NextInterestRate between specified date
'Before Calling This Function You must call InterestRate functin
'which Sets the Array of INterest rates
'If there is  no next interestrate it returns the then it returns 0
Public Function NextInterestRate() As Single

On Error GoTo ExitLine
    'Check Whether before calling this Function
    'The InterestRate Function Has called or not
    If m_NoOfInterest < 0 Then _
            Err.Raise 5002, "Interest Class", "Invalid procedure Call"
    
    If m_PresentInterestNum = m_NoOfInterest Then GoTo ExitLine
    'Return the Next Interest Rate from  Interestrate array
    m_PresentInterestNum = m_PresentInterestNum + 1
    NextInterestRate = m_InterestRate(m_PresentInterestNum)
      
ExitLine:

End Function

Public Property Get ObjectHandle()
    ObjectHandle = 0
End Property
Public Property Get ObjectName()
    ObjectName = "InterestClass"
End Property

'This Function will write the interest rate to the Interest Tab
'And Interst Rate will be in % not in fraction, And Fraction will convert to %
'In the module id we can send the
'any deposit with theie module id
'or any load with their SchemID+wis_loans(i.e. 300)
Public Function SaveInterest(ModuleID As Integer, SchemeName As String, _
        ByVal InterestRate As Single, Optional ByVal EmpIntRate As Single, _
        Optional ByVal SeniorIntRate As Single, Optional ByVal OnDate As Date) As Boolean
      
On Error GoTo ErrLine
     
    If ModuleID = 0 Then GoTo ErrLine
    If SchemeName = "" Then GoTo ErrLine
    If InterestRate <= 0 Then Exit Function
    
    If EmpIntRate = 0 Then EmpIntRate = InterestRate
    If SeniorIntRate = 0 Then SeniorIntRate = InterestRate
    
    If IsMissing(OnDate) Then OnDate = GetSysFormatDate(gStrDate)
    
    If InterestRate < 1 Then InterestRate = InterestRate * 100
    If EmpIntRate < 1 Then EmpIntRate = EmpIntRate * 100
    If SeniorIntRate < 1 Then SeniorIntRate = SeniorIntRate * 100
    
    Dim strInterest As String
    Dim Rst As ADODB.Recordset
    Dim fromDate As Date
    Dim TransID As Long
    Dim NewRecord As Boolean
    
    NewRecord = False
         
    strInterest = InterestRate & ";" & EmpIntRate & ";" & SeniorIntRate
    
    'Get the last update date of the Interest
    gDbTrans.SqlStmt = "Select * From InterestTab " & _
                  " Where SchemeName = " & AddQuotes(SchemeName, True) & _
                  " And Moduleid = " & ModuleID & " Order By TransID Desc"
    TransID = 0
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
        NewRecord = False
        fromDate = Rst("StartDate")
        TransID = FormatField(Rst("TransID"))
        'First UpDate the DataBaseTab For The Interest Date End
        gDbTrans.SqlStmt = "Update InterestTab " & _
               " Set EndDate = #" & DateAdd("d", -1, OnDate) & "# " & _
               " Where SchemeName = " & AddQuotes(SchemeName, True) & _
               " And ModuleID = " & ModuleID & _
               " And EndDate is NULL ANd TransID = " & TransID
        If DateDiff("D", fromDate, OnDate) = 0 Then
            If UCase(strInterest) = UCase(FormatField(Rst("InterestRate"))) Then SaveInterest = True: Exit Function
            NewRecord = True
            'First UpDate the DataBaseTab For The Interest Date End
            gDbTrans.SqlStmt = "Update InterestTab " & _
                   " Set InterestRate = " & AddQuotes(strInterest) & _
                   " Where SchemeName = " & AddQuotes(SchemeName, True) & _
                   " And ModuleID = " & ModuleID & _
                   " And EndDate is NULL ANd TransID = " & TransID
        End If
    Else
        NewRecord = True
        TransID = 1
        fromDate = CDate("1/1/1900")
        Set Rst = Nothing
    End If
    
    'Check For the Date if he is Entering the
    'Date earlier than the already transacted date
    If DateDiff("D", fromDate, OnDate) < 0 Then GoTo ExitLine
    
'Now Insert the same too the
'BEGIN the transaction
    gDbTrans.BeginTrans
          
    
    If Not Rst Is Nothing Then
        'Update the prevous record
        If Not gDbTrans.SQLExecute Then GoTo ExitLine
        'If we have to insert new record then
        TransID = TransID + 1
    End If
    If NewRecord Then
        'now insert the Present interest rate effective from GivenDate
        gDbTrans.SqlStmt = "Insert Into InterestTab (TransID,ModuleID," & _
                "SchemeName,InterestRate,StartDate) Values " & _
                "(" & TransID & "," & ModuleID & ", " & AddQuotes(SchemeName, True) & "," & _
                AddQuotes(strInterest) & ", #" & OnDate & "# )"
                
        If Not gDbTrans.SQLExecute Then GoTo ExitLine
    End If
    
'END the transaction
    gDbTrans.CommitTrans
    
SaveInterest = True

Exit Function

ExitLine:

gDbTrans.RollBack

Exit Function

ErrLine:
      MsgBox "Error In SaveInterest - Interest Class", , wis_MESSAGE_TITLE

End Function
