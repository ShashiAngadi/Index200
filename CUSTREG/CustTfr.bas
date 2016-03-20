Attribute VB_Name = "CustTransfer"
Option Explicit

Public Function TransferNameTab(OldDBName As String, NewDBName As String) As Boolean
Dim OldTrans As New clsDBUtils
Dim NewTrans As New clsDBUtils
Dim SqlStr As String
Dim OldRst As ADODB.Recordset

If Not OldTrans.OpenDB(OldDBName, "PRAGMANS") Then Exit Function
If Not NewTrans.OpenDB(NewDBName, "WIS!@#") Then
    OldTrans.CloseDB
    Exit Function
End If

SqlStr = "SELECT * FROM NameTab WHERE CustomerID in " & _
    "(SELECT CustomerID FROM MMMASTER) ORDER BY CustomerID"

OldTrans.SQLStmt = SqlStr
If OldTrans.Fetch(OldRst, adOpenStatic) < 1 Then
    OldTrans.CloseDB
    NewTrans.CloseDB
    Exit Function
End If

'Now TransFer the Details of NAme Tab
Dim CustomerId As Long
CustomerId = 0
'On Error Resume Next
While Not OldRst.EOF
    If CustomerId = OldRst("CustomerID") Then GoTo NextName
    CustomerId = OldRst("CustomerID")
    SqlStr = "Insert INTO NameTab (CustomerID,Title,FirstName,MIddleNAme,LastName, " & _
            " Gender,Profession,Caste,DOB,MaritalStatus,HomeAddress,OfficeAddress," & _
            "HomePhone,Officephone,eMail,Place,Reference,IsciName) " '
    SqlStr = SqlStr & " Values (" & _
            CustomerId & "," & _
            AddQuotes(IIf(IsNull(OldRst("Title")), FormatField(OldRst("Title")), OldRst("Title")), True) & "," & _
            AddQuotes(Left(OldRst("FirstName"), 25), True) & "," & _
            AddQuotes(Left(OldRst("MiddleName"), 25), True) & "," & _
            AddQuotes(Left(OldRst("LastName"), 25), True) & "," & _
            OldRst("Gender") & "," & _
            AddQuotes(IIf(IsNull(OldRst("Profession")), FormatField(OldRst("Profession")), OldRst("Profession")), True) & "," & _
            AddQuotes(IIf(IsNull(OldRst("Caste")), FormatField(OldRst("Caste")), OldRst("Caste")), True) & "," & _
            FormatDateField(OldRst("DOB")) & "," & _
            OldRst("MaritalStatus") & "," & _
            AddQuotes(Left(OldRst("HomeAddress"), 25), True) & "," & _
            AddQuotes(Left(IIf(IsNull(OldRst("OfficeAddress")), FormatField(OldRst("OfficeAddress")), OldRst("OfficeAddress")), 49), True) & "," & _
            AddQuotes(IIf(IsNull(OldRst("HomePhone")), FormatField(OldRst("HomePhone")), OldRst("HomePhone")), True) & "," & _
            AddQuotes(IIf(IsNull(OldRst("OfficePhone")), FormatField(OldRst("OfficePhone")), OldRst("OfficePhone")), True) & "," & _
            AddQuotes(IIf(IsNull(OldRst("eMail")), FormatField(OldRst("eMail")), OldRst("eMail")), True) & "," & _
            AddQuotes(IIf(IsNull(OldRst("Place")), FormatField(OldRst("Place")), OldRst("Place")), True) & "," & _
            FormatField(OldRst("Reference")) & "," & _
            AddQuotes(OldRst("IsciName"), True) & _
            " );"
    NewTrans.BeginTrans
    NewTrans.SQLStmt = SqlStr
    If Not NewTrans.SQLExecute Then
        NewTrans.RollBack
        Exit Function
    End If
    NewTrans.CommitTrans

NextName:
    OldRst.MoveNext
Wend
Set OldRst = Nothing
On Error GoTo 0

'Now Insert the Caste Details
SqlStr = "SELECT * FROM Castetab"
OldTrans.SQLStmt = SqlStr
If OldTrans.Fetch(OldRst, adOpenStatic) < 1 Then
    OldTrans.CloseDB
    NewTrans.CloseDB
    Exit Function
End If

While Not OldRst.EOF
    SqlStr = "Insert INTO CasteTab (Caste) " & _
        " Values (" & AddQuotes(FormatField(OldRst("CASTE")), True) & ")"
    NewTrans.BeginTrans
    NewTrans.SQLStmt = SqlStr
    If Not NewTrans.SQLExecute Then
        NewTrans.RollBack
        Exit Function
    End If
    NewTrans.CommitTrans

NextCaste:
    OldRst.MoveNext
Wend


'Now Insert the Place Details
SqlStr = "SELECT * FROM PLACETab"
OldTrans.SQLStmt = SqlStr
If OldTrans.Fetch(OldRst, adOpenStatic) < 1 Then
    OldTrans.CloseDB
    NewTrans.CloseDB
    Exit Function
End If

While Not OldRst.EOF
    SqlStr = "Insert INTO PLACETab (PLACE) " & _
        " Values (" & AddQuotes(FormatField(OldRst("Places")), True) & ")"
    NewTrans.BeginTrans
    NewTrans.SQLStmt = SqlStr
    If Not NewTrans.SQLExecute Then
        NewTrans.RollBack
        Exit Function
    End If
    NewTrans.CommitTrans

NextPlace:
    OldRst.MoveNext
Wend


'Now Insert the user Details
SqlStr = "SELECT * FROM UserTab"
OldTrans.SQLStmt = SqlStr
If OldTrans.Fetch(OldRst, adOpenStatic) < 1 Then
    OldTrans.CloseDB
    NewTrans.CloseDB
    Exit Function
End If

Dim Perm As Long
While Not OldRst.EOF
    SqlStr = "Insert INTO UserTab (UserID,CustomerID," & _
        "Permissions,LoginName,PassWrd,Deleted) " & _
        " Values (" & OldRst("UserId") & "," & _
        OldRst("CustomerId") & "," & _
        OldRst("Permissions") & "," & _
        AddQuotes(FormatField(OldRst("LoginName")), True) & "," & _
        AddQuotes(FormatField(OldRst("PassWord")), True) & "," & _
        False & ")"
    NewTrans.BeginTrans
    NewTrans.SQLStmt = SqlStr
    If Not NewTrans.SQLExecute Then
        NewTrans.RollBack
        If OldRst.AbsolutePosition > 1 Then GoTo NextUser
        Exit Function
    End If
    NewTrans.CommitTrans

NextUser:
    OldRst.MoveNext
    
Wend

TransferNameTab = True
OldTrans.CloseDB
NewTrans.CloseDB
Set OldTrans = Nothing
Set NewTrans = Nothing

End Function
