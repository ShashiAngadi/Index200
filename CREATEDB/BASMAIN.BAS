Attribute VB_Name = "basMain"
Option Explicit

Public gAppPath As String
Public Const gAppName = "INDEX 2000"
Public gDBTrans As clsDataBase
Public gWorkDir As String

Public gCancel As Boolean
Public gMaxCount As Long
Public gStep As Integer

Private Type TabStruct
    Field As String
    Type As String
    Length As Integer
    'Index As Boolean
    Required As Boolean
    'Primary As Boolean
    AutoIncrement As Boolean
End Type

Private Type RelnStruct
    Field As String
    SourceTable As String
    ForiegnTable As String
End Type
Public Function MakeDatabase(strdataFile As String, DBPath As String) As Boolean
#If Database Then

Dim DBName As String
Dim strRet As String
Dim dbFile As String
Dim StrPwd As String
Dim i As Integer
Set gDBTrans = New clsTransact
gWorkDir = App.Path
    ' Read the dbname from datafile.
    i = 1
    strRet = ReadFromIniFile("Databases", "DataBase" & i, strdataFile)
    If strRet = "" Then Exit Function
    ' Get the name of the database file.
    DBName = ExtractToken(strRet, "dbName")
    If DBName = "" Then Exit Function
    
    ' Check fotr thepath where database to create
    If DBPath = "" Then Exit Function

    ' Check if the file path of the database
    ' is existing.  If not create it.
    If Dir(DBPath, vbDirectory) = "" Then
        If Not MakeDirectories(DBPath) Then
            MsgBox "Error in creating the path " & DBPath _
                & " for " & DBName, vbCritical
            GoTo dbCreate_err
        End If
    Else
        ' Check if the file is already existing.
        ' If existing, get the user action.
        dbFile = StripExtn(dbFile) & ".mdb"
        If Dir(dbFile, vbNormal) <> "" Then
        Dim nRet As Integer
            'nRet = MsgBox("WARNING : " & vbCrLf & vbCrLf & "The database file '" _
                    & dbFile & "' is already existing.  If you choose to overwrite " _
                    & "this file, you will loose the existing data permanantly." _
                    & vbCrLf & vbCrLf & "Do you want overwrite this file?", _
                    vbYesNo + vbCritical + vbDefaultButton2)
            nRet = MsgBox("WARNING : " & vbCrLf & vbCrLf & "The database file '" _
                    & dbFile & "' is already existing.  If you choose to overwrite " _
                    & "this file, you will loose the existing data permanantly." _
                    & vbCrLf & vbCrLf & "Do you want overwrite this file?", _
                    vbYesNo + vbCritical + vbDefaultButton2)
            If nRet = vbYes Then
                ' Delete the existing file.
                Kill dbFile
            ElseIf nRet = vbNo Then
                GoTo dbCreate_err
            End If
        End If
    End If

    ' Create the database.
    'RaiseEvent  UpdateStatus( "Creating the database " & dbFile &  "...")
    Dim Db As Database
    Dim strLocale As String
    If Trim$(StrPwd) = "" Then
        Set Db = CreateDatabase(dbFile, dbLangGeneral)
    Else
        Set Db = CreateDatabase(dbFile, dbLangGeneral & ";pwd=" & StrPwd)
    End If

'Now Read File For The Tables & Filed
Dim TabFile As String
i = 1
Do
    'Read From The specified file to create the tables
    TabFile = ReadFromIniFile("Files", "FileName" & i, strdataFile)
    If TabFile = "" Then Exit Do
    ' Create the specified tables for this db.
    Dim J As Byte
    J = 1
    Do
        ' Read the table name.
        Dim strTblName As String
        Dim tblData() As TabStruct
        strTblName = ReadFromIniFile(StripExtn(DBName), "Table" & J, TabFile)
        If strTblName = "" Then Exit Do

        ' Read the field details for this table into an array.
        Dim K As Byte
        K = 0
        ReDim tblData(0)
        Do
            strRet = ReadFromIniFile(strTblName, _
                        "Field" & K + 1, strdataFile)
            If strRet = "" Then Exit Do

            ' Add to fields array.
            ReDim Preserve tblData(K)
            With tblData(K)
                ' Set the field name.
                .Field = ExtractToken(strRet, "FieldName")
                ' Set the field type.
                .Type = FieldTypeNum(ExtractToken(strRet, "FieldType"))

                ' Set the field length.
                .Length = Val(ExtractToken(strRet, "Length"))
                ' Check, if the required flag is set.
                .Required = IIf((UCase$(ExtractToken(strRet, _
                        "Required")) = "TRUE"), True, False)

                ' Autoincrement flag.
                .AutoIncrement = IIf((UCase$(ExtractToken(strRet, _
                        "AutoIncrement")) = "TRUE"), True, False)
            End With

            ' Increment the field count variable "k"
            K = K + 1
        Loop

        ' Create the table.
        If Not CreateTBL(Db, strTblName, tblData()) Then
            GoTo dbCreate_err
        End If

        ' If any indexes are specified, create them.
        K = 0
        Dim IndxData() As idx, IndxCount As Integer
        IndxCount = 0
        Do
            strRet = ReadFromIniFile(strTblName, _
                        "Index" & K + 1, strdataFile)
            If strRet = "" Then Exit Do
            ReDim Preserve IndxData(K)
            IndxCount = K + 1
            With IndxData(K)
                .Name = ExtractToken(strRet, "IndexName")
                .fields = ExtractToken(strRet, "Fields")
                .Primary = IIf(UCase$((ExtractToken(strRet, _
                            "Primary"))) = "TRUE", True, False)
                '.Required = IIf(UCase$((extracttoken(strRet, _
                            "Required"))) = "TRUE", True, False)
                .Unique = IIf(UCase$((ExtractToken(strRet, _
                            "Unique"))) = "TRUE", True, False)
                .IgnoreNulls = IIf(UCase$((ExtractToken(strRet, _
                            "IgnoreNulls"))) = "TRUE", True, False)
            End With
            K = K + 1
        Loop
        If IndxCount > 0 Then
            If Not CreateIndexes(Db, strTblName, IndxData()) Then
                GoTo dbCreate_err
            End If
        End If

        ' Increment the table count variable "j"
        J = J + 1
    Loop
    ' Increment the DB count variable "i"
    i = i + 1
Loop

' Set the return value.
CreateDB = True

Exit_Line:
    Exit Function

dbCreate_err:
    If Err.Number = 75 Then ' Path/File access error.
       nRet = MsgBox("Error accessing the file '" _
                & strRet & "'.", vbRetryCancel + vbCritical)
        If nRet = vbRetry Then Resume
    
    ElseIf Err Then
       
        MsgBox Err.Description, vbCritical
    End If
'    Resume

If Not gDBTrans.OpenDB(gWorkDir & "\Indx2000.mdb", "pwd") Then
     gDBTrans.CreateDB (gWorkDir & "\Indx2000.tdb")
End If



#End If
End Function

Private Sub Initialize()
'If gDBTrans Is Nothing Then
'    Set gDBTrans = New clsTransact
'End If

End Sub

Private Sub Main()
Call Initialize
'
frmCreatDb.Show vbModal
frmpath.Show

End Sub

