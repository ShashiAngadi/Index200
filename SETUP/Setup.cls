VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Property Get ObjectName()
    ObjectName = "SetupClass"
End Property


Public Function ReadSetupValue(Module As String, strKey As String, DefaultValue As String) As String
Dim DBStr As String
Dim Rst As ADODB.Recordset
Dim rst1 As ADODB.Recordset

ReadSetupValue = DefaultValue
gDbTrans.SQLStmt = "Select * from Setup where " & _
                    "ModuleData like '" & Module & "'" & _
                    " And KeyData like '" & strKey & "'"
    If gDbTrans.Fetch(rst1, adOpenDynamic) < 1 Then Exit Function
        
        DBStr = FormatField(rst1("ValueData"))
        ReadSetupValue = ""
        If DBStr <> " " Then ReadSetupValue = DBStr

End Function


Public Function WriteSetupValue(Module As String, Key As String, Value As String) As Boolean

Dim RecordExists As Boolean
Dim SetupID As Long
Dim Rst As ADODB.Recordset

'Preliminary checks
If Trim$(Module) = "" Or Trim$(Key) = "" Then
    'MsgBox "Module / Key of Setup not specified !", vbExclamation, gAppName & " - Error"
    MsgBox "Module / Key of Setup not specified !", vbExclamation, gAppName & " - Error"
    Exit Function
End If

'Check if record exists
gDbTrans.SQLStmt = "Select * from SetUp where " & _
    "ModuleData like '" & Module & "'" & _
    " and KeyData like '" & Key & "'"
If gDbTrans.Fetch(Rst, adOpenForwardOnly) > 0 Then
    RecordExists = True
    SetupID = Val(FormatField(Rst("SetupID")))
Else
    'Get new SetupID
    RecordExists = False
    gDbTrans.SQLStmt = "Select TOP 1 SetupID from Setup order by SetupID desc"
    If gDbTrans.Fetch(Rst, adOpenForwardOnly) <= 0 Then
        SetupID = 1
    Else
        SetupID = Val(FormatField(Rst("SetupID"))) + 1
    End If
End If

'Set SQL Statement accordingly
'NOTE: If value is empty then put value = " "
If Not RecordExists Then
    gDbTrans.SQLStmt = "Insert into Setup (SetupID, ModuleData, Keydata, Valuedata) values ( " & _
        SetupID & "," & _
        "'" & Module & "', " & _
        "'" & Key & "', " & _
        "'" & IIf(Trim$(Value) = "", " ", Value) & "' " & _
        ")"
Else
    gDbTrans.SQLStmt = "Update Setup Set " & _
        "Moduledata = '" & Module & "', " & _
        "KeyData = '" & Key & "', " & _
        "Valuedata = '" & IIf(Trim$(Value) = "", " ", Value) & "'" & _
        " where SetupID = " & SetupID & ";"
End If

'Begin DB transactions
gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Function
End If
gDbTrans.CommitTrans

WriteSetupValue = True

End Function

