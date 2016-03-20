Attribute VB_Name = "basKannada"
Option Explicit
'Declare Function WINTOISCII Lib "win2isc.dll" (ByVal inputstr As String) As String
Declare Function NudiStartKeyboardEngine Lib "Kannada-Nudi.dll" _
    Alias "_NudiStartKeyboardEngineVB@12" (ByVal isGlobal As Boolean, _
    ByVal isMonoLingual As Boolean, _
    ByVal needTrayIcon As Boolean) As Integer
Declare Sub NudiTurnOnScrollLock Lib "Kannada-Nudi.dll" Alias "_NudiTurnOnScrollLockVB@0" ()
Declare Function NudiStopKeyboardEngine Lib "Kannada-Nudi.dll" Alias "_NudiStopKeyboardEngineVB@0" () As Integer
Declare Sub NudiResetAllFlags Lib "Kannada-Nudi.dll" Alias "_NudiResetAllFlagsVB@0" ()
Declare Function NudiGetLastError Lib "Kannada-Nudi.dll" Alias "_NudiGetLastErrorVB@0" () As Integer

Public gFontName As String
Public gFontNameEnglish As String
Public gFontSize As Single
Public gFontSizeEnglish As Single
Public gLangOffSet As Integer
Public Const wis_KannadaOffset = 2000
Public Const wis_KannadaSamhitaOffset = 4000
Public Const wis_NoLangOffset = 0
Public gLangShree As Boolean

Private Const NUDI_ERR_ALREADY_RUNNING = -1

'
Public Function ConvertToIscii(AsciStr As String) As String

Dim StrLen As Integer
Dim IsciWord  As String
Dim I As Integer
Dim SingleChar As String * 1
StrLen = Len(AsciStr)

'IsciWord = ""
'SingleChar = ""
'For i = 1 To StrLen
'    SingleChar = Hex(Int(Asc(WINTOISCII(Mid(AsciStr, i, 1)))))
'    IsciWord = IsciWord & SingleChar
'Next i
If gLangShree Then IsciWord = ShreeToIscii(AsciStr)

ConvertToIscii = IsciWord


End Function

Public Sub KannadaInitialize()
Dim lngRetVal As Long
gFontName = "MS Sans Serif"
gFontSize = 9
gFontNameEnglish = "MS Sans Serif"
gFontSizeEnglish = 9
gLangOffSet = 0
Dim rst As ADODB.Recordset
Dim langTool As String
'Include  ..\Shared\wisReg.bas File to the project
'First Get The Lanuage Constant From the Registry
'Get the Language information From Database
Dim strRet As String
'strRet = ReadFromIniFile("Language", "Language", App.Path & "\" & constFINYEARFILE)
If gDbTrans.CommandObject Is Nothing Then
    strRet = ReadFromIniFile("Language", "Language", App.Path & "\" & constFINYEARFILE)
Else
    gDbTrans.SqlStmt = "select * From Install Where KeyData = 'Language'"
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        strRet = FormatField(rst("ValueData"))
        Call WriteToIniFile("Language", "Language", strRet, App.Path & "\" & constFINYEARFILE)
    End If
    Set rst = Nothing
    gDbTrans.SqlStmt = "select * From Install Where KeyData = 'LanguageTool'"
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        langTool = FormatField(rst("ValueData"))
        Call WriteToIniFile("Language", "LanguageTool", langTool, App.Path & "\" & constFINYEARFILE)
    End If
    Set rst = Nothing
    
    If Len(strRet) = 0 Then _
        strRet = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "Language")
End If


If Len(strRet) = 0 Then _
    strRet = ReadFromIniFile("Language", "Language", App.Path & "\" & constFINYEARFILE)

If UCase(strRet) = "KANNADA" Then
    
    gLangOffSet = wis_KannadaOffset
    langTool = ReadFromIniFile("Language", "LanguageTool", App.Path & "\" & constFINYEARFILE)
    gFontName = ReadFromIniFile("Language", "FontName", App.Path & "\" & constFINYEARFILE)
Else
'If Len(strRet) = 0 Then
    Call WriteToIniFile("Language", "Language", _
                            "English", App.Path & "\" & constFINYEARFILE)
End If
'Redington indi 2257755
If gLangOffSet Then
    gFontSize = 12
    If Len(langTool) = 0 Or UCase(langTool) = "NUDI" Then
        ''Default is NUDI
        lngRetVal = NudiStartKeyboardEngine(False, False, False)
        gFontName = "Nudi B-Akshar"
        gLangOffSet = wis_KannadaOffset
    Else
        'Stop the NUdi in case its running
        Call NudiResetAllFlags
        lngRetVal = NudiStopKeyboardEngine()
        gLangShree = True
        gFontName = ReadFromIniFile("Language", "FontName", App.Path & "\" & constFINYEARFILE)
        'Start the Samhita
        Call InitializeSamhita
        gFontSize = 13
        gLangOffSet = wis_KannadaSamhitaOffset
    End If
    'If lngRetVal = 0 Then MsgBox "Cannot Start Nudi", vbInformation
    
    'Debug.Print NudiGetLastError
End If

End Sub
Public Function LoadResourceStringS(ParamArray ResourceIDs() As Variant) As String
    Dim retValue As String
    Dim count As Integer
    retValue = LoadResString(ResourceIDs(0))
    
    For count = 0 To UBound(ResourceIDs)
        retValue = retValue & " " & LoadResString(ResourceIDs(count))
    Next
    LoadResourceStringS = retValue
End Function

Public Function GetResourceString(ParamArray ResourceIDs() As Variant) As String
    Dim retValue As String
    Dim count As Integer
    retValue = LoadResString(gLangOffSet + ResourceIDs(0))
    
    For count = 1 To UBound(ResourceIDs)
        retValue = retValue & " " & LoadResString(gLangOffSet + ResourceIDs(count))
    Next
    GetResourceString = retValue
End Function
Public Function GetResourceString_Old(ResourceID As Long) As String
    GetResourceString_Old = LoadResString(gLangOffSet + ResourceID)
End Function

Public Function GetMergeResourceString(ResourceID() As Variant) As String
    Dim retValue As String
    Dim count As Integer
    For count = 0 To UBound(ResourceID) - 1
        retValue = " " & LoadResString(gLangOffSet + ResourceID(count))
    Next
    GetMergeResourceString = Mid$(retValue, 1)
End Function

