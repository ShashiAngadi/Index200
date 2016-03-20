Attribute VB_Name = "modRegistry"
Option Explicit
Option Compare Text

Global Const gsSLASH_BACKWARD As String = "\"

''Registry API Declarations...
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx_wis Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
    ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
    ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, _
    ByRef lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32" Alias "RegEnumValueA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, _
    ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
    lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, _
    ByVal lpcbClass As Long, lpftLastWriteTime As FileTime) As Long

''Reg Data Types...
Private Const REG_NONE = 0                                          ' No value type
Private Const REG_SZ = 1                                            ' Unicode nul terminated string
Private Const REG_EXPAND_SZ = 2                                     ' Unicode nul terminated string
Private Const REG_BINARY = 3                                        ' Free form binary
Private Const REG_DWORD = 4                                         ' 32-bit number
Private Const REG_DWORD_LITTLE_ENDIAN = 4                           ' 32-bit number (same as REG_DWORD)
Private Const REG_DWORD_BIG_ENDIAN = 5                              ' 32-bit number
Private Const REG_LINK = 6                                          ' Symbolic Link (unicode)
Private Const REG_MULTI_SZ = 7                                      ' Multiple Unicode strings
Private Const REG_RESOURCE_LIST = 8                                 ' Resource list in the resource map
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9                      ' Resource list in the hardware description
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10

''Reg Create Type Values...
Private Const REG_OPTION_RESERVED = 0                               ' Parameter is reserved
Private Const REG_OPTION_NON_VOLATILE = 0                           ' Key is preserved when system is rebooted
Private Const REG_OPTION_VOLATILE = 1                               ' Key is not preserved when system is rebooted
Private Const REG_OPTION_CREATE_LINK = 2                            ' Created key is a symbolic link
Private Const REG_OPTION_BACKUP_RESTORE = 4                         ' open for backup or restore

''Reg Key Security Options...
Private Const READ_CONTROL = &H20000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Private Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Private Const KEY_EXECUTE = KEY_READ
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE _
                            + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS _
                            + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

''Return Value...
Private Const ERROR_SUCCESS = 0
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_NO_MORE_ITEMS = 259&

''Hierarchy separator
Private Const KeySeparator As String = "\"

''Registry Security Attributes TYPE...
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type
Private Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

''Reg Key ROOT Types...
Public Enum REGToolRootTypes
    HK_CLASSES_ROOT = &H80000000
    HK_CURRENT_USER = &H80000001
    HK_LOCAL_MACHINE = &H80000002
    HK_USERS = &H80000003
    HK_PERFORMANCE_DATA = &H80000004
    HK_CURRENT_CONFIG = &H80000005
    HK_DYN_DATA = &H80000006
End Enum

'   CreateRegistryKey : Creates a specified key string entry in Registry
'                                   under a specified Root Class.
'
'   [Input] :   1.  Handle of the Root Class key identified by one of the following.
'                            HKEY_CLASSES_ROOT
'                            HKEY_CURRENT_CONFIG
'                            HKEY_CURRENT_USER
'                            HKEY_LOCAL_MACHINE
'
'   [Output] :  Returns True if the specified is created.
'                   Returns False if unsuccessful in creating the key.
'
Function CreateRegistryKey(ByVal lpKeyHandle As Long, ByVal szKeyString As String) As Boolean
    Dim lRetLong As Long
    Dim hKey As Long
    Dim dwDisposition As Long
    Dim lKeyData As String

    CreateRegistryKey = False   'Initialize the return value.
    
    lRetLong = RegCreateKeyEx_wis(lpKeyHandle, szKeyString, _
                        0, vbNull, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                             0, hKey, dwDisposition)
    If lRetLong = ERROR_SUCCESS Then CreateRegistryKey = True

End Function


'       GetRegistryValue :  Gets the value of a specified key from registry.
'
'       [Input] :   1.  Handle of the Root class key.
'                      2.  Key Section name.
'                      3.  Sub key string whose value is to be fetched.
'
'       [Returns]:  The value string if successful.
'                       Null string "" if unsuccessful.
'
Function GetRegistryValue(lpKeyHandle As Long, szKeyName As String, szSubKey As String) As String
    Dim lcount As Long
    Dim lRetLng As Long
    Dim hKey As Long
    Dim hDepth As Long
    Dim lKeyValType As Long
    Dim lTmpStr As String
    Dim lKeyValSize As Long
    Dim lKeyVal As String

    'Initialize the return value.
    GetRegistryValue = ""
    
    'Open the specified key.
    lRetLng = RegOpenKeyEx(lpKeyHandle, szKeyName, 0, KEY_ALL_ACCESS, hKey)
    If (lRetLng <> ERROR_SUCCESS) Then GoTo EndLine
        
    'Initialize the string variable to fetch the value.
    lTmpStr = String$(1024, 0)
    lKeyValSize = 1024

    'Query the registry.
    lRetLng = RegQueryValueEx(hKey, szSubKey, 0, _
                         lKeyValType, lTmpStr, lKeyValSize)

    If (lRetLng <> ERROR_SUCCESS) Then GoTo EndLine
    ' Added by Rk,9/2/1998,to be removed
    If lKeyValSize = 0 Then
        lKeyValSize = 1
    End If
    'end of add
    If (Asc(Mid(lTmpStr, lKeyValSize, 1)) = 0) Then
        lTmpStr = Left(lTmpStr, lKeyValSize - 1)
    Else
        lTmpStr = Left(lTmpStr, lKeyValSize)
    End If

    Select Case lKeyValType
    Case REG_SZ
        lKeyVal = lTmpStr
    Case REG_DWORD
        For lcount = Len(lTmpStr) To 1 Step -1
            lKeyVal = lKeyVal + Hex(Asc(Mid(lTmpStr, lcount, 1)))
        Next
        lKeyVal = Format$("&h" + lKeyVal)
    End Select
    
    GetRegistryValue = lKeyVal
    lRetLng = RegCloseKey(hKey)
    Exit Function
    
EndLine:
    lRetLng = RegCloseKey(hKey)
End Function

'   DeleteRegistryKey : Deletes a specified key from the registry.
'
'   [Input]     1.  KeyHandle :  Identifies the handle of the main key. The values can be :
'                                           HKEY_CLASSES_ROOT
'                                           HKEY_CURRENT_CONFIG
'                                           HKEY_CURRENT_USER
'                                           HKEY_LOCAL_MACHINE
'
'                 2.    lpSubKey :  Key string that is to be  deleted.
'
'   [Output]    1.  Returns True when the key is not existing.
'                   2.  Returns True when the key is successfully deleted.
'                   3.  Returns False, if unable to delete key.
'
Function DeleteRegistryKey(KeyHandle As Long, lpSubKey As String) As Boolean
    Dim lRetLng As Long
    Dim hKey As Long

'--Open the specified key.
    lRetLng = RegOpenKeyEx(KeyHandle, lpSubKey, 0, KEY_ALL_ACCESS, hKey)
    If (lRetLng <> ERROR_SUCCESS) Then
        'Since the key is not present, we need not delete it.
        'We will just return true.
        DeleteRegistryKey = True
        Exit Function
    End If

'--Delete the key.
    lRetLng = RegDeleteKey(KeyHandle, lpSubKey)
    If (lRetLng <> ERROR_SUCCESS) Then
        DeleteRegistryKey = False
        Exit Function
    End If

    DeleteRegistryKey = True
End Function

Function OpenRegistryKey(ByVal lpKeyHandle As Long, szKeyString As String) As Boolean
Dim lRetLng As Long
Dim hKey As Long

    OpenRegistryKey = False 'Initialize the return value.
    
    lRetLng = RegOpenKeyEx(lpKeyHandle, szKeyString, 0, KEY_ALL_ACCESS, hKey)
    If (lRetLng = ERROR_SUCCESS) Then OpenRegistryKey = True

End Function




'   SetRegistryValue :  Sets the value of the specified key, subkey to the specified value.
'
'   [Input] :   1.  Handle of the Root Class Key.
'                 2.  Key string.
'                 3.  Sub key.
'                 4.  Value string.
'
'
Function SetRegistryValue(ByVal lpKeyHandle As Long, ByVal szKeyString As String, ByVal szSubKey As String, ByVal szValueKey As String) As Boolean

Dim RetLng As Long
Dim hKey As Long
Dim dwDisposition As Long

'First open the key
SetRegistryValue = False
    'Open the registry key
    RetLng = RegOpenKeyEx(lpKeyHandle, szKeyString, 0, KEY_ALL_ACCESS, hKey)
    If (RetLng <> ERROR_SUCCESS) Then
         RetLng = RegCreateKeyEx_wis(HKEY_LOCAL_MACHINE, szKeyString, 0, vbNull, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0, hKey, dwDisposition)
                If (RetLng <> ERROR_SUCCESS) Then
                    Exit Function
                End If
    End If
    
    'Set the value
    RetLng = RegSetValueEx(hKey, szSubKey, 0, REG_SZ, ByVal szValueKey, CLng(Len(szValueKey)))
        If RetLng <> ERROR_SUCCESS Then
            Exit Function
        End If
    RetLng = RegCloseKey(hKey)
SetRegistryValue = True

End Function


'Retrieves a key value.
Public Function GetKeyValue(ByVal KeyRoot As REGToolRootTypes, KeyName As String, ValueName As String, ByRef ValueData As String) As Boolean
    Dim i As Long                                                   ' Loop Counter
    Dim hKey As Long                                                ' Handle To An Open Registry Key
    Dim KeyValType As Long                                          ' Data Type Of A Registry Key
    Dim sTmp As String                                              ' Tempory Storage For A Registry Key Value
    Dim sReturn As String
    Dim KeyValSize As Long                                          ' Size Of Registry Key Variable
    Dim sByte As String

    If ValidKeyName(KeyName) Then
        On Error GoTo LocalErr

        ' Open registry key under KeyRoot
        Attempt RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)

        sTmp = String$(1024, 0)                                         ' Allocate Variable Space
        KeyValSize = 1024                                               ' Mark Variable Size

        ' Retrieve Registry Key Value...
        Attempt RegQueryValueEx(hKey, ValueName, 0, _
                KeyValType, sTmp, KeyValSize)                           ' Get/Create Key Value

        If (Asc(Mid(sTmp, KeyValSize, 1)) = 0) Then                     ' Win95 Adds Null Terminated String...
            sTmp = Left(sTmp, KeyValSize - 1)                           ' Null Found, Extract From String
        Else                                                            ' WinNT Does NOT Null Terminate String...
            sTmp = Left(sTmp, KeyValSize)                               ' Null Not Found, Extract String Only
        End If

        ' Determine Key Value Type For Conversion...
        Select Case KeyValType                                          ' Search Data Types...
            Case REG_SZ                                                 ' String Registry Key Data Type
                sReturn = sTmp '(Do nothing)
            Case REG_DWORD                                              ' Double Word Registry Key Data Type
                For i = Len(sTmp) To 1 Step -1                          ' Convert Each Bit
                    sByte = Hex(Asc(Mid$(sTmp, i, 1)))
                    Do Until Len(sByte) = 2
                        sByte = "0" & sByte
                    Loop
                    sReturn = sReturn & sByte                           ' Build Value Char. By Char.
                Next
                sReturn = Format$("&h" + sReturn)                       ' Convert Double Word To String
        End Select

        GetKeyValue = True
        ValueData = sReturn

LocalErr:
        On Error Resume Next
        RegCloseKey hKey
    End If
End Function

Private Sub Attempt(rc As Long)
    If (rc <> ERROR_SUCCESS) Then
        Err.Raise 5
    End If
End Sub

Private Function ValidKeyName(KeyName As String) As Boolean
    'A key name is invalid if it begins or ends with \ or contains \\
    If Left$(KeyName, 1) <> gsSLASH_BACKWARD Then
        If Right$(KeyName, 1) <> gsSLASH_BACKWARD Then
            If InStr(KeyName, gsSLASH_BACKWARD & gsSLASH_BACKWARD) = 0 Then
                ValidKeyName = True
            End If
        End If
    End If
End Function

