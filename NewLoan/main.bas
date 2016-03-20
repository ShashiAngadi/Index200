Attribute VB_Name = "basMain"
Option Explicit

Public gDbTrans As clsDBUtils

Public gAppName As String
Public gAppPath As String
Public gStrDate As String
Public gCompanyName As String
Public gCancel As Boolean
Public gWindowHandle As Long
Public gFont As StdFont

'Delimiter used through out the project
'Public gDelim As String

Public Const conEncrypt = 128

Public Enum wis_DbOperation
    Insert = 1
    Update = 2
End Enum

Public Enum wis_AccountModule
    wisMemberAccount = 1
    wisSBAccount = 2
    wisCDAccount = 3
    wisFDAccount = 4
    wisPDAccount = 5
    wisRDAccount = 6
    wisLoanAccount = 7
    wisDepositLoanAccount = 8
End Enum

Public Sub CenterMe(frmForm As Form)

With frmForm
    .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
End With
End Sub

'
'This Function is used to Decrypt the encrypted string
Public Function DecryptString(Expression)
    Dim StrData As String
    Dim RetStr As String
    'Copy the given expression to a string
    'to avoid any modification of passed argument
    
    StrData = CStr(Expression)
    
    'To derypt the datA by a value we adde while encrypting it
    If StrData = "" Then Exit Function
    Dim CharVal As Integer
    Dim Count As Integer, MaxCount As Integer
    MaxCount = Len(StrData)
    RetStr = ""
    For Count = 1 To MaxCount
        CharVal = Asc(Mid(StrData, Count, 1))
        CharVal = CharVal - Count - conEncrypt
        CharVal = IIf(CharVal < 0, CharVal + 256, CharVal)
        RetStr = RetStr & Chr(CharVal)
    Next
    
    DecryptString = RetStr
    
End Function



'
'This Function is used to encrypt the data
'If any information has to be transferred between two Banks
'or two databases located remotely.
'
Public Function EncryptString(Expression)
    Dim StrData As String
    Dim RetStr As String
    'Copy thegivenb expression to a string
    'to avoid any modification of passed argument
    
    StrData = CStr(Expression)
    
    'To Encrypt the data add a Fixed value to each Charector
    If StrData = "" Then Exit Function
    Dim sngChar As String
    Dim CharVal As Byte
    Dim Count As Integer, MaxCount As Integer
    MaxCount = Len(StrData)
    RetStr = ""
    For Count = 1 To MaxCount
        CharVal = Asc(Mid(StrData, Count, 1))
        CharVal = (CharVal + Count + conEncrypt) Mod 256
        RetStr = RetStr & Chr(CharVal)
    Next
    EncryptString = RetStr
    
    
End Function






Sub Initialise()
'gDelim = "~"
gAppName = "New Loan Module"
'Initialize the global variables
    gAppPath = App.Path
    If gDbTrans Is Nothing Then Set gDbTrans = New clsDBUtils
    
'Open the data base
    If Not gDbTrans.OpenDB(gAppPath & "\" & "test.mdb", "") Then
        If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
            End
        End If
        
    End If

Dim Rst As Recordset
gDbTrans.SQLStmt = "SELECT * FROM INstall WHERE KeyData = 'BankId'"
If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then
'    gBankID = FormatField(Rst("ValueData"))
End If


End Sub


'
'*****************************************************************************************************************
'                                   Update Last Accessed Elements
'*****************************************************************************************************************
'This function will be useful if you want to update the last accessed elements
'Eg : Last Accessed Files
'Suppose you want the last of 4 last accessed files and you have only 2 files.
' pass the other 2 elements as "" (NULL)
'
'   Girish  Desai  May 1st, 1998.
'   Modified by shashi on 24//01

Function UpdateLastAccessedElements(Str As String, strArr() As String, Optional IgnoreCase As Boolean = True)

Dim CaseVal As Integer
Dim Pos As Integer
Dim Flag As Boolean
Dim Count As Integer
Dim IgnCase As Boolean

    IgnCase = False
    If Not IsMissing(IgnoreCase) Then
        IgnCase = IgnoreCase
    End If

    If IgnCase Then
        CaseVal = vbBinaryCompare
    Else
        CaseVal = vbTextCompare
    End If

'First check out the position
    For Pos = 0 To UBound(strArr)
        If StrComp(Str, strArr(Pos), CaseVal) = 0 Then
            Flag = True
            Exit For
        End If
    Next Pos
    
    If Not Flag Then Pos = Pos - 1
    
    For Count = Pos To 1 Step -1
        strArr(Count) = strArr(Count - 1)
    Next Count

    strArr(0) = Str
    
DoEvents

End Function


''*****************************************************************************************************************
''                                   Update Last Accessed Elements
''*****************************************************************************************************************
''This function will be useful if you want keep the last accessed elements
''Eg : Last Accessed Files
''Suppose you want the last of 4 last accessed files and you have only 2 files.
'' pass the other 3rd elements as "" (NULL)
'' By shashi on 24//01
'
'Function WriteLastElements(StrArr() As String, strFileName As String)
'
'Dim CaseVal As Integer
'Dim Pos As Integer
'Dim Flag As Boolean
'Dim Count As Integer
'Dim IgnCase As Boolean
'
'    IgnCase = False
'    If Not IsMissing(IgnCase) Then
'        IgnCase = IgnCase
'    End If
'
'    If IgnCase Then
'        CaseVal = vbBinaryCompare
'    Else
'        CaseVal = vbTextCompare
'    End If
'
''First check out the position
'    For Pos = 0 To UBound(StrArr)
'        If StrComp(str, StrArr(Pos), CaseVal) = 0 Then
'            Flag = True
'            Exit For
'        End If
'    Next Pos
'
'    If Not Flag Then Pos = Pos - 1
'
'    For Count = Pos To 1 Step -1
'        StrArr(Count) = StrArr(Count - 1)
'    Next Count
'
'    StrArr(0) = str
'DoEvents
'
'End Function
'
'
Public Sub Main()
Call Initialise

Form1.Show
End Sub

