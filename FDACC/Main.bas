Attribute VB_Name = "basMain"
Option Explicit

Public gAppPath As String
Public gStrDate As String
Public Const gAppName = "INDEX2000 - FD Accounts"
Public gDBTrans As clsDBUtils
Public gCancel As Boolean
Public gWindowHandle  As Long
Public gCompanyName As String

Public gCurrUser As clsUsers
'
Public Sub Initialize()

'Initialize the global variables
    gAppPath = App.Path
    
    If gDBTrans Is Nothing Then Set gDBTrans = New clsDBUtils

'Open the data base
    If Not gDBTrans.OpenDB(gAppPath & "\..\Index 2000.MDB", "WIS!@#") Then
        If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
            End
        End If
        If Not gDBTrans.CreateDB(gAppPath & "\FDAcc.TAB", "") Then
            MsgBox "Unable to create new database !", vbCritical, gAppName & " - Error"
            On Error Resume Next
            Kill gAppPath & "\Index 2000.MDB"
            End
        End If
    End If
    gStrDate = Format(Now, "mm/dd/yyyy")
    
End Sub

Public Sub Main()

Call Initialize
'Call KannadaInitialize
KannadaInitialize
If gCurrUser Is Nothing Then Set gCurrUser = New clsUsers

gCurrUser.ShowLoginDialog

If Not gCurrUser.LoginStatus Then
    Set gCurrUser = Nothing
    gDBTrans.CloseDB
    Set gDBTrans = Nothing
    Exit Sub
End If

wisMain.Show
End Sub

Public Sub LoadPlaces(cmbObject As ComboBox)
Dim rst As Recordset
gDBTrans.SQLStmt = "Select * From PlaceTab"
cmbObject.AddItem ""
If gDBTrans.Fetch(rst, adOpenDynamic) Then
    With cmbObject
        While Not rst.EOF
            .AddItem FormatField(rst("Place"))
            rst.MoveNext
        Wend
    End With
Else
    cmbObject.AddItem "Home Town"
End If
End Sub

Public Sub LoadCastes(cmbObject As ComboBox)
Dim rst As Recordset
cmbObject.AddItem ""
gDBTrans.SQLStmt = "Select * From CasteTab"
If gDBTrans.Fetch(rst, adOpenDynamic) Then
    With cmbObject
        While Not rst.EOF
            .AddItem FormatField(rst("Caste"))
            rst.MoveNext
        Wend
    End With
Else
    cmbObject.AddItem "Indian"
End If
End Sub

Public Sub LoadGender(cmbObject As ComboBox)
Dim Gender As wis_Gender
With cmbObject
    Gender = wisNoGender
    .AddItem LoadResString(gLangOffSet + 338) ''All
    .ItemData(.NewIndex) = Gender
    
    Gender = wisMale
    .AddItem LoadResString(gLangOffSet + 385) ''mALE
    .ItemData(.NewIndex) = Gender
    
    Gender = wisFemale
    .AddItem LoadResString(gLangOffSet + 386) ''Female
    .ItemData(.NewIndex) = Gender

End With
End Sub

