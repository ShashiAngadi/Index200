Attribute VB_Name = "basMain"
Option Explicit

Public gAppPath As String
Public Const gAppName = "INDEX2000 - CA Acounts"
Public gDBTrans As clsDBUtils
Public GWindowHandle As Long
Public gStrDate As String
Public gCompanyName As String
Public gCancel As Boolean
Public gCurrUser As clsUsers

Public Enum wis_ChequeTrans
    chqIssue = 1
    chqPay = 2
    chqStop = 3
    chqLoss = 4
End Enum




Public Sub LoadPlaces(cmbObject As ComboBox)
Dim Rst As Recordset
gDBTrans.SQLStmt = "Select * From PlaceTab"
cmbObject.AddItem ""
If gDBTrans.Fetch(Rst, adOpenDynamic) Then
    With cmbObject
        While Not Rst.EOF
            .AddItem FormatField(Rst("Place"))
            Rst.MoveNext
        Wend
    End With
Else
    cmbObject.AddItem "Home Town"
End If
End Sub

Public Sub LoadCastes(cmbObject As ComboBox)
Dim Rst As Recordset
cmbObject.AddItem ""
gDBTrans.SQLStmt = "Select * From CasteTab"
If gDBTrans.Fetch(Rst, adOpenDynamic) Then
    With cmbObject
        While Not Rst.EOF
            .AddItem FormatField(Rst("Caste"))
            Rst.MoveNext
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

Public Sub Initialize()

'Initialize the global variables
    gAppPath = App.Path
    
    If gDBTrans Is Nothing Then Set gDBTrans = New clsDBUtils

'Open the data base
    If Not gDBTrans.OpenDB(gAppPath & "\CAAcc.MDB", "WIS!@#") Then
        If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
            End
        End If
        If Not gDBTrans.CreateDB(gAppPath & "\CAAcc.TAB", "") Then
            MsgBox "Unable to create new database !", vbCritical, gAppName & " - Error"
            On Error Resume Next
            Kill gAppPath & "\CAAcc.MDB"
            End
        End If
    End If



End Sub




Public Sub Main()
gLangOffSet = 5000
gStrDate = Format(Now, "mm/dd/yyyy")

Call Initialize
Call KannadaInitialize

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



