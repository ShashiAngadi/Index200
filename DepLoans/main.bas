Attribute VB_Name = "basMain"
Option Explicit
Public gDBTrans As clsDBUtils
Public gAppName As String
Public gAppPath As String
Public gCompanyName As String
Public gStrDate As String
Public gWindowHandle As Long
Public gCancel As Boolean



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


'Public Const gDelim = ";"
Public Sub Initialize()

'Initialize the global variables
    gAppPath = App.Path
    
    If gDBTrans Is Nothing Then Set gDBTrans = New clsDBUtils

'Open the data base
    If Not gDBTrans.OpenDB(gAppPath & "\..\Index 2000.MDB", "WIS!@#") Then
        If MsgBox("Unable to open the database !" & vbCrLf _
                & vbCrLf & " Creating New Database", vbQuestion _
                + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
            End
        End If
        If Not gDBTrans.CreateDB(gAppPath & "\loans.tab", "") Then
            MsgBox "Unable to create new database !", vbCritical, gAppName & " - Error"
            On Error Resume Next
            Kill gAppPath & "\loandb.MDB"
            End
        End If
    End If
End Sub
Public Sub Main()
Call Initialize
gStrDate = Now

'#If Kannada Then
    Call KannadaInitialize
'#End If
wisMain.Show
End Sub
Public Sub Show()

End Sub


