Attribute VB_Name = "basMain"
Option Explicit

Public gAppPath As String
Public Const gAppName = "INDEX2000 - SB Acounts"
Public gDBTrans As clsDBUtils
Public gCompanyName As String
Public gtranstype As wisTransactionTypes
Public gWindowHandle As Long
Public gStrDate As String
Public gCurrUser As clsUsers
Public gCancel  As Boolean



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

Sub PauseApplication(Secs As Integer)

Dim PauseTime, Start, Finish, TotalTime
        PauseTime = Secs   ' Set duration.
        Start = Timer   ' Set start time.
        Do While Timer < Start + PauseTime
            DoEvents    ' Yield to other processes.
        Loop
        Finish = Timer  ' Set end time.
        TotalTime = Finish - Start  ' Calculate total time.
End Sub


Public Function OBOfAccount(Module As Integer, OBDate As String, Optional ReportType As wisReports) As Currency
If ReportType = wisBalanceSheet Then
    gDBTrans.SQLStmt = "SELECT TOP 1 * FROM OBTab WHERE OBDate < #" & _
        OBDate & "# AND Module = " & Module & _
        " ORDER BY OBDate DESC;"
Else
    gDBTrans.SQLStmt = "SELECT TOP 1 * FROM OBTab WHERE OBDate <= #" & _
        OBDate & "# AND Module = " & Module & _
        " ORDER BY OBDate DESC;"
End If
Dim rst As ADODB.Recordset
If gDBTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
    If Module = wis_FromProfit Then
        OBOfAccount = -1
    End If
    Exit Function
End If
OBOfAccount = FormatField(rst("OBAmount"))

End Function


Public Sub Initialize()

'Initialize the global variables
    gAppPath = App.Path
    
    If gDBTrans Is Nothing Then
        Set gDBTrans = New clsDBUtils
    End If

'Open the data base
    If Not gDBTrans.OpenDB(gAppPath & "\..\Index 2000.MDB", "WIS!@#") Then
        End
        'If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
        If MsgBox(LoadResString(gLangOffSet + 605), vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
        End If
        If Not gDBTrans.CreateDB(gAppPath & "\SBAcc.TAB", "") Then
            'MsgBox "Unable to create new database !", vbCritical, gAppName & " - Error"
             MsgBox LoadResString(gLangOffSet + 605), vbCritical, gAppName & " - Error"
            On Error Resume Next
            Kill gAppPath & "\SBAcc.MDB"
            
        End If
    End If



End Sub

Public Sub Main()
'gLangOffSet = 1000
gStrDate = Format(Now, "mm/dd/yyyy")

Call Initialize
Call KannadaInitialize
If gCurrUser Is Nothing Then Set gCurrUser = New clsUsers
gCurrUser.ShowLoginDialog

wisMain.Show
End Sub


