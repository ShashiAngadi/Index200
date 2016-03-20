Attribute VB_Name = "basmain"

Option Explicit

Public gStrDate As String
Public gAppPath As String
Public Const gAppName = "INDEX 2000"
Public gDBTrans As clsTransact
Public gCurrUser As clsUsers
Public gTranstype As wisTransactionTypes
Public gUser As clsUsers
'Added On 10/5/2000 ' To Find the UserId
Public gUserID As Long
Public gCompanyName As String
Public gDlName As String

Public Sub Initialize()

'Initialize the global variables
    gAppPath = App.Path
    If gDBTrans Is Nothing Then
        Set gDBTrans = New clsTransact
    End If
  
'Get the database name
Dim DBPath As String
Dim DBFileName As String

DBPath = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\waves information systems\index 2000\settings", "server")

If DBPath = "" Then
    'Give the local path of the MDB FILE
    DBPath = App.Path
Else
    DBPath = "\\" & DBPath & "\Index 2000"
End If
'DBFileName = DBPath & "\..\Appmain\" & gAppName & ".MDB"
DBFileName = DBPath & "\Index 2000.MDB"
Debug.Assert DBFileName = ""

'for express pigmy this is the necessary path
'Open the data base
If Not gDBTrans.OpenDB(DBFileName, "PRAGMANS") Then
    If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
        End
    End If
    If Not gDBTrans.CreateDB(gAppPath & "\Index 2000.TAB", "WIS!@#") Then
        MsgBox "Unable to create new database !", vbCritical, gAppName & " - Error"
        On Error Resume Next
         Kill gAppPath & "\Index 2000.MDB"
         End
    End If
Else
    'Make a routine check to the data base
    gDBTrans.SQLStmt = "Select Count(*) as TOTUsers from UserTab"
    If gDBTrans.SQLFetch <= 0 Then
        MsgBox "Initialization Error", vbCritical, gAppName & " - Error"
        End
    End If
    If Val(gDBTrans.Rst("TotUsers")) = 0 Then  'Insert users into the databse
        gDBTrans.BeginTrans
        gDBTrans.SQLStmt = "Insert into NameTab (CustomerID, FirstName, Gender, Reference) values (1,'Administrator',0,0)"
        If Not gDBTrans.SQLExecute Then
            MsgBox "Initialization Error", vbCritical, gAppName & " - Error"
            gDBTrans.RollBack
            End
        End If
        gDBTrans.SQLStmt = "Insert into UserTab(UserID,CustomerID,LoginName,Password,Permissions) values (1,1,'admin','admin',14)"
        If Not gDBTrans.SQLExecute Then
            MsgBox "Initialization Error", vbCritical, gAppName & " - Error"
            gDBTrans.RollBack
            End
        End If
        gDBTrans.CommitTrans
    End If
    
    'Now get the Name of the Bank /Society from DataBase
    gDBTrans.SQLStmt = "Select * from Install where KeyData = " & AddQuotes("CompanyName", True)
    If gDBTrans.SQLFetch > 0 Then gCompanyName = FormatField(gDBTrans.Rst("ValueData"))

     'Now get the Name of the Dhanlaxmi Daposit from DataBase
    gDBTrans.SQLStmt = "Select * from Install where KeyData = " & AddQuotes("DLAcc", True)
    If gDBTrans.SQLFetch > 0 Then gCompanyName = FormatField(gDBTrans.Rst("ValueData"))

End If

'Get the User Up
If gUser Is Nothing Then Set gUser = New clsUsers

'Call GetServerIndianDate
gStrDate = GetSereverDate

End Sub
Public Sub Main()
'frmSplash.Show 'vbModal
'Call PauseApplication(2)
Call Initialize
Call KannadaInitialize
'Now unload the splash form
'Unload frmSplash
'frmMain.Show vbModal  'for exp pigmy only
    On Error GoTo LOGIN_ERROR
    Set gCurrUser = New clsUsers
    gCurrUser.MaxRetries = 3
    gCurrUser.CancelError = True
    gCurrUser.ShowLoginDialog
    If Not gCurrUser.LoginStatus Then
        GoTo LOGIN_ERROR
    End If
    gUserID = gCurrUser.UserID
  
    frmPDAccEx.Show vbModal
Set gUser = Nothing
Set gCurrUser = Nothing
gDBTrans.CloseDB
Set gDBTrans = Nothing
End
Exit Sub

LOGIN_ERROR:
   ' MsgBox gAppName & " could not log you on !", vbExclamation, gAppName & " - Error"
    MsgBox gAppName & " " & LoadResString(gLangOffSet + 751), vbExclamation, gAppName & " - Error"
    End
    

End Sub



