Attribute VB_Name = "basMain"
Option Explicit

Public gStrDate As String
Public gAppPath As String
Public Const gAppName = "INDEX 2000"
Public gDBTrans As clsTransact
'Public gCurrUser As clsUsers
Public gTranstype As wisTransactionTypes
'Public gUser As clsUsers
'Added On 10/5/2000 ' To Find the UserId
Public gUserID As Long

Public Function SaveInterest(ModuleID As wisModules, SchemeId As Long, InterestRate As Single, _
                                              IndianStartDate As String) As Boolean
'This Function Provides The Interest Rate Induction To All Tables
'Declare The Necessary ariables
Dim UpdateMode As Boolean
Dim TransId As Long

'Check Whether Last Transaction Was Performed On The Same Day ITself
gDBTrans.SQLStmt = "SELECT TransID from InterestTab Where StartDate = # " & IndianStartDate & " AND ModuleID = " _
                                & ModuleID & " AND Schemeid = " & SchemeId
If gDBTrans.SQLFetch > 0 Then UpdateMode = True

If UpdateMode = False Then
    
    gDBTrans.SQLStmt = "SELECT MAX (TransID) from Interest where ModuleID =" & ModuleID & _
                    " AND SchemeID = " & SchemeId
    If gDBTrans.SQLFetch = 1 Then TransId = Val(FormatField(gDBTrans.Rst(0)))
    
    If TransId > 0 Then
        gDBTrans.BeginTrans
        gDBTrans.SQLStmt = "Update InterestTab Set EndDate = #" & DateAdd("d", -1, CDate(FormatDate(IndianStartDate))) & "#" & _
                        " Where TransId = " & TransId & " And ModuleID =" & ModuleID & _
                        " AND SchemeID = " & SchemeId
        If Not gDBTrans.SQLExecute Then
            gDBTrans.RollBack
            GoTo ErrLine
        End If
        gDBTrans.CommitTrans
    End If

End If

gDBTrans.BeginTrans
If UpdateMode = True Then   'Go For UpDation
    gDBTrans.SQLStmt = "UPDATE InterestTab set InterestRate = " & InterestRate & _
        " WHERE TransID = " & TransId & " AND Moduleid = " & ModuleID & _
        " AND SchemeID = " & SchemeId & " AND StartDate = #" & FormatDate(IndianStartDate) & "#"

ElseIf UpdateMode = False Then 'Go For Insertion
    TransId = TransId + 1
    gDBTrans.SQLStmt = "INSERT INTO InterestTab (TransID ,ModuleID ,StartDate ,InterestRate, " & _
                        " SchemeID ) values (" & TransId & "," & ModuleID & ", " & _
                        " #" & FormatDate(IndianStartDate) & "#," & InterestRate & "," _
                        & SchemeId & ")"
End If
    
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        GoTo ErrLine
    End If
gDBTrans.CommitTrans

SaveInterest = True

Exit Function

ErrLine:
    'MsgBox "Unable to save settings !", vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 533), vbExclamation, gAppName & " - Error"
End Function



Public Function UndoLastInterest(ModuleID As wisModules, SchemeId As Long, InterestRate As Single, _
                                              IndianStartDate As String) As Boolean
On Error GoTo ErrLine
'undo the last transaction
Dim TransId As Long

'get the max transid for the particular Moduleid
'fire SQL
gDBTrans.SQLStmt = "SELECT MAX (TransID) from Interest where ModuleID =" & ModuleID & _
                                " AND SchemeID = " & SchemeId
If gDBTrans.SQLFetch = 1 Then TransId = Val(FormatField(gDBTrans.Rst(0)))

gDBTrans.SQLStmt = "Delete * from InterestTab where transid  = " & _
                                TransId & " AND  ModuleID =" & ModuleID & _
                                " AND SchemeID = " & SchemeId & ")"
gDBTrans.BeginTrans
If Not gDBTrans.SQLExecute Then
 gDBTrans.RollBack
 GoTo ErrLine
 End If
 gDBTrans.CommitTrans
 
gDBTrans.SQLStmt = "SELECT MAX (TransID) from Interest where ModuleID =" & ModuleID & _
                                " AND SchemeID = " & SchemeId
If gDBTrans.SQLFetch = 1 Then TransId = Val(FormatField(gDBTrans.Rst(0)))
If TransId > 0 Then
    gDBTrans.BeginTrans
    gDBTrans.SQLStmt = "Update InterestTab Set EndDate = " & Null & _
                    " Where TransId = " & TransId & " And ModuleID =" & ModuleID & _
                    " AND SchemeID = " & SchemeId
    If Not gDBTrans.SQLExecute Then
        gDBTrans.RollBack
        GoTo ErrLine
    End If
    gDBTrans.CommitTrans
End If

UndoLastInterest = True

Exit Function

ErrLine:
    'MsgBox "Unable to save settings !", vbExclamation, gAppName & " - Error"
    MsgBox LoadResString(gLangOffSet + 533), vbExclamation, gAppName & " - Error"
End Function



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
    DBPath = "\\" & DBPath
End If

DBFileName = DBPath & "\" & gAppName & ".MDB"

'Open the data base
    If Not gDBTrans.OpenDB(DBFileName, "WIS!@#") Then
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
            MsgBox "INitialization Error", vbCritical, gAppName & " - Error"
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
    End If

'Get the User Up
    'If gUser Is Nothing Then
    '    Set gUser = New clsUsers
    'End If
'Call GetServerIndianDate
gStrDate = GetSereverDate
End Sub

Public Sub Main()
'frmSplash.Show 'vbModal
'Call PauseApplication(2)
Call Initialize
Call KannadaInitialize
'Load wisMain
'Now unload the splash form
'Unload frmSplash
'wisMain.Show
frmClearing.Show


'Ask Login
    On Error GoTo LOGIN_ERROR
    'Set gCurrUser = New clsUsers
    'gCurrUser.MaxRetries = 3
    'gCurrUser.CancelError = True
    'gCurrUser.ShowLoginDialog
    'If Not gCurrUser.LoginStatus Then
     '   GoTo LOGIN_ERROR
    'End If
   ' Debug.Print " Girish CheckThis Code"
    'gUserID = gCurrUser.UserID
Exit Sub

LOGIN_ERROR:
    MsgBox gAppName & " could not log you on !", vbExclamation, gAppName & " - Error"
    End
    

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

