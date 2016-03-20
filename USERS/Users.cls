VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
        
Private m_UserName As String
Private m_UserPassword As String
Private m_UserPermissions As Long

Private m_TimeOut As Long
Private m_CancelError As Boolean
Private m_Failures As Integer

Private m_MaxRetries As Byte
Private m_LoginStatus As Boolean
Private m_UserID As Long

Private WithEvents m_frmLogin As frmLogin
Attribute m_frmLogin.VB_VarHelpID = -1
Private WithEvents m_frmUsrAdd As frmUsrAdd
Attribute m_frmUsrAdd.VB_VarHelpID = -1
Private WithEvents m_frmDateChange As frmDateChange
Attribute m_frmDateChange.VB_VarHelpID = -1

Public Event LoginStatus(Status As Boolean)

Public Property Get ObjectName()
    ObjectName = "UserClass"
End Property


Public Property Let CancelError(NewVal As Boolean)
    m_CancelError = NewVal
End Property


Public Property Get CancelError() As Boolean
    CancelError = m_CancelError
End Property

Public Property Get UserID() As Long
    UserID = m_UserID
End Property
Public Property Get Failures() As Long
    Failures = m_Failures
End Property

Public Function Login(UserName As String, Password As String, LoginDate As String)
'Dim TempUserArray(5) As String
Dim FoundUser As Boolean

'Prelim Checks
If UserName = "" Or Password = "" Then
    m_Failures = m_Failures + 1
    Exit Function
End If
    
'Initialize variables
m_UserName = ""
m_UserPassword = ""
m_UserPermissions = 0
m_UserID = 0
gDbTrans.SqlStmt = "Select * from UserTab where " & _
    " LoginName = " & AddQuotes(UserName, True) & " and " & _
    " LoginPassword = " & AddQuotes(Password, True) & " And deleted = " & False
Dim Retval As Long
Dim rst As ADODB.Recordset
Retval = gDbTrans.Fetch(rst, adOpenForwardOnly)
If Retval <= 0 Or Retval > 1 Then
    FoundUser = False
Else
    FoundUser = True
End If
    
'Return the status
    If FoundUser Then
        m_UserName = UserName
        m_UserPassword = Password
        m_UserPermissions = Val(rst("Permissions"))
        m_Failures = 0
        m_UserID = Val(FormatField(rst("UserID")))
    Else
        m_Failures = m_Failures + 1
    End If
    
    Login = FoundUser
    
End Function

Public Property Get LoginStatus() As Boolean
    LoginStatus = m_LoginStatus
End Property

Public Property Let MaxRetries(NewVal As Byte)
    m_MaxRetries = NewVal
End Property
Public Property Get MaxRetries() As Byte
    MaxRetries = m_MaxRetries
End Property

Public Property Get UserPermissions() As Long
    UserPermissions = m_UserPermissions
End Property

Public Sub ShowLoginDialog()

    Set m_frmLogin = New frmLogin
    
    m_frmLogin.Show vbModal
    If Not m_LoginStatus Then
        m_UserName = ""
        m_UserPassword = ""
        m_UserPermissions = 0
        m_UserID = 0
    End If


End Sub

Public Sub ShowDateChange()
    
    Set m_frmDateChange = New frmDateChange
    
    With m_frmDateChange
        .txtTransDate = DayBeginDate
        .Show vbModal
    End With
          
    Set m_frmDateChange = Nothing
End Sub

Public Sub ShowFinChange()
    
    Set m_frmLogin = New frmLogin
    With m_frmLogin
        
      .txtUserName.Enabled = False
      .txtUserPassword.Enabled = False
      .txtUserName.Text = gCurrUser.UserName
      .txtUserPassword.Text = gCurrUser.Userpassword
    End With
      
    m_frmLogin.Show vbModal
    
End Sub


Public Sub ShowUserDialog()
'Prelim check (who is the user ?"
If Me.UserName = "" Then Exit Sub
    
'Shows UI, based on the user permissions
    Set m_frmUsrAdd = New frmUsrAdd
    m_frmUsrAdd.Show vbModal

#If junk Then

'Load the User Details to the UI
    If Not UserDetailsLoad Then
        MsgBox "Unable to load user details !", vbExclamation, gAppName & " - Error"
        Exit Sub
    End If

'Load the details of this user
    If Not m_frmUsrAdd.LoadUserDetails(m_frmUsrAdd.cmbNames.ItemData(m_frmUsrAdd.cmbNames.ListIndex)) Then
        MsgBox "Unable to load user details !", vbExclamation, gAppName & " - Error"
        Exit Sub
    End If

'Show the form now
    m_frmUsrAdd.Show vbModal

#End If
End Sub

Public Property Get TimeOut() As Long
    TimeOut = m_TimeOut
End Property
Public Property Let TimeOut(NewVal As Long)
    m_TimeOut = NewVal
End Property
Public Property Get IsAdmin() As Boolean
    IsAdmin = m_UserPermissions And perBankAdmin
End Property
Public Property Get CanDoTransaction() As Boolean
    Dim retValue As Boolean
    retValue = m_UserPermissions And perBankAdmin
    retValue = retValue Or (m_UserPermissions And perCashier)
    retValue = retValue Or (m_UserPermissions And perClerk)
    retValue = retValue Or (m_UserPermissions And perCreateAccount)
    retValue = retValue Or (m_UserPermissions And perOnlyWaves)
    CanDoTransaction = retValue
End Property
Public Property Get CanModifyAccount() As Boolean
    Dim retValue As Boolean
    retValue = m_UserPermissions And perBankAdmin
    retValue = retValue Or (m_UserPermissions And perModifyAccount)
    retValue = retValue Or (m_UserPermissions And perOnlyWaves)
    CanModifyAccount = retValue
End Property

Public Property Get CanCreateAccount() As Boolean
    Dim retValue As Boolean
    retValue = m_UserPermissions And perBankAdmin
    retValue = retValue Or (m_UserPermissions And perCreateAccount)
    retValue = retValue Or (m_UserPermissions And perOnlyWaves)
    retValue = retValue Or (m_UserPermissions And perModifyAccount)
    CanCreateAccount = retValue
End Property
Public Property Get isPigmyAgent() As Boolean
    Dim retValue As Boolean
    retValue = m_UserPermissions And perPigmyAgent
    retValue = retValue Or (m_UserPermissions And perOnlyWaves)
    isPigmyAgent = retValue
End Property
Public Property Get CanUndoTransaction() As Boolean
    Dim retValue As Boolean
    retValue = m_UserPermissions And perBankAdmin
    retValue = retValue Or (m_UserPermissions And perOnlyWaves)
    CanUndoTransaction = retValue
End Property

Public Property Get UserName() As String
    UserName = m_UserName
End Property


Private Function XXX_UserDetailsLoad() As Boolean
#If junk Then
Dim Perms As wis_Permissions
Dim IsAdmin As Boolean

'Prelim checks
    If m_frmUsrAdd Is Nothing Then
        Exit Function
    End If

'Set admin variable
    Perms = wisAdmin
    IsAdmin = Perms And Permissions

'Set the Create new Command Button
    m_frmUsrAdd.cmdNew.Enabled = IIf(IsAdmin, True, False)
    
'Fill user list in combo
    Dim Retval As Integer
    Dim I As Integer
    Dim FullName As String
    m_frmUsrAdd.cmbNames.Clear
    gDbTrans.SqlStmt = "Select NameTab.Title, NameTab.FirstName, " & _
                            " NameTab.MiddleName, NameTab.LastName,UserTab.CustomerID, NameTab.CustomerID, " & _
                            " UserTab.LoginName from NameTab, UserTab where NameTab.CustomerID = UserTab.CustomerID"
    Retval = gDbTrans.SQLFetch
    For I = 1 To gDbTrans.Records
        
        FullName = FormatField(gDbTrans.rst(0)) & " " & _
                    FormatField(gDbTrans.rst(1)) & " " & _
                    FormatField(gDbTrans.rst(2))
        'if strcomp(gdbtrans.Rst("UserTab.LoginName")) = 0  then
        If StrComp(gDbTrans.rst("LoginName"), Me.UserName, vbBinaryCompare) = 0 Or IsAdmin Then
            m_frmUsrAdd.cmbNames.AddItem FullName
            m_frmUsrAdd.cmbNames.ItemData(m_frmUsrAdd.cmbNames.newIndex) = Val(FormatField(gDbTrans.rst("NameTab.CustomerID")))
        End If
        gDbTrans.rst.MoveNext
    Next I
        
    'm_frmUsrAdd.cmbNames.AddItem Me.UserName
    m_frmUsrAdd.cmbNames.ListIndex = 0

'Exit if name not found
    If m_frmUsrAdd.cmbNames.ListIndex < 0 Then
        Exit Function
    End If

'Permissions...
    'If m_frmUsrAdd.cmbNames.ListIndex > 0 Then
        m_frmUsrAdd.lstPermissions.Enabled = IIf(IsAdmin, True, False)
        m_frmUsrAdd.cmdSelectAll.Enabled = IIf(IsAdmin, True, False)
        m_frmUsrAdd.cmdUnselectAll.Enabled = IIf(IsAdmin, True, False)
        m_frmUsrAdd.lblPermissions.Enabled = IIf(IsAdmin, True, False)
        m_frmUsrAdd.cmdApply.Enabled = False
    'Else
    '    m_frmUsrAdd.lstPermissions.Enabled = False
    '    m_frmUsrAdd.cmdSelectAll.Enabled = False
    '    m_frmUsrAdd.cmdUnselectAll.Enabled = False
    '    m_frmUsrAdd.lblPermissions.Enabled = False
    '    m_frmUsrAdd.cmdApply.Enabled = False
    '    For I = 0 To m_frmUsrAdd.lstPermissions.ListCount - 1
    '        m_frmUsrAdd.lstPermissions.Selected(I) = True
    '    Next I
    'End If
UserDetailsLoad = True
#End If
End Function




Public Property Get Userpassword() As String
    Userpassword = m_UserPassword
End Property





Private Sub Class_Terminate()
On Error Resume Next
Unload m_frmLogin
Set m_frmLogin = Nothing
End Sub


Private Sub m_frmDateChange_DateChanged(strIndianDate As String)
If strIndianDate = "" Then Exit Sub

Dim selDate As Date
selDate = GetSysFormatDate(strIndianDate)
'now check the date whether it falls in selected finanaceila year
If DateDiff("d", FinUSFromDate, selDate) < 0 Or DateDiff("d", FinUSEndDate, selDate) > 0 Then
    If Month(selDate) = 4 Then
        '' For Open cash book allow the April month for the date
        If vbNo = MsgBox("Are you sure you want the transaction date of next financial year", _
        vbYesNo, wis_MESSAGE_TITLE) Then Exit Sub
    Else
        MsgBox "Select the Date in current financial year" 'Lokapur 08350-340373
        Exit Sub
    End If
End If

DayBeginDate = strIndianDate
gStrDate = DayBeginDate 'DayBeginUSDate

With wisMain.StatusBar1
    .Panels(2).Text = DayBeginDate
    .Panels(2).Key = "TransDate"
    .Panels(3).Text = GetWeekDayName(GetSysFormatDate(DayBeginDate))
    .Panels(3).Key = "TransDay"
End With

End Sub

Private Sub m_frmLogin_CancelClicked()
If m_CancelError Then
    'Err.Raise 20005, "Cancel Error", "Login was cancelled"
End If
    m_LoginStatus = False
End Sub

Private Sub m_frmLogin_FinYearChanged(ByVal YearID As Integer)

Dim DBFileName As String

Dim FinYearClass As clsFinChange

Set FinYearClass = New clsFinChange

'Get the DBPath from the Ini file
DBFileName = FinYearClass.GetDBNameWithPath(App.Path & "\" & constFINYEARFILE, YearID)

If DBFileName = "" Then
    MsgBox "Some files corrupted contact the vendor", vbInformation
    End
End If
'Open the database
If gDbTrans Is Nothing Then Set gDbTrans = New clsTransact
'If gDbTrans Is Nothing Then Set gDbTrans = CreateObject("Transaction.Transact")

gDbTrans.CloseDB

If Not gDbTrans.OpenDB(DBFileName, constDBPWD) Then
   MsgBox "Unable find the database.", vbInformation, wis_MESSAGE_TITLE
   End
End If

'Check for the data available in the users table
' if there were no data then insert admin, admin into the database
Call InsertUserToDB
' It will create the default view which is used through out the project
Call CreateDefaultView

End Sub

Private Sub InsertUserToDB()

Dim rst As ADODB.Recordset
Dim Retval As Integer
'Trap an error
On Error GoTo ErrLine

gDbTrans.SqlStmt = "Select Count(*) as TOTUsers from UserTab"

Retval = gDbTrans.Fetch(rst, adOpenForwardOnly)
If Retval Then Retval = FormatField(rst("TotUsers"))

If Retval = 0 Then 'Put admin / admin /
    gDbTrans.SqlStmt = "INSERT INTO NameTab " & _
                "(CustomerId,FirstName,Gender,Reference )" & _
                " VALUES ( 0, 'Waves',0,0 )"
    
    gDbTrans.BeginTrans
    Call gDbTrans.SQLExecute
    
    gDbTrans.SqlStmt = "INSERT INTO UserTab " & _
                "(UserID,CustomerId,LoginName,LoginPassword,Permissions,Deleted )" & _
                " VALUES ( 0, 0, 'mahesh','sunil'," & perOnlyWaves & ",0 )"
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.SqlStmt = "INSERT INTO UserTab " & _
                "(UserID,CustomerId,LoginName,LoginPassword,Permissions,Deleted )" & _
                " VALUES ( 1, 0, 'admin','admin'," & (perOnlyWaves - 1) & ",0 )"
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
        
    gDbTrans.CommitTrans

'Here check all the tables
Dim cmb As ComboBox
    Set cmb = frmAddGroup.cmbCumulative
    Call LoadCustomerTypes(cmb)
    cmb.Clear
    Call LoadAccountGroups(cmb)
    Call LoadParentHeads(cmb)
    cmb.Clear
    Unload frmAddGroup

End If

Exit Sub

ErrLine:
    MsgBox "InsertAdmin" & vbCrLf & Err.Description, vbCritical
    
End Sub


Private Sub m_frmLogin_FinYearSelected(ByVal YearID As Integer)

Dim DBFileName As String
Dim DBPath As String

Dim FinYearClass As clsFinChange

Set FinYearClass = New clsFinChange

'Get the DBPath from the Ini file
DBFileName = FinYearClass.GetDBNameWithPath(App.Path & "\" & constFINYEARFILE, YearID)

If DBFileName = "" Then
    MsgBox "Some files corrupted contact the vendor", vbInformation
    End
End If
'Open the database
If gDbTrans Is Nothing Then
    Set gDbTrans = New clsTransact
Else
    gDbTrans.CloseDB
End If
'If gDbTrans Is Nothing Then Set gDbTrans = CreateObject("Transaction.Transact")

'If m_Failures > 0 Then gDbTrans.CloseDB

'I am Changing that to PetroWaves.mdb
'   ---- By  Lingappa Sindhanur
If Dir(DBFileName, vbNormal) = "" Then
    DBPath = FilePath(DBFileName)
    DBFileName = DBPath & "\" & constDBName
    If Dir(DBFileName, vbNormal) <> "" Then Name DBFileName As DBPath & "\" & constDBName
    DBFileName = DBPath & "\" & constDBName
End If

If Not gDbTrans.OpenDB(DBFileName, constDBPWD) Then
   MsgBox "Unable find the database.", vbInformation, wis_MESSAGE_TITLE
   End
End If


'Now Get the Language Offset In Data BAse
gDbTrans.SqlStmt = "select * From Install Where KeyData = 'Language'"
Dim rst As Recordset
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    Dim LngOffset As Integer
    If UCase(FormatField(rst("ValueData"))) = "KANNADA" Then LngOffset = wis_KannadaOffset
    If UCase(FormatField(rst("ValueData"))) = "SAMHITA" Then LngOffset = wis_KannadaSamhitaOffset
    If LngOffset = wis_NoLangOffset Then
        gLangOffSet = LngOffset
        Call WriteToIniFile("Language", "Language", _
                IIf(LngOffset, "Kannada", "English"), App.Path & "\FinYear.Fin")
        
        If LngOffset = wis_KannadaSamhitaOffset Then
            Call WriteToIniFile("Language", "LanguageTool", _
                IIf(LngOffset = wis_KannadaSamhitaOffset, "SAMHITA", ""), App.Path & "\FinYear.Fin")
        End If
    End If
End If

'Check for the data available in the users table
' if there were no data then insert admin, admin into the database
Call InsertUserToDB
' It will create the default view which is used through out the project
Call CreateDefaultView

Set FinYearClass = Nothing
End Sub


Private Sub m_frmLogin_LoginClicked(UserName As String, Userpassword As String, LoginDate As String, UnloadDialog As Boolean)
    m_LoginStatus = Login(UserName, Userpassword, LoginDate)
    If Not m_LoginStatus Then
        'MsgBox "Invalid login !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(515), vbExclamation, gAppName & " - Error"
        
        If m_Failures >= m_MaxRetries Then
            If m_CancelError Then
                'Err.Raise 20005, "Cancel Error", "Login was cancelled"
            End If
            UnloadDialog = True
        Else
            UnloadDialog = False
        End If
    Else
        UnloadDialog = True
    End If
End Sub

Private Sub m_frmUsrAdd_GetUserDetails(LoginName As String, Password As String, Permissions As Long)
    LoginName = m_UserName
    Password = m_UserPassword
    Permissions = m_UserPermissions
End Sub

Private Sub m_frmUsrAdd_GetUserID(l_UserId As Long)
    l_UserId = m_UserID
End Sub

