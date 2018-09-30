VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmUsrAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Information"
   ClientHeight    =   7335
   ClientLeft      =   2385
   ClientTop       =   1560
   ClientWidth     =   7380
   Icon            =   "Usradd.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   400
      Left            =   6000
      TabIndex        =   19
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame fraUser 
      Height          =   6075
      Left            =   210
      TabIndex        =   0
      Top             =   540
      Width           =   7005
      Begin VB.ComboBox cmbNamesEnglish 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3600
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   400
         Left            =   4170
         TabIndex        =   17
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   400
         Left            =   2880
         TabIndex        =   16
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmdDetail 
         Caption         =   "Details"
         Height          =   400
         Left            =   5070
         TabIndex        =   15
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cmbNames 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   930
         Width           =   4575
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6390
         TabIndex        =   2
         Top             =   930
         Width           =   315
      End
      Begin VB.CommandButton cmdUnselectAll 
         Caption         =   "Unselect all"
         Height          =   400
         Left            =   5580
         TabIndex        =   13
         Top             =   4530
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select all"
         Height          =   400
         Left            =   5580
         TabIndex        =   12
         Top             =   4020
         Width           =   1215
      End
      Begin VB.ListBox lstPermissions 
         Height          =   2535
         Left            =   210
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   2700
         Width           =   5175
      End
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1740
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1770
         Width           =   1455
      End
      Begin VB.TextBox txtLoginName 
         Height          =   345
         Left            =   1740
         TabIndex        =   5
         Top             =   1350
         Width           =   1455
      End
      Begin VB.TextBox txtConfirm 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   4830
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1770
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   0
         X2              =   6580
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   60
         X2              =   6580
         Y1              =   2310
         Y2              =   2310
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         Caption         =   "Add Users"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   870
         TabIndex        =   14
         Top             =   360
         Width           =   3825
      End
      Begin VB.Image img 
         Height          =   435
         Left            =   240
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblPermissions 
         Caption         =   "Permissions (Select all that apply)"
         Height          =   225
         Left            =   210
         TabIndex        =   10
         Top             =   2370
         Width           =   3255
      End
      Begin VB.Label lblUserReg 
         Caption         =   "Register user:"
         Height          =   315
         Left            =   210
         TabIndex        =   1
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label lblConfirmPassword 
         Caption         =   "Confirm password:"
         Height          =   285
         Left            =   3300
         TabIndex        =   8
         Top             =   1830
         Width           =   1515
      End
      Begin VB.Label lblLoginPassword 
         Caption         =   "Login password:"
         Height          =   285
         Left            =   210
         TabIndex        =   6
         Top             =   1770
         Width           =   1485
      End
      Begin VB.Label lblLoginName 
         Caption         =   "Login name:"
         Height          =   285
         Left            =   210
         TabIndex        =   4
         Top             =   1380
         Width           =   1185
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   6615
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   11668
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "User"
            Key             =   "user"
            Object.Tag             =   "user"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Trans"
            Key             =   "trans"
            Object.Tag             =   "trans"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraTrans 
      Height          =   6135
      Left            =   240
      TabIndex        =   20
      Top             =   480
      Width           =   6975
      Begin VB.CommandButton cmdEndDate 
         Caption         =   "..."
         Height          =   315
         Left            =   6555
         TabIndex        =   28
         Top             =   840
         Width           =   315
      End
      Begin VB.CommandButton cmdStDate 
         Caption         =   "..."
         Height          =   315
         Left            =   2970
         TabIndex        =   27
         Top             =   840
         Width           =   315
      End
      Begin VB.TextBox txtEndDate 
         Height          =   315
         Left            =   5145
         TabIndex        =   26
         Top             =   840
         Width           =   1230
      End
      Begin VB.TextBox txtStartDate 
         Height          =   315
         Left            =   1605
         TabIndex        =   25
         Top             =   840
         Width           =   1290
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   4815
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   8493
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   6060
         TabIndex        =   22
         Top             =   240
         Width           =   795
      End
      Begin VB.ComboBox cmbUser 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label lblDate2 
         AutoSize        =   -1  'True
         Caption         =   "&Ending date :"
         Height          =   195
         Left            =   3615
         TabIndex        =   30
         Top             =   900
         Width           =   1605
      End
      Begin VB.Label lblDate1 
         AutoSize        =   -1  'True
         Caption         =   "&Starting date :"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   885
         Width           =   990
      End
      Begin VB.Label lblUser 
         Caption         =   "Register user:"
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   390
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmUsrAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_FromUSDate As Date
Private m_ToUSDate As Date
Private m_rstTrans As Recordset
Dim M_SlNo As Integer
Dim m_rowNo As Integer
Private m_selUserID As Long
Private m_UserID As Long
Private m_Permissions As wis_Permissions

Public Event UserModified(LoginName As String, Password As String, Permissions As Long)
'Public Event GetUserDetails(LoginName As String, Password As String, Permissions As Long)
Public Event GetUserID(l_UserId As Long)

Private m_NewUserID As Long

Private m_CustReg As clsCustReg


Private Function EnableControls(Enable As Boolean)

Dim Perms As wis_Permissions
Dim UserPerms As wis_Permissions

'Routine enables controls on the UI based on the conditions

'First disable all controls
    lstPermissions.BackColor = wisGray
    lblPermissions.Enabled = False:
    
    If cmbNames.ListIndex < 0 Then Exit Function
    UserPerms = cmbNames.ItemData(cmbNames.ListIndex)

'First check if admin is the user
    Perms = perBankAdmin
    If cmbNames.ItemData(cmbNames.ListIndex) = 0 Then  'Disable Permissions section because
                                            'admin's permissions can never be modified
        'optEmployee.Enabled = False
        'optPigmy.Enabled = False
        'lstPermissions.Enabled = False
        lstPermissions.BackColor = wisGray
        cmdSelectAll.Enabled = False: cmdUnselectAll.Enabled = False
    Else
        'optEmployee.Enabled = True
        'optPigmy.Enabled = True
        'lstPermissions.Enabled = False
        lstPermissions.BackColor = wisGray
        cmdSelectAll.Enabled = False: cmdUnselectAll.Enabled = False
    End If
        
If cmbNames.ItemData(cmbNames.ListIndex) = 0 Then
    cmdApply.Enabled = False
    Exit Function
End If

'Next check if this user has permissions to modify
    Perms = perCreateAccount
    If (m_Permissions And Perms) Or m_Permissions = perBankAdmin Then   'User can modify anybody's account
        cmdNew.Enabled = True
        lstPermissions.BackColor = vbWhite
        cmdSelectAll.Enabled = True: cmdUnselectAll.Enabled = True
        lstPermissions.Tag = "True"
    Else
        cmdNew.Enabled = False
        lstPermissions.BackColor = wisGray
        lstPermissions.Tag = "False"
        cmdSelectAll.Enabled = False: cmdUnselectAll.Enabled = False
    End If
    
    If m_Permissions = perBankAdmin Then
        cmdNew.Enabled = True
        txtPassword.ToolTipText = txtPassword.Text
        lstPermissions.Enabled = True: lstPermissions.BackColor = vbWhite
        cmdSelectAll.Enabled = True: cmdUnselectAll.Enabled = True
    End If
cmdApply.Enabled = False

End Function

Private Sub InitGrid()
If gCurrUser.IsAdmin Then
    grd.Clear
    grd.Cols = 5
    grd.Rows = 18
    grd.TextMatrix(0, 0) = GetResourceString(33) ' "slno"
    grd.TextMatrix(0, 1) = GetResourceString(33) ' "Date"
    grd.TextMatrix(0, 2) = GetResourceString(36) ' "Account number"
    grd.TextMatrix(0, 3) = GetResourceString(205) ' "Customer"
    grd.TextMatrix(0, 4) = GetResourceString(38) ' "Trans Type"
    
    Dim Wid As Single
    Dim I As Integer
    Wid = (grd.Width - 185) / grd.Cols
    Dim ColCount As Integer
    For ColCount = 0 To grd.Cols - 1
        Wid = GetSetting(App.EXEName, "UserTrans", _
                "ColWidth" & ColCount, 1 / grd.Cols) * grd.Width
        If Wid > grd.Width * 0.9 Then Wid = grd.Width / grd.Cols
        grd.ColWidth(ColCount) = Wid
            
    Next ColCount
End If
End Sub

Private Sub LoadUsers()

'Request User Details from the Calling Application
    Dim l_UserId As Long
    Dim rst As Recordset
    Dim Perms As wis_Permissions
    Dim I As Integer
    
    RaiseEvent GetUserID(l_UserId)
    m_UserID = l_UserId

'Based on the userid, check out his permissions and fill up name combo accordingly
    Dim Retval As Long
    Dim TotalPerms As Long
    
    gDbTrans.SqlStmt = "Select * from UserTab " & _
            " WHERE UserID = " & l_UserId
    Retval = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If Retval <= 0 Or Retval > 1 Then
        'MsgBox "Unable to locate the user !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(610), vbExclamation, gAppName & " - Error"
    Else
        TotalPerms = Val(rst("Permissions"))
        m_Permissions = TotalPerms
        If m_Permissions And perBankAdmin Then cmdRemove.Enabled = True
    End If
    Perms = m_Permissions
    cmbNames.Clear
    cmbNamesEnglish.Clear
ReCheck:
    gDbTrans.SqlStmt = "Select * from UserTab where UserID = " & l_UserId
    If (Perms And perBankAdmin) Or (Perms = perOnlyWaves) Then
        'Load all the user names from database
        gDbTrans.SqlStmt = "Select * from UserTab " & _
            " Where Deleted = " & False & " And Permissions <> " & perOnlyWaves
        cmbNames.AddItem GetResourceString(149), 0
        cmbNamesEnglish.AddItem LoadResString(149)
    End If
        
ExecuteLine:
    Retval = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If Retval <= 0 Then Exit Sub
    If m_CustReg Is Nothing Then Set m_CustReg = New clsCustReg
    For I = 1 To rst.recordCount
        If Val(rst("UserID")) = 0 Then
            cmbNames.AddItem "Administrator"
            cmbNamesEnglish.AddItem "Administrator"
        Else
            Dim nameEnglish As String
            cmbNames.AddItem m_CustReg.CustomerNameNew(Val(rst("CustomerID")), nameEnglish)
            cmbNames.ItemData(cmbNames.newIndex) = Val(rst("UserID"))
            
            cmbNamesEnglish.AddItem nameEnglish
            cmbNamesEnglish.ItemData(cmbNames.newIndex) = Val(rst("UserID"))
            
            cmbUser.AddItem m_CustReg.CustomerName(Val(rst("CustomerID")))
            cmbUser.ItemData(cmbUser.newIndex) = Val(rst("UserID"))
            

        End If
        rst.MoveNext
    Next I
    
    If I > 1 Then cmbNames.ListIndex = 0
    
'Load details to UI
If Not LoadToUI(l_UserId) Then GoTo ErrLine

'Set the list index
For I = 0 To cmbNames.ListCount - 1
    If cmbNames.ItemData(I) = l_UserId Then cmbNames.ListIndex = I: Exit For
Next I

'Position then List Box
'lstPermissions.Top = 2950
'lstPermissions.Height = 1700

Call EnableControls(True)
    
Exit Sub

ErrLine:
Me.MousePointer = vbDefault
Screen.MousePointer = vbDefault

End Sub

Private Sub LoadUserPermission()
'While changing the List values It may go into the Ifinite Loop
Static InLoop As Boolean

Dim Perms As wis_Permissions
Dim tmpPerms As wis_Permissions
    If cmbNames.ListIndex < 0 Then Exit Sub
    
    Perms = cmbNames.ItemData(cmbNames.ListIndex)
    lstPermissions.Enabled = True
    
    'Full Permission
    tmpPerms = perBankAdmin
    lstPermissions.Selected(0) = IIf(tmpPerms And Perms, True, False)
    'Create Account Permission
    tmpPerms = perCreateAccount
    lstPermissions.Selected(1) = IIf(tmpPerms And Perms, True, False)
    'VIEW Transaction Permission
    tmpPerms = perReadOnly
    lstPermissions.Selected(2) = IIf(tmpPerms And Perms, True, False)
    'MAKE Transaction Permission
    tmpPerms = perClerk
    lstPermissions.Selected(3) = IIf(tmpPerms And Perms, True, False)
    'UNDo Permission
    tmpPerms = perBankAdmin
    lstPermissions.Selected(4) = IIf(tmpPerms And Perms, True, False)
    
    tmpPerms = perPigmyAgent ': optPigmy.value = False
    'If tmpPerms And Perms Then optPigmy.value = True

End Sub


Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

lblCaption.Caption = GetResourceString(149)
lblUserReg.Caption = GetResourceString(150)
lblUser.Caption = GetResourceString(150)
cmdDetail.Caption = GetResourceString(295)  'Details
cmdLoad.Caption = GetResourceString(3)  'Load

lblLoginName.Caption = GetResourceString(151, 35)
lblLoginPassword.Caption = GetResourceString(151, 153)
lblConfirmPassword.Caption = GetResourceString(154, 153)

lblPermissions.Caption = GetResourceString(156)
cmdSelectAll.Caption = GetResourceString(25)
cmdUnselectAll.Caption = GetResourceString(26)
cmdRemove.Caption = GetResourceString(12)
cmdApply.Caption = GetResourceString(6)
cmdClose.Caption = GetResourceString(11)

lblDate1.Caption = GetResourceString(109)
lblDate2.Caption = GetResourceString(110)
'TabStrip1.Tabs(1).Caption = GetResourceString(3) '156
TabStrip1.Tabs(2).Caption = GetResourceString(38)
End Sub


Private Sub UpdateTransGrid()
    Dim transType As wisTransactionTypes
    With grd
        While Not m_rstTrans.EOF
            If .Rows < m_rowNo + 3 Then .Rows = m_rowNo + 4
            M_SlNo = M_SlNo + 1: m_rowNo = m_rowNo + 1
            .Row = m_rowNo
            .Col = 0: .Text = M_SlNo
            .Col = 1: .Text = FormatField(m_rstTrans("TransDate"))
            .Col = 2: .Text = FormatField(m_rstTrans("AccNum"))
            .Col = 3: .Text = FormatField(m_rstTrans("Name"))
            transType = FormatField(m_rstTrans("TransType"))
            .Col = 4: .Text = GetResourceString(IIf(transType = wDeposit Or transType = wContraDeposit, 271, 272))
            
            m_rstTrans.MoveNext
        Wend
        'm_rstTrans.MoveNext
    End With
End Sub

Private Sub UpdateUserTransactions(tableName1 As String, AccountName As String)
On Error GoTo Exit_Line
    Dim CurrRow As Integer
    CurrRow = m_rowNo
'Get the each transcations
    Dim mastrTable As String
    Dim transTable As String
    
    gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name,TransDate" & _
            " FROM " & tableName1 & "Trans A," & tableName1 & "Master B,QryName C " & _
            " WHERE TransDate >= #" & m_FromUSDate & "# AND TransDate <= #" & m_ToUSDate & "#" & _
            " And A.AccID = B.AccID and A.UserID = " & m_selUserID & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
    
    If gDbTrans.Fetch(m_rstTrans, adOpenDynamic) > 0 Then
        m_rowNo = m_rowNo + 1
        grd.Row = m_rowNo
        grd.Col = 1: grd.Text = AccountName
        UpdateTransGrid
    End If
    transTable = tableName1 & "IntTrans"
    If UCase(tableName1) = "SB" Or UCase(tableName1) = "CA" Then transTable = tableName1 & "PLTrans"
        
    gDbTrans.SqlStmt = "SELECT AccNum,A.AccId,Amount,TransType,Name,TransDate" & _
            " FROM " & transTable & " A," & tableName1 & "Master B,QryName C " & _
            " WHERE TransDate >= #" & m_FromUSDate & "# AND TransDate <= #" & m_ToUSDate & "#" & _
            " And A.AccID = B.AccID and A.UserID = " & m_selUserID & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
    
    If gDbTrans.Fetch(m_rstTrans, adOpenDynamic) > 0 Then
        m_rowNo = m_rowNo + 1
        grd.Row = m_rowNo
        grd.Col = 1: grd.Text = AccountName & " " & GetResourceString(47)
        UpdateTransGrid
    End If
    
    
Exit_Line:
    'If CurrRow = m_rowNo Then m_rowNo = m_rowNo - 1
End Sub


Private Sub UpdateUserLoanTransactions(SchemeID As Integer, AccountName As String)
    Dim CurrRow As Integer
    CurrRow = m_rowNo
'Get the each transcations
    
    gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId,Amount,TransType,Name,TransDate" & _
                " FROM LoanTrans A,LoanMaster B,QryName C " & _
                " WHERE TransDate >= #" & m_FromUSDate & "# AND TransDate <= #" & m_ToUSDate & "#" & _
                " And A.LoanID = B.LoanID and A.UserID = " & m_selUserID & _
                " ANd SchemeID = " & SchemeID & _
                " And C.CustomerId = B.CustomerID ORDER BY Transdate,val(AccNum)"
        
    
    If gDbTrans.Fetch(m_rstTrans, adOpenDynamic) > 0 Then
        m_rowNo = m_rowNo + 1
        grd.Row = m_rowNo
        grd.Col = 3: grd.Text = AccountName: grd.CellFontBold = True
        UpdateTransGrid
    End If
    Dim intType As Integer
Dim FieldName As String
Dim headName As String

intType = 0

StartInt:
FieldName = IIf(intType = 0, "IntAmount", IIf(intType = 1, "PenalIntAmount", "MiscAmount"))
headName = AccountName & " " & _
           GetResourceString(IIf(intType = 0, 47, IIf(intType = 1, 345, 327)))

    gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId," & _
                FieldName & "  as amount,TransType,Name,TransDate" & _
                " FROM LoanIntTrans A,LoanMaster B,QryName C " & _
                " WHERE " & FieldName & " > 0 And TransDate >= #" & m_FromUSDate & "# AND TransDate <= #" & m_ToUSDate & "#" & _
                " And A.LoanID = B.LoanID " & _
                " and A.UserID = " & m_selUserID & " and SchemeID = " & SchemeID & _
                " And C.CustomerId = B.CustomerID ORDER BY Transdate, val(AccNum)"
        
    If gDbTrans.Fetch(m_rstTrans, adOpenDynamic) > 0 Then
        m_rowNo = m_rowNo + 1
        grd.Row = m_rowNo
        grd.Col = 3: grd.Text = headName: grd.CellFontBold = True
        UpdateTransGrid
    End If
    
    intType = intType + 1
    If intType < 3 Then GoTo StartInt
    
    
Exit_Line:
    'If CurrRow = m_rowNo Then m_rowNo = m_rowNo - 1
End Sub

Private Sub cmbNames_Click()

Dim l_UserId As Long

If cmbNames.ListIndex < 0 Then Exit Sub

Dim Perms As wis_Permissions
Perms = gCurrUser.UserPermissions
cmdRemove.Enabled = cmbNames.ListIndex
l_UserId = cmbNames.ItemData(cmbNames.ListIndex)

If ((Perms And perBankAdmin) Or (Perms And perOnlyWaves)) _
    And cmbNames.ListIndex = 0 Then Exit Sub
'Call LoadUserPermission

If m_NewUserID = l_UserId Then

    'Get the Login Name
    Dim rst As Recordset
    Dim LoginName As String

'Select a login name * Password by default
    l_UserId = 0
    'Get the UserLogin name & password
    Do
        l_UserId = l_UserId + 1
        gDbTrans.SqlStmt = "Select LoginName from UserTab" & _
            " WHERE LoginName = 'User" & Format(l_UserId, "000") & "'"
        If gDbTrans.Fetch(rst, adOpenForwardOnly) = 0 Then Exit Do
    Loop
    LoginName = "user" & Format(l_UserId, "000")
    txtLoginName = LoginName
    txtPassword = LoginName
    txtConfirm = LoginName
    ''Initially allw only reasd permissions
    For l_UserId = 1 To lstPermissions.ListCount - 2
        lstPermissions.Selected(l_UserId) = False
    Next
    lstPermissions.Selected(lstPermissions.ListCount - 1) = True
    Exit Sub

End If

Call LoadToUI(l_UserId)
Call EnableControls(True)

If m_NewUserID Then cmdNew.Enabled = False

End Sub

Private Sub cmbUser_Change()
    If m_selUserID <> cmbUser.ItemData(cmbUser.ListIndex) Then Call InitGrid
    m_selUserID = cmbUser.ItemData(cmbUser.ListIndex)
End Sub

Private Sub cmbUser_Click()
    Call InitGrid
End Sub


Private Sub cmdApply_Click()
    
'Check if name was specified
    If cmbNames.ListIndex < 0 Then
        'MsgBox "User name not selected !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(611), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If

'Check for login name
    If Trim$(txtLoginName.Text) = "" Then
        'MsgBox "Login name not specified !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(516), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If
    
'Check for valid Login names
    If InStr(1, txtLoginName.Text, " ", vbBinaryCompare) > 0 Then
        'MsgBox "Login name should not have spaces!", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(516), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If

'Check for Valid password
    If InStr(1, txtPassword.Text, " ", vbBinaryCompare) > 0 Then
        'MsgBox "Password should not be blank!", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(515), vbExclamation, gAppName & " - Error"
        Exit Sub
    End If
    
'Compare passwords
    If StrComp(txtPassword.Text, txtConfirm.Text, vbBinaryCompare) <> 0 Then
        'MsgBox "Passwords do not match !", vbExclamation
        MsgBox GetResourceString(516), vbExclamation
        Exit Sub
    End If

'Compute permissions
    Dim I As Integer
    Dim Perms  As wis_Permissions
    Dim tmpPerms As wis_Permissions
    
    For I = 0 To lstPermissions.ListCount - 1
        If lstPermissions.Selected(I) = True Then _
                Perms = Perms Or lstPermissions.ItemData(I)
    Next I

Dim l_UserId As Long

l_UserId = cmbNames.ItemData(cmbNames.ListIndex)

If (Perms And perPigmyAgent) And (Perms > (perPigmyAgent Or perReadOnly)) Then
    Debug.Print "Kannada"
    MsgBox "Pigmy Agent Can not have other permissions", vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
    
'Check if new rec (Never a new record)
    Dim NewRec As Boolean
    Dim isThereAnyAdmin As Boolean
    Dim rst As ADODB.Recordset
    isThereAnyAdmin = Perms And perBankAdmin
    
    gDbTrans.SqlStmt = "Select UserID from UserTab where " & _
            " USerID = " & l_UserId
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        isThereAnyAdmin = True
        NewRec = True
    ElseIf Not isThereAnyAdmin Then
        ' Before assignng the Permission Check at least one Admin Has to Be there
        'isThereAnyAdmin = True
        NewRec = False
        gDbTrans.SqlStmt = "Select * from UserTab Where " & _
            " UserID <> " & l_UserId & " AND Deleted = " & False & _
            " And Permissions < " & perOnlyWaves
            
        tmpPerms = perBankAdmin
        If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
            rst.MoveFirst
            Do While rst.EOF = False
                If FormatField(rst("Permissions")) And tmpPerms Then
                    If FormatField(rst("Permissions")) <> perPigmyAgent Then
                        isThereAnyAdmin = True
                        If isThereAnyAdmin Then Exit Do
                    End If
                End If
                rst.MoveNext
            Loop
        End If
    End If

' Check for the existance of admin in the data base
    If Not isThereAnyAdmin Then
        MsgBox "There is no employee as administrator " & vbCrLf & _
                "So you can not do this operation"
        Exit Sub 'GoTo ExitLine
    End If


'Insert into the DB
    Dim SqlStr As String
    If NewRec Then

        SqlStr = "Insert into UserTab (UserID,CustomerId," & _
                " LoginName, LoginPassword, " & _
                " Permissions,CreateDate,Deleted) values (" & _
                l_UserId & ", " & _
                m_CustReg.CustomerID & "," & _
                AddQuotes(Trim$(txtLoginName.Text), True) & ", " & _
                AddQuotes(Trim$(txtPassword.Text), True) & ", " & _
                Perms & ", #" & gStrDate & "#, False )"
    Else
        SqlStr = "Update UserTab set " & _
                " LoginName = " & AddQuotes(Trim$(txtLoginName.Text), True) & ", " & _
                " loginPassword = " & AddQuotes(Trim$(txtPassword.Text), True) & "," & _
                " Permissions = " & Perms & _
                " WHERE UserID = " & l_UserId & ";"
    End If
    
    gDbTrans.BeginTrans
    'iF ANY ALTERATIONS MADE TO THE CUSTOMER
    If m_NewUserID Then
        If m_NewUserID = l_UserId Then Call m_CustReg.SaveCustomer
    Else
        m_CustReg.ModuleID = wis_Users
        Call m_CustReg.SaveCustomer
    End If
    gDbTrans.SqlStmt = SqlStr
    If Not gDbTrans.SQLExecute Then
        'MsgBox "Unable to perform transaction !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(535), vbExclamation, gAppName & " - Error"
        gDbTrans.RollBack
        Exit Sub
    End If
    
    
    gDbTrans.CommitTrans
    
    If NewRec Then m_NewUserID = 0

ExitLine:
cmdApply.Enabled = False
End Sub

Private Sub cmdClose_Click()

Unload Me
End Sub

Private Sub cmdDetails_Click()

End Sub

Private Sub cmdDetail_Click()
    Dim l_UserId As Long
    l_UserId = cmbNames.ItemData(cmbNames.ListIndex)
    If l_UserId < 1 Then Exit Sub
    Load frmEmpoyee
    frmEmpoyee.UserID = l_UserId
    frmEmpoyee.Show 1
End Sub

Private Sub cmdEndDate_Click()
With Calendar
    .Left = Me.Left + fraTrans.Left + cmdEndDate.Left - .Width
    .Top = Me.Top + fraTrans.Top + cmdEndDate.Top + 300
    .selDate = txtEndDate.Text
    .Show vbModal, Me
    If .selDate <> "" Then txtEndDate.Text = Calendar.selDate
End With

End Sub

Private Sub cmdLoad_Click()
If cmbUser.ListIndex < 0 Then Exit Sub
m_selUserID = cmbUser.ItemData(cmbUser.ListIndex)
If m_selUserID < 0 Then Exit Sub

If Me.txtStartDate.Enabled Then
    If Not DateValidate(txtStartDate.Text, "/", True) Then
        'MsgBox "Invalid date specified !" & vbCrLf & vbCrLf & "Please specify in DD/MM/YYYY format", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501) & vbCrLf & vbCrLf & "Please specify in DD/MM/YYYY format", vbExclamation, gAppName & " - Error"
        ActivateTextBox txtStartDate
        Exit Sub
    End If
End If
If txtEndDate.Enabled Then
    If Not DateValidate(txtEndDate.Text, "/", True) Then
        'MsgBox "Invalid date specified !" & vbCrLf & vbCrLf & "Please specify in DD/MM/YYYY format", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(501) & vbCrLf & vbCrLf & "Please specify in DD/MM/YYYY format", vbExclamation, gAppName & " - Error"
        ActivateTextBox txtEndDate
        Exit Sub
    End If
End If
m_FromUSDate = GetSysFormatDate(txtStartDate.Text)
m_ToUSDate = GetSysFormatDate(txtEndDate.Text)

Call InitGrid
Dim AccountName As String
m_rowNo = 0
M_SlNo = 0
AccountName = GetResourceString(351)
Call UpdateUserTransactions("Mem", AccountName)

AccountName = GetResourceString(421)
Call UpdateUserTransactions("SB", AccountName)

AccountName = GetResourceString(422)
Call UpdateUserTransactions("CA", AccountName)

AccountName = GetResourceString(423)
Call UpdateUserTransactions("FD", AccountName)


AccountName = GetResourceString(424)
Call UpdateUserTransactions("RD", AccountName)

AccountName = GetResourceString(425)
Call UpdateUserTransactions("PD", AccountName)

Call UpdateBKCCUserTransactions

Call UpdateUserTransactionsDepositLoan
''Update for Loan transactions
gDbTrans.SqlStmt = "Select * From LoanScheme"
Dim rstLoanSchemes As Recordset
If gDbTrans.Fetch(rstLoanSchemes, adOpenDynamic) > 0 Then
    While Not rstLoanSchemes.EOF
        Call UpdateUserLoanTransactions(FormatField(rstLoanSchemes("SchemeID")), FormatField(rstLoanSchemes("SchemeName")))
        rstLoanSchemes.MoveNext
    Wend
End If
End Sub

Private Sub UpdateUserTransactionsDepositLoan()
Dim AccName As String
AccName = GetResourceString(43, 58)
On Error GoTo Exit_Line
    Dim CurrRow As Integer
    CurrRow = m_rowNo
'Get the each transcations
    Dim mastrTable As String
    Dim transTable As String
    
   gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId,Amount,TransType,Name,TransDate" & _
            " FROM DepositLoanTrans A,DepositLoanMaster B,QryName C " & _
            " WHERE TransDate >= #" & m_FromUSDate & "# AND TransDate <= #" & m_ToUSDate & "#" & _
            " And A.LoanID = B.LoanID and A.UserID = " & m_selUserID & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate,val(AccNum)"
   
    
    If gDbTrans.Fetch(m_rstTrans, adOpenDynamic) > 0 Then
        m_rowNo = m_rowNo + 1
        grd.Row = m_rowNo
        grd.Col = 3: grd.Text = AccName: grd.CellFontBold = True
        UpdateTransGrid
    End If
    
Dim intType As Integer
Dim FieldName As String
Dim headName As String

intType = 0

StartInt:
FieldName = IIf(intType = 0, "Amount", IIf(intType = 1, "PenalAmount", "MiscAmount"))
headName = AccName & " " & _
        GetResourceString(IIf(intType = 0, 47, IIf(intType = 1, 345, 327)))

    gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId, " & FieldName & " as Amount,TransType,Name,TransDate" & _
            " FROM DepositLoanIntTrans A,DepositLoanMaster B,QryName C " & _
            " WHERE TransDate >= #" & m_FromUSDate & "# AND TransDate <= #" & m_ToUSDate & "#" & _
            " And A.LoanID = B.LoanID and A.UserID = " & m_selUserID & _
            " And " & FieldName & " > 0 " & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate,val(AccNum)"
    
    If gDbTrans.Fetch(m_rstTrans, adOpenDynamic) > 0 Then
        m_rowNo = m_rowNo + 1
        grd.Row = m_rowNo
        grd.Col = 3: grd.Text = headName: grd.CellFontBold = True
        UpdateTransGrid
    End If
    
    intType = intType + 1
    If intType < 3 Then GoTo StartInt
    
Exit_Line:
    'If CurrRow = m_rowNo Then m_rowNo = m_rowNo - 1
End Sub

Private Sub UpdateBKCCUserTransactions()

On Error GoTo Exit_Line
    Dim CurrRow As Integer
    CurrRow = m_rowNo
'Get the each transcations
    Dim AccName As String
    
Dim Deposit As Boolean

Deposit = True

Repeat:

AccName = GetResourceString(229, IIf(Deposit, 43, 58))
  gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId,Amount,TransType,Name,TransDate" & _
        " FROM BKCCTrans A,BkccMaster B,QryName C " & _
        " WHERE TransDate >= #" & m_FromUSDate & "# AND TransDate <= #" & m_ToUSDate & "#" & _
        " And Amount > 0 and A.LoanID = B.LoanID And Deposit = " & Deposit & _
        " And C.CustomerId = B.CustomerID and A.UserID = " & m_selUserID & " ORDER BY Transdate,val(AccNum)"

    
    If gDbTrans.Fetch(m_rstTrans, adOpenDynamic) > 0 Then
        m_rowNo = m_rowNo + 1
        grd.Row = m_rowNo
        grd.Col = 3: grd.Text = AccName: grd.CellFontBold = True
        UpdateTransGrid
    End If
    
Dim intType As Integer
Dim FieldName As String
Dim headName As String

intType = 0

StartInt:
FieldName = IIf(intType = 0, "IntAmount", IIf(intType = 1, "PenalIntAmount", "MiscAmount"))
headName = AccName & " " & _
           GetResourceString(IIf(intType = 0, 47, IIf(intType = 1, 345, 327)))

    gDbTrans.SqlStmt = "SELECT AccNum,A.LoanId, " & FieldName & " as Amount,TransType,Name,TransDate" & _
            " FROM BKCCIntTrans A,DepositLoanMaster B,QryName C " & _
            " WHERE TransDate >= #" & m_FromUSDate & "# AND TransDate <= #" & m_ToUSDate & "#" & _
            " And A.LoanID = B.LoanID and A.UserID = " & m_selUserID & _
            " And " & FieldName & " > 0 " & _
            " And C.CustomerId = B.CustomerID ORDER BY Transdate"
    
    If gDbTrans.Fetch(m_rstTrans, adOpenDynamic) > 0 Then
        m_rowNo = m_rowNo + 1
        grd.Row = m_rowNo
        grd.Col = 3: grd.Text = headName: grd.CellFontBold = True
        UpdateTransGrid
    End If
    
    intType = intType + 1
    If intType < 3 Then GoTo StartInt
    
    If Deposit = True Then
        intType = 0
        Deposit = False
        GoTo Repeat
    End If
    
Exit_Line:
    'If CurrRow = m_rowNo Then m_rowNo = m_rowNo - 1
End Sub

Private Sub cmdNew_Click()

Dim l_UserId As Long
Dim rst As ADODB.Recordset
With cmbNames
    If .ListIndex < 0 Then Exit Sub
    l_UserId = .ItemData(.ListIndex)
End With

If m_CustReg Is Nothing Then Set m_CustReg = New clsCustReg

If l_UserId Then
    gDbTrans.SqlStmt = "Select * from USerTab where USerID = " & l_UserId
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        'MsgBox "Error creating new user !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(616), vbExclamation, gAppName & " - Error"
        Exit Sub
    Else
        l_UserId = FormatField(rst("CustomerId"))
        m_CustReg.LoadCustomerInfo (l_UserId)
        m_CustReg.ShowDialog
        Exit Sub
    End If
End If

'Show dialog for new customer
'Dim Perms As wis_Permissions
'Perms = gCurrUser.UserPermissions
    
'Get new userid
    gDbTrans.SqlStmt = "Select Max(UserID) from USerTab"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <= 0 Then
        'MsgBox "Error creating new user !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(616), vbExclamation, gAppName & " - Error"
        Exit Sub
    Else
        m_NewUserID = FormatField(rst(0)) + 1
    End If
    
    m_CustReg.NewCustomer
    m_CustReg.ModuleID = wis_Users
    m_CustReg.ShowDialog
    If m_CustReg.CustomerID <= 0 Then Exit Sub
    
    'Now check whether this user already there and marked as Delete then
    gDbTrans.SqlStmt = "Select * from UserTab" & _
                " WHERE CustomerID = " & m_CustReg.CustomerID
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then  'This person already registered as user
        If FormatField(rst("DELETED")) Then
            gDbTrans.SqlStmt = "UPDATE UserTab set Permissions = 0," & _
                " Deleted = False WHERE CustomerID = " & m_CustReg.CustomerID
            gDbTrans.BeginTrans
            If Not gDbTrans.SQLExecute Then
                gDbTrans.RollBack
                'MsgBox "Unable to register new user !", vbExclamation, gAppName & " - Error"
                MsgBox GetResourceString(617), vbExclamation, gAppName & " - Error"
                Exit Sub
            End If
            gDbTrans.CommitTrans
            Exit Sub
        End If
    End If
    
    'get the New User ID
    gDbTrans.SqlStmt = "Select Max(UserID) From UserTab"
    l_UserId = 1
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then _
        l_UserId = FormatField(rst(0)) + 1
    m_NewUserID = l_UserId
    
    m_CustReg.ModuleID = wis_Users
    
'    gDbTrans.BeginTrans
'    Call m_CustReg.SaveCustomer
    With cmbNames
        .AddItem m_CustReg.FullName
        .ItemData(.newIndex) = l_UserId
        .ListIndex = .newIndex
    End With
    With cmbNamesEnglish
        .AddItem m_CustReg.FullNameEnglish
        .ItemData(.newIndex) = l_UserId
        .ListIndex = .newIndex
    End With
'    gDbTrans.SQLStmt = "Insert into UserTab " & _
            " (UserID,CustomerID,LoginName," & _
            "LoginPassword,Permissions) " & _
            " values (" & l_UserID & "," & _
            m_CustReg.CustomerId & "," & _
            AddQuotes(LoginName, True) & "," & _
            AddQuotes(Password, True) & "," & _
            "0)"
                        
'    If Not gDbTrans.SQLExecute Then
'        gDbTrans.RollBack
'        'MsgBox "Unable to register new user !", vbExclamation, gAppName & " - Error"
'        MsgBox GetResourceString(617), vbExclamation, gAppName & " - Error"
'        Exit Sub
'    End If
'    gDbTrans.CommitTrans

End Sub

Private Sub cmdRemove_Click()

If m_Permissions = perPigmyAgent Then Exit Sub
If m_Permissions And perBankAdmin = 0 Then Exit Sub

Dim isThereAnyAdmin As Boolean
Dim tmpPerms As wis_Permissions
Dim rst As ADODB.Recordset

Dim l_UserId As Long
With cmbNames
    l_UserId = .ItemData(.ListIndex)
End With

If l_UserId = gUserID Then
    If MsgBox("You are removing your account " & vbCrLf & _
        "Do you want to continue?", vbInformation + vbDefaultButton2, _
        wis_MESSAGE_TITLE) = vbNo Then Exit Sub

End If

'Before assignng the Permission Check at least one Admin Has to Be there
isThereAnyAdmin = False
gDbTrans.SqlStmt = "Select * from UserTab Where " & _
        " UserID <> " & l_UserId & " AND Deleted = " & False & _
        " And Permissions <> " & perOnlyWaves
        
tmpPerms = perBankAdmin
If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
    Do While rst.EOF = False
        tmpPerms = FormatField(rst("Permissions"))
        If tmpPerms And perBankAdmin Then
            If tmpPerms <> perPigmyAgent Then
                isThereAnyAdmin = True
                Exit Do
            End If
        End If
        rst.MoveNext
    Loop
End If

' Check for the existance of admin in the data base
If Not isThereAnyAdmin Then
    MsgBox "There is no other employee as administrator " & vbCrLf & _
            "So you can not do this operation"
    Exit Sub 'GoTo ExitLine
End If

gDbTrans.BeginTrans

gDbTrans.SqlStmt = "UPDATE UserTab SET Deleted = " & True & _
    " WHERE UserID = " & l_UserId
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    MsgBox "Unable to remove the user", vbInformation, wis_MESSAGE_TITLE
    Exit Sub
End If
gDbTrans.CommitTrans

Call LoadUsers

End Sub

Private Sub cmdSelectAll_Click()
Dim I As Integer
For I = 0 To lstPermissions.ListCount - 1
    lstPermissions.Selected(I) = True
Next I
cmdApply.Enabled = True
End Sub

Private Sub cmdStDate_Click()
With Calendar
    .selDate = gStrDate
    .Left = Me.Left + fraTrans.Left + cmdStDate.Left
    .Top = Me.Top + fraTrans.Top + cmdStDate.Top + 300
    .selDate = txtStartDate.Text
    .Show vbModal, Me
    If .selDate <> "" Then txtStartDate.Text = Calendar.selDate
End With

End Sub

Private Sub cmdUnselectAll_Click()
Dim I As Integer
For I = 0 To lstPermissions.ListCount - 1
    lstPermissions.Selected(I) = False
Next I
cmdApply.Enabled = True
End Sub


Private Sub Form_Load()
Dim blnVal As Boolean
Dim BckCol As Long
Dim I As Integer
Dim rst As ADODB.Recordset

'Center the form
    Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    Caption = gAppName & " - Add Users"
     
'Set icon for the form caption
    Icon = LoadResPicture(161, vbResIcon)
    'set kannada fonts
    Call SetKannadaCaption
    img.Picture = LoadResPicture(141, vbResIcon)
    
'Fill the list box with permissions
    Dim Perms As wis_Permissions
    With lstPermissions
        Perms = perBankAdmin
        .AddItem "Administrator": .ItemData(.newIndex) = Perms
        
        Perms = perCreateLedger
        .AddItem "Create Ledger Account": .ItemData(.newIndex) = Perms
        
        Perms = perCreateAccount
        .AddItem "Create Customer Account": .ItemData(.newIndex) = Perms
        
        Perms = perModifyAccount
        .AddItem "Modify Customer Account": .ItemData(.newIndex) = Perms
        
        Perms = perClerk
        .AddItem "Clerk": .ItemData(.newIndex) = Perms
        
        Perms = perCashier
        .AddItem "Cashier": .ItemData(.newIndex) = Perms
        
        Perms = perPassingOfficer
        .AddItem "Passing Officer": .ItemData(.newIndex) = Perms
                
        Perms = perPigmyAgent
        .AddItem "Pigmy Agent": .ItemData(.newIndex) = Perms
        
        Perms = perReadOnly
        .AddItem "Read Only": .ItemData(.newIndex) = Perms

    End With

Call LoadUsers
TabStrip1.Tabs(1).Selected = True
txtStartDate.Text = gStrDate: txtEndDate.Text = gStrDate

ErrLine:
    
End Sub

Private Function LoadToUI(l_UserId As Long) As Boolean
Dim blnVal As Boolean
Dim BckCol As Long
Dim TotalPerms As Long
Dim Retval As Long
Dim I As Integer
Dim rst As ADODB.Recordset
    
'Query the data base for the user details
    gDbTrans.SqlStmt = "Select * from UserTab where UserID = " & l_UserId
    Retval = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If Retval <= 0 Or Retval > 1 Then
        'MsgBox "Unable to locate the user !", vbExclamation, gAppName & " - Error"
        MsgBox GetResourceString(610), vbExclamation, gAppName & " - Error"
        GoTo ErrLine
    End If
    
''Set back colour and Enabled variables accordingly based on results obtained
    txtLoginName.Text = rst("LoginName")
    txtPassword.Text = rst("LoginPassword")
    txtConfirm.Text = rst("LoginPassword")
    TotalPerms = rst("Permissions")

    
'Set the right permissions
    'Done deliberately...
'    optEmployee.value = True
'    optPigmy.value = True
    
    Dim Perm As wis_Permissions
    Perm = perPigmyAgent
    
    For I = 0 To lstPermissions.ListCount - 1
        'lstPermissions.Selected(i) = False
'        If TotalPerms And lstPermissions.ItemData(i) Then
'            lstPermissions.Selected(i) = True
'        End If
        lstPermissions.Selected(I) = TotalPerms And lstPermissions.ItemData(I)
    Next I
    If TotalPerms = perBankAdmin Then lstPermissions.Selected(0) = True
    
LoadToUI = True
Exit Function
ErrLine:

End Function

Private Sub Form_Resize()

fraTrans.Height = fraUser.Height
fraTrans.Width = fraUser.Width
fraTrans.Top = fraUser.Top
fraTrans.Left = fraUser.Left

'cmdSelectAll.Top = lstPermissions.Top + lstPermissions.Height - 75 - cmdSelectAll.Height
'cmdUnselectAll.Top = cmdSelectAll.Top - cmdUnselectAll.Height - 175

'Frame1.Height = cmdSelectAll.Top + cmdSelectAll.Height + 100
'Frame1.Height = lstPermissions.Top + lstPermissions.Height + 75
'Me.cmdApply.Top = Frame1.Top + Frame1.Height + 50
'Me.cmdClose.Top = cmdApply.Top
'Me.cmdRemove.Top = cmdApply.Top
'Me.Height = Me.cmdRemove.Top + cmdApply.Height + 500

End Sub

Private Sub Form_Unload(cancel As Integer)
If gCurrUser.IsAdmin Then
    Dim ColCount As Integer
    For ColCount = 0 To grd.Cols - 1
        Call SaveSetting(App.EXEName, "UserTrans", _
                "ColWidth" & ColCount, grd.ColWidth(ColCount) / grd.Width)
    Next ColCount

End If
End Sub

Private Sub lstPermissions_Click()
'If the First Check is there then check all other'
'Because first is Full Permission
'While changing the Lsti values It may go into the Ifinite Loop
Static InLoop As Boolean

cmdApply.Enabled = IIf(lstPermissions.Tag = "True", True, False)

If InLoop Then Exit Sub
InLoop = True

Dim count As Integer
Dim Perms As wis_Permissions
Perms = perBankAdmin
If (m_Permissions And Perms) = 0 Then Call LoadUserPermission

'If checked the first then check all other options
If lstPermissions.Selected(0) Then
'lstPermissions.SelCount
End If
InLoop = False
Exit Sub

If lstPermissions.ListIndex = 0 Then
    If lstPermissions.Selected(0) = True Then
        For count = 1 To lstPermissions.ListCount - 1
            lstPermissions.Selected(count) = False
        Next
    End If
Else
    If lstPermissions.Selected(0) Then
        lstPermissions.Selected(lstPermissions.ListIndex) = False
        Exit Sub
    End If
End If

cmdApply.Enabled = True

InLoop = False
End Sub


Private Sub optEmployee_Click()
lblPermissions.Enabled = True
lstPermissions.Enabled = True
cmdSelectAll.Enabled = True
cmdUnselectAll.Enabled = True
lstPermissions.BackColor = vbWhite
cmdApply.Enabled = True
End Sub

Private Sub optPigmy_Click()
lblPermissions.Enabled = False
lstPermissions.Enabled = False
cmdSelectAll.Enabled = False
cmdUnselectAll.Enabled = False
lstPermissions.BackColor = wisGray
cmdApply.Enabled = True
End Sub


Private Sub lstPermissions_ItemCheck(Item As Integer)

Dim count As Integer

With lstPermissions
    If .Selected(Item) Then
        For count = Item + 1 To .ListCount - 2
            .Selected(count) = True
        Next
    End If
    If Item = 7 Then Exit Sub
    If Item And .Selected(Item) = False Then .Selected(0) = False
End With



End Sub

Private Sub TabStrip1_Click()
   ' If Not gCurrUser.IsAdmin Then Exit Sub
    If TabStrip1.SelectedItem.Index = 1 Then
        fraUser.Visible = True
        fraUser.ZOrder 0
        fraTrans.Visible = False
    Else
        fraUser.Visible = False
        fraTrans.ZOrder 0
        fraTrans.Visible = True
    End If
End Sub

Private Sub txtConfirm_Change()
If Trim$(txtLoginName.Text) = "" Then
    cmdApply.Enabled = False
Else
    cmdApply.Enabled = True
End If

End Sub

Private Sub txtLoginName_Change()

If Trim$(txtLoginName.Text) = "" Then
    cmdApply.Enabled = False
Else
    cmdApply.Enabled = True
End If


End Sub


Private Sub txtPassword_Change()
If Trim$(txtLoginName.Text) = "" Then
    cmdApply.Enabled = False
Else
    cmdApply.Enabled = True
End If


End Sub


