VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAddGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum wis_GroupType
    grpCustomer = 1
    grpCaste = 2
    grpPlace = 3
    grpAccount = 4
    grpDeposit = 5
    grpProduct = 6
    grpUnit = 7
    grpLoanPurpose = 8
    grpFarmer = 9
    grpMember = 10
    grpAllDeposit = 11
End Enum

Private Enum RecordStatus
    recNothing = 0
    recCancel = 1
    recAdd = 2
    recDelete = 3
    recUpdate = 4
    recInitialize = 5
End Enum

Private m_Status As RecordStatus
Private m_GroupType As wis_GroupType
Private M_ModuleID As wisModules
Private m_AllowDuplicate As Boolean

'Here we are declaring the Varible
'for TableName, Field NAme & thier value
Private m_TableName As String
Private m_Fields() As String
Private m_FieldValues() As String
Private m_strField As String
Private m_strFieldVal As String


Private m_PrevName As String
Private m_PrevNameEnglish As String
Private m_GroupName As String
Private m_GroupNameEnglish As String
Private m_GroupID As Long

Public Event ItemAdded(strAddItem As String, NewID As Long)
Public Event ItemDeleting(strDelItem As String, cancel As Integer)
Public Event ItemDeleted(strDelItem As String)

Private WithEvents m_frmGroup As frmAddGroup
Attribute m_frmGroup.VB_VarHelpID = -1

Public Property Get ObjectName()
    ObjectName = "AddGroup"
End Property

Private Sub AddCaste()
    Dim rst As Recordset
    gDbTrans.SqlStmt = "SELECT CasteID, Caste,CasteEnglish FROM CasteTab"
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)

Dim strMsg As String
Dim CasteId As Integer

If m_Status = recAdd Then
    CasteId = 0
    gDbTrans.SqlStmt = "Select Max(CasteID) From CasteTab"
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
        CasteId = FormatField(rst(0))
    CasteId = CasteId + 1
    
    gDbTrans.SqlStmt = "Insert Into CasteTab" & _
            "(CasteID,Caste,CasteEnglish ) " & _
            " Values " & _
            "( " & CasteId & ", " & AddQuotes(m_GroupName, True) & ", " & AddQuotes(m_GroupNameEnglish, True) & ")"
            
    m_GroupID = CasteId
    'MsgBox "Saved the details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(528)

ElseIf m_Status = recDelete Then

'Check in the Name Tab whether this Cast s using or Not
gDbTrans.SqlStmt = "Select Count(*) from NameTab where Caste = " & AddQuotes(m_GroupName, True)
If gDbTrans.Fetch(rst, adOpenDynamic) Then
    If FormatField(rst(0)) > 0 Then
        MsgBox GetResourceString(532), vbOKOnly, wis_MESSAGE_TITLE
        Exit Sub
    End If
End If
gDbTrans.SqlStmt = "Select Count(*) from ShgMaster where Caste = " & AddQuotes(m_GroupName, True)
If gDbTrans.Fetch(rst, adOpenDynamic) Then
    If FormatField(rst(0)) > 0 Then
        MsgBox GetResourceString(532), vbOKOnly, wis_MESSAGE_TITLE
        Exit Sub
    End If
End If

    Dim cancel As Integer
    RaiseEvent ItemDeleting(m_GroupName, cancel)
    If cancel Then Exit Sub
    gDbTrans.SqlStmt = "Delete * from CasteTab " & _
        "where Caste = " & AddQuotes(m_GroupName, True)
    'MsgBox "Group Deleted", vbInformation
    strMsg = GetResourceString(677)

ElseIf m_Status = recUpdate Then
    gDbTrans.BeginTrans
    gDbTrans.SqlStmt = "UPDATE NameTab " & _
                   " SET Caste = " & AddQuotes(m_GroupName, True) & _
                   " ,CasteEnglish = " & AddQuotes(m_GroupNameEnglish, True) & _
                   " WHERE Caste = " & AddQuotes(m_PrevName)
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBack: gDbTrans.BeginTrans
    
    gDbTrans.SqlStmt = "UPDATE ShgMaster " & _
                   " SET Caste = " & AddQuotes(m_GroupName, True) & _
                   " WHERE Caste = " & AddQuotes(m_PrevName)
    
    If gDbTrans.SQLExecute Then
        gDbTrans.CommitTrans
    Else
        gDbTrans.RollBack
    End If
    
    gDbTrans.SqlStmt = " UPDATE CasteTab " & _
                   " SET Caste = " & AddQuotes(m_GroupName, True) & _
                   " ,CasteEnglish = " & AddQuotes(m_GroupNameEnglish, True) & _
                   " WHERE CasteID = " & m_GroupID

    'MsgBox "Updated the Details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(707)

End If

gDbTrans.BeginTrans
If gDbTrans.SQLExecute Then
    gDbTrans.CommitTrans
Else
    gDbTrans.RollBack
End If


If Len(strMsg) Then MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE
If m_Status = recUpdate Then RaiseEvent ItemDeleted(m_GroupName)
If m_Status = recAdd Or m_Status = recUpdate Then RaiseEvent ItemAdded(m_GroupName, m_GroupID)
    
Exit_Line:
    
End Sub

Private Sub AddCustomerType()
    
    Dim rst As Recordset

Restart:
    gDbTrans.SqlStmt = "SELECT CustType,CustTypeName, CustTypeNameEnglish  From CustomerType"
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
    
    With m_frmGroup
        .txtItem = ""
        .txtItemEnglish = ""
        .Caption = "Add New Customer Type"
        .lblItem.Caption = GetResourceString(205, 253) '"Customer Type
        .lblItemEnglish.Caption = GetResourceString(205, 253, 468) '"Customer Type
        .lblList.Caption = GetResourceString(205, 253) 'Customer Type List
        .chkCumulative.Visible = True
        .chkCumulative.Caption = GetResourceString(252) & " UI" 'Cumulative
        .chkCumulative.Value = vbUnchecked
        
        If Not rst Is Nothing Then
          With .cmbList
            .Clear
            .AddItem ""
            .ItemData(.newIndex) = 0
            While Not rst.EOF
                .AddItem FormatField(rst("CustTypeName"))
                .ItemData(.newIndex) = FormatField(rst("CustType"))
                rst.MoveNext
            Wend
          End With
        End If
    End With
    
    Set rst = Nothing

Dim UIType As Byte
Dim strMsg As String

With m_frmGroup
    .Show 1
    UIType = IIf(.chkCumulative.Value = vbChecked, 0, 1)
End With
If m_Status = recCancel Then
    Unload m_frmGroup
    Exit Sub
End If

If m_Status = recAdd Then
    Dim NewID As Long
    gDbTrans.SqlStmt = "SELECT MAx(CustType) From CustomerType"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then NewID = FormatField(rst(0))
    Set rst = Nothing
    NewID = NewID + 1
    gDbTrans.SqlStmt = "Insert Into CustomerType " & _
            " (CustType, CustTypeName,CustTypeNameEnglish,UIType ) " & _
            " Values (" & NewID & ", " & _
            AddQuotes(m_GroupName, True) & ", " & _
            AddQuotes(m_GroupNameEnglish, True) & ", " & _
            UIType & ")"
            
    m_GroupID = NewID
     'MsgBox "Saved the details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(528)

ElseIf m_Status = recDelete Then
    
    gDbTrans.SqlStmt = " SELECT count(*) FROM Nametab " & _
                   " WHERE CustType = " & m_GroupID
                   
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        If FormatField(rst(0)) > 0 Then
            'MsgBox "Group can not be deleted while product from this group exists", vbInformation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(772)
            Exit Sub
        End If
    End If
    
    Dim cancel As Integer
    RaiseEvent ItemDeleting(m_GroupName, cancel)
    If cancel Then Exit Sub
    
    gDbTrans.SqlStmt = "Delete * from CustomerType " & _
        "where CustTypeName= " & AddQuotes(m_GroupName, True)
        
        
    'MsgBox "Group Deleted", vbInformation
    strMsg = GetResourceString(677)

ElseIf m_Status = recUpdate Then

    gDbTrans.SqlStmt = " UPDATE CustomerType " & _
                   " SET CustTypeName = " & AddQuotes(m_GroupName, True) & _
                   " ,CustTypeNameEnglish = " & AddQuotes(m_GroupNameEnglish, True) & _
                   " ,UIType = " & UIType & " WHERE CustType  = " & m_GroupID

    'MsgBox "Updated the Details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(707)

End If

gDbTrans.BeginTrans
If gDbTrans.SQLExecute Then
    gDbTrans.CommitTrans
Else
    gDbTrans.RollBack
End If

If Len(strMsg) Then MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE


If m_Status = recUpdate Then RaiseEvent ItemDeleted(m_GroupName)
If m_Status = recAdd Or m_Status = recUpdate Then RaiseEvent ItemAdded(m_GroupName, m_GroupID)

GoTo Restart

End Sub

Private Sub AddFarmerType()
    
    Dim rst As Recordset

Restart:
    gDbTrans.SqlStmt = "SELECT FarmerTypeId,TypeName, TypeEnglish From FarmerTypeTab"
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
    
    With m_frmGroup
        .txtItem = ""
        .txtItemEnglish = ""
        .Caption = "Add New Type"
        .lblItem.Caption = GetConfigValue("FarmerType", GetResourceString(378)) & " " & GetResourceString(253) '"Customer Type
        .lblItemEnglish.Caption = GetConfigValue("FarmerType", GetResourceString(378)) & " " & GetResourceString(253, 468) '"Customer Type
        .lblList.Caption = GetConfigValue("FarmerType", GetResourceString(378)) & " " & GetResourceString(253) 'Customer Type List
        .chkCumulative.Visible = False
        '.chkCumulative.Caption = GetResourceString(252) & " UI" 'Cumulative
        '.chkCumulative.Value = vbUnchecked
        
        If Not rst Is Nothing Then
          With .cmbList
            .Clear
            .AddItem ""
            .ItemData(.newIndex) = 0
            While Not rst.EOF
                .AddItem FormatField(rst("TypeName"))
                .ItemData(.newIndex) = FormatField(rst("FarmerTypeID"))
                rst.MoveNext
            Wend
          End With
        End If
    End With
    
    Set rst = Nothing

Dim strMsg As String

With m_frmGroup
    .Show 1
    'UIType = IIf(.chkCumulative.Value = vbChecked, 0, 1)
End With
If m_Status = recCancel Then
    Unload m_frmGroup
    Exit Sub
End If

If m_Status = recAdd Then
    Dim NewID As Long
    gDbTrans.SqlStmt = "SELECT Max(FarmerTypeId) From FarmerTypeTab"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then NewID = FormatField(rst(0))
    Set rst = Nothing
    NewID = NewID + 1
    gDbTrans.SqlStmt = "Insert Into FarmerTypeTab " & _
            " (FarmerTypeId, TypeName,TypeEnglish ) " & _
            " Values (" & NewID & ", " & _
            AddQuotes(m_GroupName, True) & "," & AddQuotes(m_GroupNameEnglish, True) & ")"
            
    m_GroupID = NewID
     'MsgBox "Saved the details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(528)

ElseIf m_Status = recDelete Then
    
    gDbTrans.SqlStmt = " SELECT count(*) FROM Nametab " & _
                   " WHERE FarmerTypeId = " & m_GroupID
                   
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        If FormatField(rst(0)) > 0 Then
            'MsgBox "Group can not be deleted while product from this group exists", vbInformation, wis_MESSAGE_TITLE
            MsgBox GetResourceString(772)
            Exit Sub
        End If
    End If
    
    Dim cancel As Integer
    RaiseEvent ItemDeleting(m_GroupName, cancel)
    If cancel Then Exit Sub
    
    gDbTrans.SqlStmt = "Delete * from FarmerTypeTab " & _
        "where TypeName= " & AddQuotes(m_GroupName, True)
        
    'MsgBox "Group Deleted", vbInformation
    strMsg = GetResourceString(677)

ElseIf m_Status = recUpdate Then

    gDbTrans.SqlStmt = " UPDATE FarmerTypeTab " & _
                   " SET TypeName = " & AddQuotes(m_GroupName, True) & _
                   " WHERE FarmerTypeId  = " & m_GroupID

    'MsgBox "Updated the Details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(707)

End If

gDbTrans.BeginTrans
If gDbTrans.SQLExecute Then
    gDbTrans.CommitTrans
Else
    gDbTrans.RollBack
End If

If Len(strMsg) Then MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE


If m_Status = recUpdate Then RaiseEvent ItemDeleted(m_GroupName)
If m_Status = recAdd Or m_Status = recUpdate Then RaiseEvent ItemAdded(m_GroupName, m_GroupID)

GoTo Restart

End Sub

Private Sub AddLoanPurpose()
    
    Dim rst As Recordset
    Dim strPrevName As String
    
    gDbTrans.SqlStmt = "SELECT PurposeId,Purpose,PurposeEnglish From LoanPurpose"
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
    
    With m_frmGroup
        .Caption = "Add new Loan purpose"
        .lblItem.Caption = GetResourceString(80, 221) '"Loan purpose
        .lblItemEnglish.Caption = GetResourceString(80, 221, 468) '"Loan purpose
        .lblList.Caption = GetResourceString(80, 221) '"Loan purpose
        .chkCumulative.Visible = False
        
        gDbTrans.SqlStmt = "SELECT * From LoanScheme"
        If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then .cmbCumulative.Visible = True
        With .cmbCumulative
            .AddItem GetResourceString(338, 18) '"All Loans
            .ItemData(.newIndex) = 0
            .ListIndex = 0
            If Not rst Is Nothing Then
                While Not rst.EOF
                    .AddItem FormatField(rst("SchemeName"))
                    .ItemData(.newIndex) = FormatField(rst("SchemeID"))
                    rst.MoveNext
                Wend
            End If
        End With
        
        With .cmbList
            .Clear
            .AddItem ""
            If Not rst Is Nothing Then
                .ItemData(.newIndex) = 0
                While Not rst.EOF
                    .AddItem FormatField(rst("Purpose"))
                    .ItemData(.newIndex) = FormatField(rst("PurposeId"))
                    rst.MoveNext
                Wend
            End If
        End With
    End With
    
    Set rst = Nothing

Dim Cumulative As Integer

With m_frmGroup
    .Show 1
    
'    Cumulative = IIf(.chkCumulative.Value = vbChecked, Inst_Yearly, 0)
    If .cmbCumulative.ListIndex >= 0 Then _
            Cumulative = .cmbCumulative.ItemData(.cmbCumulative.ListIndex)
    
End With

If m_Status = recCancel Then
    Unload m_frmGroup
    Exit Sub
End If

Dim strMsg As String
gDbTrans.BeginTrans

If m_Status = recAdd Then
    Dim NewID As Long
    gDbTrans.SqlStmt = "SELECT Max(DepositID) From DepositName"
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then NewID = FormatField(rst(0))
    If NewID < wisDeposit_FD Then NewID = wisDeposit_FD
    
    Set rst = Nothing
    NewID = NewID + 1
    gDbTrans.SqlStmt = "Insert Into DepositName (" & _
            " DepositID, DepositName,DepositNameEnglish,Cumulative,UserID) " & _
            " Values (" & NewID & ", " & _
            AddQuotes(m_GroupName, True) & ", " & _
            AddQuotes(m_GroupNameEnglish, True) & ", " & _
            Cumulative & "," & gUserID & " )"
    strMsg = GetResourceString(528)
    m_GroupID = NewID
    
ElseIf m_Status = recDelete Then
    Dim cancel As Integer
    RaiseEvent ItemDeleting(m_GroupName, cancel)
    If cancel Then Exit Sub
    
    gDbTrans.SqlStmt = "Delete * from DepositName " & _
            "where DepositName = " & AddQuotes(m_GroupName, True)
    'MsgBox "Group Deleted", vbInformation
    strMsg = GetResourceString(677)

ElseIf m_Status = recUpdate Then
    
    gDbTrans.SqlStmt = "UPDATE DepositName " & _
                " SET DepositName = " & AddQuotes(m_GroupName, True) & _
                " ,DepositNameEnglish = " & AddQuotes(m_GroupNameEnglish, True) & _
                " WHERE DepositID = " & m_GroupID

    'Now make the same changes to the DepositLedger head,Interest ledger head
    'Deposit Loan Ledger Head
    'change the name of the ledger heads
    If Not UpdateLedgerHeads Then GoTo Hell
    
    'MsgBox "Updated the Details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(707)
End If

If gDbTrans.SQLExecute Then
    gDbTrans.CommitTrans
Else
    gDbTrans.RollBack
End If

If Len(strMsg) Then MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE

If m_Status = recUpdate Or m_Status = recDelete Then RaiseEvent ItemDeleted(m_PrevName)
If m_Status = recAdd Or m_Status = recUpdate Then _
        RaiseEvent ItemAdded(m_GroupName, m_GroupID)

Hell:
    Call gDbTrans.RollBack
    
End Sub


Private Sub AddDeposit()
    
    Dim rst As Recordset
    Dim strPrevName As String
    Dim Cumulative As Byte

With m_frmGroup
    Cumulative = IIf(.chkCumulative.Value = vbChecked, Inst_Yearly, 0)
    If .cmbCumulative.ListIndex >= 0 Then _
            Cumulative = .cmbCumulative.ItemData(.cmbCumulative.ListIndex)
End With

Dim strMsg As String
gDbTrans.BeginTrans

If m_Status = recAdd Then
    Dim NewID As Long
    gDbTrans.SqlStmt = "SELECT Max(DepositID) From DepositName"
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then NewID = FormatField(rst(0))
    
    If NewID < wisDeposit_FD Then NewID = wisDeposit_FD
    
    Set rst = Nothing
    NewID = NewID + 1
    gDbTrans.SqlStmt = "Insert Into DepositName (" & _
            " DepositID, DepositName,DepositNameEnglish,Cumulative) " & _
            " Values (" & NewID & ", " & _
            AddQuotes(m_GroupName, True) & ", " & _
            AddQuotes(m_GroupNameEnglish, True) & ", " & _
            Cumulative & ")"
    strMsg = GetResourceString(528)
    m_GroupID = NewID
    
ElseIf m_Status = recDelete Then
    
    gDbTrans.SqlStmt = "Select Count(*) from FDMaster where DepositID = " & m_GroupID
    If gDbTrans.Fetch(rst, adOpenDynamic) Then
        If FormatField(rst(0)) > 0 Then
            MsgBox GetResourceString(532), vbOKOnly, wis_MESSAGE_TITLE
            Exit Sub
        End If
    End If
    
    Dim cancel As Integer
    RaiseEvent ItemDeleting(m_GroupName, cancel)
    If cancel Then Exit Sub
    
    gDbTrans.SqlStmt = "Delete * from DepositName " & _
            "where DepositName = " & AddQuotes(m_GroupName, True)
        
    'MsgBox "Group Deleted", vbInformation
    strMsg = GetResourceString(677)

ElseIf m_Status = recUpdate Then
    
    gDbTrans.SqlStmt = "Select * from DepositName where DepositID = " & m_GroupID
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        m_PrevNameEnglish = FormatField(rst("DepositNameEnglish"))
    Else
        MsgBox GetResourceString(532), vbOKOnly, wis_MESSAGE_TITLE
        Exit Sub
    End If
    
    'Now make the same changes to the DepositLedger head,Interest ledger head
    'Deposit Loan Ledger Head
    'change the name of the ledger heads
    If Not UpdateLedgerHeads Then GoTo Hell
    
    gDbTrans.SqlStmt = "UPDATE DepositName " & _
                " SET DepositName = " & AddQuotes(m_GroupName, True) & _
                " ,DepositNameEnglish = " & AddQuotes(m_GroupNameEnglish, True) & _
                " WHERE DepositID = " & m_GroupID

    'MsgBox "Updated the Details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(707)
End If

If gDbTrans.SQLExecute Then
    gDbTrans.CommitTrans
Else
    gDbTrans.RollBack
End If

If Len(strMsg) Then MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE

If m_Status = recUpdate Or m_Status = recDelete Then RaiseEvent ItemDeleted(m_PrevName)
If m_Status = recAdd Or m_Status = recUpdate Then RaiseEvent ItemAdded(m_GroupName, m_GroupID)

Hell:
    Call gDbTrans.RollBack
    
End Sub


Private Sub AddMemberType()
    
    Dim rst As Recordset
    Dim strPrevName As String
    Dim Cumulative As Byte

Dim strMsg As String
gDbTrans.BeginTrans

If m_Status = recAdd Then
    Dim NewID As Integer
    gDbTrans.SqlStmt = "SELECT Max(MemberType) From MemberTypeTab"
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then NewID = FormatField(rst(0))
    
    Set rst = Nothing
    NewID = NewID + 1
    gDbTrans.SqlStmt = "Insert Into MemberTypeTab (" & _
            " MemberType, MemberTypeName,MemberTypeNameEnglish) " & _
            " Values (" & NewID & ", " & _
            AddQuotes(m_GroupName, True) & ", " & _
            AddQuotes(m_GroupNameEnglish, True) & ")"
    strMsg = GetResourceString(528)
    m_GroupID = NewID
    
ElseIf m_Status = recDelete Then
    
    gDbTrans.SqlStmt = "Select Count(*) from MemMaster where MemberType = " & m_GroupID
    If gDbTrans.Fetch(rst, adOpenDynamic) Then
        If FormatField(rst(0)) > 0 Then
            MsgBox GetResourceString(532), vbOKOnly, wis_MESSAGE_TITLE
            Exit Sub
        End If
    End If
    
    Dim cancel As Integer
    RaiseEvent ItemDeleting(m_GroupName, cancel)
    If cancel Then Exit Sub
    
    gDbTrans.SqlStmt = "Delete * from MemberTypeTab " & _
            " where MemberType = " & m_GroupID
    strMsg = GetResourceString(677)

ElseIf m_Status = recUpdate Then
    
    gDbTrans.SqlStmt = "Select * from MemberTypeTab where MemberType = " & m_GroupID
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        m_PrevNameEnglish = FormatField(rst("MemberTypeNameEnglish"))
    Else
        MsgBox GetResourceString(532), vbOKOnly, wis_MESSAGE_TITLE
        Exit Sub
    End If
    
    'Now make the same changes to the DepositLedger head,Interest ledger head
    'Deposit Loan Ledger Head
    'change the name of the ledger heads
    If Not UpdateLedgerHeads Then GoTo Hell
    
    gDbTrans.SqlStmt = "UPDATE MemberTypeTab " & _
                " SET MemberTypeName = " & AddQuotes(m_GroupName, True) & _
                " ,MemberTypeNameEnglish = " & AddQuotes(m_GroupNameEnglish, True) & _
                " WHERE MemberType = " & m_GroupID

    'MsgBox "Updated the Details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(707)
End If

If gDbTrans.SQLExecute Then
    gDbTrans.CommitTrans
Else
    gDbTrans.RollBack
End If

If Len(strMsg) Then MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE

If m_Status = recUpdate Or m_Status = recDelete Then RaiseEvent ItemDeleted(m_PrevName)
If m_Status = recAdd Or m_Status = recUpdate Then _
        RaiseEvent ItemAdded(m_GroupName, m_GroupID)

Hell:
    Call gDbTrans.RollBack
    
End Sub


Private Function AddDepositAccountType() As Boolean
    
    Dim rst As Recordset
    Dim strPrevName As String
    Dim Cumulative As Byte
    Dim strMsg As String
    Dim resStringId As Integer
    Dim Deptype As Integer
    Dim depTableName As String
    Dim depTransTableName As String
    Dim recordCount As Integer
    
    Call GetTableNameForModule(M_ModuleID, depTableName, depTransTableName, , , resStringId)

    
If m_Status = recNothing Then 'recInitialize
    gDbTrans.SqlStmt = "Select * From DepositTypeTab where ModuleID = " & M_ModuleID
    recordCount = gDbTrans.Fetch(rst, adOpenDynamic)
    If recordCount = 0 Then
        gDbTrans.BeginTrans
    
        'There are no Deposit Types Before, Please check is there any Deposits are there
        gDbTrans.SqlStmt = "Select CustomerID From " & depTableName
        recordCount = gDbTrans.Fetch(rst, adOpenDynamic)
        'If FormatField(rst("DepositType")) = 0 Then
        If recordCount > 0 Then
            If MsgBox("Existing Deposit types will be converted as module" & vbCrLf & _
                "Do you want to continue?", vbYesNo + vbQuestion, wis_MESSAGE_TITLE) = vbNo Then
                     AddDepositAccountType = False
                     Exit Function
            End If
            If MsgBox("This will create new Ledger heads for each Deposit type" & vbCrLf & _
                "Are you sure, you want to continue?", vbYesNo + vbQuestion + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then _
                    AddDepositAccountType = False: Exit Function
            ''Now Convert the Existing Accounts
            gDbTrans.BeginTrans
            ''Update Deposit Loans
            Deptype = 0
            Select Case M_ModuleID
                Case wis_SBAcc
                    Deptype = wisDeposit_SB
                Case wis_CAAcc
                    Deptype = wisDeposit_CA
                Case wis_RDAcc
                    Deptype = wisDeposit_RD
                Case wis_PDAcc
                    Deptype = wisDeposit_PD
                Case wis_Deposits
                    Deptype = wisDeposit_FD
            End Select
            gDbTrans.SqlStmt = "Update DepositLoanMaster Set DepositType = DepositType + " & 1 & " where DepositType = " & Deptype
            Call gDbTrans.SQLExecute
            gDbTrans.SqlStmt = "Update PledgeDeposit Set DepositType = DepositType + " & 1 & " where DepositType = " & Deptype
            Call gDbTrans.SQLExecute
            '
            Deptype = Deptype + 1
            ''Update the AccType in bankHeadID's
            gDbTrans.SqlStmt = "Update BankHeadIDs Set AccType = AccType + " & Deptype & " where AccType = " & M_ModuleID
            Call gDbTrans.SQLExecute
            'Create the Deposit Type
            gDbTrans.SqlStmt = "Update " & depTableName & " Set DepositType =  " & Deptype
            Call gDbTrans.SQLExecute
            
            gDbTrans.SqlStmt = "Insert into DepositTypeTab " & _
                " (DepositType, ModuleID, DepositTypeName,DepositTypeNameEnglish)" & _
                " Values ( " & Deptype & "," & M_ModuleID & "," & _
                AddQuotes(GetResourceString(102, resStringId)) & "," & _
                AddQuotes(LoadResourceStringS(102, resStringId)) & ")"
            Call gDbTrans.SQLExecute
            
            'Update the Head Name also
            m_PrevName = GetResourceString(resStringId)
            m_GroupName = GetResourceString(102, resStringId)
            m_PrevNameEnglish = ""
            m_GroupNameEnglish = LoadResourceStringS(102, resStringId)
            Call UpdateLedgerHeads(True)
            
            'Deposit Interest Paid
            m_PrevName = GetResourceString(resStringId, 487)
            m_GroupName = GetResourceString(102, resStringId, 487)
            m_GroupNameEnglish = LoadResourceStringS(102, resStringId, 487)
            Call UpdateLedgerHeads(True)
            
             'Payable Interest
            m_PrevName = GetResourceString(resStringId, 375, 47)
            m_GroupName = GetResourceString(102, resStringId, 375, 47)
            m_GroupNameEnglish = LoadResourceStringS(102, resStringId, 375, 47)
            Call UpdateLedgerHeads(True)
            
            'Interest Provision
            m_PrevName = GetResourceString(resStringId, 450)
            m_GroupName = GetResourceString(102, resStringId, 450)
            m_GroupNameEnglish = LoadResourceStringS(102, resStringId, 450)
            Call UpdateLedgerHeads(True)
            
            'Deposit Loan
            m_PrevName = GetResourceString(resStringId, 58)
            m_GroupName = GetResourceString(102, resStringId, 58)
            m_GroupNameEnglish = LoadResourceStringS(102, resStringId, 58)
            Call UpdateLedgerHeads(True)
            
            'Deposit Loan Interest
            m_PrevName = GetResourceString(resStringId, 58, 483)
            m_GroupName = GetResourceString(102, resStringId, 58, 483)
            m_GroupNameEnglish = LoadResourceStringS(102, resStringId, 58, 483)
            Call UpdateLedgerHeads(True)

            gDbTrans.CommitTrans
            
            AddDepositAccountType = True
            Exit Function
            
        End If
    End If
End If


    gDbTrans.BeginTrans
    
    If m_Status = recAdd Then
        Dim NewID As Integer
        gDbTrans.SqlStmt = "SELECT Max(DepositType) From DepositTypeTab where ModuleID = " & M_ModuleID
        
        If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then NewID = FormatField(rst(0))
        
        Set rst = Nothing
        NewID = NewID + 1
        gDbTrans.SqlStmt = "Insert Into DepositTypeTab (" & _
                " DepositType, ModuleID, DepositTypeName,DepositTypeNameEnglish) " & _
                " Values (" & NewID & ", " & M_ModuleID & "," & _
                AddQuotes(m_GroupName, True) & ", " & _
                AddQuotes(m_GroupNameEnglish, True) & ")"
        strMsg = GetResourceString(528)
        m_GroupID = NewID
        
    ElseIf m_Status = recDelete Then
        
        gDbTrans.SqlStmt = "Select Count(*) from " & depTableName & " where DepositType = " & m_GroupID
        If gDbTrans.Fetch(rst, adOpenDynamic) Then
            If FormatField(rst(0)) > 0 Then
                MsgBox GetResourceString(532), vbOKOnly, wis_MESSAGE_TITLE
                Exit Function
            End If
        End If
        
        Dim cancel As Integer
        RaiseEvent ItemDeleting(m_GroupName, cancel)
        If cancel Then Exit Function
        
        gDbTrans.SqlStmt = "Delete * from DepositTypeTab " & _
                " where DepositType = " & m_GroupID & " And ModuleId = " & M_ModuleID
        strMsg = GetResourceString(677)
    
    ElseIf m_Status = recUpdate Then
        
        gDbTrans.SqlStmt = "Select * from DepositTypeTab where DepositType = " & m_GroupID & " And ModuleID = " & M_ModuleID
        If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
            m_PrevNameEnglish = FormatField(rst("DepositTypeNameEnglish"))
        Else
            MsgBox GetResourceString(532), vbOKOnly, wis_MESSAGE_TITLE
            Exit Function
        End If
        
        'Now make the same changes to the DepositLedger head,Interest ledger head
        'Deposit Loan Ledger Head
        'change the name of the ledger heads
        If Not UpdateLedgerHeads Then GoTo Hell
        
        gDbTrans.SqlStmt = "UPDATE DepositTypeTab " & _
                    " SET DepositTypeName = " & AddQuotes(m_GroupName, True) & _
                    " ,MDepositTypeNameEnglish = " & AddQuotes(m_GroupNameEnglish, True) & _
                    " WHERE DepositType = " & m_GroupID & " And ModuleID = " & M_ModuleID
    
        'MsgBox "Updated the Details ", vbInformation, wis_MESSAGE_TITLE
        strMsg = GetResourceString(707)
    End If
    
    If Len(gDbTrans.SqlStmt) > 0 Then
        If gDbTrans.SQLExecute Then
            gDbTrans.CommitTrans
        Else
            gDbTrans.RollBack
        End If
    Else
        gDbTrans.RollBack
    End If
    
    If Len(strMsg) Then MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE
    
    If m_Status = recUpdate Or m_Status = recDelete Then RaiseEvent ItemDeleted(m_PrevName)
    If m_Status = recAdd Or m_Status = recUpdate Then _
            RaiseEvent ItemAdded(m_GroupName, m_GroupID)
    
AddDepositAccountType = True

Hell:
        Call gDbTrans.RollBack

End Function

Private Function AddNew()

Dim rst As ADODB.Recordset

Dim intCumulative As Integer

ReDim m_Fields(2)
ReDim m_FieldValues(2)

Select Case m_GroupType
    Case grpAccount
        m_TableName = "AccountGroup"
        m_Fields(0) = "AccGroupID"
        m_Fields(1) = "GroupName"
        With m_frmGroup
            .Caption = "Add/Remove Account Group"
            .lblItem.Caption = GetResourceString(36, 157) 'Group
            .lblItemEnglish.Caption = GetResourceString(36, 157, 468) 'Group
            .lblList.Caption = GetResourceString(36, 157) 'Group
            .chkCumulative.Visible = False
        End With

    Case grpCaste
        m_TableName = "CasteTab"
        m_Fields(0) = "CasteID"
        m_Fields(1) = "Caste"
        With m_frmGroup
            .Caption = "Add/Remove Caste"
            .lblItem.Caption = GetResourceString(111)
            .lblItemEnglish.Caption = GetResourceString(111, 468)
            .lblList.Caption = GetResourceString(111)
            .chkCumulative.Visible = False
        End With
    
    Case grpCustomer
        
        m_TableName = "CustomerType"
        m_Fields(0) = "CustType"
        m_Fields(1) = "CustTypeName"
        With m_frmGroup
            .Caption = "Add/Remove Customer type"
            .lblItem.Caption = GetResourceString(205, 253) '"Customer Type
            .lblItemEnglish.Caption = GetResourceString(205, 253, 468) '"Customer Type
            .lblList.Caption = GetResourceString(205, 253) 'Customer Type List
            .chkCumulative.Visible = True
            .chkCumulative.Caption = GetResourceString(252) & " UI" 'Cumulative
            .chkCumulative.Value = vbUnchecked
        End With
    
    Case grpFarmer
        
        m_TableName = "FarmerTypeTab"
        m_Fields(0) = "FarmerTypeId"
        m_Fields(1) = "TypeName"
        m_Fields(2) = "TypeEnglish"
        With m_frmGroup
            .Caption = "Add/Remove Farmer type"
            .lblItem.Caption = GetConfigValue("FarmerTypeName", GetResourceString(378)) & " " & GetResourceString(253)  '"Customer Type
            .lblItemEnglish.Caption = GetConfigValue("FarmerTypeName", GetResourceString(378)) & " " & GetResourceString(253, 468) '"Customer Type
            .lblList.Caption = GetConfigValue("FarmerTypeName", GetResourceString(378)) & " " & GetResourceString(253)  'Customer Type List
            .chkCumulative.Visible = False
            '.chkCumulative.Caption = GetResourceString(252) & " UI" 'Cumulative
            .chkCumulative.Value = vbUnchecked
        End With
    
    Case grpDeposit
        ReDim m_Fields(3)
        ReDim m_FieldValues(3)

        m_TableName = "DepositName"
        m_Fields(0) = "DepositID"
        m_Fields(1) = "DepositName"
        m_Fields(2) = "Cumulative"
        m_Fields(3) = "DepositNameEnglish"
        
        With m_frmGroup
            .Caption = "Add new deposit scheme"
            .lblItem.Caption = GetResourceString(43, 35) '"deposit Name
            .lblItemEnglish.Caption = GetResourceString(43, 35, 468) '"deposit Name
            .lblList.Caption = GetResourceString(45) 'Deposits
            .chkCumulative.Visible = True
            .chkCumulative.Caption = GetResourceString(320) 'Cumulative
            .cmbCumulative.Visible = True
            With .cmbCumulative
                .Clear
                .AddItem ""
                .ItemData(.newIndex) = 0
                .AddItem GetResourceString(463) 'Monthly
                .ItemData(.newIndex) = Inst_Monthly
                .AddItem GetResourceString(413) 'Bi Monthly
                .ItemData(.newIndex) = Inst_BiMonthly
                .AddItem GetResourceString(414) 'Quarterly
                .ItemData(.newIndex) = Inst_Quartery
                .AddItem "Half Yearly" 'GetResourceString(463) 'halfe
                .ItemData(.newIndex) = Inst_HalfYearly
                .AddItem GetResourceString(208) 'Yearly
                .ItemData(.newIndex) = Inst_Yearly
            End With
    
        End With
    
    Case grpLoanPurpose
        ReDim m_Fields(3)
        ReDim m_FieldValues(3)
    
        m_TableName = "LoanPurpose"
        m_Fields(0) = "PurposeID"
        m_Fields(1) = "Purpose"
        m_Fields(2) = "SchemeId"
        m_Fields(3) = "PurposeEnglish"
        With m_frmGroup
            .Caption = "Add new Loan purpose"
            .lblItem.Caption = GetResourceString(80, 221) '"LOan purpose
            .lblItemEnglish.Caption = GetResourceString(80, 221, 468) '"LOan purpose
            .lblList.Caption = GetResourceString(80, 221) '"LOan purpose
            .chkCumulative.Visible = False
            
            gDbTrans.SqlStmt = "SELECT * From LoanScheme"
            If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then .cmbCumulative.Visible = True
            With .cmbCumulative
                .Clear
                .AddItem GetResourceString(338, 18) '"All Loans
                .ItemData(.newIndex) = 0
                .ListIndex = 0
            End With
            If Not rst Is Nothing Then
              With .cmbCumulative
                While Not rst.EOF
                    .AddItem FormatField(rst("SchemeName"))
                    .ItemData(.newIndex) = FormatField(rst("SchemeID"))
                    rst.MoveNext
                Wend
              End With
            End If
        End With
        
    Case grpPlace
        m_TableName = "PlaceTab"
        m_Fields(0) = "PlaceID"
        m_Fields(1) = "Place"
        With m_frmGroup
            .Caption = "Add/Remove Place"
            .lblItem.Caption = GetResourceString(112)      'PLace
            .lblItemEnglish.Caption = GetResourceString(112, 468)     'PLace
            .lblList.Caption = GetResourceString(112) 'PLace
            .chkCumulative.Visible = False
        End With
    
    Case grpProduct
        m_TableName = "ProductGroup"
        m_Fields(0) = "GroupID"
        m_Fields(1) = "GroupName"
        With m_frmGroup
            .Caption = "Add/Remove Product group"
            .lblItem.Caption = GetResourceString(157, 35)
            .lblItemEnglish.Caption = GetResourceString(157, 35, 468)
            .lblList.Caption = GetResourceString(157, 35)
            .chkCumulative.Visible = False
        End With
    
    Case grpUnit
        m_TableName = "Units"
        m_Fields(0) = "UnitID"
        m_Fields(1) = "UnitName"
        With m_frmGroup
            .Caption = "Add/Remove Units"
            .lblItem.Caption = GetResourceString(161)
            .lblItemEnglish.Caption = GetResourceString(161, 468)
            .lblList.Caption = GetResourceString(161)
            .chkCumulative.Visible = False
        End With
    Case grpMember
        Dim rstMemType As ADODB.Recordset
        gDbTrans.SqlStmt = "Select distinct(MemberType) From MemMaster"
        intCumulative = gDbTrans.Fetch(rst, adOpenDynamic)
        If intCumulative > 1 Then
            gDbTrans.SqlStmt = "Select * From MemberTypeTab"
            intCumulative = gDbTrans.Fetch(rst, adOpenDynamic)
            If intCumulative <= 1 Then
                If FormatField(rst("MemberType")) = 0 Then
                    If MsgBox("Existing member types will be converted as module" & vbCrLf & _
                        "Do you want to continue?", vbYesNo + vbQuestion, wis_MESSAGE_TITLE) = vbNo Then
                            Exit Function
                    End If
                    If MsgBox("This will create new Ledger heads for each member type" & vbCrLf & _
                        "Are you sure, you want to continue?", vbYesNo + vbQuestion + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then _
                        Exit Function
                    '''Now COnvert the Existing mebere
                    gDbTrans.SqlStmt = "Select distinct(MemberType) From MemMaster"
                    intCumulative = gDbTrans.Fetch(rst, adOpenDynamic)
            
                    gDbTrans.BeginTrans
                    gDbTrans.SqlStmt = "Delete * from MemberTypeTab"
                    Call gDbTrans.SQLExecute
                    
                    While rst.EOF = False
                        intCumulative = FormatField(rst(0))
                        gDbTrans.SqlStmt = "Insert into MemberTypeTab " & _
                            " (MemberType, MemberTypeName,MemberTypeNameEnglish)" & _
                            " Values (" & intCumulative & "," & _
                            AddQuotes(GetResourceString(101 + intCumulative, 49)) & "," & _
                            AddQuotes(LoadResourceStringS(101 + intCumulative, 49)) & ")"
                        
                        Call gDbTrans.SQLExecute
                        'Move to Next member Type
                        rst.MoveNext
                    Wend
                    If rst.recordCount = 1 Then
                        rst.MoveFirst
                        intCumulative = FormatField(rst(0))
                        'Share Account
                        m_PrevName = GetResourceString(53, 36)
                        m_GroupName = GetResourceString(101 + intCumulative, 49, 53, 36)
                        m_PrevNameEnglish = "" 'LoadResourceStringS(53, 36)
                        m_GroupNameEnglish = LoadResourceStringS(101 + intCumulative, 49, 53, 36)
                        If Not UpdateLedgerHeads(True) Then
                            gDbTrans.RollBack
                            gDbTrans.BeginTrans
                        End If
                        'Membership  Fee
                        m_PrevName = GetResourceString(79, 191)
                        m_GroupName = GetResourceString(101 + intCumulative, 49, 191)
                        'm_PrevNameEnglish = LoadResourceStringS(79, 191)
                        m_GroupNameEnglish = LoadResourceStringS(101 + intCumulative, 49, 191)
                        If Not UpdateLedgerHeads(True) Then
                            gDbTrans.RollBack
                            gDbTrans.BeginTrans
                        End If
                        
                        'Share Fee
                        m_PrevName = GetResourceString(53, 191)
                        m_GroupName = GetResourceString(101 + intCumulative, 49, 53, 191)
                        'm_PrevNameEnglish = LoadResourceStringS(53, 191)
                        m_GroupNameEnglish = LoadResourceStringS(101 + intCumulative, 49, 53, 191)
                        If Not UpdateLedgerHeads(True) Then
                            gDbTrans.RollBack
                            gDbTrans.BeginTrans
                        End If
                        
                        'SHARE DIVEDEND
                        m_PrevName = GetResourceString(49, 200)
                        m_GroupName = GetResourceString(101 + intCumulative, 49, 200)
                        'm_PrevNameEnglish = LoadResourceStringS(49, 200)
                        m_GroupNameEnglish = LoadResourceStringS(101 + intCumulative, 49, 200)
                        If Not UpdateLedgerHeads Then
                            gDbTrans.RollBack
                            gDbTrans.BeginTrans
                        End If
                        
                        ''Update the AccTYpe in bankHeadID's
                        gDbTrans.SqlStmt = "Update BankHeadIDs Set AccType = AccType + 1  where AccType = " & wis_Members
                        Call gDbTrans.SQLExecute
                    
                    End If
                    
                    
                    gDbTrans.CommitTrans
                    
                    
                End If
            End If
        ElseIf intCumulative > 1 Then
        
        End If
        
        m_TableName = "MemberTypeTab"
        m_Fields(0) = "MemberType"
        m_Fields(1) = "MemberTypeName"
        With m_frmGroup
            .Caption = "Add/Remove Member Types"
            .lblItem.Caption = GetResourceString(101)
            .lblItemEnglish.Caption = GetResourceString(101, 468)
            .lblList.Caption = GetResourceString(101)
            .chkCumulative.Visible = False
        End With
        
    Case grpAllDeposit
        
        If Not AddDepositAccountType Then Exit Function
        Dim resStringId As Integer
        Call GetTableNameForModule(M_ModuleID, , , , , resStringId)
        
        With m_frmGroup
            .Caption = "Add/Remove member deposit Types"
            .lblItem.Caption = GetResourceString(resStringId, 253)
            .lblItemEnglish.Caption = GetResourceString(resStringId, 253, 468)
            .lblList.Caption = GetResourceString(resStringId, 253) & GetResourceString(92)
            .chkCumulative.Visible = False
        End With
        
        'Dim rstMemType As ADODB.Recordset
        m_TableName = "DepositTypeTab"
        m_Fields(0) = "DepositType"
        m_Fields(1) = "DepositTypeName"
        
End Select

Restart:

    gDbTrans.SqlStmt = "SELECT " & m_Fields(0) & "," & m_Fields(1) & " From " & m_TableName
    If m_GroupType = grpAllDeposit Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " Where ModuleID= " & M_ModuleID
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
    
    With m_frmGroup.cmbList
        .Clear
        .AddItem ""
        .ItemData(.newIndex) = 0
        If Not rst Is Nothing Then
            While Not rst.EOF
                .AddItem FormatField(rst(1))
                .ItemData(.newIndex) = FormatField(rst(0))
                rst.MoveNext
            Wend
        End If
    End With
    
    Set rst = Nothing

With m_frmGroup
    .Show 1
    
    intCumulative = IIf(.chkCumulative.Value = vbChecked, 1, 0)
    If .cmbCumulative.Visible = True And .cmbCumulative.ListIndex >= 0 Then _
            intCumulative = .cmbCumulative.ItemData(.cmbCumulative.ListIndex)
    
End With
    
If m_Status = recCancel Then
    Unload m_frmGroup
    Exit Function
End If

    
Select Case m_GroupType
    Case grpMember
        Call AddMemberType
        GoTo Restart
    Case grpDeposit
        intCumulative = IIf(m_frmGroup.chkCumulative.Value = vbChecked, Inst_Yearly, 0)
        If m_frmGroup.cmbCumulative.ListIndex >= 0 Then _
                intCumulative = m_frmGroup.cmbCumulative.ItemData(m_frmGroup.cmbCumulative.ListIndex)
        Call AddDeposit
        GoTo Restart
    Case grpLoanPurpose
        intCumulative = 0
        If m_frmGroup.cmbCumulative.ListIndex >= 0 Then _
                intCumulative = m_frmGroup.cmbCumulative.ItemData(m_frmGroup.cmbCumulative.ListIndex)
    Case grpCustomer
        intCumulative = IIf(m_frmGroup.chkCumulative.Value = vbChecked, 0, 1)
        
        
    Case grpCaste
        Call AddCaste
        GoTo Restart
    Case grpPlace
        Call AddPlace
        GoTo Restart
    Case grpAllDeposit
        Call AddDepositAccountType
        GoTo Restart
End Select

Dim strMsg  As String

If m_Status = recAdd Then
    'If Len(m_Fields(2)) < 1 Then m_Fields(2) = m_Fields(1) & "English"
    Dim NewID As Long
    gDbTrans.SqlStmt = "SELECT Max(" & m_Fields(0) & ") From " & m_TableName
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then NewID = FormatField(rst(0))
    Set rst = Nothing
    NewID = NewID + 1
    If m_GroupType = grpDeposit Or m_GroupType = grpCustomer Then
        If m_GroupType = grpDeposit And NewID < wisDeposit_FD Then NewID = wisDeposit_FD + NewID
        gDbTrans.SqlStmt = "Insert Into " & m_TableName & _
                "( " & m_Fields(0) & "," & m_Fields(1) & "," & m_Fields(2) & "," & m_Fields(3) & " ) " & _
                " VALUES (" & NewID & ", " & _
                AddQuotes(m_GroupName, True) & ", " & intCumulative & ", " & AddQuotes(m_GroupNameEnglish, True) & _
                 ")"
    
    Else 'If UBound(m_Fields) = 2 Then
        If Len(Trim$(m_Fields(2))) < 1 Then m_Fields(2) = m_Fields(1) & "English"
        gDbTrans.SqlStmt = "Insert Into " & m_TableName & _
                "( " & m_Fields(0) & "," & m_Fields(1) & "," & m_Fields(2) & " ) Values " & _
                "(" & NewID & ", " & AddQuotes(m_GroupName, True) & ", " & AddQuotes(m_GroupNameEnglish, True) & ")"
    
    'Else
     '   gDbTrans.SqlStmt = "Insert Into " & m_TableName & _
                "( " & m_Fields(0) & "," & m_Fields(1) & "," & m_Fields(2) & " ) " & _
                " VALUES (" & NewID & ", " & _
                AddQuotes(m_GroupName, True) & ", " & AddQuotes(m_GroupNameEnglish, True) & ", " & _
                intCumulative & ")"
    End If
    
    m_GroupID = NewID
    
    'MsgBox "Saved the details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(528)
    
ElseIf m_Status = recDelete Then
    
    Dim cancel As Integer
    RaiseEvent ItemDeleting(m_GroupName, cancel)
    If cancel Then Exit Function
    
    gDbTrans.SqlStmt = "Delete * from " & m_TableName & _
        " where " & m_Fields(1) & " = " & AddQuotes(m_GroupName, True)
            'MsgBox "Group Deleted", vbInformation
    strMsg = GetResourceString(677)

ElseIf m_Status = recUpdate Then
    
    If m_GroupType = grpDeposit Or m_GroupType = grpLoanPurpose Then _
        m_Fields(2) = ""
    
    If Len(m_Fields(2)) < 1 Then m_Fields(2) = m_Fields(1) & "English"
    gDbTrans.SqlStmt = "UPDATE " & m_TableName & _
            " SET " & m_Fields(1) & " = " & AddQuotes(m_GroupName, True) & "," & m_Fields(2) & " = " & AddQuotes(m_GroupNameEnglish, True) & _
            " WHERE " & m_Fields(0) & " = " & m_GroupID

    'MsgBox "Updated the Details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(707)

End If

gDbTrans.BeginTrans
If gDbTrans.SQLExecute Then
    gDbTrans.CommitTrans
Else
    gDbTrans.RollBack
End If

If Len(strMsg) Then MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE

If m_Status = recUpdate Or m_Status = recDelete Then RaiseEvent ItemDeleted(m_GroupName)
If m_Status = recAdd Or m_Status = recUpdate Then RaiseEvent ItemAdded(m_GroupName, m_GroupID)
    
GoTo Restart
End Function

Private Sub AddAccountGroup()
    Dim rst As ADODB.Recordset
Restart:
    gDbTrans.SqlStmt = "SELECT AccGroupId,GroupName,GroupNameEnglish From AccountGroup"
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
    With m_frmGroup
        .Caption = "Add/Remove Account Group"
        .lblItem.Caption = GetResourceString(36, 157) 'Group
        .lblList.Caption = GetResourceString(36, 157) 'Group
        .chkCumulative.Visible = False
        
        If Not rst Is Nothing Then
          With .cmbList
            .AddItem ""
            .ItemData(.newIndex) = 0
            While Not rst.EOF
                .AddItem FormatField(rst("GroupName"))
                .ItemData(.newIndex) = FormatField(rst("AccGroupID"))
                rst.MoveNext
            Wend
          End With
        End If
    End With
    
    Set rst = Nothing

m_frmGroup.Show 1
    
If m_Status = recCancel Then
    Unload m_frmGroup
    Exit Sub
End If

Dim strMsg  As String

If m_Status = recAdd Then
    Dim NewID As Long
    gDbTrans.SqlStmt = "SELECT MAx(AccGroupId) From Accountgroup"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then NewID = FormatField(rst(0))
    Set rst = Nothing
    NewID = NewID + 1
    gDbTrans.SqlStmt = "Insert Into AccountGroup Values " & _
            " (" & NewID & ", " & AddQuotes(m_GroupName, True) & ", " & AddQuotes(m_GroupNameEnglish, True) & ")"
    m_GroupID = NewID
    
    'MsgBox "Saved the details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(528)
ElseIf m_Status = recDelete Then
    
    Dim cancel As Integer
    RaiseEvent ItemDeleting(m_GroupName, cancel)
    If cancel Then Exit Sub
    
    gDbTrans.SqlStmt = "Delete * from AccountGroup " & _
        "where GroupName= " & AddQuotes(m_GroupName, True)
            'MsgBox "Group Deleted", vbInformation
    strMsg = GetResourceString(677)

ElseIf m_Status = recUpdate Then

    gDbTrans.SqlStmt = " UPDATE AccountGroup" & _
                   " SET GroupName = " & AddQuotes(m_GroupName, True) & _
                   " ,GroupNameEnglish = " & AddQuotes(m_GroupNameEnglish, True) & _
                   " WHERE AccGroupID = " & m_GroupID

    'MsgBox "Updated the Details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(707)

End If

gDbTrans.BeginTrans
If gDbTrans.SQLExecute Then
    gDbTrans.CommitTrans
Else
    gDbTrans.RollBack
End If

If Len(strMsg) Then MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE

If m_Status = recUpdate Then RaiseEvent ItemDeleted(m_GroupName)
If m_Status = recAdd Or m_Status = recUpdate Then RaiseEvent ItemAdded(m_GroupName, m_GroupID)
    
GoTo Restart
End Sub

Private Sub AddProductGroup()
    Dim rst As Recordset
    gDbTrans.SqlStmt = " SELECT GroupID,GroupName " & _
                  " FROM ProductGroup " & _
                  " ORDER BY GroupID"
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
    
    With m_frmGroup
        
        .Caption = "Add/Remove Product group"
        .lblItem.Caption = GetResourceString(157, 35)
        .lblItemEnglish.Caption = GetResourceString(157, 35, 468)
        .lblList.Caption = GetResourceString(157, 35)
        .chkCumulative.Visible = False
        If Not rst Is Nothing Then
          With .cmbList
            .AddItem ""
            .ItemData(.newIndex) = 0
            While Not rst.EOF
                .AddItem FormatField(rst("GroupName"))
                .ItemData(.newIndex) = FormatField(rst("GroupID"))
                rst.MoveNext
            Wend
          End With
        End If
        
    End With

    Set rst = Nothing

m_frmGroup.Show 1
    
If m_Status = recCancel Then
    Unload m_frmGroup
    Exit Sub
End If


Dim strMsg As String
Dim lngGroupID As Long


If m_Status = recAdd Then

    gDbTrans.SqlStmt = "SELECT MAX(GroupID) FROM ProductGroup"
    
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
    lngGroupID = FormatField(rst.Fields(0)) + 1
    
    'insert into the database
    gDbTrans.SqlStmt = " INSERT INTO ProductGroup (GroupID,GroupName,GroupNameEnglish) " & _
                      " VALUES ( " & _
                      lngGroupID & "," & _
                      AddQuotes(m_GroupName, True) & "," & _
                      AddQuotes(m_GroupNameEnglish, True) & " ) "
    
    'MsgBox "Saved the details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(528)
    
ElseIf m_Status = recDelete Then
    
    gDbTrans.SqlStmt = " SELECT ProductName FROM Products " & _
                   " WHERE GroupID = " & m_GroupID
                   
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then
        'MsgBox "Group can not be deleted while product from this group exists", vbInformation, wis_MESSAGE_TITLE
        MsgBox GetResourceString(772)
        Exit Sub
    End If
    
    Dim cancel As Integer
    RaiseEvent ItemDeleting(m_GroupName, cancel)
    If cancel Then Exit Sub
    
    Debug.Print "Kannada"
    If MsgBox("Do you want to Delete this group ? ", _
            vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    gDbTrans.SqlStmt = " DELETE FROM ProductGroup " & _
                       " WHERE GroupID = " & m_GroupID
    
    'MsgBox "Group Deleted", vbInformation
    strMsg = GetResourceString(677)

ElseIf m_Status = recUpdate Then

    gDbTrans.SqlStmt = " UPDATE ProductGroup " & _
                   " SET GroupName = " & AddQuotes(m_GroupName, True) & _
                   " ,GroupNameEnglish = " & AddQuotes(m_GroupNameEnglish, True) & _
                   " WHERE GroupID = " & m_GroupID

    'MsgBox "Updated the Details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(707)
End If

If m_Status = recNothing Then Exit Sub

gDbTrans.BeginTrans
If gDbTrans.SQLExecute Then
    gDbTrans.CommitTrans
Else
    gDbTrans.RollBack
End If

If Len(strMsg) Then MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE

If m_Status = recDelete Then RaiseEvent ItemDeleted(m_GroupName)
If m_Status = recAdd Or m_Status = recUpdate Then RaiseEvent ItemAdded(m_GroupName, lngGroupID)
    
    
End Sub

Private Sub AddUnitGroup()
    Dim rst As Recordset
    gDbTrans.SqlStmt = " SELECT UnitID,UnitName " & _
                  " FROM Units " & _
                  " ORDER BY UnitID"

    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
    
    With m_frmGroup
        
        .Caption = "Add/Remove unit"
        .lblItem.Caption = GetResourceString(161, 35)
        .lblList.Caption = GetResourceString(161, 35)
        .chkCumulative.Visible = False
        If Not rst Is Nothing Then
          With .cmbList
            .AddItem ""
            .ItemData(.newIndex) = 0
            While Not rst.EOF
                .AddItem FormatField(rst("UnitName"))
                .ItemData(.newIndex) = FormatField(rst("UnitID"))
                rst.MoveNext
            Wend
          End With
        End If
        
    End With

    Set rst = Nothing

m_frmGroup.Show 1
    
If m_Status = recCancel Then
    Unload m_frmGroup
    Exit Sub
End If

Dim lngUnitID As Long
Dim strMsg As String

If m_Status = recAdd Then


    gDbTrans.SqlStmt = " SELECT MAX(UnitID) FROM Units"
    
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
    lngUnitID = FormatField(rst.Fields(0)) + 1
    
    'insert into the database
    gDbTrans.SqlStmt = " INSERT INTO Units (UnitID,UnitName) " & _
                  " VALUES ( " & _
                  lngUnitID & "," & _
                  AddQuotes(m_GroupName, True) & " ) "

    'MsgBox "Saved the details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(528)
    
ElseIf m_Status = recDelete Then
    Dim cancel As Integer
    RaiseEvent ItemDeleting(m_GroupName, cancel)
    If cancel Then Exit Sub
    
    Debug.Print "Kannada"
    If MsgBox("Do you want to Delete this Unit ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    gDbTrans.SqlStmt = " DELETE FROM Units " & _
                    " WHERE UnitID = " & m_GroupID
    
    'MsgBox "Deleted the Unit", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(678)

ElseIf m_Status = recUpdate Then

    gDbTrans.SqlStmt = " UPDATE Units " & _
                   " SET UnitName = " & AddQuotes(m_GroupName, True) & _
                   " WHERE UnitID = " & m_GroupID

    'MsgBox "Updated the Details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(707)

End If

If m_Status = recNothing Then Exit Sub

gDbTrans.BeginTrans
If gDbTrans.SQLExecute Then
    gDbTrans.CommitTrans
Else
    gDbTrans.RollBack
End If

If Len(strMsg) Then MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE

If m_Status = recDelete Then RaiseEvent ItemDeleted(m_GroupName)
If m_Status = recAdd Or m_Status = recUpdate Then RaiseEvent ItemAdded(m_GroupName, 0)
    
End Sub

Public Property Get GroupType() As wis_GroupType
    GroupType = m_GroupType
End Property

Public Sub ShowAddGroup(GroupType As wis_GroupType, Optional ModuleID As wisModules = wis_None)
    m_GroupType = GroupType
    m_AllowDuplicate = False 'AllowDuplicate , Optional AllowDuplicate As Boolean = False
    M_ModuleID = ModuleID
    
    m_Status = recNothing
        
    If Not m_frmGroup Is Nothing Then
        Unload m_frmGroup
    Else
        Set m_frmGroup = New frmAddGroup
    End If
    
    Load m_frmGroup
    
    Call AddNew
    
    Exit Sub
    
    'Now Add it to the Group

    If m_GroupType = grpAccount Then Call AddAccountGroup
    If m_GroupType = grpCaste Then Call AddCaste
    If m_GroupType = grpPlace Then Call AddPlace
    If m_GroupType = grpCustomer Then Call AddCustomerType
    If m_GroupType = grpDeposit Then Call AddDeposit
    If m_GroupType = grpProduct Then Call AddProductGroup
    If m_GroupType = grpUnit Then Call AddUnitGroup
    If m_GroupType = grpLoanPurpose Then Call AddLoanPurpose
    If m_GroupType = grpFarmer Then Call AddFarmerType
    
    If m_Status <> recNothing Then Unload m_frmGroup
    
End Sub

Private Sub AddPlace()
    Dim rst As Recordset
    gDbTrans.SqlStmt = "SELECT PlaceID,Place FROm PlaceTab"
    Call gDbTrans.Fetch(rst, adOpenForwardOnly)
    
Dim PlaceID As Integer
Dim strMsg As String

If m_Status = recAdd Then
    PlaceID = 0
    gDbTrans.SqlStmt = "Select Max(PlaceID) From PlaceTab"
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
        PlaceID = FormatField(rst(0))
    PlaceID = PlaceID + 1
    
    gDbTrans.SqlStmt = "Insert Into PlaceTab (PlaceID,Place,PlaceEnglish)" & _
            " Values ( " & PlaceID & ", " & _
            AddQuotes(m_GroupName, True) & "," & _
            AddQuotes(m_GroupNameEnglish, True) & ")"
    
    gDbTrans.BeginTrans
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBack: GoTo Exit_Line
    gDbTrans.CommitTrans
    
    strMsg = GetResourceString(528)
    
ElseIf m_Status = recDelete Then

    'Check in the Name Tab whether this Cast s using or Not
    gDbTrans.SqlStmt = "Select Count(*) from NameTab where Place = " & AddQuotes(m_GroupName, True)
    If gDbTrans.Fetch(rst, adOpenDynamic) Then
        If FormatField(rst(0)) > 0 Then
            MsgBox GetResourceString(532), vbOKOnly, wis_MESSAGE_TITLE
            Exit Sub
        End If
    End If
    gDbTrans.SqlStmt = "Select Count(*) from ShgMaster where Place = " & AddQuotes(m_GroupName, True)
    If gDbTrans.Fetch(rst, adOpenDynamic) Then
        If FormatField(rst(0)) > 0 Then
            MsgBox GetResourceString(532), vbOKOnly, wis_MESSAGE_TITLE
            Exit Sub
        End If
    End If

    Dim cancel As Integer
    RaiseEvent ItemDeleting(m_GroupName, cancel)
    If cancel Then Exit Sub
    
    gDbTrans.SqlStmt = "Delete * from PlaceTab " & _
            " where Place= " & AddQuotes(m_GroupName, True)
    
    gDbTrans.BeginTrans
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBack: GoTo Exit_Line
    gDbTrans.CommitTrans
        
    'MsgBox "Group Deleted", vbInformation
    strMsg = GetResourceString(677)

ElseIf m_Status = recUpdate Then
    
    gDbTrans.BeginTrans
    
    gDbTrans.SqlStmt = "UPDATE NameTab " & _
                   " SET Place = " & AddQuotes(m_GroupName, True) & _
                   " ,PlaceEnglish = " & AddQuotes(m_GroupNameEnglish, True) & _
                   " WHERE Place = " & AddQuotes(m_PrevName)
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBack: GoTo Exit_Line
    
    gDbTrans.SqlStmt = "UPDATE ShgMaster " & _
                   " SET Place = " & AddQuotes(m_GroupName, True) & _
                   " WHERE Place = " & AddQuotes(m_PrevName)
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBack: GoTo Exit_Line
        
    gDbTrans.SqlStmt = "UPDATE PlaceTab " & _
                   " SET Place = " & AddQuotes(m_GroupName, True) & _
                   " ,PlaceEnglish = " & AddQuotes(m_GroupNameEnglish, True) & _
                   " WHERE PlaceID = " & m_GroupID

    If Not gDbTrans.SQLExecute Then gDbTrans.RollBack: GoTo Exit_Line
    
    gDbTrans.CommitTrans
    'MsgBox "Updated the Details ", vbInformation, wis_MESSAGE_TITLE
    strMsg = GetResourceString(707)
End If


If Len(strMsg) Then MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE

If m_Status = recUpdate Then RaiseEvent ItemDeleted(m_GroupName)
If m_Status = recAdd Or m_Status = recUpdate Then RaiseEvent ItemAdded(m_GroupName, m_GroupID)

Exit_Line:

End Sub

Private Function UpdateLedgerHeads(Optional searchExact As Boolean = False) As Boolean

UpdateLedgerHeads = False

'Get the Whether those heads are there are not
Dim rst As Recordset
Dim headID As Long
Dim AccType As Integer

Dim OldName As String
Dim OldNameEnglish As String
Dim NewName As String
Dim NewNameEnglish As String
Dim TableName As String
'Ledger Head
OldName = Trim$(m_PrevName)
NewName = Trim$(m_GroupName)

'gDbTrans.SqlStmt = "Select * From BankHeadIDs WHERE HeadName like '%" & AddQuotes(OldName, False) & "%'"
gDbTrans.SqlStmt = "Select HeadID,HeadName,HeadNameEnglish, 'BankHeadIDs' as TableName From BankHeadIDs WHERE HeadName like '%" & AddQuotes(OldName, False) & "%'" & _
    " UNION " & _
    "Select HeadID,HeadName,HeadNameEnglish, 'Heads' as TableName From Heads WHERE HeadName like '%" & AddQuotes(OldName, False) & "%'"
    

If searchExact Then
    gDbTrans.SqlStmt = "Select * From BankHeadIDs WHERE HeadName = " & AddQuotes(OldName)
    gDbTrans.SqlStmt = "Select HeadID,HeadName,HeadNameEnglish, 'BankHeadIDs' as TableName From BankHeadIDs WHERE HeadName =" & AddQuotes(OldName) & _
    " UNION " & _
    "Select HeadID,HeadName,HeadNameEnglish, 'Heads' as TableName From Heads WHERE HeadName = " & AddQuotes(OldName)

    NewNameEnglish = m_GroupNameEnglish
End If
If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then UpdateLedgerHeads = True: Exit Function

'AccType = FormatField(rst("AccType"))
    
'All other heads
'Change the name of the Head _
'which are related to the same account group
While Not rst.EOF
    headID = rst("HeadId")
    OldName = FormatField(rst("HeadName"))
    OldNameEnglish = FormatField(rst("HeadNameEnglish"))
    TableName = FormatField(rst("TableName"))
    
    NewName = Replace$(OldName, m_PrevName, m_GroupName)
    If Len(m_PrevNameEnglish) Then _
        NewNameEnglish = Replace$(OldNameEnglish, m_PrevNameEnglish, m_GroupNameEnglish)
    
    gDbTrans.SqlStmt = "Update " & TableName & " Set " & _
        " HeadName = " & AddQuotes(NewName, True) & "," & _
        " HeadNameEnglish = " & AddQuotes(NewNameEnglish, True) & _
        " WHERE HeadID = " & headID
    
    If Not gDbTrans.SQLExecute Then Exit Function
    
    rst.MoveNext
Wend

gDbTrans.SqlStmt = "Select * From ParentHeads " & _
    " WHERE HeadName like '%" & AddQuotes(m_PrevName, False) & "%'"
If searchExact Then
    gDbTrans.SqlStmt = "Select * From ParentHeads WHERE ParentName = " & AddQuotes(OldName)
    NewNameEnglish = m_GroupNameEnglish
End If
If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
    
    While Not rst.EOF
        headID = rst("ParentId")
        OldName = FormatField(rst("ParentName"))
        OldNameEnglish = FormatField(rst("ParentNameEnglish"))
        
        NewName = Replace$(OldName, m_PrevName, m_GroupName)
        If Len(m_PrevNameEnglish) Then _
            NewNameEnglish = Replace$(OldNameEnglish, m_PrevNameEnglish, m_GroupNameEnglish)
                
        gDbTrans.SqlStmt = "Update ParentHeads Set " & _
            " ParentName = " & AddQuotes(NewName, True) & _
            " , ParentNameEnglish = " & AddQuotes(NewNameEnglish, True) & _
            " WHERE ParentID = " & headID
        
        If Not gDbTrans.SQLExecute Then Exit Function
        rst.MoveNext
    Wend
    
 End If
 UpdateLedgerHeads = True
End Function

Private Sub m_frmGroup_AddClick(StrName As String, cancel As Integer, strNameEnglish As String)
    
    Dim count As Integer
    Dim MaxCount As Integer
    
    With m_frmGroup.cmbList
        If Not m_AllowDuplicate Then
            MaxCount = .ListCount - 1
            'Check FOr the Duplicate Entry
            For count = 0 To MaxCount
                If StrComp(.List(count), StrName, vbTextCompare) = 0 Then Exit For
            Next
            If count <= MaxCount Then
                Debug.Print "Kannada"
                MsgBox "Duplicate record not allowed", vbInformation, wis_MESSAGE_TITLE
                Exit Sub
            End If
        End If
    End With
    
    m_Status = recAdd
    m_GroupName = StrName
    m_GroupNameEnglish = strNameEnglish

End Sub

Private Sub m_frmGroup_CancelClick(cancel As Boolean)
    m_Status = recCancel
End Sub


Private Sub m_frmGroup_ItemSelected(Id As Long)
m_frmGroup.cmdAdd.Caption = GetResourceString(171)

Dim rst As Recordset

If m_GroupType = grpCustomer Then
    gDbTrans.SqlStmt = "Select UIType,CustTypeNameEnglish From CustomerType " & _
            " Where CustType = " & Id
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
            m_frmGroup.chkCumulative = IIf(FormatField(rst(0)) = 0, vbChecked, vbUnchecked)
            m_frmGroup.txtItemEnglish = FormatField(rst(1))
    End If
ElseIf m_GroupType = grpFarmer Then
    gDbTrans.SqlStmt = "Select FarmerTypeId,TypeEnlgish From FarmerTypeTab " & _
            " Where FarmerTypeId = " & Id
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
            m_frmGroup.chkCumulative = IIf(FormatField(rst(0)) = 0, vbChecked, vbUnchecked)
            m_frmGroup.txtItemEnglish = FormatField(rst(1))
    End If
ElseIf m_GroupType = grpDeposit Then
    gDbTrans.SqlStmt = "Select Cumulative,DepositNameEnglish From DepositNAme " & _
            " Where DepositID = " & Id
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        Dim Cumulative As wisInstallmentTypes
        With m_frmGroup
            .chkCumulative = IIf(FormatField(rst(0)), vbChecked, vbUnchecked)
            Cumulative = FormatField(rst(0))
            If Cumulative < Inst_Monthly Then Cumulative = Inst_Yearly
            If Cumulative = Inst_Monthly Then .cmbCumulative.ListIndex = 1
            If Cumulative = Inst_BiMonthly Then .cmbCumulative.ListIndex = 2
            If Cumulative = Inst_Quartery Then .cmbCumulative.ListIndex = 3
            If Cumulative = Inst_HalfYearly Then .cmbCumulative.ListIndex = 4
            If Cumulative = Inst_Yearly Then .cmbCumulative.ListIndex = 5
            
            .txtItemEnglish = FormatField(rst("DepositNameEnglish"))
        End With
    End If
Else
    If m_GroupType = grpAllDeposit Then _
        gDbTrans.SqlStmt = "Select DepositType,DepositTypeName, DepositTypeNameEnglish From DepositTypeTab " & _
            " Where DepositType = " & Id & " And ModuleId = " & M_ModuleID
    If m_GroupType = grpAccount Then _
        gDbTrans.SqlStmt = "SELECT AccGroupId,GroupName,GroupNameEnglish From AccountGroup" & _
            " Where AccGroupId = " & Id
    If m_GroupType = grpCaste Then _
        gDbTrans.SqlStmt = "SELECT CasteID, Caste,CasteEnglish FROM CasteTab" & _
            " Where CasteId = " & Id
    If m_GroupType = grpPlace Then _
        gDbTrans.SqlStmt = "SELECT PlaceID,Place,PlaceEnglish FROm PlaceTab Where PlaceId = " & Id
    'If m_GroupType = grpCustomer Then Call AddCustomerType
    'If m_GroupType = grpDeposit Then Call AddDeposit
    If m_GroupType = grpProduct Then _
        gDbTrans.SqlStmt = " SELECT GroupID,GroupName,GroupNameEnglish FROM ProductGroup " & _
                  " Where GroupId = " & Id
    If m_GroupType = grpUnit Then _
        gDbTrans.SqlStmt = " SELECT UnitID,UnitName,UnitNameEnglish FROM Units " & _
                  " Wherr UnitID = " & Id
    If m_GroupType = grpLoanPurpose Then _
        gDbTrans.SqlStmt = "SELECT PurposeId,Purpose,PurposeEnglish From LoanPurpose " & _
                " Where PurposeId = " & Id
    If m_GroupType = grpMember Then _
        gDbTrans.SqlStmt = "SELECT MemberTYpe,MemberTYpeName,MemberTypeNameEnglish FROM MemberTypeTab " & _
                  " Where MemberTYpe = " & Id
    'If m_GroupType = grpFarmer Then Call AddFarmerType
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then m_frmGroup.txtItemEnglish.Text = FormatField(rst(2))
End If
End Sub


Private Sub m_frmGroup_RemoveClick(StrName As String, Id As Long, cancel As Integer)
    
gDbTrans.SqlStmt = ""

If GroupType = grpCustomer Then
    'Before Deleting this CustomerType
    'Make sure that this customer type is not in use by the customer table
    gDbTrans.SqlStmt = "Select COUNT(*) From NameTab WHERE " & _
            "CustType = (Select CustType From CustomerType " & _
                "Where CustTypeName = " & AddQuotes(StrName, True) & ")"
ElseIf GroupType = grpPlace Then
    gDbTrans.SqlStmt = "Select COUNT(*) From NameTab WHERE " & _
            " FarmerTypeId = " & Id
ElseIf GroupType = grpPlace Then
    gDbTrans.SqlStmt = "Select COUNT(*) From NameTab WHERE " & _
            "Place = " & AddQuotes(StrName, True)
ElseIf GroupType = grpCaste Then
    gDbTrans.SqlStmt = "Select COUNT(*) From NameTab WHERE " & _
            "Caste = " & AddQuotes(StrName, True)

ElseIf GroupType = grpDeposit Then
    gDbTrans.SqlStmt = "Select COUNT(*) From FDMaster WHERE " & _
            "DepositType = (Select DepositID From DepositName " & _
                "Where DepositName = " & AddQuotes(StrName, True) & ")"

ElseIf GroupType = grpAllDeposit Then
    'gDbTrans.SqlStmt = "Select COUNT(*) From DepositTypeTab WHERE " & _
            "DepositType = (Select DepositID From DepositName " & _
                "Where DepositName = " & AddQuotes(StrName, True) & ")"

ElseIf GroupType = grpProduct Then
    gDbTrans.SqlStmt = "Select COUNT(*) From RelationMaster WHERE " & _
            "GroupID = (Select GroupID From ProductGroup " & _
                "Where GroupName = " & AddQuotes(StrName, True) & ")"

ElseIf GroupType = grpUnit Then
    gDbTrans.SqlStmt = "Select COUNT(*) From RelationMaster WHERE " & _
            "UnitID = (Select UnitID From Units " & _
                "Where UnitName = " & AddQuotes(StrName, True) & ")"

End If

If Len(gDbTrans.SqlStmt) Then
    Dim rstTemp As Recordset
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
        If FormatField(rstTemp(0)) > 0 Then
            Set rstTemp = Nothing
            MsgBox "This Group is in use ", vbInformation, wis_MESSAGE_TITLE
            Debug.Print "Kannada"
            cancel = True
            Exit Sub
        End If
    End If
End If

RaiseEvent ItemDeleting(StrName, cancel)

m_GroupName = StrName
m_PrevName = StrName
m_GroupID = Id

'If the User has canceled this procedure then Exit the sub
m_Status = recDelete
If cancel Then m_Status = recNothing

'Else Remove The Item
End Sub


Private Sub m_frmGroup_UpDateClick(StrName As String, Id As Long, cancel As Integer, strNameEnglish As String)
    
Dim count As Integer
Dim MaxCount As Integer
m_PrevName = ""
With m_frmGroup.cmbList
    If Not m_AllowDuplicate Then
        MaxCount = .ListCount - 1
        'Check FOr the Duplicate Entry
        For count = 0 To MaxCount
          'get the existing name of the head
          If Id = .ItemData(count) Then m_PrevName = .List(count)
          
          If count <> .ListIndex Then _
            If StrComp(.List(count), StrName, vbTextCompare) = 0 Then Exit For
        Next
        If count <= MaxCount Then
            m_PrevName = ""
            Debug.Print "Kannada"
            MsgBox "Duplicate record not allowed", vbInformation, wis_MESSAGE_TITLE
            Exit Sub
        End If
    End If
End With

m_Status = recUpdate
m_GroupName = StrName
m_GroupNameEnglish = strNameEnglish
m_GroupID = Id

End Sub


