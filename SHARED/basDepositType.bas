Attribute VB_Name = "basDepositType"
'Public Function SelectDepositType(ModuleID As wisModules, cancel As boolen, Optional ByRef haveMultiDeposits As Boolean, Optional ByRef DepositName, Optional ByRef DepositNameInEnglish) As Integer
Private Function SelectDepositType(ModuleID As wisModules, cancel As Boolean, ByRef haveMultiDeposits As Boolean) As Integer
    
    Dim selectDep As New clsSelectDeposit
    Dim multiDeposit As Boolean
    Dim Deptype As Integer
    Deptype = selectDep.SelectDeposit(ModuleID, grpAllDeposit, haveMultiDeposits, cancel)
    Set selectDep = Nothing
    
    If Deptype > -1 Then
        'm_frmSBAcc.DepositType = DepType
        'm_frmSBAcc.MultipleDeposit = multiDeposit
        'If m_DepositType <> DepType And m_frmSBAcc.IsFormLoaded Then m_frmSBAcc.txtAccNo = ""
        
    End If
    'm_DepositType = DepType
    SelectDepositType = Deptype
    
End Function
Public Sub SetDepositCheckBoxCaption(ByVal ModuleID As wisModules, ByRef chkDep As CheckBox, ByRef cmbDep As ComboBox)
    
    'Check the No Of Deposits
    Dim SBDepNames() As String
    SBDepNames = GetDepositTypesList(ModuleID)
    If UBound(SBDepNames) = 1 Then
        chkDep.Caption = GetResourceString(421, 271)  'GetDepositName(ModuleID, 0) & " " & GetResourceString(271)
        cmbDep.Visible = False
    Else
        chkDep.Caption = GetResourceString(271)
        'chkDep.Caption = ""
        cmbDep.Visible = True
        Call LoadDepositTypesToCombo(wis_SBAcc, cmbDep, True)
    End If
    
End Sub

Public Sub LoadDepositTypesToCombo(ByVal ModuleID As wisModules, ByRef cmbDeposit As ComboBox, Optional ClearCombo As Boolean = False)
    
    If ModuleID = wis_Deposits Then
        Call LoadFixedDepositTypes(cmbDeposit)
        Exit Sub
    End If
    
    
    If ClearCombo Then cmbDeposit.Clear
    Dim RstDep As Recordset
    gDbTrans.SqlStmt = "Select * from DepositTypeTab where ModuleID = " & ModuleID
    If gDbTrans.Fetch(RstDep, adOpenDynamic) > 0 Then
        While Not RstDep.EOF
            cmbDeposit.AddItem FormatField(RstDep("DepositTypeName"))
            cmbDeposit.ItemData(cmbDeposit.newIndex) = FormatField(RstDep("DepositType"))
            RstDep.MoveNext
        Wend
    Else
        Dim resId As Integer
        Call GetTableNameForModule(ModuleID, , , , , resId)
        cmbDeposit.AddItem GetResourceString(resId)
        cmbDeposit.ItemData(cmbDeposit.newIndex) = 0
    End If
    
End Sub
Public Function GetDepositTypeOfAccount(ModuleID As wisModules, AccId As Long, ByRef DepositName As String, ByRef DepositNameInEnglish As String) As String
     Dim TableName As String
     Dim rst As Recordset
     Dim Deptype As Integer
     TableName = GetMasterTableName(ModuleID)
     gDbTrans.SqlStmt = "Select * From " & TableName & " Where AccID = " & AccId
     If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        Deptype = FormatField(rst("DepositType"))
        DepositName = GetDepositName(ModuleID, Deptype, DepositNameInEnglish)
        GetDepositTypeOfAccount = Deptype
     End If
    
End Function
Public Function GetDepositTypesList(ModuleID As wisModules) As String()
    Dim depNames() As String
    Dim RstDep As Recordset
    gDbTrans.SqlStmt = "Select * from DepositTYpeTab where ModuleID = " & ModuleID
    If gDbTrans.Fetch(RstDep, adOpenDynamic) > 0 Then
        ReDim Preserve depNames(0)
        While Not RstDep.EOF
            ReDim Preserve depNames(UBound(depNames) + 1)
            depNames(UBound(depNames) - 1) = FormatField(RstDep("DepositTypeName"))
            'cmbDeposit.ItemData(cmbDeposit.newIndex) = FormatField(rstDep("DepositTypeName"))
            RstDep.MoveNext
        Wend
    Else
        Dim resId As Integer
        ReDim Preserve depNames(1)
        Call GetTableNameForModule(ModuleID, , , , , resId)
        depNames(0) = GetResourceString(resId)
        'cmbDeposit.ItemData(cmbDeposit.newIndex) = 0
    End If
    GetDepositTypesList = depNames
End Function
Public Function GetDepositTypeIDFromHeadID(AccHeadID As Long) As Integer
    Dim rstHeads As Recordset
    Dim RstDep As Recordset
    Dim headName As String
    Dim ParentID As Long
    Dim ModuleID As wisModules
    Dim retValue As Integer
    Dim Deptype As Integer
    
    gDbTrans.SqlStmt = "Select HeadName,AccType From BankHeadIds where Headid = " & AccHeadID
    
    If gDbTrans.Fetch(rstHeads, adOpenDynamic) > 0 Then
        'ParentID = FormatField(rstHeads("ParentID"))
        headName = FormatField(rstHeads("HeadName"))
        ModuleID = FormatField(rstHeads("AccTYpe"))
        ModuleID = ModuleID - ModuleID Mod 100
        
        Select Case ModuleID 'ParentID
            Case wis_SBAcc, wis_CAAcc, wis_PDAcc, wis_RDAcc 'parMemberDeposit, parMemDepLoan
                
                gDbTrans.SqlStmt = "Select * From DepositTypeTab where DepositTypeName = " & AddQuotes(headName, True)
                If gDbTrans.Fetch(RstDep, adOpenDynamic) > 0 Then
                    retValue = FormatField(RstDep("ModuleID"))
                    Deptype = FormatField(RstDep("DepositTYpe"))
                End If
            Case wis_Deposits
                gDbTrans.SqlStmt = "Select * From DepositName where DepositName = " & AddQuotes(headName, True)
                If gDbTrans.Fetch(RstDep, adOpenDynamic) > 0 Then
                    retValue = wis_Deposits
                    Deptype = FormatField(RstDep("DepositId"))
                End If
                
            Case wis_Loans
                gDbTrans.SqlStmt = "SELECT SchemeID,SchemeName from LoanScheme Order by SchemeName "
                If gDbTrans.Fetch(RstDep, adOpenDynamic) > 0 Then
                    retValue = wis_Loans
                    DepositType = FormatField(RstDep("SchemeID"))
                End If
            Case wis_Members
                gDbTrans.SqlStmt = "Select * From MemberTypeTab where MemberTypeName = " & AddQuotes(headName, True)
                If gDbTrans.Fetch(RstDep, adOpenDynamic) > 0 Then
                    DepositType = FormatField(RstDep("MemeberType"))
                End If
        End Select
    End If
    
    GetDepositTypeIDFromHeadID = Deptype
End Function
Public Function GetDepositName(ModuleID As wisModules, DepositType As Integer, Optional ByRef NameInEnglish As String) As String
    Dim depName As String
    Dim resStringId As Integer
    Dim rst As Recordset
    gDbTrans.SqlStmt = "Select * from DepositTypeTab Where ModuleID = " & ModuleID & " And DepositType =  " & DepositType
    If DepositType > 0 Then
        If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
            depName = FormatField(rst("DepositTypeName"))
            NameInEnglish = FormatField(rst("DepositTypeNameEnglish"))
        End If
    ElseIf DepositType = 0 Then
        Select Case ModuleID
            Case wis_CAAcc
               resStringId = 422
            Case wis_PDAcc
                resStringId = 425
            Case wis_RDAcc
                resStringId = 424
            Case wis_SBAcc
                resStringId = 421
            Case wis_Deposits
                resStringId = 423
            Case Else
                Err.Raise 400, , "Invalid Module"
        End Select
        
        depName = GetResourceString(resStringId)
        NameInEnglish = LoadResourceStringS(resStringId)
        
    End If
        
    GetDepositName = depName
End Function
Private Function GetDepositTypeID(ModuleID As wisModules, Optional ByRef MutliDeposit As Boolean) As Integer
    Dim rst As Recordset
    Dim TableName As String
    TableName = GetMasterTableName(ModuleID)
    gDbTrans.SqlStmt = "Select "
    
    GetAccountGroupTypeID = 0
End Function
Public Sub GetTableNameForModule(ByVal ModuleID As wisModules, Optional ByRef depMasterTableName As String, Optional ByRef depTransTable As String, Optional ByRef depPLTableName As String, Optional ByRef depIntPayableTableName As String, Optional ByRef resStringId As Integer)
   depIntPayableTableName = ""
    Select Case ModuleID
        Case wis_CAAcc
            depMasterTableName = "CAMaster"
            depTransTableName = "CATrans"
            depPLTableName = "CAPLTrans"
            resStringId = 422
        Case wis_PDAcc
            depMasterTableName = "PDMaster"
            depTransTableName = "PDTrans"
            depPLTableName = "PDIntTrans"
            depIntPayableTableName = "PDIntPayable"
            resStringId = 425
        Case wis_RDAcc
            depMasterTableName = "RDMaster"
            depTransTableName = "RDTrans"
            depPLTableName = "RDIntTrans"
            depIntPayableTableName = "RDIntPayable"
            resStringId = 424
        Case wis_SBAcc
            depMasterTableName = "SBMaster"
            depTransTableName = "SBTrans"
            depPLTableName = "SbPLTrans"
            resStringId = 421
        Case wis_Deposits
            depMasterTableName = "FDMaster"
            depTransTableName = "FDTrans"
            depPLTableName = "FDIntTrans"
            depIntPayableTableName = "FDIntPayable"
            resStringId = 421
        Case Else
            Err.Raise 400, , "Invalid Module"
    End Select
End Sub
Public Function GetMasterTableName(ModuleID As wisModules) As String
    Dim TableName As String
    Call GetTableNameForModule(ModuleID, TableName)
    GetMasterTableName = TableName
End Function
Public Function GetTransTableName(ModuleID As wisModules) As String
    Dim TableName As String
    Call GetTableNameForModule(ModuleID, , TableName)
    GetTransTableName = TableName
End Function
Public Function GetPLTransTableName(ModuleID As wisModules) As String
    Dim TableName As String
    Call GetTableNameForModule(ModuleID, , , TableName)
    GetPLTransTableName = TableName
End Function
Public Function GetIntPaybleTableName(ModuleID As wisModules) As String
    Dim TableName As String
    Call GetTableNameForModule(ModuleID, , , , TableName)
    GetIntPaybleTableName = TableName
End Function

