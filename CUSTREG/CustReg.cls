VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------
' clsRegister   :   Registers the details of a customer.
'                   This module will provide all the
'                   functionality such as displaying a dialog
'                   for getting client details and
'                   updating them to database.
'
' Last updated  :   30/12/1999 (dd/mm/yyyy)
' Updated by    :   Ravindranath M.
' Version No.   :   1.0.1
'--------------------------------------------------------------

Option Explicit

' Objects used by this module.
Private WithEvents RegDialog As frmCustReg
Attribute RegDialog.VB_VarHelpID = -1
Private WithEvents LookUpDialog As frmLookUp
Attribute LookUpDialog.VB_VarHelpID = -1
Private WithEvents m_frmCustInfo As frmCustInfo
Attribute m_frmCustInfo.VB_VarHelpID = -1

' Member variables...
Private m_Reference As Long 'Added on 2nd March, 2000
Private M_ModuleID As Long

Private m_CustomerID As Long
Private m_Title As String
Private m_FirstName As String
Private m_MiddleName As String
Private m_LastName As String
Private m_InstName As String
Private m_EnglishName As String
Private m_EnglishInstName As String

Private m_DOB As String
Private m_Guardian As String
Private m_Gender As Integer
Private m_MaritalStatus As Integer
Private m_Profession As String
Private m_Caste As String
Private m_Place As String
Private m_HomeAddress As String
Private m_OfficeAddress As String
Private m_HomePhone As String
Private m_OfficePhone As String
Private m_eMail As String
Private m_UIType As Byte
Private m_CustType As Byte
Private m_FarmerType As Byte

'Private m_dbFile As String
Private m_Modified As Boolean
Private m_CustomerLoaded As Boolean

'Introduced By Shashi on 23/4/2002
Private m_InsertCustomer  As Boolean

' Initialization status.
Private m_InitStatus As Integer


Public Property Get ObjectName()
    ObjectName = "CustClass"
End Property
Public Function CustomerNameNew(customerID As Long, ByRef CustomerNameEnglish As String)
    Dim rst As ADODB.Recordset
    gDbTrans.SQLStmt = "Select Title + ' ' + FirstName + ' ' + MiddleName + ' ' + LastName as Name,FullName " & _
            " From NameTab Where CustomerId = " & customerID
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <> 1 Then
        CustomerNameNew = " "
        CustomerNameEnglish = ""
    Else
        CustomerNameNew = FormatField(rst("Name"))
        CustomerNameEnglish = FormatField(rst("FullName"))
        If Len(CustomerNameEnglish) < 1 And gLangOffSet = 0 Then CustomerNameEnglish = CustomerNameNew
    End If
End Function
Public Function CustomerName(customerID As Long)
    Dim rst As ADODB.Recordset
    gDbTrans.SQLStmt = "Select Title + ' ' + FirstName + ' ' + MiddleName + ' ' + LastName as Name,FullName " & _
            " From NameTab Where CustomerId = " & customerID
    If gDbTrans.Fetch(rst, adOpenForwardOnly) <> 1 Then
        CustomerName = " "
    Else
        CustomerName = FormatField(rst("Name"))
    End If
End Function


Private Function GetNewCutomerId() As Long
Dim rst As ADODB.Recordset

'If we are creating two customers From the two places at same time
'Then THis function returns the same customer id
'to both functions and while saving the customer it may give error
'so it is better to read from one palce where
'we can update the customer id as we fetch this function


'TO SOLVE THIS PROBLEM
' ON TESTING BASIS I'LL KEEP THE CUSTOMER IN
'A TABLE called FreecustID

Dim custId As Long
Dim count As Integer
Dim transHasBegun  As Boolean

    gDbTrans.SQLStmt = "SELECT * FROM FreeCustId Where Selected = False " & _
        " And FreeID > 0 And  FreeID not in (Select distinct customerID from NameTab) Order By FreeID"
    If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then
        'INsert 10 new CutomerID
        gDbTrans.SQLStmt = "SELECT Max(CustomerID) From NameTab"
        Call gDbTrans.Fetch(rst, adOpenForwardOnly)
        custId = FormatField(rst(0))
        transHasBegun = gDbTrans.BeginTrans
        For count = 1 To 10
            gDbTrans.SQLStmt = "INSERT INTo FreeCustID ( FreeID, Selected) " & _
                " VALUES (" & custId + count & ", False )"
            
            If Not gDbTrans.SQLExecute Then
                If transHasBegun Then gDbTrans.RollBack
                Exit Function
            End If
        Next
        If transHasBegun Then gDbTrans.CommitTrans: transHasBegun = False
        gDbTrans.SQLStmt = "SELECT * FROM FreeCustId Where " & _
                        " Selected =False Order By FreeID "
        If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
    End If
    
    custId = Val(FormatField(rst("FreeID")))
    gDbTrans.SQLStmt = "UPDATE FreeCustID set Selected = True " & _
        " WHERE FreeID = " & custId
    transHasBegun = gDbTrans.BeginTrans
    Call gDbTrans.SQLExecute
    If transHasBegun Then gDbTrans.CommitTrans
    GetNewCutomerId = custId
    
    Exit Function
    
TestEnd:

' Get the new Customer ID from the Database.
    gDbTrans.SQLStmt = "SELECT MAX(CustomerID) FROM NameTab"
    custId = 1
    Dim Lret As Long
    Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If Lret < 1 Then
        Err.Raise 2000 + 2, , "Error In Customer Table"
    ElseIf Lret > 0 Then
        custId = Val(FormatField(rst(0))) + 1
    End If
    
    GoTo ExitLine
'Changed By shashi on 18/3/2002
    
    gDbTrans.SQLStmt = "SELECT count(*) as CustCount FROM NameTab"
    Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    If FormatField(rst(0)) = custId - 1 Then GoTo ExitLine
    gDbTrans.SQLStmt = "SELECT CustomerID FROM NameTab Order By CustomerId Asc"
    Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)
    custId = 1
    Do
        If rst.EOF Then GoTo ExitLine
        If custId <> rst("CustomerID") Then GoTo ExitLine
        custId = custId + 1
        rst.MoveNext
    Loop
    
ExitLine:

GetNewCutomerId = custId
    
transHasBegun = gDbTrans.BeginTrans
gDbTrans.SQLStmt = "UPDATE Install set ValueData = '" & Val(custId + 1) & "'" & _
    " WHERE KEYDATA = 'NewCustomerID'"
If gDbTrans.SQLExecute Then
    If transHasBegun Then gDbTrans.CommitTrans
Else
    If transHasBegun Then gDbTrans.RollBack
End If

End Function
Public Property Let ModuleID(NewVal As wisModules)

If NewVal <> wis_CAAcc And _
    NewVal <> wis_CustReg And _
    NewVal <> wis_Deposits And _
    NewVal <> wis_Members And _
    NewVal <> wis_PDAcc And _
    NewVal <> wis_RDAcc And _
    NewVal <> wis_SBAcc And _
    NewVal <> wis_Users And _
    NewVal <> wis_MatAcc And _
    NewVal <> wis_Loans And _
    NewVal <> wis_None Then
        Err.Raise "Invalid Module ID specified !"
End If
M_ModuleID = NewVal
m_Reference = m_Reference Or NewVal

End Property
' Clears the current customer details.
Private Sub Clear()

m_Reference = 0

m_Title = ""
m_FirstName = ""
m_MiddleName = ""
m_LastName = ""
m_EnglishName = ""
m_DOB = ""
m_Guardian = ""
m_Gender = -1
m_MaritalStatus = -1
m_Profession = ""
m_Caste = ""
m_HomeAddress = ""
m_OfficeAddress = ""
m_HomePhone = ""
m_OfficePhone = ""
m_eMail = ""
m_CustomerLoaded = False
m_UIType = 0
m_CustType = 0

'Delete the Photos of any unsaved
Call DeleteUnregisteredImageFiles

If RegDialog Is Nothing Then Set RegDialog = New frmCustReg
RegDialog.NewCustomer = True
Call LoadToForm


m_InsertCustomer = True

End Sub
Public Property Get FullNameEnglish() As String
    FullNameEnglish = Trim$(m_EnglishName)
End Property
Public Property Get FullName() As String
    
If m_UIType = 0 Then
    FullName = Trim$(m_Title & " " & m_FirstName & " " & m_MiddleName & " " & m_LastName)
Else
    FullName = Trim$(m_Title & " " & m_InstName)
End If
    
    
End Property
Public Property Get FormValue() As frmCustReg
    Set FormValue = RegDialog
End Property

' Gets the customer details for a specified Account number.
Public Function LoadCustomerInfo(customerID As Long) As Boolean
On Error GoTo Err_line

Dim StrSql As String, rst As ADODB.Recordset
Dim Lret As Long
gDbTrans.SQLStmt = "SELECT * FROM NameTab WHERE CustomerID = " & customerID
Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)

If RegDialog Is Nothing Then Set RegDialog = New frmCustReg

If Lret > 0 Then
    ' Update the variables.
    m_InsertCustomer = False
    
    With RegDialog
        m_Reference = FormatField(rst("Reference")) Or M_ModuleID
        m_CustomerID = FormatField(rst("CustomerID"))
        m_Title = FormatField(rst("Title"))
        m_FirstName = FormatField(rst("FirstName"))
        m_MiddleName = FormatField(rst("MiddleName"))
        m_LastName = FormatField(rst("LastName"))
        m_EnglishName = FormatField(rst("FullName"))
                
        m_DOB = FormatField(rst("DOB"))
        If m_DOB = "1/1/100" Then m_DOB = ""
        'm_Guardian = FormatField(gDBTrans.Rst("Guardian"))
        'm_Guardian = "Create field in Name Table"
        m_Gender = FormatField(rst("Gender"))
        m_MaritalStatus = Val(FormatField(rst("MaritalStatus")))
        m_Profession = FormatField(rst("Profession"))
        m_Caste = FormatField(rst("Caste"))

        ' Update the home address details...
        m_HomeAddress = FormatField(rst("HomeAddress"))
        m_OfficeAddress = FormatField(rst("OfficeAddress"))
        m_Place = FormatField(rst("Place"))

        'Phone nos
        m_HomePhone = FormatField(rst("HomePhone"))
        m_OfficePhone = FormatField(rst("OfficePhone"))

        ' eMail address, if any...
        m_eMail = FormatField(rst("eMail"))
        
        
        'Load the Farmer Type
        With .cmbFarmerType
            Dim FarmerType As Integer
            FarmerType = FormatField(rst("FarmerTypeID"))
            Dim I As Integer
            For I = 0 To .ListCount - 1
                If .ItemData(I) = FarmerType Then
                    .ListIndex = I
                    Exit For
                End If
            Next
        End With
        
        m_CustType = FormatField(rst("CustType"))
        'Now Get the Informattion Of the UI Type
        gDbTrans.SQLStmt = "SELECT UITYPE From CustomerType " & _
            " WHERE CustType = " & m_CustType
        If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then _
                m_UIType = FormatField(rst("UITYPE"))
                
        .cmdLookUp.Enabled = False
        .cmdOk.Default = True
        
        .NewCustomer = False
        .setAccNo (m_CustomerID)
        .PhotoInitialize
    
    End With

    m_CustomerLoaded = True  ' Set the Existing Flag
    LoadToForm      ' Load the details to the Interface dialog.
    
    LoadCustomerInfo = True
Else
    RegDialog.NewCustomer = True
    'MsgBox "No details available for Customer ID: " _
        & CustomerID, vbExclamation, wis_MESSAGE_TITLE
    MsgBox GetResourceString(662) & customerID, vbExclamation, wis_MESSAGE_TITLE

End If

Exit Function

Err_line:
    If Err.Number = 383 Then    'Combo text read only.
        Resume Next
    ElseIf Err Then
        'MsgBox "Error while getting the customer details " _
                & "for account number " & CustomerID & vbCrLf _
                & Err.Description, vbCritical
        
        MsgBox GetResourceString(651) & customerID & vbCrLf _
                & Err.Description, vbCritical
    End If
'Resume
End Function


Private Sub LoadToForm()
On Error Resume Next
Dim count As Byte
    ' Update the variables.
    With RegDialog
        '.cmdLookup.Caption = IIf(m_CustomerLoaded, "&Clear", "&Lookup")
        If m_Title = "" Then
            .cmbTitle.ListIndex = -1
        Else
            count = 0
            With .cmbTitle
                Do
                    If .List(count) = m_Title Then
                        .ListIndex = count
                        Exit Do
                    End If
                    count = count + 1
                    If count = .ListCount Then Exit Do
                Loop
                .Text = m_Title
            End With
        End If
        
        'If m_UIType = 1 Then
            If m_Title = "" Then
                .cmbInstTitle.ListIndex = -1
            Else
                With .cmbInstTitle
                    count = 0
                    Do
                        If .List(count) = m_Title Then
                            .ListIndex = count
                            Exit Do
                        End If
                        count = count + 1
                        If count = .ListCount Then Exit Do
                    Loop
                    '.Text = m_Title
                End With
            End If
            
            .txtInstName = m_FirstName & " " & m_MiddleName & _
                 " " & m_LastName
            .txtInstHead = m_Profession
            .txtEstd = m_DOB
            m_InstName = .txtInstName
            'm_
        'End If
        
        .txtFirstName = m_FirstName
        .txtMiddleName = m_MiddleName
        .txtLastName = m_LastName
        .txtEnglishName = m_EnglishName
        .txtDOB.Text = m_DOB
        If m_DOB <> "" Then 'Exit Sub
            If DateDiff("yyyy", GetSysFormatDate(.txtDOB.Text), gStrDate) < 18 Then
                .txtGuardian = m_Guardian
                .txtGuardian.Enabled = True
                .txtGuardian.BackColor = vbWhite
                .lblGuardian.Enabled = True
            Else
                .txtGuardian.Enabled = False
                .txtGuardian.BackColor = .BackColor
                .lblGuardian.Enabled = False
            End If
        End If
        
        If m_Gender Then .cmbGender.ListIndex = m_Gender - 1
        'If m_Gender >= 0 Then .cmbGender.ListIndex = m_Gender
        
        .cmbMaritalStatus.ListIndex = m_MaritalStatus
        'If m_MaritalStatus >= 0 Then .cmbMaritalStatus.ListIndex = m_MaritalStatus
        .txtProfession = m_Profession
        .txtCaste = m_Caste
        If m_Caste = "" Then
            .cmbCaste.ListIndex = -1
        Else
            count = 0
            Do
                If .cmbCaste.List(count) = m_Caste Then
                    .cmbCaste.ListIndex = count
                    Exit Do
                End If
                count = count + 1
                If count = .cmbCaste.ListCount Then Exit Do
            Loop
'            .cmbCaste.Text = m_Caste
        End If
        ' Update the home address details...
        .HomeAddress = m_HomeAddress
        .OfficeAddress = m_OfficeAddress
        .txtHomeCity.Text = m_Place
        If m_Place = "" Then
            .cmbHomeCity.ListIndex = -1
        Else
            With .cmbHomeCity
                count = 0
                Do
                    If .List(count) = m_Place Then
                        .ListIndex = count
                        Exit Do
                    End If
                    count = count + 1
                    If count = .ListCount Then Exit Do
                Loop
                '.cmbHomeCity.Text = m_Place
            End With
        End If
        'Home and office phones
        .txtHomePhone.Text = m_HomePhone
        .txtOffPhone.Text = m_OfficePhone
        
        ' eMail address, if any...
        .txtEmail.Text = m_eMail
        
        ''Detail of the Institutions UI
        .fraCust(m_UIType).ZOrder 0
        With .cmbCustType
            count = 0
            Do
                If .ItemData(count) = m_CustType Then
                    .ListIndex = count
                    Exit Do
                End If
                count = count + 1
                If count = .ListCount Then Exit Do
            Loop
        End With
        
        .optUI(0).Enabled = False
        .optUI(1).Enabled = False
        .cmbCustType.Enabled = False
        If gCurrUser.UserPermissions = perOnlyWaves Or m_CustType = 0 Then
            .optUI(0) = False
            .optUI(1) = False
            .optUI(0).Enabled = True
            .optUI(1).Enabled = True
            .cmbCustType.Enabled = True
        End If
        
    End With
    
End Sub
Public Sub NewCustomer()
On Error GoTo Err_line
' Load the details to form.
Call Clear
If m_CustomerID Then
    'Before assigning the new customer
    'update the status Old Customer
    'If the old customer has saved then
    'the detail of that id has been deleted
    'In such case the below code will not affect
    gDbTrans.SQLStmt = "UPDate FreeCustID Set Selected = False " & _
            " WHERE FreeId = " & m_CustomerID
    Dim transHasBegun As Boolean
    transHasBegun = gDbTrans.BeginTrans
    Call gDbTrans.SQLExecute
    If transHasBegun Then gDbTrans.CommitTrans
    m_CustomerID = 0

End If
If RegDialog Is Nothing Then Set RegDialog = New frmCustReg
With RegDialog.cmdLookUp
    .Enabled = True
    .Default = True
End With

m_Modified = False  ' Set the Modified flag to False.

Exit Sub
Err_line:
    If Err Then
        'MsgBox "Error in initializing the Customer details." _
            & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
        MsgBox "Error in initializing the Customer details." _
            & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    End If
'Resume
End Sub

'Displays the Customer Registration dialog.
Public Sub ShowDialog()
If RegDialog Is Nothing Then Set RegDialog = New frmCustReg
If m_CustomerLoaded Then
    RegDialog.cmdLookUp.Caption = GetResourceString(8)    ' Clear
Else
    RegDialog.cmdLookUp.Caption = GetResourceString(17)    ' Lookup
End If
RegDialog.Show vbModal
End Sub

Public Property Get ObjectHandle() As Long
    If Not RegDialog Is Nothing Then _
            ObjectHandle = RegDialog.hwnd: Exit Property
    If Not m_frmCustInfo Is Nothing Then _
            ObjectHandle = m_frmCustInfo.hwnd
End Property

Public Sub ShowCustInfo()
    If m_frmCustInfo Is Nothing Then Set m_frmCustInfo = New frmCustInfo
    
    m_frmCustInfo.Show
    gWindowHandle = m_frmCustInfo.hwnd
    
End Sub

' Prepares an SQL statement for data INSERTION/UPDATION and returns the same.
' INPUT:
'   UpdateType: 1. Insertion;  2. Updation
'
' Returns:
'   SQL statement string.
' Calling Procedure: UpdateName routine
'
Private Function GetSQL(UpdateType As Integer) As String

Dim StrSql As String
Dim IsciName As String
Dim Gender As wis_Gender
Dim FarmerTypeID As Integer

'Convert The Customer Name into ISCii format
With RegDialog
    On Error Resume Next
    IsciName = Left((.txtFirstName + " " + .txtLastName), 20)
    If .txtEnglishName.Text = "" Then
        If gLangOffSet = 0 Then
            .txtEnglishName.Text = .txtFirstName.Text + " " + .txtLastName.Text
        ElseIf gLangShree Then
            .txtEnglishName.Text = ConvertToEnglish(.txtFirstName.Text + " " + .txtLastName.Text)
        End If
    End If
    On Error GoTo 0
End With

If gLangOffSet <> wis_NoLangOffset Then IsciName = Left(ConvertToIscii(IsciName), 20)

' If Date of Birth Not speicfied then Put a dummy value there
If Trim$(m_DOB) = "" Then m_DOB = "1/1/100"

'**************
With RegDialog
    If .cmbGender.ListIndex = -1 Then
        Gender = wisNoGender
    Else
        Gender = .cmbGender.ItemData(.cmbGender.ListIndex)
    End If
    
    FarmerTypeID = 0
    If .cmbFarmerType.ListIndex >= 0 Then _
        FarmerTypeID = .cmbFarmerType.ItemData(.cmbFarmerType.ListIndex)
    
    Select Case UpdateType
        Case wis_INSERT
            'Set the ModuleID
            'RefID = m_ModuleID
            If m_CustomerID = 0 Then m_CustomerID = GetNewCutomerId
            StrSql = "INSERT INTO nametab (CustomerID,Title, FirstName, " _
                & "MiddleName, LastName, Gender, DOB, MaritalStatus, " _
                & "Profession, caste, HomeAddress, OfficeAddress, " _
                & "HomePhone, OfficePhone, eMail, Place, CustType, "
            
            If FarmerTypeID > 0 Then StrSql = StrSql & " FarmerTypeID, "
            StrSql = StrSql & " Reference,IsciName,FullName) Values ("
            
            StrSql = StrSql _
                & m_CustomerID & ", " _
                & AddQuotes(.cmbTitle.Text, True) & ", " _
                & AddQuotes(.txtFirstName.Text, True) & ", " _
                & AddQuotes(.txtMiddleName.Text, True) & ", " _
                & AddQuotes(.txtLastName.Text, True) & ", " _
                & Gender & ", #" & GetSysFormatDate(m_DOB) & "#, " _
                & .cmbMaritalStatus.ListIndex & ", " _
                & AddQuotes(.txtProfession.Text, True) & ", " _
                & AddQuotes(.txtCaste.Text, True) & ", " _
                & AddQuotes(.HomeAddress, True) & ", " _
                & AddQuotes(.OfficeAddress, True) & ", " _
                & AddQuotes(.txtHomePhone.Text, True) & ", " _
                & AddQuotes(.txtOffPhone.Text, True) & ", " _
                & AddQuotes(Left(.txtEmail.Text, 25), True) & ", " _
                & AddQuotes(.txtHomeCity.Text, True) & ", " _
                & m_CustType & ", "
            If FarmerTypeID > 0 Then StrSql = StrSql & FarmerTypeID & ", "
            StrSql = StrSql & m_Reference & ", " _
                & AddQuotes(IsciName, True) & ", " _
                & AddQuotes(.txtEnglishName.Text, True) _
                & ")"

        Case wis_UPDATE
            StrSql = "UPDATE nameTab SET " _
                & "Title = " & AddQuotes(.cmbTitle.Text, True) & ", " _
                & "FirstName = " & AddQuotes(.txtFirstName.Text, True) & ", " _
                & "MiddleName = " & AddQuotes(.txtMiddleName.Text, True) & ", " _
                & "LastName = " & AddQuotes(.txtLastName.Text, True) & ", " _
                & "Gender = " & Gender & ", " _
                & "DOB = #" & GetSysFormatDate(m_DOB) & "#, MaritalStatus = " _
                & .cmbMaritalStatus.ListIndex & ", Profession = " _
                & AddQuotes(.txtProfession.Text, True) & ", " _
                & "Caste = " & AddQuotes(.txtCaste.Text, True) & ", " _
                & "HomeAddress = " & AddQuotes(.HomeAddress, True) & ", " _
                & "OfficeAddress = " & AddQuotes(.OfficeAddress, True) & ", " _
                & "HomePhone = " & AddQuotes(.txtHomePhone.Text, True) & ", " _
                & "OfficePhone = " & AddQuotes(.txtOffPhone.Text, True) & ", " _
                & "eMail = " & AddQuotes(.txtEmail.Text, True) & ", " _
                & "Place = " & AddQuotes(.txtHomeCity.Text, True) & ", " _
                & "CustType = " & m_CustType & "," & "Reference = " & m_Reference & "," _
                & "IsciName = " & AddQuotes(IsciName, True) & "," _
                & "FullName = " & AddQuotes(.txtEnglishName.Text, True)
                If FarmerTypeID > 0 Then StrSql = StrSql & ",FarmerTypeID = " & FarmerTypeID
                StrSql = StrSql & " where Customerid = " & m_CustomerID & ";"
    End Select
    
End With
GetSQL = StrSql
End Function

Public Function SaveCustomer() As Boolean
Dim inTransaction As Boolean
Dim transHasBegun  As Boolean

'If the details of the customer are not modified, exit.
If Not m_Modified Then
    SaveCustomer = True
    GoTo Exit_Line
End If

'Throw error if ModuleID is not set
If CustomerLoaded Then
    If M_ModuleID = 0 Then
        Err.Raise 20001, "CustReg", "Invalid module id set !"
    End If
End If


On Error GoTo SaveCustomer_err
Dim nRet As Integer

With gDbTrans
    'If Not CustomerLoaded Then
    If m_InsertCustomer Then
        ' Get the SQL statement for data insertion.
        .SQLStmt = GetSQL(wis_INSERT)
    Else
        .SQLStmt = GetSQL(wis_UPDATE)
    End If

'Temporary block of code, to be redefined later
    'If Not CustomerLoaded Then
        transHasBegun = gDbTrans.BeginTrans
        inTransaction = True
    'End If
    
    If Not .SQLExecute Then
        If transHasBegun Then gDbTrans.RollBack
        'MsgBox "Error adding the Customer details.", _
                vbCritical, wis_MESSAGE_TITLE
        MsgBox "Error adding the Customer details.", _
                vbCritical, wis_MESSAGE_TITLE
        GoTo SaveCustomer_err
    End If

'Temporary block of code, to be redefined later
    If m_InsertCustomer Then
        'Now Save the Photo & Sign
        Call SaveUnregisteredImageFiles(m_CustomerID)
        
        m_InsertCustomer = False
        gDbTrans.SQLStmt = "DELETE * From FreeCustID WHERE FreeId = " & m_CustomerID
        If Not .SQLExecute Then
            If transHasBegun Then gDbTrans.RollBack
            'MsgBox "Error adding the Customer details.", _
                    vbCritical, wis_MESSAGE_TITLE
            MsgBox "Error adding the Customer details.", _
                    vbCritical, wis_MESSAGE_TITLE
            GoTo SaveCustomer_err
        End If
    End If
    If transHasBegun Then gDbTrans.CommitTrans: transHasBegun = False
    inTransaction = False

End With

SaveCustomer = True
m_CustomerLoaded = True
m_Modified = False  ' Reset the modified flag.
'MsgBox "Saved the Client detals.", vbInformation, wis_MESSAGE_TITLE


Exit_Line:
    If transHasBegun Then gDbTrans.RollBack
    
    Exit Function

SaveCustomer_err:
    
    If Not CustomerLoaded And transHasBegun Then gDbTrans.RollBack
    ' Handle errors here...
    If Err Then MsgBox "UpdateName: " & Err.Description, vbCritical
    
    GoTo Exit_Line
    
End Function
Private Sub Class_Initialize()

'Reset the module id
    m_Reference = 0
    
' Create an instance of Register dialog.
'If RegDialog Is Nothing Then Set RegDialog = New frmCustReg

' Initialize the new customer details.
 'Me.NewCustomer
 'RegDialog.NewCustomer = True
 m_InsertCustomer = True
End Sub
Private Sub Class_Terminate()
' Destroy the dbTransaction class object.
'Set dbObj = Nothing
Dim transHasBegun  As Boolean

If customerID > 0 And CustomerLoaded Then
    transHasBegun = gDbTrans.BeginTrans
    gDbTrans.SQLStmt = "UPDATE FreeCustID set Selected = False " & _
        " WHERE FreeId = " & m_CustomerID
    If gDbTrans.SQLExecute Then
        If transHasBegun Then gDbTrans.CommitTrans
    Else
        If transHasBegun Then gDbTrans.RollBack
    End If
End If
    
If Not RegDialog Is Nothing Then
    Unload RegDialog
    Set RegDialog = Nothing
End If
If Not LookUpDialog Is Nothing Then
    Unload LookUpDialog
    Set LookUpDialog = Nothing
End If
End Sub
Public Property Get customerID() As Long
    customerID = m_CustomerID
End Property
Public Property Get Modified() As Boolean
    Modified = m_Modified
End Property

Public Property Let Modified(NewVal As Boolean)
    m_Modified = NewVal
End Property

Private Sub LookUpDialog_SelectClick(strSelection As String)
'MsgBox strSelection
If Not LoadCustomerInfo(Val(strSelection)) Then
    'MsgBox "Unable to load customer details !", vbCritical, gAppName & " - Error"
    MsgBox GetResourceString(555), vbCritical, gAppName & " - Error"
    Exit Sub
End If

'  If the customer details are loaded successfully,
' set the flag "m_CustomerLoaded" to True.
' And set the Inser stament as false
m_InsertCustomer = False
m_CustomerLoaded = True
With RegDialog.cmdLookUp
    .Caption = GetResourceString(8)
    .Default = False
End With

End Sub

Private Sub m_frmCustInfo_WindowClosed()
    gWindowHandle = 0
    Set m_frmCustInfo = Nothing
End Sub

Private Sub RegDialog_CancelClick()
'    m_CustomerID = 0
    Call LoadToForm
End Sub

Private Sub RegDialog_ClearClick()
Call Clear
RegDialog.cmdLookUp.Default = True
m_CustomerLoaded = False
m_InsertCustomer = True

End Sub

Private Sub RegDialog_LookUpClick(strSearch As String)
Screen.MousePointer = vbHourglass
On Error GoTo Err_line
Dim Lret As Long
Dim rst As ADODB.Recordset

' Query the database to get all the customer names...
With gDbTrans
    .SQLStmt = "SELECT CustomerId, " & _
        " Title + ' ' + FirstName + ' ' + MIddleName + ' ' + LastName AS Name, " & _
        " Profession FROM NameTab "
    If Trim(strSearch) <> "" Then
        .SQLStmt = .SQLStmt & " Where (FirstName like '" & strSearch & "%' " & _
            " Or MiddleName like '" & strSearch & "%' " & _
            " Or LastName like '" & strSearch & "%')"
        .SQLStmt = .SQLStmt & " Order by IsciName"
    Else
        .SQLStmt = .SQLStmt & " Order by CustomerId"
    End If
    Lret = .Fetch(rst, adOpenStatic)
    If Lret <= 0 Then
        Screen.MousePointer = vbDefault
        'MsgBox "No data available!", vbExclamation
        MsgBox GetResourceString(278), vbExclamation
        GoTo Exit_Line
    End If
End With


' Create a report dialog.
If LookUpDialog Is Nothing Then Set LookUpDialog = New frmLookUp

With LookUpDialog
    .m_SelItem = ""
    ' Fill the data to report dialog.
    If Not FillView(.lvwReport, rst, True) Then
        'MsgBox "Error filling the customer details.", vbCritical
        MsgBox "Error filling the customer details.", vbCritical
        GoTo Exit_Line
    End If
    Screen.MousePointer = vbDefault
    ' Display the dialog.
    .Show vbModal
   End With

Exit_Line:
    Screen.MousePointer = vbDefault
    Exit Sub

Err_line:
    If Err Then
        'MsgBox "Data lookup: " & vbCrLf _
            & Err.Description, vbCritical
        MsgBox "Data lookup: " & vbCrLf _
            & Err.Description, vbCritical
    End If
    GoTo Exit_Line
End Sub

Private Sub RegDialog_OKClick()
Dim count As Integer
Dim MaxCount As Integer
Dim strNames() As String
Dim rst As Recordset
' Check the values of RegDialog against the
' member variables. If any of these are changed
' set the Modified Flag to True.


With RegDialog

    With .cmbCustType
        'Check the UI Type
        m_UIType = Val(.Tag)
        
        'Get the Customer type
        If .ListIndex < 0 Then .ListIndex = 0
        m_CustType = .ItemData(.ListIndex)
    End With

    
    'Check the Institutions name.
    If m_UIType = 0 Then
        ' Check the Title.
        If StrComp(m_Title, .cmbTitle.Text, vbTextCompare) <> 0 Then
            m_Title = .cmbTitle.Text
            m_Modified = True
        End If
        ' Check the First name.
        If StrComp(m_FirstName, .txtFirstName.Text, vbTextCompare) <> 0 Then
            m_FirstName = Trim$(.txtFirstName.Text)
            m_Modified = True
        End If
        ' Check the Middle name.
        If StrComp(m_MiddleName, .txtMiddleName.Text, vbTextCompare) <> 0 Then
            m_MiddleName = Trim$(.txtMiddleName.Text)
            m_Modified = True
        End If
        ' Check the Last name.
        If StrComp(m_LastName, .txtLastName.Text, vbTextCompare) <> 0 Then
            m_LastName = Trim$(.txtLastName.Text)
            m_Modified = True
        End If
        ' Check the Age value.
        If m_DOB <> (.txtDOB.Text) Then
            m_DOB = Trim$(.txtDOB.Text)
            m_Modified = True
        End If
        ' Check the Gender.
        If m_Gender <> .cmbGender.ListIndex Then
            If .cmbGender.ListIndex >= 0 Then m_Gender = .cmbGender.ListIndex
            'update nametab for gender
            m_Modified = True
        End If
    
        ' Check Marital Status.
        If m_MaritalStatus <> .cmbMaritalStatus.ListIndex Then
            m_MaritalStatus = .cmbMaritalStatus.ListIndex
            m_Modified = True
        End If
    
        ' Check the Profession.
        If StrComp(m_Profession, .txtProfession.Text, vbTextCompare) <> 0 Then
            m_Profession = Trim$(.txtProfession.Text)
            m_Modified = True
        End If
    
        ' Check the Caste.
        If StrComp(m_Caste, .txtCaste.Text, vbTextCompare) <> 0 Then
            m_Caste = Trim$(.txtCaste.Text)
            m_Modified = True
        End If
    
    Else
        
        ' Check the Title.
        If StrComp(m_Title, .cmbInstTitle.Text, vbTextCompare) <> 0 Then
            m_Title = .cmbInstTitle.Text
            m_Modified = True
            'Exit Sub
        End If
        If StrComp(m_InstName, .txtInstName.Text, vbTextCompare) <> 0 Then
            m_InstName = Trim$(.txtInstName.Text)
            m_Modified = True
        End If
        strNames() = Split(.txtInstName)
        MaxCount = UBound(strNames)
        If MaxCount > 2 Then
            For count = 1 To MaxCount - 2 ' To 2 Step -1
                strNames(0) = strNames(0) & " " & strNames(count)
            Next
            strNames(1) = strNames(MaxCount - 1)
            strNames(2) = strNames(MaxCount)
        End If
        ReDim Preserve strNames(2)
        m_FirstName = Left(strNames(0), 250)
        m_MiddleName = Left(strNames(1), 100)
        m_LastName = Left(strNames(2), 100)
        
        ' Check estabklished value.
        If m_DOB <> (.txtEstd.Text) Then
            m_DOB = Trim$(.txtEstd.Text)
            m_Modified = True
        End If
        ' Check the Profession.
        If StrComp(m_Profession, .txtInstHead.Text, vbTextCompare) <> 0 Then
            m_Profession = Trim$(.txtInstHead.Text)
            m_Modified = True
        End If
    End If
    
    
    ' Check the Home Address.
    If StrComp(m_HomeAddress, .HomeAddress, vbTextCompare) <> 0 Then
        m_HomeAddress = Trim$(.HomeAddress)
        m_Modified = True
    End If

    ' Check the Office Address.
    If StrComp(m_OfficeAddress, .OfficeAddress, vbTextCompare) <> 0 Then
        m_OfficeAddress = Trim$(.OfficeAddress)
        m_Modified = True
    End If

    ' Check the Home Telephone Number.
    If StrComp(m_HomePhone, .txtHomePhone.Text, vbTextCompare) <> 0 Then
        m_HomePhone = Trim$(.txtHomePhone)
        m_Modified = True
    End If

    ' Check the Office Telephone Number.
    If StrComp(m_OfficePhone, .txtOffPhone.Text, vbTextCompare) <> 0 Then
        m_OfficePhone = Trim$(.txtOffPhone)
        m_Modified = True
    End If

    ' Check the eMail text.
    If StrComp(m_eMail, .txtEmail.Text, vbTextCompare) <> 0 Then
        m_eMail = Trim$(.txtEmail.Text)
        m_Modified = True
    End If

End With

m_CustomerLoaded = True
'm_CustomerLoaded = False
'Get the CustomerID
If m_InsertCustomer And m_CustomerID <= 0 Then m_CustomerID = GetNewCutomerId

'm_CustomerID = GetNewCutomerId
'If Not SaveCustomer() Then
'    'MsgBox "Unable to save customer !", vbCritical, gAppName & " - Error"
'    MsgBox GetResourceString(617), vbCritical, gAppName & " - Error"
'    'gDBTrans.RollBack
'    Exit Sub
'End If
'm_CustomerLoaded = True
End Sub
Public Property Get CustomerLoaded() As Boolean
    CustomerLoaded = m_CustomerLoaded
End Property
Public Property Get IsNewCustomerLoaded() As Boolean
    IsNewCustomerLoaded = m_InsertCustomer
End Property

Private Sub RegDialog_windowclosed()
On Error Resume Next
Set RegDialog = Nothing
End Sub
