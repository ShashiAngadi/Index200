VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ===================================
' This is PDL used throughout the Class
' ===================================
' Handle Error
' On error Goto Hell:
' Declarations
' Validations
' Exit
' Hell:

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private WithEvents m_frmLedger As frmLedger
Attribute m_frmLedger.VB_VarHelpID = -1
'Private WithEvents m_frmLedger As frmLedger1

Private m_HeadID As Long
Private m_OpBalance As Currency
Private m_LedgerName As String
Private m_ParentID As Long
Private m_dbOperation As wis_DBOperation

' To Handle Grid Functions
Private m_GrdFunctions As clsGrdFunctions

Public Property Get ObjectName()
    ObjectName = "LedgerClass"
End Property


'This function Will fetch the Particulars of the TransID Given
Public Function GetTransIDParticulars(TransID As Long) As String

' Handle Error
On Error GoTo Hell:

GetTransIDParticulars = ""

' Declarations
Dim rstParticulars As ADODB.Recordset

If TransID = 0 Then Exit Function

gDbTrans.SqlStmt = " SELECT Particulars " & _
                   " FROM TransParticulars " & _
                   " WHERE TransID=" & TransID
                   
If gDbTrans.Fetch(rstParticulars, adOpenForwardOnly) < 1 Then Exit Function
                
GetTransIDParticulars = FormatField(rstParticulars.Fields("Particulars"))

Set rstParticulars = Nothing

Exit Function

Hell:

    MsgBox "GetTransIDParticulars : " & vbCrLf & Err.Description
    Set rstParticulars = Nothing
    
End Function
'This will return the VoucherName of the given VoucherType
Public Function GetVoucherName(VoucherType As Wis_VoucherTypes) As String

GetVoucherName = ""

Select Case VoucherType
    Case Sales
        GetVoucherName = "Sales"
    Case payment
        GetVoucherName = "Payment"
    Case Purchase
        GetVoucherName = "Purchase"
    Case Receipt
        GetVoucherName = "Receipt"
    Case CONTRA
        GetVoucherName = "Contra"
    Case FreePurchase
        GetVoucherName = "Free Purchase"
    Case FreeSales
        GetVoucherName = "Free Sales"
    Case Journal
        GetVoucherName = "Journal"
    Case RejectionsIn
        GetVoucherName = "RejectionsIn"
    Case RejectionsOut
        GetVoucherName = "RejectionsOut"
End Select

End Function

'This will load all Heads to the given Recordset from the Base Table

Public Sub LoadHeadstoRecordset(rstHeads As ADODB.Recordset)

' Handle Error
On Error GoTo Hell:

gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM Heads"
                   
Call gDbTrans.Fetch(rstHeads, adOpenDynamic)

Exit Sub

Hell:
    
    
End Sub

Private Sub SetHeaderFooter(ByVal Balance As Currency, grd As MSFlexGrid, AccountType As wis_AccountType, MsgString As String)

With grd

'    m_GrdFunctions.fMoreRows (2)
    
    '.Row = .Row + 2
    .Col = 1: .Text = MsgString: .CellFontBold = True
               
      Select Case AccountType
        Case Asset, Loss, ItemPurchase
        
            If Balance >= 0 Then .Col = 3: .CellForeColor = vbBlack
            If Balance < 0 Then .Col = 4: .CellForeColor = vbRed
        
        Case Liability, Profit, ItemSales
        
            If Balance >= 0 Then .Col = 4: .CellForeColor = vbBlack
            If Balance < 0 Then .Col = 3: .CellForeColor = vbRed

        Case Loss
        
            If Balance >= 0 Then .Col = 3: .CellForeColor = vbBlack
            If Balance < 0 Then .Col = 4: .CellForeColor = vbRed

            
        Case Profit
        
            If Balance >= 0 Then .Col = 4: .CellForeColor = vbBlack
            If Balance < 0 Then .Col = 3: .CellForeColor = vbRed

            
        Case ItemSales
        
            If Balance >= 0 Then .Col = 4: .CellForeColor = vbBlack
            If Balance < 0 Then .Col = 3: .CellForeColor = vbRed
            
        Case ItemPurchase
        
            If Balance >= 0 Then .Col = 3: .CellForeColor = vbBlack
            If Balance < 0 Then .Col = 4: .CellForeColor = vbRed

    End Select
               
    .Text = FormatCurrency(Abs(Balance)): .CellFontBold = True
    
End With

End Sub
Public Function IsHeadIDBalanceDefault(ByVal Balance As Currency, headID As Long) As Boolean

On Error GoTo Hell:

' Declarations
Dim AccountType As wis_AccountType
Dim AccTransClass As clsAccTrans

' get the AccountTYpe

Set AccTransClass = New clsAccTrans

AccountType = AccountType = AccTransClass.GetAccountType(headID)

Set AccTransClass = Nothing

' Here default is True
IsHeadIDBalanceDefault = True

' Now set the Variable

Select Case AccountType

    Case Asset
    
        If Balance < 0 Then IsHeadIDBalanceDefault = False
    
    Case Liability
        
        If Balance < 0 Then IsHeadIDBalanceDefault = False

    Case Loss
        
        If Balance < 0 Then IsHeadIDBalanceDefault = False
        
    Case Profit
    
        If Balance < 0 Then IsHeadIDBalanceDefault = False

    Case ItemSales
    
        If Balance < 0 Then IsHeadIDBalanceDefault = False
        
    Case ItemPurchase
    
        If Balance < 0 Then IsHeadIDBalanceDefault = False

End Select
               
Exit Function

Hell:
        
End Function

Private Function AddRecords() As wis_FunctionReturned

AddRecords = Failure

If Validated <> Success Then Exit Function

If m_dbOperation = Insert Then SaveLedger
If m_dbOperation = Update Then UpdateLedger
  

AddRecords = Success

End Function

Public Sub ShowLedger()

' Check the Form's Status
If m_frmLedger Is Nothing Then Set m_frmLedger = New frmLedger

m_dbOperation = Insert

Load m_frmLedger
If ParentID Then
    With m_frmLedger
        Dim count As Integer
        For count = 0 To .cmbParent.ListCount - 1
            If .cmbParent.ItemData(count) = m_ParentID Then
                .cmbParent.ListIndex = count
                Exit For
            End If
        Next
    End With
End If
m_frmLedger.Show vbModal


End Sub

Private Function SaveLedger() As wis_FunctionReturned

On Error GoTo NoSaveError:

Dim rstHeads As ADODB.Recordset

Dim curOpBalance As Currency
Dim headID As Long
Dim OpAmericanDate As String

SaveLedger = Failure

' check the form's status
If m_frmLedger Is Nothing Then Exit Function

With m_frmLedger

    If .cmbParent.ListIndex = -1 Then Exit Function
    
    OpAmericanDate = GetSysFormatDate(FinIndianFromDate)
    curOpBalance = CCur(Val(.txtOpBalance.Text))
    If .chkNegBal = vbChecked Then curOpBalance = curOpBalance * -1
    
    SaveLedger = False
    
    ParentID = .cmbParent.ItemData(.cmbParent.ListIndex)
    
    'Get the Maximum Head From the database
    gDbTrans.SqlStmt = " SELECT MAX(HeadID) FROM Heads " & _
                       " WHERE ParentID = " & ParentID
    
    Call gDbTrans.Fetch(rstHeads, adOpenForwardOnly)
    
    headID = FormatField(rstHeads.Fields(0)) + 1
    
    If headID < ParentID Then headID = headID + ParentID
    If gLangOffSet = 0 Then .txtLedgerEnglish.Text = .txtLedgerName.Text
    'Insert the heads inot the database
    gDbTrans.SqlStmt = " INSERT INTO Heads (HeadID,HeadName,ParentID,HeadNameEnglish,IsContraHead) " & _
                       " VALUES ( " & _
                       headID & "," & _
                       AddQuotes(.txtLedgerName.Text, True) & "," & _
                       ParentID & "," & _
                       AddQuotes(.txtLedgerEnglish.Text, True) & "," & _
                       .chkContra.Value & ")"
    
    gDbTrans.BeginTrans
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    'Insert the opening balance
    gDbTrans.SqlStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                     " VALUES ( " & _
                     headID & "," & _
                     "#" & OpAmericanDate & "#," & _
                     curOpBalance & ")"
                     
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.CommitTrans
    
End With

SaveLedger = Success

' Clear the Controls

ClearControls
Set rstHeads = Nothing

Exit Function

NoSaveError:
        
    ' Clear up the  Transactions if any
    Set rstHeads = Nothing
    SaveLedger = FatalError

End Function


' This function will checks the HeadID & parentID in the database
' If the headID is not avaialble in the heads Table it will create the
' Head with given HeadID & parentID
' Inputs :
'           HeadID as long
'           ParentID as long
' OutPut :
Public Function GetHeadIDCreated(ByVal ParentID As Long, ByVal headName As String, ByVal opBalance As Currency, ByVal headNameEnglish) As Long

'Trap an error
On Error GoTo ErrLine

'Declare the variables
Dim rstHeads As ADODB.Recordset
Dim UsOpDate As String
Dim headID As Long
Dim HeadNameInEnglish As String

HeadNameInEnglish = headNameEnglish
'initialise the function
GetHeadIDCreated = -1

'Validate the inputs
If ParentID = 0 Then Exit Function
If headName = "" Then Exit Function

'Check the given Heads in the database
gDbTrans.SqlStmt = " SELECT HeadID,HeadName,HeadNameEnglish" & _
                   " FROM Heads " & _
                   " WHERE HeadName = " & AddQuotes(headName) & _
                   " AND ParentID = " & ParentID
                
'if exists then exit function
If gDbTrans.Fetch(rstHeads, adOpenForwardOnly) > 0 Then
    GetHeadIDCreated = FormatField(rstHeads.Fields(0))
    If Len(FormatField(rstHeads.Fields("HeadNameEnglish"))) < 1 Then
        gDbTrans.SqlStmt = "UPDATE Heads  SET [HeadNameEnglish] =" & AddQuotes(HeadNameInEnglish, True) & _
                      " WHERE HeadID =  " & FormatField(rstHeads.Fields(0))
        Call gDbTrans.SQLExecute
    End If
    Exit Function
End If

'Get the Maximum Head From the database
gDbTrans.SqlStmt = " SELECT MAX(HeadID)" & _
                   " FROM Heads " & _
                   " WHERE ParentID = " & ParentID

Call gDbTrans.Fetch(rstHeads, adOpenForwardOnly)

headID = FormatField(rstHeads.Fields(0)) + 1
    
If headID < ParentID Then headID = headID + ParentID

' Else create the head with given headID & parentiD
'Insert the heads into the database
gDbTrans.SqlStmt = " INSERT INTO Heads (HeadID,HeadName,ParentID,HeadNameEnglish) " & _
                  " VALUES ( " & _
                  headID & "," & _
                  AddQuotes(headName) & "," & _
                  ParentID & "," & _
                  AddQuotes(HeadNameInEnglish) & ")"

If Not gDbTrans.SQLExecute Then Exit Function


UsOpDate = GetSysFormatDate(FinIndianFromDate)
'Insert the opening balance
gDbTrans.SqlStmt = " INSERT INTO OpBalance (HeadID,OpDate,opAmount) " & _
                 " VALUES ( " & _
                 headID & "," & _
                 "#" & UsOpDate & "#," & _
                 opBalance & ")"
If Not gDbTrans.SQLExecute Then Exit Function

GetHeadIDCreated = headID

Set rstHeads = Nothing

Exit Function

ErrLine:
    MsgBox "CreateHead: " & vbCrLf & Err.Description, vbCritical, wis_MESSAGE_TITLE
    
    GetHeadIDCreated = -1
    
    Set rstHeads = Nothing
    
    Exit Function
        
    
End Function


Private Function UpdateLedger() As wis_FunctionReturned

On Error GoTo NoUpdateError:

Dim curOpBalance As Currency
Dim LedgerName As String
Dim rstHead As Recordset

Dim OpAmericanDate As String

UpdateLedger = Failure

'check the form's status

If m_frmLedger Is Nothing Then Exit Function

With m_frmLedger

    If .cmbParent.ListIndex = -1 Then Exit Function
    
    OpAmericanDate = GetSysFormatDate(FinIndianFromDate)
    curOpBalance = CCur(Val(.txtOpBalance.Text))
    If .chkNegBal = vbChecked Then curOpBalance = curOpBalance * -1
    
    LedgerName = Trim$(.txtLedgerName.Text)
    ParentID = .cmbParent.ItemData(.cmbParent.ListIndex)
    
    gDbTrans.SqlStmt = " UPDATE Heads SET " & _
                       " HeadName = " & AddQuotes(LedgerName, True) & "," & _
                       " HeadNameEnglish = " & AddQuotes(Trim$(.txtLedgerEnglish.Text), True) & "," & _
                       " IsContraHead = " & .chkContra.Value & _
                       " WHERE HeadID = " & m_HeadID
    
    gDbTrans.BeginTrans
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    'Insert the opening balance
    gDbTrans.SqlStmt = " UPDATE OpBalance SET " & _
                       " OpAmount = " & curOpBalance & _
                       " WHERE OpDate = #" & OpAmericanDate & "#" & _
                       " AND HeadID = " & m_HeadID
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    'If it is head of bank ledgers like SB,CA,FD then
    'update the same in the bankindexHeaID
    gDbTrans.SqlStmt = " UPDATE BankHeadIds SET " & _
                       " AliasName = " & AddQuotes(LedgerName, True) & "," & _
                       " HeadNameEnglish = " & AddQuotes(Trim$(.txtLedgerEnglish.Text), True) & _
                       " WHERE HeadID = " & m_HeadID
    
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    
    gDbTrans.CommitTrans
    
End With

UpdateLedger = Success

'Clear the controls
ClearControls

Exit Function

NoUpdateError:

    UpdateLedger = FatalError

End Function

Private Sub ClearControls()

' check the form's status
If m_frmLedger Is Nothing Then Exit Sub

With m_frmLedger

    '.cmbParent.ListIndex = -1
    .txtLedgerName.Text = ""
    .txtLedgerEnglish.Text = ""
    .txtOpBalance.Text = ""
    .chkNegBal.Value = vbUnchecked
    .chkContra.Value = vbUnchecked
    .cmdClear.Enabled = False
    '.lvwLedger.ColumnHeaders.Clear
    .cmdOk.Caption = GetResourceString(10)
    .cmbParent.Locked = False
    Call LoadHeadsToListView(ParentID, FinIndianFromDate)
    ActivateTextBox .txtLedgerName
End With

m_dbOperation = Insert

End Sub

Private Sub LoadHeadsToListView(ByVal ParentID As Long, ByVal AsOnIndainDate As String)

Dim rstHeads As ADODB.Recordset

' check the form's status
If m_frmLedger Is Nothing Then Exit Sub

If ParentID = 0 Then Exit Sub
If Not DateValidate(AsOnIndainDate, "/", True) Then Exit Sub

m_frmLedger.lvwLedger.ListItems.Clear

gDbTrans.SqlStmt = " SELECT a.HeadID,HeadName,OpAmount,HeadNameEnglish,IsContraHead " & _
                   " FROM Heads a,OpBalance b " & _
                   " WHERE a.ParentID =  " & ParentID & _
                   " AND a.HeadID=b.HeadID" & _
                   " AND b.OpDate=" & "#" & FinUSFromDate & "#"  'GetSysFormatDate(AsOnIndainDate)
                   
If gDbTrans.Fetch(rstHeads, adOpenForwardOnly) < 1 Then Exit Sub

If Not FillViewNew(m_frmLedger.lvwLedger, rstHeads, "HeadID", True) Then Exit Sub

Set rstHeads = Nothing

End Sub

Private Function Validated() As wis_FunctionReturned

On Error GoTo ValidError:

Dim rst As ADODB.Recordset

Validated = Failure

' Check the Form's Status
If m_frmLedger Is Nothing Then Exit Function

With m_frmLedger

    If .cmbParent.ListIndex = -1 Then Err.Raise vbObjectError + 513, , "Select Parent Name "
    
    If Not TextBoxCurrencyValidate(.txtOpBalance, True, True) Then Err.Raise vbObjectError + 513, , "Invalid Opening Balance Specified"
    
    Validated = Success
    
    If m_dbOperation = Update Then Exit Function
    
    Validated = Failure
    
    'Check for the existence of the head
    gDbTrans.SqlStmt = " SELECT HeadName FROM Heads " & _
                       " WHERE HeadName = " & AddQuotes(.txtLedgerName.Text, True) & _
                       " AND ParentID = " & .cmbParent.ItemData(.cmbParent.ListIndex)
    
    If gDbTrans.Fetch(rst, adOpenForwardOnly) > 0 Then Err.Raise vbObjectError + 513, , "Duplicate Entry "
        
    Set rst = Nothing

End With

Validated = Success
Set rst = Nothing

Exit Function

ValidError:
    
    MsgBox "Ledger Creation :" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
    Validated = FatalError
    Set rst = Nothing
   
End Function

Private Sub Class_Initialize()

Set m_GrdFunctions = New clsGrdFunctions

End Sub

Private Sub Class_Terminate()

On Error Resume Next

Set m_frmLedger = Nothing
Set m_GrdFunctions = Nothing

End Sub


Private Sub m_frmLedger_CancelClick()

On Error Resume Next

Set m_frmLedger = Nothing
    
End Sub

Public Property Get ObjectHandle() As Long
    If Not m_frmLedger Is Nothing Then _
                ObjectHandle = m_frmLedger.hwnd
End Property

Private Sub m_frmLedger_ClearClick()
    Call ClearControls
End Sub

Private Sub m_frmLedger_LookupClick(ParentID As Long)

Call LoadHeadsToListView(ParentID, FinIndianFromDate)
    
End Sub
Private Sub m_frmLedger_LvwLedgerClick(headID As Long)

m_HeadID = headID

Dim rstHeads As Recordset

gDbTrans.SqlStmt = " SELECT a.HeadID,HeadName,OpAmount,HeadNameEnglish,IsContraHead " & _
                   " FROM Heads a,OpBalance b " & _
                   " WHERE a.HeadID =  " & headID & _
                   " AND a.HeadID=b.HeadID" & _
                   " AND b.OpDate=" & "#" & FinUSFromDate & "#"
                   
If gDbTrans.Fetch(rstHeads, adOpenForwardOnly) < 1 Then Exit Sub

With m_frmLedger
    
    LedgerName = FormatField(rstHeads("HeadName")) '.txtLedgerName.Text
    opBalance = Abs(FormatField(rstHeads("OpAmount"))) 'Val(.txtOpBalance.Text)
    .chkNegBal.Value = vbUnchecked
    If FormatField(rstHeads("OpAmount")) < 1 Then .chkNegBal.Value = vbChecked
    .chkContra.Value = vbUnchecked
    If FormatField(rstHeads("IsContraHead")) = 1 Then .chkContra.Value = vbChecked
    m_dbOperation = Update
    .cmdClear.Enabled = True
End With

End Sub

Private Sub m_frmLedger_OKClick()

Call AddRecords

End Sub



Private Sub m_frmLookUp_SelectClick(strSelection As String)

m_HeadID = CLng(strSelection)

m_dbOperation = Update

End Sub
Public Property Get opBalance() As Currency
 opBalance = m_OpBalance
End Property

Public Property Let opBalance(ByVal NewValue As Currency)
    m_OpBalance = NewValue
End Property

Public Property Get LedgerName() As String
    LedgerName = m_LedgerName
End Property

Public Property Let LedgerName(ByVal NewValue As String)
    m_LedgerName = NewValue
End Property

Public Property Get ParentID() As Long
    ParentID = m_ParentID
End Property

Public Property Let ParentID(ByVal NewValue As Long)
    m_ParentID = NewValue
End Property
