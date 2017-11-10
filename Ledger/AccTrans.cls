VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAccTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CrHeads"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"DrHeads"
Option Explicit

Private m_tmpTest As String

Private m_Forced As Boolean

Private m_frmAccTrans As frmAccTrans
Private m_AllTrans As AllTrans
Private m_GrdFunctions As clsGrdFunctions

' Local Variable(s) to Hold Property Value(s)
Private M_SlNo As Integer

Private m_HeadName As String
Private m_TransDate As String

Private m_Debit As Currency
Private m_Credit As Currency
Private m_Balance As Currency
Private m_HeadID As Long
Private m_PrevHeadID As Long
Private m_TransID As Long

Private m_isUpdatingGrid As Boolean
Private m_UpdatingTrans As Boolean

Public Enum wis_LoadLedgers
    LoadALL = 0
    LoadExceptCash = 1
    LoadOnlyCash = 2
    LoadCashCreditros = 4
End Enum

' For the Combos
Private Const VoucherCombo = 0
Private Const ParentCombo = 1
Private Const LedgerCombo = 2
'Sda 17/12/02
Private Const TransGrid = 0
Private Const LedgerGrid = 1

' For the Options
Private Const FromOpt = 0
Private Const ToOpt = 1

' For the Add Commands
Private Const AddToGrid = 0
Private Const DeleteGrid = 1

'
Private Sub CheckTransId()

Dim rstTrans As ADODB.Recordset

gDbTrans.SqlStmt = " SELECT Sum(Debit),Sum(Credit)" & _
                   " FROM AccTrans " & _
                   " WHERE TransID = " & TransID
                 
Call gDbTrans.Fetch(rstTrans, adOpenForwardOnly)

If FormatField(rstTrans(0)) <> FormatField(rstTrans(1)) Then _
    MsgBox "CheckTransID :" & TransID & vbCrLf & "Total Not Same !!!"
    
End Sub

' This will Convert Existing  ParentHeads data to New data which is in the Parents Table
Private Function ConvertParentIDtoNewParentID() As Boolean

On Error GoTo Hell:

ConvertParentIDtoNewParentID = False

Dim rstParents As ADODB.Recordset

Dim Parent As Long
Dim ParentID As Long
Dim NewParentID As Long

gDbTrans.SqlStmt = " SELECT ParentID,Parent,NewParentID " & _
                   " FROM Parents"

Call gDbTrans.Fetch(rstParents, adOpenForwardOnly)

With rstParents

    Do While Not .EOF
        
        Parent = .Fields("Parent")
        ParentID = .Fields("ParentID")
        NewParentID = .Fields("NewParentID")
        
        gDbTrans.SqlStmt = " UPDATE ParentHeads " & _
                         " SET ParentID=" & NewParentID & "," & _
                         " Parent = " & Parent & _
                         " WHERE ParentID = " & ParentID
                         
        If Not gDbTrans.SQLExecute Then Exit Function
                
        gDbTrans.SqlStmt = "UPDATE Heads " & _
                         " SET ParentID = " & NewParentID & _
                         " WHERE ParentID = " & ParentID
                         
        If Not gDbTrans.SQLExecute Then Exit Function
                        
        .MoveNext
    Loop
    
End With

ConvertParentIDtoNewParentID = True

Exit Function

Hell:
    ConvertParentIDtoNewParentID = False
    
End Function
' This will Convert Existing  HeadsID data to New Heads ID
Private Function ConvertHeadsIDtoNewHeadsID() As Boolean

On Error GoTo Hell:

ConvertHeadsIDtoNewHeadsID = False

Dim rstHeads As ADODB.Recordset

Dim headID As Long
Dim NewHeadID As Long

gDbTrans.SqlStmt = "SELECT HeadID" & _
                  " FROM Heads" & _
                  " ORDER BY ParentID"

Call gDbTrans.Fetch(rstHeads, adOpenForwardOnly)

NewHeadID = 5

With rstHeads

    Do While Not .EOF
        
        headID = .Fields("HeadID")
        NewHeadID = NewHeadID + 1
        
        gDbTrans.SqlStmt = "UPDATE Heads" & _
                           " SET HeadID = " & NewHeadID & _
                           " WHERE HeadID = " & headID
        If Not gDbTrans.SQLExecute Then Exit Function
                
        gDbTrans.SqlStmt = "UPDATE AccTrans " & _
                         " SET HeadID=" & NewHeadID & _
                         " WHERE HeadID=" & headID
                         
        If Not gDbTrans.SQLExecute Then Exit Function
        
        
        gDbTrans.SqlStmt = "UPDATE CompanyCreation" & _
                         " SET HeadID = " & NewHeadID & _
                         " WHERE HeadID = " & headID
                         
        If Not gDbTrans.SQLExecute Then Exit Function
        
        
        gDbTrans.SqlStmt = " UPDATE InvoiceDetails" & _
                         " SET HeadID=" & NewHeadID & _
                         " WHERE HeadID=" & headID
                         
        If Not gDbTrans.SQLExecute Then Exit Function
               
        
        gDbTrans.SqlStmt = " UPDATE OpBalance " & _
                         " SET HeadID=" & NewHeadID & _
                         " WHERE HeadID=" & headID
                         
        If Not gDbTrans.SQLExecute Then gDbTrans.RollBack
            
        gDbTrans.SqlStmt = " UPDATE Purchase " & _
                         " SET HeadID=" & NewHeadID & _
                         " WHERE HeadID=" & headID
                         
        If Not gDbTrans.SQLExecute Then Exit Function
        
                
        gDbTrans.SqlStmt = " UPDATE RelationMaster " & _
                         " SET HeadID=" & NewHeadID & _
                         " WHERE HeadID=" & headID
                         
        If Not gDbTrans.SQLExecute Then Exit Function
                        
        gDbTrans.SqlStmt = " UPDATE Sales " & _
                         " SET HeadID=" & NewHeadID & _
                         " WHERE HeadID=" & headID
                         
        If Not gDbTrans.SQLExecute Then Exit Function
                
        .MoveNext
    Loop
    
End With

ConvertHeadsIDtoNewHeadsID = True

Exit Function

Hell:
    ConvertHeadsIDtoNewHeadsID = False
    
End Function
Public Property Let ForcedTransaction(NewBool As Boolean)
'This transaction is for the transaction
'which will take from the Bank Class
'Beacause some time we need to insert the Negative amount
'in such case if forcedtransaction is set to true then
'it allow to to the transaction with negative amount
    m_Forced = NewBool
End Property

'
Public Sub frmClearClicked()

If m_frmAccTrans Is Nothing Then Exit Sub

ClearCollections
    
ClearAccTransControls

' With the Collections

headID = 0
SlNo = 0
Debit = 0
Credit = 0
UpdatingTrans = False

End Sub
' This function will delete the transaction from the Table
' It takes no arguments
' Retruns true if success
'
Public Function frmDeleteIDClicked() As Boolean

frmDeleteIDClicked = False

' Validations
If m_frmAccTrans Is Nothing Then Exit Function
If m_frmAccTrans.DBOperation <> DeleteRec Then Exit Function
If TransID = 0 Then Exit Function

If gCurrUser.UserPermissions And perBankAdmin = 0 Then Exit Function

If MsgBox("Do You Want to Delete the Transaction !", vbQuestion + vbYesNo) = vbNo Then Exit Function

gDbTrans.BeginTrans
If DeleteTransID <> Success Then gDbTrans.RollBack: Exit Function
gDbTrans.CommitTrans

' Clear all the COntrols in the form
ClearAccTransControls

' Clear the COllections
ClearCollections

headID = 0
SlNo = 0
Debit = 0
Credit = 0
TransID = 0

frmDeleteIDClicked = True

End Function
Public Sub frmDeleteClicked()

If MsgBox("Do You want to delete the Current Entry !", vbQuestion + vbYesNo) = vbNo Then Exit Sub

Call DeleteProcedures

End Sub

'This handle the option buttons  for the further voucher entries
Private Sub FurtherVoucherEntries(ByVal DebitAmount As Currency, ByVal CreditAmount As Currency)

Dim voucherTypes As Wis_VoucherTypes

' Validations
If m_frmAccTrans Is Nothing Then Exit Sub

' This Will Fetch The Current Voucher Type Selected by the User
With m_frmAccTrans

    Call LoadParentHeads(.cmbTab0Ledger(ParentCombo))
    
    .optTab0From(FromOpt).Enabled = True
    .optTab0From(ToOpt).Enabled = True
    
    With .cmbTab0Ledger(VoucherCombo)
       voucherTypes = .ItemData(.ListIndex)
       If voucherTypes = Journal Then Call LoadParentHeadsForVoucher(m_frmAccTrans.cmbTab0Ledger(ParentCombo), LoadExceptCash)
    End With
    
    If DebitAmount > CreditAmount Then
        .optTab0From(FromOpt).Value = (voucherTypes <> Journal)
    Else
        .optTab0From(ToOpt).Value = True
    End If
    
        
End With

End Sub

'
Public Function GetVoucherNametoArray()

Dim VoucherName(0 To 10) As String
Dim Vouchers As Wis_VoucherTypes
Dim Item As Integer

'GetVoucherNametoArray = VoucherName
'ReDim GetVoucherNametoArray(0 To 10)

Vouchers = CONTRA
VoucherName(Vouchers) = "Contra"
Vouchers = FreePurchase
VoucherName(Vouchers) = "Free Purchase"
Vouchers = FreeSales
VoucherName(Vouchers) = "Free Sales"
Vouchers = Journal
VoucherName(Vouchers) = "Journal"
Vouchers = payment
VoucherName(Vouchers) = "Payment"
Vouchers = Purchase
VoucherName(Vouchers) = "Purchase"
Vouchers = Receipt
VoucherName(Vouchers) = "Receipt"
Vouchers = RejectionsIn
VoucherName(Vouchers) = "Rejections In"
Vouchers = RejectionsOut
VoucherName(Vouchers) = "Rejections Out"
Vouchers = Sales
VoucherName(Vouchers) = "Sales"
Vouchers = VouNothing
VoucherName(Vouchers) = "Nothing"

GetVoucherNametoArray = VoucherName(Vouchers)


End Function

' When in the frmAccTrans user clicks on the Add comand
' This Procedures is Called
Private Function AddProcedures() As wis_FunctionReturned

' Handle Error
On Error GoTo Hell:

' Declarations
Dim VoucherType As Wis_VoucherTypes

AddProcedures = Failure

' Validate the Controls
If m_frmAccTrans Is Nothing Then Exit Function

If ValidateLedgers <> Success Then Exit Function

' Check if the HeadID Exists the Collection
If ALlTransHeadIDExists(headID) Then Err.Raise vbObjectError + 513, , "Head Already Selected!"

' Update the Grid
If UpDateGrids <> Success Then Exit Function

' Update the Collections
If UpdateCollections <> Success Then Exit Function

Call PrepareAccTrans

AddProcedures = Success

Exit Function

Hell:
    
    MsgBox "AddProcedures() " & Err.Description
    AddProcedures = FatalError
    
End Function

' Here starts the debit entries
' Start the Loop and start inserting the Variables
Private Function AllTransCRInsertToTable(ByVal VoucherType As Wis_VoucherTypes, _
                                      ByVal USTransDate As Date _
                                      ) As wis_FunctionReturned

On Error GoTo Hell:

Dim ObjKeys As Variant
    
Dim MaxCount As Integer
Dim Item As Integer

AllTransCRInsertToTable = Failure

MaxCount = AllTrans.count
ObjKeys = AllTrans.GetKeys

For Item = 0 To MaxCount - 1
    With AllTrans(ObjKeys(Item))
        If .CrAmount > 0 Or (m_Forced And .CrAmount < 0) Then
            ' Insert the Variables
            gDbTrans.SqlStmt = "INSERT INTO AccTrans " & _
                     " (HeadID,TransID,VoucherType,TransDate,Debit,Credit)" & _
                     " VALUES (" & _
                     .headID & "," & _
                     TransID & "," & _
                     VoucherType & "," & _
                     "#" & USTransDate & "#," & _
                      .DrAmount & "," & _
                      .CrAmount & ")"
            If Not gDbTrans.SQLExecute Then Exit Function
            .DrAmount = 0
            
        End If
    End With
Next Item

AllTransCRInsertToTable = Success

Exit Function

Hell:
     
    AllTransCRInsertToTable = FatalError
    
End Function

' Here starts the debit entries
' Start the Loop and start inserting the Variables
Private Function AllTransDRInsertToTable(ByVal VoucherType As Wis_VoucherTypes, _
                                      ByVal USTransDate As String _
                                      ) As wis_FunctionReturned

On Error GoTo Hell:

Dim ObjKeys As Variant
    
Dim MaxCount As Integer
Dim Item As Integer

AllTransDRInsertToTable = Failure

MaxCount = AllTrans.count
ObjKeys = AllTrans.GetKeys

For Item = 0 To MaxCount - 1
    With AllTrans(ObjKeys(Item))
        If .DrAmount > 0 Or (m_Forced And .DrAmount < 0) Then
            'Insert the Variables
            gDbTrans.SqlStmt = "INSERT INTO AccTrans " & _
                     " (HeadID,TransID,VoucherType,TransDate,Debit,Credit)" & _
                     " VALUES (" & _
                     .headID & "," & _
                     TransID & "," & _
                     VoucherType & "," & _
                     "#" & USTransDate & "#," & _
                      .DrAmount & "," & _
                      .CrAmount & ")"
            If Not gDbTrans.SQLExecute Then Exit Function
            .CrAmount = 0
        End If
    End With
Next Item


AllTransDRInsertToTable = Success

Exit Function

Hell:
     
    AllTransDRInsertToTable = FatalError
    
End Function
Private Function IsCashTransaction() As wis_FunctionReturned
IsCashTransaction = Failure

On Error GoTo Hell

Dim VoucherType As Wis_VoucherTypes

With m_frmAccTrans
    Dim I As Integer
    I = .cmbTab0Ledger(VoucherCombo).ListIndex
    VoucherType = .cmbTab0Ledger(VoucherCombo).ItemData(I)
End With

'IF Voucher Type is Contra Then Exit From This Function
If VoucherType = Journal Then Exit Function
If (gCurrUser.UserPermissions And perOnlyWaves) Then Exit Function

'If VoucherType = Contra Then Exit Function

'Validate the Ledger
'If ValidateLedgers <> Success Then Exit Function

IsCashTransaction = Success
Exit Function

Hell:
IsCashTransaction = FatalError
End Function

Private Function IsAnyTransaction() As wis_FunctionReturned
IsAnyTransaction = Failure

On Error GoTo Hell

Dim VoucherType As Wis_VoucherTypes

With m_frmAccTrans
    Dim I As Integer
    I = .cmbTab0Ledger(VoucherCombo).ListIndex
    VoucherType = .cmbTab0Ledger(VoucherCombo).ItemData(I)
End With

'Validate the Ledger
If ValidateLedgers <> Success Then Exit Function

IsAnyTransaction = Success
Exit Function

Hell:
IsAnyTransaction = FatalError
End Function


Private Function IsTransIDFromInventory(ByVal TransID As Long) As Boolean

On Error GoTo Hell:

Dim rstID As ADODB.Recordset

IsTransIDFromInventory = False

gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM IDFromInventory " & _
                   " WHERE TransID=" & TransID

If gDbTrans.Fetch(rstID, adOpenForwardOnly) < 1 Then Exit Function

IsTransIDFromInventory = True

Set rstID = Nothing

Exit Function

Hell:

End Function

'This Function Checks whether the given Tranasction id
'related to the bank transaction
Private Function IsBankHeadTransaction(ByVal TransactionID As Long) As Boolean

IsBankHeadTransaction = False

Dim rstTrans As Recordset

gDbTrans.SqlStmt = "SELECT * FROM AccTrans " & _
            " WHERE TransID =  " & TransactionID & _
            " And HeadID <> " & wis_CashHeadID

If gDbTrans.Fetch(rstTrans, adOpenForwardOnly) < 1 Then
    IsBankHeadTransaction = False
    Set rstTrans = Nothing
    Exit Function
End If
    
Dim rstTemp As Recordset
While Not rstTrans.EOF
    gDbTrans.SqlStmt = "SELECT * FRom ParentHeads " & _
            " WHERE ParentID = (SELECT ParentID " & _
                " From Heads Where HeadID = " & rstTrans("HeadID") & " )"
    Call gDbTrans.Fetch(rstTemp, adOpenDynamic)
    If FormatField(rstTemp("UserCreated")) > 2 Then
        Set rstTrans = Nothing
        Set rstTemp = Nothing
        IsBankHeadTransaction = True
        Exit Function
    End If
    rstTrans.MoveNext
Wend

Set rstTemp = Nothing

End Function



' This sub is check all transid and its treatment,along with vouchers
' This sub checks the transactions with current validations
Public Sub MyExternalValidations()

Dim rstTrans As ADODB.Recordset
Dim VoucherType As Wis_VoucherTypes

gDbTrans.SqlStmt = " SELECT DISTINCT(TransID),VoucherType" & _
                    " FROM AccTrans"
       
Call gDbTrans.Fetch(rstTrans, adOpenForwardOnly)

Do While Not rstTrans.EOF
    
    Call MyLoadtoCollection(rstTrans.Fields("TransID"))
    
    VoucherType = rstTrans.Fields("VoucherType")
    
    If AllTransVoucherCheck(VoucherType) <> Success Then Debug.Print rstTrans.Fields("TransID")
    
    ClearCollections
    
    rstTrans.MoveNext
    
Loop


End Sub

Private Sub MyLoadtoCollection(ByVal lngTransID As Long)

Dim rstTrans As ADODB.Recordset

If lngTransID = 0 Then Exit Sub '

' Get the Records to Recordset
gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM AccTrans " & _
                   " WHERE TransID=" & lngTransID
                   
If gDbTrans.Fetch(rstTrans, adOpenDynamic) < 1 Then Exit Sub

With rstTrans

    ' Start the Loop
    Do While Not .EOF
    
        'Feed data to Controls
        headID = FormatField(.Fields("HeadID"))
        Debit = FormatField(.Fields("Debit"))
        Credit = FormatField(.Fields("Credit"))
        TransDate = FormatField(.Fields("TransDate"))
        
        Balance = GetBalance(headID, .Fields(TransDate), .Fields(TransID))
        
        'Update the Collections
        If UpdateCollections <> Success Then Exit Sub
        
        .MoveNext
        
    Loop
End With

End Sub


Public Function NewDbConvertions() As Boolean

NewDbConvertions = False

If Not ConvertParentIDtoNewParentID Then Exit Function
If Not ConvertHeadsIDtoNewHeadsID Then Exit Function

NewDbConvertions = True

End Function
Public Property Get ObjectName()
    ObjectName = "AccTrans"
End Property

Private Sub PrepareAccTrans()

With m_frmAccTrans
    '.txtTab0Amount.Text = FormatCurrency(DrCrTotalsSame(True))
    .txtTab0Amount.Value = DrCrTotalsSame(True)
    
    ' This will check the cash id exists if not by default adds to combo box
    If ALlTransHeadIDExists(wis_CashHeadID) Then
       ' if added then simply clear the combobox
        .cmbTab0Ledger(LedgerCombo).Clear
    Else
        SetListIndextoCash
    End If
    
    .cmbTab0Ledger(VoucherCombo).Locked = True
    .txtTab0CurrentDate.Locked = True
    .cmdTab0Date.Enabled = False
    
    If .cmbTab0Ledger(VoucherCombo).ListIndex = 2 Then  'Contra Transcation
        
    End If
    '.cmbTab0Ledger(ParentCombo).SetFocus
    '.grdLedger(TransGrid).Enabled = True
    '.cmdTab0AddtoGrid(AddToGrid).Caption = "Add"
    .cmdTab0AddtoGrid(AddToGrid).Caption = GetResourceString(10) 'Add
    .cmdTab0AddtoGrid(DeleteGrid).Enabled = False
    
    If .txtTab0Amount.Value = 0 Then
        If AllTrans.count > 0 Then
            .cmdOk.Enabled = True
            .txtParticulars.SetFocus
            .cmdTab0AddtoGrid(AddToGrid).Enabled = False
        End If
    End If
End With

End Sub

' This sub will Set the Combo's List Index to Cash ID
' So that payment voucher entry made easy
Private Sub SetListIndextoCash()

Dim voucherTypes As Wis_VoucherTypes
Dim lngIndex As Long

' Validations
If m_frmAccTrans Is Nothing Then Exit Sub
' This Will Fetch The Current Voucher Type Selected by the User
With m_frmAccTrans
    With .cmbTab0Ledger(ParentCombo)
        For lngIndex = 0 To .ListCount - 1
            If wis_CashParentID = .ItemData(lngIndex) Then
                .ListIndex = lngIndex
                Exit For
            End If
        Next lngIndex
    End With
End With

End Sub


' This will handle the ParentHeads after the Each Add
Private Sub HandleVouchersAfterAdd(ByVal DebitAmout As Currency, ByVal CreditAmount As Currency)

If SlNo = 1 Then
    If m_frmAccTrans.DBOperation = Insert Then AfterFirstVoucherEntry
    If m_frmAccTrans.DBOperation = Update Then HandleVouchers
ElseIf SlNo > 1 Then
    Call FurtherVoucherEntries(DebitAmout, CreditAmount)
End If

End Sub

'This will handle the ParentHeads after the Each Add
Private Sub AfterFirstVoucherEntry()

Dim voucherTypes As Wis_VoucherTypes
Dim WavesUser As Boolean
WavesUser = (gCurrUser.UserPermissions And perOnlyWaves)
' Validations
If m_frmAccTrans Is Nothing Then Exit Sub

'This Will Fetch The Current Voucher Type Selected by the User
With m_frmAccTrans

''''''''''''''''
''Changed on April 13, 2014 for testing, can be deleted later
If .cmbTab0Ledger(VoucherCombo).ListIndex = 2 Then WavesUser = True
''''''''''''''''

    '.optTab0From(FromOpt).Enabled = True
    '.optTab0From(ToOpt).Enabled = True
    'Modified on 8/11/2004
    .optTab0From(FromOpt).Enabled = WavesUser
    .optTab0From(ToOpt).Enabled = WavesUser
    
    With .cmbTab0Ledger(VoucherCombo)
       voucherTypes = .ItemData(.ListIndex)
    End With
        
    ' Please Don't change the case creiteria
    'Altered according to Index 2000  'Sahshi 4/1/2003
    Select Case voucherTypes
        Dim LedgerType As wis_LoadLedgers
        Case payment
            .optTab0From(FromOpt).Value = True
            LedgerType = LoadALL
            Call LoadParentHeadsForVoucher(.cmbTab0Ledger(ParentCombo), LedgerType)
        
        Case Receipt
            .optTab0From(ToOpt).Value = True
            LedgerType = LoadALL
            Call LoadParentHeadsForVoucher(.cmbTab0Ledger(ParentCombo), LedgerType)
        
        Case CONTRA
            .optTab0From(ToOpt).Value = True
            LedgerType = LoadOnlyCash
            Call LoadParentHeadsForVoucher(.cmbTab0Ledger(ParentCombo), LedgerType)
        
        Case Journal
            .optTab0From(ToOpt).Value = True
            LedgerType = LoadExceptCash
            Call LoadParentHeadsForVoucher(.cmbTab0Ledger(ParentCombo), LedgerType)
        
    End Select

End With

End Sub

Public Property Let TempTest(NewValue As String)
    m_tmpTest = NewValue
End Property

'This will get the Transdate from collection
' and will check the amount
' if exists it will add the amount
Private Function TransDateDataExists(ByVal USTransDate As Date) As Long
' Handle Error
On Error GoTo Hell:

' Declare Variables
Dim rstTrans As ADODB.Recordset
Dim FromHeadID As Long
Dim ToHeadID As Long
Dim ContrORCashId As Long
Dim TransID As Long

Dim DrTotal As Currency
Dim CrTotal As Currency

Dim Item As Long
Dim MaxCount As Long

Dim ObjKeys As Variant

' Intialise the Function
TransDateDataExists = 0

If AllTrans Is Nothing Then Exit Function

' Check the Collection Dr Amount and Cr Amount Total

MaxCount = AllTrans.count
ObjKeys = AllTrans.GetKeys

For Item = 0 To MaxCount - 1
    With AllTrans(ObjKeys(Item))
        If .DrAmount <> 0 Then ToHeadID = .headID
        If .CrAmount <> 0 Then FromHeadID = .headID
    End With
    
Next Item

'SELECT A.TransID, A.HeadID, A.Debit, A.Credit,
'B.HeadID, B.Debit, B.Credit, B.TransDate, B.VoucherType
'FROM AccTrans AS A INNER JOIN AccTrans AS B ON A.TransID = B.TransID;

gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM QryAccTransMerge" & _
                   " WHERE TransDate = #" & USTransDate & "#" & _
                   " AND A.HEADID = " & ToHeadID & _
                   " AND B.HEADID = " & FromHeadID
 
If gDbTrans.Fetch(rstTrans, adOpenStatic) < 1 Then Exit Function

DrTotal = rstTrans.Fields("a.Debit").Value
CrTotal = rstTrans.Fields("b.Credit").Value

For Item = 0 To MaxCount - 1
    With AllTrans(ObjKeys(Item))
        If .DrAmount <> 0 Then
            .DrAmount = .DrAmount + DrTotal
            DrTotal = .DrAmount
        End If
        
        If .CrAmount <> 0 Then
            .CrAmount = .CrAmount + CrTotal
            CrTotal = .CrAmount
        End If
        
    End With
    
Next Item

TransID = rstTrans.Fields("TransID")

gDbTrans.SqlStmt = " UPDATE AccTrans " & _
                  " SET Debit = " & DrTotal & _
                  " WHERE HeadID = " & ToHeadID & _
                  " AND TransID = " & TransID
                  
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Function
End If

gDbTrans.SqlStmt = " UPDATE AccTrans " & _
                  " SET Credit = " & CrTotal & _
                  " WHERE HeadID = " & FromHeadID & _
                  " AND TransID = " & TransID
                  
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Function
End If

TransDateDataExists = TransID

Exit Function

Hell:
    TransDateDataExists = 0
    
End Function
'This will get the Transdate from collection
' and will check the amount
' if exists it will add the amount
Private Function UndoTransDateData(ByVal USTransDate As String) As Long
' Handle Error
On Error GoTo Hell:

' Declare Variables
Dim rstTrans As ADODB.Recordset
Dim FromHeadID As Long
Dim ToHeadID As Long
Dim TransID As Long

Dim DrTotal As Currency
Dim CrTotal As Currency

Dim Item As Long
Dim MaxCount As Long

Dim ObjKeys As Variant

' Intialise the Function
UndoTransDateData = 0

If AllTrans Is Nothing Then Exit Function

' Check the Collection Dr Amount and Cr Amount Total

MaxCount = AllTrans.count
ObjKeys = AllTrans.GetKeys

For Item = 0 To MaxCount - 1
    With AllTrans(ObjKeys(Item))
        
        If .DrAmount > 0 Then ToHeadID = .headID
        
        If .CrAmount > 0 Then FromHeadID = .headID
        
    End With
    
Next Item

gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM QryAccTransMerge" & _
                   " WHERE TransDate = #" & USTransDate & "#" & _
                   " AND A.HEADID=" & ToHeadID & _
                   " AND B.HEADID=" & FromHeadID
                   
If gDbTrans.Fetch(rstTrans, adOpenStatic) < 1 Then Exit Function

DrTotal = rstTrans.Fields("a.Debit").Value
CrTotal = rstTrans.Fields("b.Credit").Value

For Item = 0 To MaxCount - 1
    
    With AllTrans(ObjKeys(Item))
        
        If .DrAmount > 0 Then
            .DrAmount = DrTotal - .DrAmount
            DrTotal = .DrAmount
        End If
        
        If .CrAmount > 0 Then
            .CrAmount = CrTotal - .CrAmount
            CrTotal = .CrAmount
        End If
        
    End With
    
Next Item

TransID = rstTrans.Fields("TransID")

gDbTrans.BeginTrans

If DrTotal > 0 And CrTotal > 0 Then
    gDbTrans.SqlStmt = " UPDATE AccTrans " & _
                      " SET Debit= " & DrTotal & _
                      " WHERE HeadID=" & ToHeadID & _
                      " AND TransID=" & TransID
                      
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    gDbTrans.SqlStmt = " UPDATE AccTrans " & _
                      " SET Credit= " & CrTotal & _
                      " WHERE HeadID=" & FromHeadID & _
                      " AND TransID=" & TransID
                      
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
Else
    ' Becasue we have Reached o we are simply deleting
    m_TransID = TransID
    gDbTrans.BeginTrans
    If DeleteTransID = Success Then gDbTrans.CommitTrans Else gDbTrans.RollBack
    
End If

gDbTrans.CommitTrans

UndoTransDateData = TransID

Exit Function

Hell:
    
    UndoTransDateData = 0
    
End Function

Private Property Get UpdatingTrans() As Boolean
UpdatingTrans = m_UpdatingTrans
End Property

Private Property Let UpdatingTrans(NewValue As Boolean)
m_UpdatingTrans = NewValue
With m_frmAccTrans
    '.grdLedger(TransGrid).Enabled = Not m_UpdatingTrans
    .cmbTab0Ledger(VoucherCombo).Locked = m_UpdatingTrans
    .cmbTab0Ledger(ParentCombo).Locked = m_UpdatingTrans
    .cmbTab0Ledger(LedgerCombo).Locked = m_UpdatingTrans
    .cmdTab0Date.Enabled = Not NewValue
    .txtTab0CurrentDate.Locked = m_UpdatingTrans
    If m_UpdatingTrans Then .DBOperation = Update Else .DBOperation = Insert
End With

End Property

' This Will Check the Vouchers ans recpective entries
Private Function ValidateVouchers(ByVal VoucherType As Wis_VoucherTypes) As wis_FunctionReturned

' Handle error
On Error GoTo Hell:

ValidateVouchers = Failure

' This Will Fetch The Current Voucher Type Selected by the User

If DrVoucherCheck(VoucherType) <> Success Then Exit Function
If CrVoucherCheck(VoucherType) <> Success Then Exit Function

ValidateVouchers = Success

Exit Function

Hell:
    ValidateVouchers = FatalError
    
End Function

Private Sub LoadParentHeadsForVoucher(ctrlComboBox As ComboBox, ByVal LoadLedgers As wis_LoadLedgers)

On Error GoTo NoLoadParents:

Dim rstParent As ADODB.Recordset
Dim SqlStr As String
ctrlComboBox.Clear

'Check for the user permissions
If (gCurrUser.UserPermissions And perOnlyWaves) Then
    SqlStr = " SELECT ParentName,ParentID " & _
                    " FROM ParentHeads WHERE UserCreated <> 0 "
Else
    SqlStr = " SELECT ParentName,ParentID " & _
                    " FROM ParentHeads WHERE UserCreated <= 2 "
End If

Select Case LoadLedgers

    Case LoadALL
    
        gDbTrans.SqlStmt = SqlStr & " ORDER BY ParentName "
                    
    Case LoadExceptCash
    
        'gDbTrans.SQLStmt = " SELECT ParentName,ParentID " & _
                           " FROM ParentHeads " & _
                           " WHERE ParentID <> " & wis_BanksParentID & _
                           " AND ParentID <> " & wis_BanksODParentID & _
                           " AND ParentID <> " & wis_CashParentID & _
                           " ORDER BY ParentName "
      
      '''Changed For INdex 2000 NEW
       gDbTrans.SqlStmt = SqlStr & " AND ParentID <> " & wis_CashParentID & _
                           " ORDER BY ParentName "
      
    Case LoadOnlyCash
    
         gDbTrans.SqlStmt = SqlStr & " AND ParentID = " & wis_BanksParentID & _
                            " OR ParentID = " & wis_BanksODParentID & _
                            " OR ParentID = " & wis_CashParentID & _
                            " ORDER BY ParentName "
                
    Case LoadCashCreditros
    
         gDbTrans.SqlStmt = SqlStr & " AND ParentID = " & wis_BanksParentID & _
                            " OR ParentID = " & wis_BanksODParentID & _
                            " OR ParentID = " & wis_CashParentID & _
                            " OR ParentID=" & wis_CreditorsParentID & _
                            " OR ParentID=" & wis_DebitorsParentID & _
                            " ORDER BY ParentName "
        
End Select

SqlStr = ""

Call gDbTrans.Fetch(rstParent, adOpenForwardOnly)

Do While Not rstParent.EOF
    With ctrlComboBox
        .AddItem FormatField(rstParent.Fields("ParentName"))
        .ItemData(.newIndex) = FormatField(rstParent.Fields("ParentID"))
    End With
    'Move to the next record
    rstParent.MoveNext
Loop

Set rstParent = Nothing

Exit Sub

NoLoadParents:
    ctrlComboBox.Clear
    Set rstParent = Nothing
End Sub

Sub ClearAccTransControls()

With m_frmAccTrans

    .InitTab0Grid
    
    'Call LoadVouchersToCombo(.cmbTab0Ledger(VoucherCombo))
    .cmbTab0Ledger(VoucherCombo).ListIndex = -1
'    Call LoadParentHeads(.cmbTab0Ledger(ParentCombo))
    
    .cmbTab0Ledger(VoucherCombo).Locked = False
    .txtTab0CurrentDate.Locked = False
    .cmdTab0Date.Enabled = True
    .cmbTab0Ledger(LedgerCombo).Clear
    .txtTab0Amount.Value = 0
    .txtParticulars.Text = ""
    .txtParticulars.Enabled = True
    
    '.lblTab0Balance(0).Caption = GetResourceString(42) '"Balance"
    
'    .cmdOk.Enabled = False
    .cmdTab0AddtoGrid(AddToGrid).Enabled = True
    .cmdTab0AddtoGrid(AddToGrid).Caption = GetResourceString(10) '"Add"
    .cmdTab0AddtoGrid(DeleteGrid).Enabled = False
    .lblLedgerName.Caption = ""
    .grdLedger(.tabTrans.SelectedItem.Index - 1).Enabled = True
    .DBOperation = Insert
    On Error Resume Next
    .cmbTab0Ledger(VoucherCombo).SetFocus
    Err.Clear
End With

End Sub

' This fetch the data from Base table and load it to Grid
'
Public Sub ShowLedgerToGrid(ByVal headID As Long, grd As MSFlexGrid, _
                           DateToConsider As Boolean, _
                           fromDate As String, toDate As String)

' Handle Error
On Error GoTo Hell:

' Declarations
Dim rstLedger As ADODB.Recordset
Dim rstTrans As ADODB.Recordset
Dim rstParticulars As ADODB.Recordset

Dim headName As String
Dim TransDate As String
Dim VoucherName As String

Dim AccountType As wis_AccountType
Dim VoucherType As Wis_VoucherTypes

Dim Debit As Currency
Dim Credit As Currency
Dim opBalance As Currency
Dim DebitTotal As Currency
Dim CreditTotal As Currency
Dim ClosingBalance As Currency

Dim HeadIDDebit As Currency
Dim HeadIDCredit As Currency

Dim SlNo As Long
Dim TransID As Long
Dim ToHeadID As Long
Dim PrevTransID As Long

Dim Particulars As String

Dim LedgerClass As clsLedger

' Validations
If headID = 0 Then Exit Sub

' Set the Grid
Set m_GrdFunctions.fGrd = grd

Set LedgerClass = New clsLedger

grd.Clear

' Refresh the grid for the Ledger
Call InitTab1Grid(grd)

' First put the Opening Balance

If DateToConsider Then opBalance = GetOpBalance(headID, fromDate)
If Not DateToConsider Then opBalance = GetOpBalance(headID, FinUSFromDate)

AccountType = GetAccountType(headID)

grd.Row = 1
Call SetHeaderFooter(opBalance, grd, AccountType, GetResourceString(284))

' Set the Sql Statement
StartTimer

If DateToConsider Then _
    gDbTrans.SqlStmt = " SELECT TransID,Debit,Credit" & _
                        " FROM AccTrans " & _
                        " WHERE HeadID = " & headID & _
                        " AND TransDate BETWEEN #" & GetSysFormatDate(fromDate) & "#" & _
                        " AND #" & GetSysFormatDate(toDate) & "#" & _
                        " ORDER BY TransDate,TransID"
    
If Not DateToConsider Then _
    gDbTrans.SqlStmt = " SELECT TransID,Debit,Credit,TransDate" & _
                        " FROM AccTrans " & _
                        " WHERE HeadID = " & headID & _
                        " ORDER BY TransDate,TransID"
                        
If Not gDbTrans.CreateView("QryLedger") Then Exit Sub
    
Dim RowsToEnter As Long
gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM QryLedger"
                
'Fetch the Data to Recordset based on the Condition
RowsToEnter = gDbTrans.Fetch(rstTrans, adOpenStatic)

gDbTrans.SqlStmt = " SELECT a.Debit,a.Credit,a.VoucherType,a.TransID,a.TransDate,c.HeadName" & _
                    " FROM AccTrans a, QryLedger b,Heads c " & _
                    " WHERE a.HeadID <> " & headID & _
                    " AND a.TransID=b.TransID" & _
                    " AND a.HeadID=c.HeadID"
                    
Call gDbTrans.Fetch(rstLedger, adOpenStatic)


'gDbTrans.SQLStmt = " SELECT a.TransID,Particulars " & _
                   " FROM TransParticulars a , QryLedger b" & _
                   " WHERE a.TransID=b.TransID" & _
                   " ORDER BY a.TransID"
                   
'Call gDbTrans.Fetch(rstParticulars, adOpenStatic)

' Kept for the Reference
'
''    .Col = 0: .Text = "Date": .CellFontBold = True
''    .Col = 1: .Text = "Ledger Name": .CellFontBold = True
''    .Col = 2: .Text = "Voucher": .CellFontBold = True
''    .Col = 3: .Text = "Dr ": .CellFontBold = True
''    .Col = 4: .Text = "Cr ": .CellFontBold = True
''    .Col = 5: .Text = "Balance ": .CellFontBold = True

SlNo = grd.Row + 1

DebitTotal = 0
CreditTotal = 0
PrevTransID = 0

' This will create bulk rows
If RowsToEnter > 0 Then m_GrdFunctions.fMoreRows (RowsToEnter + 2)
    
Do While Not rstTrans.EOF
    
     'TransID = FormatField(rstTrans.Fields("TransID"))
     TransID = rstTrans.Fields("TransID")
     
     ' This will fetch the particulars if any
     'Particulars = ""
    
    
     'HeadIDDebit = FormatField(rstTrans.Fields("Debit"))
     HeadIDDebit = rstTrans.Fields("Debit")
     'HeadIDCredit = FormatField(rstTrans.Fields("Credit"))
     HeadIDCredit = rstTrans.Fields("Credit")
     
     rstLedger.Filter = adFilterNone
     If HeadIDDebit > 0 Then
         rstLedger.Filter = "TransID=" & TransID & " AND Credit > " & 0
     Else
         rstLedger.Filter = "TransID=" & TransID & " AND Debit > " & 0
     End If
                   
     RowsToEnter = rstLedger.recordCount
     
     ' This will update the bulk rows
     m_GrdFunctions.fMoreRows (RowsToEnter + 2)
     
     Do While Not rstLedger.EOF
         
         Debit = 0
         Credit = 0
         
         Do
            
            If HeadIDDebit > 0 Then
                If HeadIDDebit < rstLedger.Fields("Credit") Then
                    Debit = HeadIDDebit
                Else
                    Debit = rstLedger.Fields("Credit")
                End If
                'Credit = rstLedger.Fields("Debit")
                Exit Do
            End If
            
            If HeadIDCredit > 0 Then
                If HeadIDCredit < rstLedger.Fields("Debit") Then
                    Credit = HeadIDCredit
                Else
                    Credit = rstLedger.Fields("Debit")
                End If
                'Debit = rstLedger.Fields("Credit")
                Exit Do
            End If
            
         Loop
         
         'ToHeadID = rstLedger.Fields("HeadID")
         
         headName = rstLedger.Fields("HeadName")
         VoucherType = rstLedger.Fields("VoucherType")
               
         VoucherName = LedgerClass.GetVoucherName(VoucherType)
         TransDate = FormatField(rstLedger.Fields("TransDate"))
         
         DebitTotal = DebitTotal + Debit
         CreditTotal = CreditTotal + Credit
         
         ClosingBalance = GetClosingBalanceonAccountType(opBalance, DebitTotal, CreditTotal, AccountType)
         
         With grd
          
             'Call m_GrdFunctions.fMoreRows(2)
             
             SlNo = SlNo + 1
             .Row = SlNo
             .RowData(SlNo) = TransID
             .Col = 0: .Text = TransDate
             .Col = 1: .Text = headName ': .CellFontBold = True
             .Col = 2: .Text = VoucherName
             .Col = 3: .Text = FormatCurrency(Debit)
             .Col = 4: .Text = FormatCurrency(Credit)
             .Col = 5: .Text = FormatCurrency(ClosingBalance)
             
         End With
         
         rstLedger.MoveNext
         
     Loop
    
    rstTrans.MoveNext
        
Loop


m_GrdFunctions.fMoreRows (2)
grd.Row = grd.Row + 2

ClosingBalance = GetClosingBalanceonAccountType(opBalance, DebitTotal, CreditTotal, AccountType)
Call SetHeaderFooter(ClosingBalance, grd, AccountType, GetResourceString(285))

StopTimer

Set LedgerClass = Nothing
Set rstLedger = Nothing
Set rstTrans = Nothing
Set rstParticulars = Nothing

Exit Sub

Hell:
        
    Set LedgerClass = Nothing
    Set rstLedger = Nothing
    Set rstTrans = Nothing
    Set rstParticulars = Nothing
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

'This will refresh the Grid for to Show  ledger
Public Sub InitTab1Grid(grdLedger As MSFlexGrid)

With grdLedger

    .Clear
    .AllowUserResizing = flexResizeBoth
    .Rows = 5
    .Cols = 6
    .FixedCols = 1
    .FixedRows = 1
    
    .Row = 0
     
    .Col = 0: .CellFontBold = True: .Text = GetResourceString(37) ' "Date"
    .Col = 1: .CellFontBold = True: .Text = GetResourceString(36) ' "Ledger Name"
    .Col = 2: .CellFontBold = True: .Text = GetResourceString(41)  '"Voucher"
    .Col = 3: .CellFontBold = True: .Text = GetResourceString(277)  '"Dr "
    .Col = 4: .CellFontBold = True: .Text = GetResourceString(276)  '"Cr "
    .Col = 5: .CellFontBold = True: .Text = GetResourceString(42)  '"Balance"
    
    .ColWidth(0) = 1275
    .ColWidth(1) = 1100
    .ColWidth(2) = 825
    .ColWidth(3) = 900
    .ColWidth(4) = 900
    .ColWidth(5) = 1200
    
    .Row = 1

End With

End Sub

' This will read the Property TransID and Delete the Records
' Dependencies : TransID
Private Function DeleteTransID() As wis_FunctionReturned

' Handle Error
On Error GoTo Hell:

' Initialise Function
DeleteTransID = Failure

' Should Never Happen
If TransID = 0 Then Exit Function

' Delete from AccTrans
gDbTrans.SqlStmt = " DELETE * " & _
                   " FROM AccTrans " & _
                   " WHERE TransID = " & TransID

'gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then gDbTrans.RollBack

gDbTrans.SqlStmt = " DELETE * " & _
                   " FROM TransParticulars " & _
                   " WHERE TransID = " & TransID

If Not gDbTrans.SQLExecute Then Exit Function

DeleteTransID = Success

Exit Function

Hell:
    DeleteTransID = FatalError
    
End Function

'This function will check the Debit side total and Credit side total
Private Function DrCrTotalsSame(isHandleVouchers As Boolean) As Currency

' Handle Error
On Error GoTo Hell:

' Declare Variables
Dim DrTotal As Currency
Dim CrTotal As Currency

Dim Item As Integer
Dim MaxCount As Integer

Dim ObjKeys As Variant

' Intialise the Function
DrCrTotalsSame = 0

'Check the Collection Dr Amount and Cr Amount Total
MaxCount = AllTrans.count
ObjKeys = AllTrans.GetKeys
For Item = 0 To MaxCount - 1
    DrTotal = DrTotal + AllTrans(ObjKeys(Item)).DrAmount
    CrTotal = CrTotal + AllTrans(ObjKeys(Item)).CrAmount
Next Item

' If Equal Then make function's value to true

If DrTotal = CrTotal Then DrCrTotalsSame = 0
If DrTotal > CrTotal Then DrCrTotalsSame = DrTotal - CrTotal
If DrTotal < CrTotal Then DrCrTotalsSame = CrTotal - DrTotal

If Not isHandleVouchers Then Exit Function

' This will handle the Vocuhers and Option Buttons
Call HandleVouchersAfterAdd(DrTotal, CrTotal)

Exit Function

Hell:

    MsgBox "DrCrTotalSame :" & vbCrLf & Err.Description
    
    
End Function

'This function will check the validation of the Voucher
Private Function AllTransVoucherCheck(ByVal VoucherType As Wis_VoucherTypes) As wis_FunctionReturned

' Handle Error
On Error GoTo Hell:

' Intialise the Function
AllTransVoucherCheck = Failure

' Validate the Inputs and Classes
If AllTrans Is Nothing Then Exit Function

If DrVoucherCheck(VoucherType) <> Success Then Exit Function
If CrVoucherCheck(VoucherType) <> Success Then Exit Function

AllTransVoucherCheck = Success

Exit Function

Hell:
    
    AllTransVoucherCheck = FatalError
    
    
End Function

Private Function DrVoucherCheck(ByVal VoucherType As Wis_VoucherTypes) As wis_FunctionReturned

' Handle Error
On Error GoTo Hell:

' Declare Variables
Dim headID As Long
Dim ParentID As Long
Dim Item As Long
Dim MaxCount As Long

Dim isCashFound As Boolean
Dim isPurchaseFound As Boolean

Dim ObjKeys As Variant

Dim objTrans As Trans

Const CURRENTSIDE = "Debit Side: " & vbCrLf

' Intialise the Function
DrVoucherCheck = Failure

' Validate the Inputs and Classes
If AllTrans Is Nothing Then Exit Function

' Check the Collection Dr Amount and Cr Amount Total
isCashFound = False
isPurchaseFound = False

Set objTrans = New Trans
MaxCount = AllTrans.count
ObjKeys = AllTrans.GetKeys
For Item = 0 To MaxCount - 1

    Set objTrans = AllTrans.Item(ObjKeys(Item))
    
    'Validation Goes Hear
    ' Please Don't Change the Case Creiteria
    If objTrans.DrAmount <> 0 Then
        Debug.Assert objTrans.DrAmount > 0
        headID = objTrans.headID
        ParentID = objTrans.ParentID
    
        Select Case VoucherType
        
            Case payment
                'If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or _
                        (ParentID = wis_BanksODParentID) Then
                If (ParentID = wis_CashParentID) Then _
                Err.Raise vbObjectError + 513, , CURRENTSIDE & "Payment Voucher Can not Contain Cash or Bank !!"
                
            Case Receipt
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then isCashFound = True
                            
            Case Sales
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Or _
                    (ParentID = wis_CreditorsParentID) Or (ParentID = wis_DebitorsParentID) Then
                                     
                     isCashFound = True
                End If
                
            Case Purchase
               If ParentID = wis_PurchaseParentID Then isPurchaseFound = True
                
            Case FreePurchase
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Free Voucher Can not Contain Cash or Bank !!"
                    
            Case FreeSales
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Free Voucher Can not Contain Cash or Bank !!"
                
            Case CONTRA
            
                If (ParentID <> wis_CashParentID) And (ParentID <> wis_BanksParentID) And (ParentID <> wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Contra Voucher Can not Contain other than Cash !!"
            
            Case Journal
        
                'If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Journal Voucher Can not Contain Cash or Bank !!"
                If ParentID = wis_CashParentID Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Journal Voucher Can not Contain Cash !!"
    
            Case RejectionsIn
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "RejectionsIN Voucher Can not Contain Cash or Bank !!"
                
            Case RejectionsOut
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "RejectionsOut Voucher Can not Contain Cash or Bank !!"
                
        End Select
        
    End If
    
Next Item

Set objTrans = Nothing

Select Case VoucherType

    Case Sales
        If Not isCashFound Then Err.Raise vbObjectError + 513, , CURRENTSIDE & "Sales Voucher Must have Creditros or Cash or Bank"
                
    Case Purchase
        
        If Not isPurchaseFound Then Err.Raise vbObjectError + 513, , CURRENTSIDE & "Purchase Voucher Must have Purchase Account"
    
    Case Receipt
    
        'For the time being we are allowing this for the current year's transaction
        'we are commenting the following line
        'in the next fin year we have to acivate the following line
        'Date 15/01/2005
        If Not isCashFound Then Err.Raise vbObjectError + 513, , CURRENTSIDE & "Payment or Receipt Voucher Should contain Cash or Bank !!"
        
End Select

DrVoucherCheck = Success

Exit Function


Hell:
    
    If Err Then MsgBox "Voucher Check : " & Err.Description, vbCritical, wis_MESSAGE_TITLE
    
    DrVoucherCheck = FatalError
    'Resume
    
End Function

'This function will check the validation of the Voucher
Private Function CrVoucherCheck(ByVal VoucherType As Wis_VoucherTypes) As wis_FunctionReturned

' Handle Error
On Error GoTo Hell:
' Declare Variables

Dim headID As Long
Dim ParentID As Long
Dim Item As Long
Dim MaxCount As Long

Dim ObjKeys As Variant

Dim objTrans As Trans
Dim isCashFound As Boolean
Dim isSalesFound As Boolean

Const CURRENTSIDE = "Credit Side: " & vbCrLf

' Intialise the Function
CrVoucherCheck = Failure

' Validate the Inputs and Classes
'THe Below condition never occurs 'SDA 16/12/2002
If AllTrans Is Nothing Then Exit Function

' Check the Collection Dr Amount and Cr Amount Total

isCashFound = False
isSalesFound = False

Set objTrans = New Trans
MaxCount = AllTrans.count
ObjKeys = AllTrans.GetKeys

For Item = 0 To MaxCount - 1

    Set objTrans = AllTrans.Item(ObjKeys(Item))
    
    'Validation Goes Hear
    If objTrans.CrAmount <> 0 Then
        Debug.Assert objTrans.CrAmount > 0
        headID = objTrans.headID
        ParentID = objTrans.ParentID
        
        ' Please Don't Change the Case Creiteria
        Select Case VoucherType
        
            Case payment
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then isCashFound = True
                    
            Case Receipt
            
                'If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Receipt Voucher Can not Contain Cash or Bank !!"
                If ParentID = wis_CashParentID Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Receipt Voucher Can not Contain Cash or Bank !!"
                
            Case Sales
                
                If (ParentID = wis_SalesParentID) Then isSalesFound = True
                
            Case Purchase
                
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Or _
                    (ParentID = wis_CreditorsParentID) Or (ParentID = wis_DebitorsParentID) Then
                                     
                     isCashFound = True
                End If
                
            Case FreePurchase
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Free Voucher Can not Contain Cash or Bank !!"
                    
            Case FreeSales
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Free Voucher Can not Contain Cash or Bank !!"
            Case CONTRA
            
                If (ParentID <> wis_CashParentID) And (ParentID <> wis_BanksParentID) And (ParentID <> wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Contra Voucher Can not Contain Other than Cash or Bank !!"
                
            Case Journal
        
                'If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Journal Voucher Can not Contain Cash or Bank !!"
                If ParentID = wis_CashParentID Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "Journal Voucher Can not Contain Cash or Bank !!"
                
            Case RejectionsIn
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "RejectionsIN Voucher Can not Contain Cash or Bank !!"
                
            Case RejectionsOut
            
                If (ParentID = wis_CashParentID) Or (ParentID = wis_BanksParentID) Or (ParentID = wis_BanksODParentID) Then _
                    Err.Raise vbObjectError + 513, , CURRENTSIDE & "RejectionsOut Voucher Can not Contain Cash or Bank !!"
                
        End Select
        
    End If
    
Next Item

Set objTrans = Nothing

Select Case VoucherType
    Case payment
        'For the time being we are allowing this for the current year's transaction
        'we are commenting the following line
        'in the next fin year we have to acivate the following line
        'Date 15/01/2005
        
        If Not isCashFound Then Err.Raise vbObjectError + 513, , CURRENTSIDE & "Payment Voucher Should contain Cash or Bank !!"
    Case Purchase
        If Not isCashFound Then Err.Raise vbObjectError + 513, , CURRENTSIDE & "Purchase Voucher Should contain Cash or Bank or Debitor!!"
    Case Sales
        If Not isSalesFound Then Err.Raise vbObjectError + 513, , CURRENTSIDE & "Sales Voucher Should contain Sales Account !"
End Select

CrVoucherCheck = Success

Exit Function

Hell:
    If Err Then MsgBox "Voucher Check : " & Err.Description, vbCritical, wis_MESSAGE_TITLE
    CrVoucherCheck = FatalError
    Set objTrans = Nothing
    
End Function

' This will handle all the proceedings of the Vouchers
' Quite a Responsible Sub !!
Private Sub HandleVouchers()

Dim voucherTypes As Wis_VoucherTypes

Dim WavesUser As Boolean
WavesUser = (gCurrUser.UserPermissions And perOnlyWaves)

If m_frmAccTrans Is Nothing Then Exit Sub

'This Will Fetch The Current Voucher Type Selected by the User
With m_frmAccTrans
    
    With .cmbTab0Ledger(VoucherCombo)
        If .ListIndex = -1 Then Exit Sub
        voucherTypes = .ItemData(.ListIndex)
    End With

    If SlNo = 0 Then
        .optTab0From(FromOpt).Enabled = False
        .optTab0From(ToOpt).Enabled = False
        .cmbTab0Ledger(ParentCombo).Clear
        .cmbTab0Ledger(LedgerCombo).Clear
    End If
    
    .cmdTab0AddtoGrid(AddToGrid).Caption = GetResourceString(10)  'Add
    .cmdTab0AddtoGrid(DeleteGrid).Caption = GetResourceString(14)   'Delete
    .cmdOk.Caption = GetResourceString(7) 'Save
    .cmdTab0AddtoGrid(AddToGrid).Visible = True
    .cmdTab0AddtoGrid(DeleteGrid).Visible = True
    'now assgn the oringanal tab index of the particulars textbox
    .txtParticulars.TabIndex = Val(.txtParticulars.Tag)
    
    'Please don't change the case creteria
    Select Case voucherTypes
        Case payment
            'Compatable To Index 20000
            .optTab0From(ToOpt).Value = True
            Call LoadParentHeadsForVoucher(.cmbTab0Ledger(ParentCombo), LoadExceptCash)
            .cmdOk.Caption = GetResourceString(IIf(WavesUser, 7, 10)) 'Add
            .cmdOk.Enabled = Not WavesUser
            .cmdTab0AddtoGrid(AddToGrid).Visible = WavesUser
            .cmdTab0AddtoGrid(DeleteGrid).Visible = WavesUser
            .txtParticulars.TabIndex = IIf(WavesUser, Val(.txtParticulars.Tag), .cmdOk.TabIndex - 1)
        Case Receipt
            '.optTab0From(FromOpt).value = True
            '.optTab0From(FromOpt).Enabled = True
            'Call LoadParentHeadsForVoucher(.cmbTab0Ledger(ParentCombo), LoadExceptCash)
            
            'Compatable To Index 20000
            .optTab0From(FromOpt).Value = True
            .optTab0From(FromOpt).Enabled = WavesUser
            Call LoadParentHeadsForVoucher(.cmbTab0Ledger(ParentCombo), LoadExceptCash)
            .cmdOk.Caption = GetResourceString(IIf(WavesUser, 7, 10)) 'Add
            .cmdOk.Enabled = Not WavesUser
            .cmdTab0AddtoGrid(AddToGrid).Visible = WavesUser
            .cmdTab0AddtoGrid(DeleteGrid).Visible = WavesUser
            .txtParticulars.TabIndex = IIf(WavesUser, Val(.txtParticulars.Tag), .cmdOk.TabIndex - 1)
        
        Case CONTRA
            '.optTab0From(FromOpt).value = True
            '.optTab0From(FromOpt).Enabled = True
            'Call LoadParentHeadsForVoucher(.cmbTab0Ledger(ParentCombo), LoadOnlyCash)
            
            'Compatable To INdex 20000
            .optTab0From(FromOpt).Value = True
            '.optTab0From(FromOpt).Enabled = True
            .optTab0From(FromOpt).Enabled = WavesUser
            Call LoadParentHeadsForVoucher(.cmbTab0Ledger(ParentCombo), LoadExceptCash)
        Case Journal
            .cmdOk.Enabled = False
            .optTab0From(FromOpt).Value = True
            .optTab0From(FromOpt).Enabled = True
            Call LoadParentHeadsForVoucher(.cmbTab0Ledger(ParentCombo), LoadExceptCash)
            
        Case Sales
            .optTab0From(ToOpt).Value = True
            .optTab0From(ToOpt).Enabled = True
            Call LoadParentHeadsForVoucher(.cmbTab0Ledger(ParentCombo), LoadCashCreditros)
            .txtParticulars.TabIndex = IIf(WavesUser, Val(.txtParticulars.Tag), .cmdOk.TabIndex - 1)
            
        Case Purchase
            .optTab0From(FromOpt).Value = True
            .optTab0From(FromOpt).Enabled = True
            Call LoadParentHeadsForVoucher(.cmbTab0Ledger(ParentCombo), LoadCashCreditros)
            .txtParticulars.TabIndex = IIf(WavesUser, Val(.txtParticulars.Tag), .cmdOk.TabIndex - 1)
            
        Case FreePurchase
            .optTab0From(FromOpt).Value = True
            .optTab0From(FromOpt).Enabled = True
            Call LoadParentHeadsForVoucher(.cmbTab0Ledger(ParentCombo), LoadExceptCash)
            .txtParticulars.TabIndex = IIf(WavesUser, Val(.txtParticulars.Tag), .cmdOk.TabIndex - 1)
            
        Case FreeSales
            .optTab0From(FromOpt).Value = True
            .optTab0From(FromOpt).Enabled = True
            Call LoadParentHeadsForVoucher(.cmbTab0Ledger(ParentCombo), LoadExceptCash)
            .txtParticulars.TabIndex = IIf(WavesUser, Val(.txtParticulars.Tag), .cmdOk.TabIndex - 1)
            
    End Select
End With
End Sub

Private Sub LoadGridFromFrm(ByVal RowNum As Integer)

' Handle Error
On Error GoTo Hell:

' Declare Varaiables
Dim Amount As Currency
Dim ParentName As String
Dim DRCRType As wis_DrCrType
Dim ParentID As Long

Dim count As Integer
Dim Item As Integer

'Validate the Controls
With m_frmAccTrans
    With .grdLedger(TransGrid)
        .Row = RowNum
        ' Get the HeadID from the Row
        headID = .RowData(RowNum)
        
        'if user has clicked in the blank row
        If headID = 0 Then Exit Sub
        
        ' After to Check if HeadID has changed
        PrevHeadID = headID
        .Col = 0
        SlNo = Val(.Text)
        .Col = 1
        headName = .Text
        .Col = 2
        Debit = Val(.Text)
        .Col = 3
        Credit = Val(.Text)
    End With
    
    ' Reallocate the ParentHeads
    Call DrCrTotalsSame(True)
        
    ' Get the ParentName
    ParentName = GetParentName(headID)
    ParentID = GetParentID(headID)

    count = .cmbTab0Ledger(ParentCombo).ListCount - 1
    For Item = 0 To count
        If .cmbTab0Ledger(ParentCombo).ItemData(Item) = ParentID Then
            .cmbTab0Ledger(ParentCombo).ListIndex = Item
            Exit For
        End If
    Next Item
    
    ' load ParentHead's Heads to Combo
    Call LoadLedgersToCombo(.cmbTab0Ledger(LedgerCombo), ParentID)
        
    ' Set the LedgerName Combo to HeadName
    
    count = .cmbTab0Ledger(LedgerCombo).ListCount - 1
    For Item = 0 To count
        If .cmbTab0Ledger(LedgerCombo).ItemData(Item) = headID Then
            .cmbTab0Ledger(LedgerCombo).ListIndex = Item
            .lblTab0Balance(0).Caption = GetBalance(headID, GetSysFormatDate(TransDate), TransID)
            Exit For
        End If
    Next Item

    ' Set the Option Buttons and get the Amount
    If IsCashTransaction = Success And Not IsBankHeadTransaction(TransID) Then
        .optTab0From(FromOpt).Enabled = False
        .optTab0From(ToOpt).Enabled = False
    Else
        .optTab0From(FromOpt).Enabled = True
        .optTab0From(ToOpt).Enabled = True
    End If
    If Credit > 0 Then
        .optTab0From(FromOpt).Value = True
        .txtTab0Amount.Value = Credit
    End If
    
    If Debit > 0 Then
        .optTab0From(ToOpt).Value = True
        .txtTab0Amount.Value = Debit
    End If
    
    ' Make the Delete command visible
    .cmdTab0AddtoGrid(DeleteGrid).Enabled = True
    
    ''Code For Index 2000 Compitability
    .cmdTab0AddtoGrid(DeleteGrid).Visible = True    'By SDA
    .cmdTab0AddtoGrid(DeleteGrid).Enabled = True    'By SDA

'    .cmdTab0AddtoGrid(AddToGrid).Visible = True     'By SDA
'    .cmdTab0AddtoGrid(AddToGrid).Enabled = True     'By SDA
    
    'Make the Add Command to Update Combo
    .cmdTab0AddtoGrid(AddToGrid).Caption = GetResourceString(171) '"Update"
    .cmdOk.Caption = GetResourceString(IIf(IsCashTransaction, 171, 7)) '"Update"
    
    .cmdTab0AddtoGrid(AddToGrid).Enabled = True
    '.cmdOk.Enabled = False
                
    ' Disable the grdledger
    '.grdLedger(TransGrid).Enabled = False

End With

m_isUpdatingGrid = True

Exit Sub

Hell:

End Sub

'This will swap the value of the Option Buttons of the m_frmAccTrans
Private Sub SwapOptionButtons()

On Error GoTo Hell:

' Validation

If m_frmAccTrans Is Nothing Then Exit Sub

With m_frmAccTrans

    If .optTab0From(FromOpt).Value Then
        .optTab0From(ToOpt).Value = True
    ElseIf .optTab0From(ToOpt).Value Then
        .optTab0From(FromOpt).Value = True
    End If
    
End With

Exit Sub

Hell:

    
End Sub

' When in the frmAccTrans user clicks on the Update comand
' This procedures is called
Private Function UpdateProcedures() As wis_FunctionReturned

' Handle Error
On Error GoTo Hell:

UpdateProcedures = Failure

' Declarations

' Validate the Controls
If m_frmAccTrans Is Nothing Then Exit Function
If ValidateLedgers <> Success Then Exit Function

If PrevHeadID <> headID Then
    If ALlTransHeadIDExists(headID) Then
        MsgBox "Head Selected Already Exists"
        Exit Function
    End If
End If

If AllTransHeadsUpdate(headID, Debit, Credit) <> Success Then Exit Function
'Now Check whether this transction Cah Or Contra
' Update the Grid
If UpDateGrids() <> Success Then Exit Function
' Reset the form for the next entry
Call PrepareAccTrans

m_isUpdatingGrid = False


UpdateProcedures = Success

Exit Function

Hell:
    MsgBox "UpdateProcedures()" & vbCrLf & Err.Description
    UpdateProcedures = FatalError
    
End Function

' When in the frmAccTrans user clicks on the Delete comand
' This procedures is called
Private Sub DeleteProcedures()

' Handle Error
On Error GoTo Hell:

' Declarations
Dim RowNum As Integer
Dim MaxRows As Integer
Dim CurrentRow As Integer

Dim isContinue As Boolean

' Validate the Controls
If m_frmAccTrans Is Nothing Then Exit Sub
If ValidateLedgers <> Success Then Exit Sub


'First Check Whthter this is cash TrasnCtion or Not
'If it  ic cash then Delete the TrasnCtion Here Itselef Then
'Clear the All grid
If IsCashTransaction = Success Then
    'delete the transction
    gDbTrans.BeginTrans
    If DeleteTransID <> Success Then
        gDbTrans.RollBack
        GoTo Hell
    End If
    gDbTrans.CommitTrans
    
    'Then Delet delete the Collection
    Call ClearCollections

    ' Clear all the COntrols in the form
    Call ClearAccTransControls
    
    UpdatingGrid = False
    UpdatingTrans = False
End If

'First Remove Key from the Collection
If ALlTransHeadIDExists(PrevHeadID) Then AllTransHeadIDRemove (PrevHeadID)
'If ALlTransHeadIDExists(HeadID) Then AllTransHeadIDRemove (HeadID)
With m_frmAccTrans
    ' Update the Grid
    With .grdLedger(TransGrid)  '.grdLedger(.tabTrans.SelectedItem.Index - 1)
        .Col = 0
        RowNum = Val(.Text)
        .RemoveItem (RowNum)
        'This is the Property to Update
        .Col = 0
        .Row = 1
        SlNo = 0
        If Val(.Text) > 0 Then isContinue = True
        Do While isContinue
            SlNo = SlNo + 1
            .Col = 0: .Text = SlNo
            .Row = .Row + 1
            .Col = 0
            If Val(.Text) = 0 Then isContinue = False
        Loop
    End With
    
    ' Check the Collection Dr Amount and Cr Amount Total and it should be same
    If DrCrTotalsSame(False) = 0 And SlNo > 0 Then
        .cmdTab0AddtoGrid(AddToGrid).Enabled = False
        .cmdOk.Enabled = True
    Else
        .cmdTab0AddtoGrid(AddToGrid).Enabled = True
        .cmdOk.Enabled = False
    End If
    
    'Now bring back the normal conditions
    .cmdTab0AddtoGrid(AddToGrid).Caption = GetResourceString(10) '"Add"
    .cmdTab0AddtoGrid(DeleteGrid).Enabled = False
    .grdLedger(.tabTrans.SelectedItem.Index - 1).Enabled = True
    
    If SlNo = 0 Then
        .cmbTab0Ledger(VoucherCombo).ListIndex = -1
        .cmbTab0Ledger(VoucherCombo).Locked = False
        .txtTab0CurrentDate.Locked = False
        .cmdTab0Date.Enabled = True
        If gOnLine Then .cmdTab0Date.Enabled = False

        .optTab0From(FromOpt).Enabled = False
        .optTab0From(ToOpt).Enabled = True
    End If
End With

m_isUpdatingGrid = False

Exit Sub

Hell:

End Sub

Private Property Let Credit(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Credit = 5
    m_Credit = vData
End Property

Public Property Let UpdatingGrid(ByVal Data As Boolean)
    m_isUpdatingGrid = Data
End Property

Private Property Get Credit() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Credit
    Credit = m_Credit
End Property

Public Property Get UpdatingGrid() As Boolean
    UpdatingGrid = m_isUpdatingGrid
End Property

Private Property Let Debit(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Debit = 5
    m_Debit = vData
End Property

Private Property Get Debit() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Debit
    Debit = m_Debit
End Property
Private Property Let Balance(ByVal vData As Currency)
    m_Balance = vData
End Property
Private Property Get Balance() As Currency
    Balance = m_Balance
End Property
Private Property Let TransDate(ByVal vData As String)
    m_TransDate = vData
End Property
Private Property Get TransDate() As String
    TransDate = m_TransDate
End Property

Private Property Let headName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HeadName = 5
    m_HeadName = vData
End Property

Private Property Get headName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HeadName
    headName = m_HeadName
End Property
Private Property Let headID(ByVal vData As Long)
    m_HeadID = vData
End Property

Private Property Get headID() As Long
    headID = m_HeadID
End Property
Private Property Let PrevHeadID(ByVal vData As Long)
    m_PrevHeadID = vData
End Property
Private Property Get PrevHeadID() As Long
    PrevHeadID = m_PrevHeadID
End Property

Public Property Let TransID(ByVal vData As Long)
    m_TransID = vData
End Property
Public Property Get TransID() As Long
    TransID = m_TransID
End Property


Private Property Let SlNo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SlNo = 5
    M_SlNo = vData
End Property


Private Property Get SlNo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SlNo
    SlNo = M_SlNo
End Property

Private Sub ClearCollections()

On Error Resume Next
Set m_AllTrans = Nothing

'Shashi 21/12/2002
'Clear the Updation process if it is doing it
UpdatingGrid = False

'Shashi transid
'to avoide the wating of transid
'we are insertin the same transid which is deleted
'so if we had not set it to zero theres is cahnce of
'deleting any previous transacion
TransID = 0
End Sub
Private Function ALlTransHeadIDExists(ByVal headID As Long) As Boolean
   
On Error GoTo Hell:

    ALlTransHeadIDExists = False
    ALlTransHeadIDExists = AllTrans.Exists("Tr" & headID)
                
Exit Function

Hell:
    
End Function
Private Property Get AllTrans() As AllTrans
    If m_AllTrans Is Nothing Then Set m_AllTrans = New AllTrans
    Set AllTrans = m_AllTrans
End Property
Private Function AllTransGetKeys() As Variant

    If m_AllTrans Is Nothing Then Set m_AllTrans = New AllTrans
        
    Set AllTransGetKeys = m_AllTrans.GetKeys
    
    
End Function
Private Property Set AllTrans(ByVal vData As AllTrans)
    Set m_AllTrans = vData
End Property

' Returns the MaxTransID from the Base Table
' Inputs :
'           HeadID as long
' Retrurns TransID as long
Public Function GetMaxTransID(ByVal headID As Long) As Long
' Handle Error
On Error GoTo TransIDFailed:

' Declare Variable
Dim rstTransID As ADODB.Recordset

Dim TableName As String

' Intialise the Function
GetMaxTransID = 0

' Validate the inputs
If headID = 0 Then Exit Function

' Default Name
TableName = "AccTrans"

'If HeadID = wis_CashHeadID Then TableName = "CashTrans"

' Fetch the MaxTransID from Acctrans AND FROM Cash Trans
gDbTrans.SqlStmt = " Select Max(TransID) as AccTransID " & _
                   " FROM " & TableName & _
                   " WHERE Headid=" & headID
                   
If gDbTrans.Fetch(rstTransID, adOpenForwardOnly) < 1 Then Exit Function

' Set getMaxtransid
GetMaxTransID = FormatField(rstTransID.Fields("AccTransID"))

Set rstTransID = Nothing

Exit Function

TransIDFailed:
    
    Set rstTransID = Nothing
    
End Function


'
' Returns the MaxTransID from the Base Table
' Returns TransID as long
'
Public Function GetMaxTransIDFromTable() As Long
' Handle Error
On Error GoTo TransIDFailed:
' Declare Variable
Dim rstTransID As ADODB.Recordset
' Intialise the Function
GetMaxTransIDFromTable = 0
' Fetch the MaxTransID from Acctrans AND FROM Cash Trans
gDbTrans.SqlStmt = " Select Max(TransID) as AccTransID " & _
                   " FROM AccTrans"

If gDbTrans.Fetch(rstTransID, adOpenForwardOnly) < 1 Then Exit Function
' Set getMaxtransid
GetMaxTransIDFromTable = FormatField(rstTransID.Fields(0))
Set rstTransID = Nothing
Exit Function
TransIDFailed:
    Set rstTransID = Nothing
End Function

'Returns the Opening Balance from the Base Table
' Inputs :
'           HeadID as long
'           TransID as long
'
' Retrurns Balance
Public Function GetOpBalance(ByVal headID As Long, ByVal AsOnDate As Date) As Currency
' Handle Error
On Error GoTo TransIDFailed:
' Declare Variable
Dim rstBalance As ADODB.Recordset
' Intialise the Function
GetOpBalance = 0
' Validate the inputs
If headID = 0 Then Exit Function
If Not DateValidate(CStr(AsOnDate), "/", False) Then Exit Function
'Dim AsOnDate As Date
'AsOnDate = FormatDate(AsOnIndainDate)

' Fetch the Balance from Acctrans AND FROM Cash Trans
If CDate(AsOnDate) = CDate(FinUSFromDate) Then
    gDbTrans.SqlStmt = " SELECT OpAmount " & _
                       " FROM OpBalance " & _
                       " WHERE Headid=" & headID & _
                       " AND Opdate = " & "#" & AsOnDate & "#"
    
    If gDbTrans.Fetch(rstBalance, adOpenForwardOnly) < 1 Then Exit Function
    GetOpBalance = FormatField(rstBalance.Fields(0))
Else
    AsOnDate = DateAdd("d", -1, AsOnDate)
    GetOpBalance = GetBalance(headID, AsOnDate)
End If

Set rstBalance = Nothing
Exit Function

TransIDFailed:
        Set rstBalance = Nothing
End Function

' Returns the Balance from the Base Table
' Inputs :
'           HeadID as long
'
'
' Retrurns Balance
'
Private Function GetBalance(ByVal headID As Long, _
                           ByVal AsOnDate As Date, _
                           Optional TransID As Long) As Currency
                           
' Handle Error
On Error GoTo TransIDFailed:
' Declare Variable
Dim rstBalance As ADODB.Recordset
Dim TableName As String
Dim opBalance As Currency
Dim DebitTotal As Currency
Dim CreditTotal As Currency
Dim AccountType As wis_AccountType
' Intialise the Function
GetBalance = 0
' Validate the inputs
If headID = 0 Then Exit Function

Dim fromDate As Date

fromDate = FinUSFromDate

' Get the Opening Balance
opBalance = GetOpBalance(headID, FinUSFromDate)

' Get the Sum of Debit and Sum of Credit as on date
gDbTrans.SqlStmt = " SELECT SUM(a.Debit) as DebitTotal,SUM(a.Credit) as CreditTotal" & _
                   " FROM AccTrans a" & _
                   " WHERE TransDate BETWEEN #" & fromDate & "#" & _
                   " AND #" & AsOnDate & "#" & _
                   " AND a.HeadID = " & headID
                   
If gDbTrans.Fetch(rstBalance, adOpenForwardOnly) < 1 Then Exit Function
DebitTotal = FormatField(rstBalance.Fields("DebitTotal"))
CreditTotal = FormatField(rstBalance.Fields("CreditTotal"))
Set rstBalance = Nothing

If TransID Then
    'If there are more transaction on the same day then the
    'result of the balance will include the balances after theis transction
    'so to get the balance after a particular tranction this is usefull
    'Sda 7/1/03
    gDbTrans.SqlStmt = "SELECT SUM(a.Debit) as DebitTotal,SUM(a.Credit) as CreditTotal" & _
                   " FROM AccTrans a" & _
                   " WHERE TransDate = #" & AsOnDate & "#" & _
                   " AND a.HeadID = " & headID & _
                   " AND TransID >= " & TransID
    If gDbTrans.Fetch(rstBalance, adOpenForwardOnly) > 0 Then
        DebitTotal = DebitTotal - FormatField(rstBalance.Fields("DebitTotal"))
        CreditTotal = CreditTotal - FormatField(rstBalance.Fields("CreditTotal"))
        Set rstBalance = Nothing
    End If
End If

' Get AccountType
AccountType = GetAccountType(headID)
'Depending opon the AccountType calculate the Balance
GetBalance = GetClosingBalanceonAccountType(opBalance, DebitTotal, CreditTotal, AccountType)

Exit Function

TransIDFailed:
        Set rstBalance = Nothing
End Function
' Returns the Balance from the Base Table
' Inputs :
'           HeadID as long
'
'
' Retrurns Balance
'
Public Function GetBalanceFromRecordset(ByVal headID As Long, rstBalance As ADODB.Recordset, ByVal AccountType As wis_AccountType) As Currency
' Handle Error
On Error GoTo TransIDFailed:
' Declare Variable
Dim TransID As Long
Dim opBalance As Currency
Dim DebitTotal As Currency
Dim CreditTotal As Currency
' Intialise the Function
GetBalanceFromRecordset = 0
' Validate the inputs
If headID = 0 Then Exit Function
' Get the Opening Balance
opBalance = GetOpBalance(headID, FinUSFromDate)
' Get the Sum of Debit and Sum of Credit as on date

DebitTotal = 0
CreditTotal = 0
If gDbTrans.FindRecord(rstBalance, "HeadID=" & headID) Then
    DebitTotal = FormatField(rstBalance.Fields("DebitTotal"))
    CreditTotal = FormatField(rstBalance.Fields("CreditTotal"))
End If
'Depending opon the AccountType calculate the Balance
GetBalanceFromRecordset = GetClosingBalanceonAccountType(opBalance, DebitTotal, CreditTotal, AccountType)
Exit Function
TransIDFailed:
End Function

' Returns the ClosingBalance after Calculation
' Retrurns ClosingBalance
Public Function GetClosingBalance(ByVal headID As Long, _
                                  ByVal opBalance As Currency, _
                                  ByVal DebitTotal As Currency, _
                                  ByVal CreditTotal As Currency) As Currency
' Handle Error
On Error GoTo TransIDFailed:
' Declare Variable
Dim AccountType As wis_AccountType
' Intialise the Function
GetClosingBalance = 0
' Validate the inputs
If headID = 0 Then Exit Function
' Get AccountType
AccountType = GetAccountType(headID)
'Depending opon the AccountType calculate the Balance
GetClosingBalance = _
GetClosingBalanceonAccountType(opBalance, DebitTotal, CreditTotal, AccountType)
Exit Function
TransIDFailed:
End Function
' Returns the ClosingBalance after Calculation
'
' Retrurns ClosingBalance
'
Public Function GetClosingBalanceonAccountType(ByVal opBalance As Currency, _
                                    ByVal DebitTotal As Currency, _
                                    ByVal CreditTotal As Currency, _
                                    ByVal AccountType As wis_AccountType) As Currency
' Handle Error
On Error GoTo TransIDFailed:
' Declare Variable
' Intialise the Function
GetClosingBalanceonAccountType = 0
'Depending opon the AccountType calculate the Balance
Select Case AccountType
    Case Asset, Loss, ItemPurchase
        GetClosingBalanceonAccountType = opBalance - CreditTotal + DebitTotal
    Case Liability, Profit, ItemSales
        GetClosingBalanceonAccountType = opBalance + CreditTotal - DebitTotal
    Case Loss
        GetClosingBalanceonAccountType = opBalance - CreditTotal + DebitTotal
    Case Profit
        GetClosingBalanceonAccountType = opBalance + CreditTotal - DebitTotal
    Case ItemPurchase
        GetClosingBalanceonAccountType = opBalance - CreditTotal + DebitTotal
    Case ItemSales
        GetClosingBalanceonAccountType = opBalance + CreditTotal - DebitTotal
End Select
Exit Function
TransIDFailed:
End Function


' Returns the AccountType from the Base Table
' Inputs :
'           HeadID as long
'
' Returns AccountType
'
Public Function GetAccountType(ByVal headID As Long) As wis_AccountType

' Handle Error
On Error GoTo Hell:
' Declare Variable
Dim rstAccountType As ADODB.Recordset
' Intialise the Function
GetAccountType = 0
' Validate the inputs
If headID = 0 Then Exit Function
' Fetch the MaxTransID from Acctrans AND FROM Cash Trans
gDbTrans.SqlStmt = " Select AccountType " & _
                   " FROM Heads a,ParentHeads b " & _
                   " WHERE a.ParentID=b.ParentID" & _
                   " AND a.HeadID=" & headID
If gDbTrans.Fetch(rstAccountType, adOpenForwardOnly) < 1 Then Exit Function
' Set getMaxtransid
GetAccountType = FormatField(rstAccountType.Fields(0))
Set rstAccountType = Nothing
Exit Function

Hell:
        Set rstAccountType = Nothing
End Function
' Returns the HeadName from the Base Table
' Inputs :
'           HeadID as long
'
' Returns HeadName as string
'
Public Function GetHeadName(ByVal headID As Long) As String

' Handle Error
On Error GoTo Hell:
' Declare Variable
Dim rstHeadName As ADODB.Recordset
' Intialise the Function
GetHeadName = ""
' Validate the inputs
If headID = 0 Then Exit Function
' Fetch the Head Name from Heads
gDbTrans.SqlStmt = " SELECT HeadName " & _
                   " FROM Heads" & _
                   " WHERE HeadID=" & headID
If gDbTrans.Fetch(rstHeadName, adOpenForwardOnly) < 1 Then Exit Function
' Set HeadName
GetHeadName = FormatField(rstHeadName.Fields(0))
Set rstHeadName = Nothing
Exit Function
Hell:
        Set rstHeadName = Nothing
End Function

'This function Saves the given data in the table
Public Function SaveVouchers(ByVal VoucherType As Wis_VoucherTypes, _
                             ByVal TransDate As Date, _
                             Optional ByVal Particulars As String, _
                             Optional ByVal isDataFromInventory As Boolean) _
                             As wis_FunctionReturned
    
' Handle Errors
On Error GoTo CouldNotSave:
' Declare Variables
Dim DrTotal As Currency
Dim CrTotal As Currency
Dim Balance As Currency
Dim UpdatedBalance As Currency
Dim Item As Long
Dim MaxCount As Long
Dim headID As Long

Dim ObjKeys As Variant

' Intialise the Function
SaveVouchers = Failure
' Validate the Inputs and Classes
'Need Not to validate 'SDA
If AllTrans Is Nothing Then Exit Function
If ValidateVouchers(VoucherType) <> Success Then Exit Function
' Check the Collection Dr Amount and Cr Amount Total and it should be same
If DrCrTotalsSame(False) > 0 Then Exit Function

' Get the MaxTransID from the Table
' If m_frmAccTrans.DbOperation = Insert Then TransID = GetMaxTransIDFromTable + 1
Dim TmpId As Long


'TransID = IIf(TransID = 0, GetMaxTransIDFromTable + 1, TransID)
If m_TransID <= 0 Then m_TransID = GetMaxTransIDFromTable + 1

' Please Don't Change the Case Creiteria
Select Case VoucherType
    Case payment, Sales, RejectionsIn
        If AllTransDRInsertToTable(VoucherType, TransDate) <> Success Then Exit Function
        If AllTransCRInsertToTable(VoucherType, TransDate) <> Success Then Exit Function
    Case Receipt, Purchase, RejectionsOut
        If AllTransCRInsertToTable(VoucherType, TransDate) <> Success Then Exit Function
        If AllTransDRInsertToTable(VoucherType, TransDate) <> Success Then Exit Function
    Case FreePurchase, FreeSales
        If AllTransDRInsertToTable(VoucherType, TransDate) <> Success Then Exit Function
        If AllTransCRInsertToTable(VoucherType, TransDate) <> Success Then Exit Function
    
    Case CONTRA
        If AllTransCRInsertToTable(VoucherType, TransDate) <> Success Then Exit Function
        If AllTransDRInsertToTable(VoucherType, TransDate) <> Success Then Exit Function
    Case Journal
        If AllTransCRInsertToTable(VoucherType, TransDate) <> Success Then Exit Function
        If AllTransDRInsertToTable(VoucherType, TransDate) <> Success Then Exit Function
End Select

' Particulars is given it will be inserted in the different table
If Len(Particulars) > 0 Then
    gDbTrans.SqlStmt = "INSERT INTO TransParticulars (TransID,Particulars)" & _
                     " VALUES(" & TransID & "," & AddQuotes(Particulars) & ")"
                     
    If Not gDbTrans.SQLExecute Then Exit Function
End If
If isDataFromInventory Then ' This will Keep Track of the Data's From Inventory Module
    gDbTrans.SqlStmt = " INSERT INTO IdFromInventory" & _
                     " (TransID) " & _
                     " VALUES ( " & TransID & ")"
    If Not gDbTrans.SQLExecute Then Exit Function
End If
SaveVouchers = Success

If m_frmAccTrans Is Nothing Then ClearCollections

Exit Function
CouldNotSave:
    SaveVouchers = FatalError
End Function

' This function Update the given data in the table
Public Function UpdateVouchers(ByVal VoucherType As Wis_VoucherTypes, ByVal TransDate As Date, Optional ByVal Particulars As String) As wis_FunctionReturned
' Handle Errors
On Error GoTo CouldNotSave:
' Declare Variables
Dim DrTotal As Currency
Dim CrTotal As Currency
Dim Balance As Currency
Dim UpdatedBalance As Currency
Dim Item As Long
Dim MaxCount As Long
Dim headID As Long
Dim ObjKeys As Variant

' Intialise the Function
UpdateVouchers = Failure

' Validate the Inputs and Classes
If AllTrans Is Nothing Then Exit Function
If ValidateVouchers(VoucherType) <> Success Then Exit Function
' Convert the TransDdate to US format

'TransID = TransDateDataExists(TransDate)
If TransID Then DeleteTransID
UpdateVouchers = SaveVouchers(VoucherType, TransDate)
'UpdateVouchers = Success


Exit Function

CouldNotSave:
    UpdateVouchers = FatalError
    
End Function
' This function Update the given data in the table
Public Function UndoVouchers(ByVal VoucherType As Wis_VoucherTypes, ByVal USTransDate As Date, Optional ByVal Particulars As String) As wis_FunctionReturned
' Handle Errors
On Error GoTo CouldNotSave:
' Declare Variables
Dim TransID As Long
' Intialise the Function
UndoVouchers = Failure
' Validate the Inputs and Classes
If AllTrans Is Nothing Then Exit Function
If ValidateVouchers(VoucherType) <> Success Then Exit Function
' Convert the TransDdate to US format
TransID = UndoTransDateData(USTransDate)
UndoVouchers = Success
ClearCollections
Exit Function
CouldNotSave:
    UndoVouchers = FatalError
End Function



'Returns the UpdatedBalance
' Inputs :
'           HeadID as long
'           Balance as currency
'           DrCrType as wis_DRCRType
'
' Retrurns Updated Balance
'
Private Function GetUpdatedBalance(ByVal headID As Long, ByVal Amount As Currency, ByVal Balance As Currency, DRCRType As wis_DrCrType) As Currency
' Handle Error
On Error GoTo Hell:
' Declare Variable
Dim rstHeadType As ADODB.Recordset
Dim AccountType As wis_AccountType
' Intialise the Function
GetUpdatedBalance = 0
' Validate the inputs
If headID = 0 Then Exit Function
'If Balance = 0 Then Exit Function
AccountType = GetAccountType(headID)
Select Case AccountType
    Case Loss
        If DRCRType = enumDebit Then GetUpdatedBalance = Balance + Amount
        If DRCRType = enumCredit Then GetUpdatedBalance = Balance - Amount
    Case Liability
        If DRCRType = enumDebit Then GetUpdatedBalance = Balance - Amount
        If DRCRType = enumCredit Then GetUpdatedBalance = Balance + Amount
    Case ItemSales
        If DRCRType = enumDebit Then GetUpdatedBalance = Balance - Amount
        If DRCRType = enumCredit Then GetUpdatedBalance = Balance + Amount
    Case Asset
        If DRCRType = enumDebit Then GetUpdatedBalance = Balance + Amount
        If DRCRType = enumCredit Then GetUpdatedBalance = Balance - Amount
    Case Profit
        If DRCRType = enumDebit Then GetUpdatedBalance = Balance - Amount
        If DRCRType = enumCredit Then GetUpdatedBalance = Balance + Amount
    Case ItemPurchase
        If DRCRType = enumDebit Then GetUpdatedBalance = Balance + Amount
        If DRCRType = enumCredit Then GetUpdatedBalance = Balance - Amount
End Select
Set rstHeadType = Nothing
Exit Function
Hell:
Set rstHeadType = Nothing
End Function

'This sub will add the given data to DrHeads Collection
Public Function AllTransHeadsAdd(ByVal headID As Long, ByVal DrAmount As Currency, ByVal CrAmount As Currency) As wis_FunctionReturned
On Error GoTo CouldNotAdd:
Dim ParentID As Long
AllTransHeadsAdd = Failure
' check the collection
If AllTrans Is Nothing Then Exit Function
    
' check if the headid exists in the both collection
If ALlTransHeadIDExists(headID) Then
    MsgBox "Head is already Exists in Collection !"
    Exit Function
End If

ParentID = GetParentID(headID)
Call AllTrans.Add(headID, DrAmount, CrAmount, ParentID, "Tr" & headID)
AllTransHeadsAdd = Success

Exit Function

CouldNotAdd:
    If Err.Number = 457 Then MsgBox "Head is already Exists in Collection !"
    AllTransHeadsAdd = FatalError

End Function
'This sub will add the given data to DrHeads Collection
Public Function AllTransHeadsUpdate(ByVal headID As Long, ByVal DrAmount As Currency, ByVal CrAmount As Currency) As wis_FunctionReturned
On Error GoTo CouldNotAdd:
Dim ParentID As Long
AllTransHeadsUpdate = Failure
' check the collection
If AllTrans Is Nothing Then Exit Function
If PrevHeadID = 0 Then Exit Function
ParentID = GetParentID(headID)
Call AllTrans.Update(headID, DrAmount, CrAmount, ParentID, PrevHeadID, "Tr" & headID)
AllTransHeadsUpdate = Success
Exit Function
CouldNotAdd:
    If Err.Number = 457 Then MsgBox " Head is already Exists in Collection !"
    AllTransHeadsUpdate = FatalError
End Function





'This sub will remove the based the given data given data to DrHeads Collection
Public Sub AllTransHeadIDRemove(ByVal headID As Long)
On Error GoTo CouldNotRemove:
If AllTrans Is Nothing Then Exit Sub
    
Call AllTrans.Remove("Tr" & headID)
Exit Sub
CouldNotRemove:
    If Err.Number = 5 Then MsgBox "Head does not Exist in the Collection!"
End Sub

'
Public Sub ShowAccTrans()
' set the frm to new
If m_frmAccTrans Is Nothing Then Set m_frmAccTrans = New frmAccTrans
' Load m_frmAccTrans
' Set the frm's grd to grid functions
If m_GrdFunctions Is Nothing Then Set m_GrdFunctions = New clsGrdFunctions
' set the grid
Set m_GrdFunctions.fGrd = m_frmAccTrans.grdLedger(0)
' by default insert
m_frmAccTrans.DBOperation = Insert
gWindowHandle = m_frmAccTrans.hwnd
m_frmAccTrans.Show vbModal
gWindowHandle = 0
Set m_frmAccTrans = Nothing
Set m_GrdFunctions = Nothing

End Sub

'
Public Sub SetAccTrans(acctrans As frmAccTrans)
' set the frm to new
Set m_frmAccTrans = acctrans
'Load m_frmAccTrans
' set the frm's grd to grid functions
Set m_GrdFunctions = New clsGrdFunctions
' set the grid
'Set m_GrdFunctions.fGrd = m_frmAccTrans.grdLedger(m_ActiveTab)
End Sub

'
'This function will update the base Collections
Private Function UpdateCollections() As wis_FunctionReturned
On Error GoTo Hell:
' Declarations
Dim Balance As Currency
Dim UpdatedBalance As Currency
Dim TransID As Long
Dim Amount As Currency
Dim DrCrEntry As wis_DrCrType
' Initialise the Function
UpdateCollections = Failure
' Check the Form
'If m_frmAccTrans Is Nothing Then Exit Function
UpdateCollections = AllTransHeadsAdd(headID, Debit, Credit)
'UpdateCollections = Success
Exit Function
Hell:
    UpdateCollections = FatalError
End Function

' This Function will update the Form's grid
' If Failed Retruns True
Private Function UpDateGrids() As wis_FunctionReturned
'Handle Errors
On Error GoTo Hell:
' Declarations
Dim UpdatedBalance As Currency
Dim Amount As Currency
' Initialise the Function
UpDateGrids = Failure
' Check the Form
If m_frmAccTrans Is Nothing Then Exit Function
If m_GrdFunctions Is Nothing Then Exit Function
If m_GrdFunctions.fGrd Is Nothing Then Set m_GrdFunctions.fGrd = m_frmAccTrans.grdLedger(TransGrid)

UpdatedBalance = GetClosingBalance(headID, Balance, Debit, Credit)

With m_frmAccTrans.grdLedger(TransGrid)
    If UpdatingGrid Then
        SlNo = .Row
    Else
        SlNo = SlNo + 1
    End If
    m_GrdFunctions.fMoreRows (2)
    .Row = SlNo
    .RowData(.Row) = headID
    .Col = 0: .Text = SlNo
    .Col = 1: .Text = headName
    .Col = 2: .Text = FormatCurrency(Debit)
    .Col = 3: .Text = FormatCurrency(Credit)
    .Col = 4: .Text = FormatCurrency(UpdatedBalance)
End With

UpDateGrids = Success

Exit Function

Hell:

    UpDateGrids = FatalError
    
End Function

'This function will validate the all the controls for the from
Private Function ValidateLedgers() As wis_FunctionReturned
' Handle the Error
On Error GoTo Hell:
' Declarations
Dim StrAmount As String
' Intialise the function
ValidateLedgers = Failure
' Validate the From
If m_frmAccTrans Is Nothing Then Exit Function
With m_frmAccTrans
    If .cmbTab0Ledger(VoucherCombo).ListIndex = -1 Then Err.Raise vbObjectError + 513, , "Select the Voucher"
    'Check for the Heads Selected
    If .cmbTab0Ledger(ParentCombo).ListIndex = -1 Then Err.Raise vbObjectError + 513, , "Select the Parent Head"
    If .cmbTab0Ledger(LedgerCombo).ListIndex = -1 Then Err.Raise vbObjectError + 513, , "Select the Ledger Head"
    ' Check the Validate of the Date
    If Not TextBoxDateValidate(.txtTab0CurrentDate, "/", True, True) Then Exit Function
    StrAmount = .txtTab0Amount.Value
    'If Not TextBoxCurrencyValidate(.txtTab0Amount,  False, True) Then Exit Function
    'now update the Propertries
    Credit = 0
    Debit = 0
    headID = 0
    headName = ""
    Balance = 0
    TransDate = ""
    If .optTab0From(FromOpt).Enabled Then
        If .optTab0From(FromOpt) Then Debit = Val(StrAmount)
        If .optTab0From(ToOpt) Then Credit = Val(StrAmount)
    Else
        If .optTab0From(FromOpt) Then Credit = Val(StrAmount)
        If .optTab0From(ToOpt) Then Debit = Val(StrAmount)
    End If
    ' get the id and headname
    With .cmbTab0Ledger(LedgerCombo)
        headID = .ItemData(.ListIndex)
        headName = .Text
    End With
    TransDate = .txtTab0CurrentDate.Text
    Balance = Val(.lblTab0Balance(0).Caption)
End With
If Debit > 0 Or Credit > 0 Then ValidateLedgers = Success
Exit Function
Hell:
    ValidateLedgers = FatalError
    'Resume
End Function
'This function will validate the all the controls for the from
Private Function ValidateLedgersNew() As wis_FunctionReturned
' Handle the Error
On Error GoTo Hell:
' Declarations
Dim StrAmount As String
Dim strMsg As String
' Intialise the function
ValidateLedgersNew = Failure
' Validate the From
If m_frmAccTrans Is Nothing Then Exit Function
With m_frmAccTrans
    If .cmbTab0Ledger(VoucherCombo).ListIndex = -1 Then Err.Raise vbObjectError + 513, , "Select the Voucher"
    If .cmbTab0Ledger(ParentCombo).ListIndex = -1 Then Err.Raise vbObjectError + 513, , "Select the Parent Head"
    If .cmbTab0Ledger(LedgerCombo).ListIndex = -1 Then Err.Raise vbObjectError + 513, , "Select the Ledger Head"
    ' Check the Validate of the Date
    If Not TextBoxDateValidate(.txtTab0CurrentDate, "/", True, True) Then Exit Function
    'If Not TextBoxCurrencyValidate(.txtTab0Amount, True, True) Then Exit Function
    'now update the Propertries
    Credit = 0
    Debit = 0
    headID = 0
    headName = ""
    Balance = 0
    TransDate = ""
    If .optTab0From(FromOpt) Then Credit = Val(StrAmount)
    If .optTab0From(ToOpt) Then Debit = Val(StrAmount)
    ' get the id and headname
    With .cmbTab0Ledger(LedgerCombo)
        headID = .ItemData(.ListIndex)
        headName = .Text
    End With
    TransDate = .txtTab0CurrentDate.Text
    Balance = Val(.lblTab0Balance(0).Caption)
End With
ValidateLedgersNew = Success
Exit Function

Hell:
        
    If Err Then MsgBox "ValidateLedgersNew : " & vbCrLf & Err.Description
    
    ValidateLedgersNew = FatalError
    
End Function
' This function returns the ParentName from the given Headid
' Input is Headid as long
' Returns ParentName string
'
' Pradeep
'
Public Function GetParentName(ByVal headID As Long) As String
' Handle Error
On Error GoTo NoParentName:
' Declare Variables
Dim rstParentName As ADODB.Recordset
' Intialiase the Variable
GetParentName = ""
' Check the Input Received if Zero then Exit
If headID = 0 Then Exit Function
' set the sqlstmt
gDbTrans.SqlStmt = " SELECT ParentName " & _
                   " FROM Heads a,ParentHeads b " & _
                   " WHERE a.HeadID=" & headID & _
                   " AND a.ParentID=b.ParentID"
' Now fetch the record
If gDbTrans.Fetch(rstParentName, adOpenForwardOnly) < 1 Then Exit Function
' Here is the ParentID!
GetParentName = FormatField(rstParentName.Fields("ParentName"))
Set rstParentName = Nothing
Exit Function
NoParentName:
    Set rstParentName = Nothing
End Function

Public Property Get ObjectHandle() As Long
    If Not m_frmAccTrans Is Nothing Then _
            ObjectHandle = m_frmAccTrans.hwnd
End Property

Private Sub Class_Terminate()
On Error Resume Next
    Set m_AllTrans = Nothing
    Set m_frmAccTrans = Nothing
End Sub

'
Public Sub frmAddClicked()
If m_frmAccTrans Is Nothing Then Exit Sub
Call AddProcedures
End Sub



' This fetch the data from Base table and load it to Grid
'
Public Sub ShowNewLedgerToGrid(ByVal headID As Long, grd As MSFlexGrid, _
                           DateToConsider As Boolean, _
                           fromDate As String, toDate As String)
' Handle Error
On Error GoTo Hell:
' Declarations
Dim rstLedger As ADODB.Recordset
Dim rstTrans As ADODB.Recordset
Dim rstParticulars As ADODB.Recordset
Dim headName As String
Dim TransDate As String
Dim VoucherName As String
Dim AccountType As wis_AccountType
Dim VoucherType As Wis_VoucherTypes
Dim Debit As Currency
Dim Credit As Currency
Dim opBalance As Currency
Dim DebitTotal As Currency
Dim CreditTotal As Currency
Dim ClosingBalance As Currency
Dim HeadIDDebit As Currency
Dim HeadIDCredit As Currency
Dim SlNo As Long
Dim TransID As Long
Dim ToHeadID As Long
Dim PrevTransID As Long
Dim Particulars As String
Dim LedgerClass As clsLedger
' Validations
If headID = 0 Then Exit Sub
' Set the Grid
Set m_GrdFunctions.fGrd = grd
Set LedgerClass = New clsLedger
grd.Clear
' Refresh the grid for the Ledger
Call InitTab1Grid(grd)
' First put the Opening Balance
If DateToConsider Then opBalance = GetOpBalance(headID, GetSysFormatDate(fromDate))
If Not DateToConsider Then opBalance = GetOpBalance(headID, FinUSFromDate)
AccountType = GetAccountType(headID)
grd.Row = 1
Call SetHeaderFooter(opBalance, grd, AccountType, GetResourceString(284)) '"Opening Balance")
' Set the Sql Statement
StartTimer
If DateToConsider Then _
    gDbTrans.SqlStmt = " SELECT TransID,Debit,Credit,VoucherType" & _
                        " FROM AccTrans " & _
                        " WHERE HeadID = " & headID & _
                        " AND TransDate BETWEEN #" & GetSysFormatDate(fromDate) & "#" & _
                        " AND #" & GetSysFormatDate(toDate) & "#" & _
                        " ORDER BY TransDate,TransID"
If Not DateToConsider Then _
    gDbTrans.SqlStmt = " SELECT TransID,Debit,Credit,TransDate,VoucherType" & _
                        " FROM AccTrans " & _
                        " WHERE HeadID = " & headID & _
                        " ORDER BY TransDate,TransID"
If Not gDbTrans.CreateView("QryLedger") Then Exit Sub
Dim RowsToEnter As Long
gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM QryLedger"
'Fetch the Data to Recordset based on the Condition
RowsToEnter = gDbTrans.Fetch(rstTrans, adOpenStatic)
gDbTrans.SqlStmt = " SELECT a.Debit,a.Credit,a.VoucherType,a.TransID,a.TransDate,c.HeadName" & _
                    " FROM AccTrans a, QryLedger b,Heads c " & _
                    " WHERE a.HeadID <> " & headID & _
                    " AND a.TransID = b.TransID" & _
                    " AND a.HeadID = c.HeadID"
Call gDbTrans.Fetch(rstLedger, adOpenStatic)

'gDbTrans.SQLStmt = " SELECT a.TransID,Particulars " & _
                   " FROM TransParticulars a , QryLedger b" & _
                   " WHERE a.TransID=b.TransID" & _
                   " ORDER BY a.TransID"
                   
'Call gDbTrans.Fetch(rstParticulars, adOpenStatic)

' Kept for the Reference
'
''    .Col = 0: .Text = "Date": .CellFontBold = True
''    .Col = 1: .Text = "Ledger Name": .CellFontBold = True
''    .Col = 2: .Text = "Voucher": .CellFontBold = True
''    .Col = 3: .Text = "Dr ": .CellFontBold = True
''    .Col = 4: .Text = "Cr ": .CellFontBold = True
''    .Col = 5: .Text = "Balance ": .CellFontBold = True

SlNo = grd.Row + 1
DebitTotal = 0
CreditTotal = 0
PrevTransID = 0
' This will create bulk rows
If RowsToEnter > 0 Then m_GrdFunctions.fMoreRows (RowsToEnter + 2)
Do While Not rstTrans.EOF
     TransID = rstTrans.Fields("TransID")
     ' This will fetch the particulars if any
     HeadIDDebit = rstTrans.Fields("Debit")
     HeadIDCredit = rstTrans.Fields("Credit")
     rstLedger.Filter = adFilterNone
     If HeadIDDebit > 0 Then
         rstLedger.Filter = "TransID = " & TransID '& " AND Credit > " & 0
     Else
         rstLedger.Filter = "TransID = " & TransID '& " AND Debit > " & 0
     End If
     RowsToEnter = rstLedger.recordCount
     ' This will update the bulk rows
     m_GrdFunctions.fMoreRows (RowsToEnter + 2)
     If RowsToEnter > 1 Then
        Debit = 0
        Credit = 0
        headName = "As Per Details"
        VoucherType = rstLedger.Fields("VoucherType")
        Debit = HeadIDDebit
        Credit = HeadIDCredit
        VoucherName = LedgerClass.GetVoucherName(VoucherType)
        TransDate = FormatField(rstLedger.Fields("TransDate"))
        DebitTotal = DebitTotal + Debit
        CreditTotal = CreditTotal + Credit
        ClosingBalance = GetClosingBalanceonAccountType(opBalance, DebitTotal, CreditTotal, AccountType)
        With grd
            'Call m_GrdFunctions.fMoreRows(2)
            SlNo = SlNo + 1
            .Row = SlNo
            .RowData(SlNo) = TransID
            .Col = 0: .Text = TransDate
            .Col = 1: .Text = headName ': .CellFontBold = True
            .Col = 2: .Text = VoucherName
            .Col = 3: .Text = FormatCurrency(Debit)
            .Col = 4: .Text = FormatCurrency(Credit)
            .Col = 5: .Text = FormatCurrency(ClosingBalance)
        End With
        Do While Not rstLedger.EOF
            Debit = 0
            Credit = 0
            headName = rstLedger.Fields("HeadName")
            VoucherType = rstLedger.Fields("VoucherType")
            Debit = rstLedger.Fields("Credit")
            Credit = rstLedger.Fields("Debit")
            VoucherName = LedgerClass.GetVoucherName(VoucherType)
            If Debit > 0 Then headName = headName & " " & Debit & " " & "Dr"
            If Credit > 0 Then headName = headName & " " & Credit & " " & "Cr"
            With grd
                'Call m_GrdFunctions.fMoreRows(2)
                SlNo = SlNo + 1
                .Row = SlNo
                .RowData(SlNo) = TransID
                .Col = 1: .Text = headName ': .CellFontBold = True
            End With
            rstLedger.MoveNext
        Loop
    Else
        Debit = 0
        Credit = 0
        headName = rstLedger.Fields("HeadName")
        VoucherType = rstLedger.Fields("VoucherType")
        Debit = rstLedger.Fields("Credit")
        Credit = rstLedger.Fields("Debit")
        VoucherName = LedgerClass.GetVoucherName(VoucherType)
        TransDate = FormatField(rstLedger.Fields("TransDate"))
        DebitTotal = DebitTotal + Debit
        CreditTotal = CreditTotal + Credit
        ClosingBalance = GetClosingBalanceonAccountType(opBalance, DebitTotal, CreditTotal, AccountType)
        With grd
            'Call m_GrdFunctions.fMoreRows(2)
            SlNo = SlNo + 1
            .Row = SlNo
            'If VoucherType = Contra Then
                'Check for the Index Heads If it has Index head s then Need not
                'to delete it
            Dim rstTemp As Recordset
            gDbTrans.SqlStmt = "SELECT HEADID FROM HEADS WHERE HEADID " & _
                " IN (SELECT Distinct HeadID FROM ACCTrans where TransID = " & TransID & " )"
            If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then
                If Not rstTemp.EOF And Not rstTemp.BOF Then .RowData(SlNo) = TransID
            End If
            'End If
            .Col = 0: .Text = TransDate
            .Col = 1: .Text = headName ': .CellFontBold = True
            .Col = 2: .Text = VoucherName
            .Col = 3: .Text = FormatCurrency(Debit)
            .Col = 4: .Text = FormatCurrency(Credit)
            .Col = 5: .Text = FormatCurrency(ClosingBalance)
        End With
    End If
    rstTrans.MoveNext
Loop
m_GrdFunctions.fMoreRows (2)
grd.Row = grd.Row + 2
ClosingBalance = GetClosingBalanceonAccountType(opBalance, DebitTotal, CreditTotal, AccountType)
Call SetHeaderFooter(ClosingBalance, grd, AccountType, GetResourceString(285))
StopTimer
Set LedgerClass = Nothing
Set rstLedger = Nothing
Set rstTrans = Nothing
Set rstParticulars = Nothing

Exit Sub
Hell:
    'Resume
    Debug.Assert Err.Number = 0
    Set LedgerClass = Nothing
    Set rstLedger = Nothing
    Set rstTrans = Nothing
    Set rstParticulars = Nothing
End Sub

Public Sub frmGridClicked(ByVal RowNum As Integer, ByVal TabIndex As Integer)

If m_frmAccTrans Is Nothing Then Exit Sub


If TabIndex = 0 Then
    'If this transcationis of cash then need not to update this so exit
    If m_frmAccTrans.grdLedger(TabIndex).RowData(RowNum) = wis_CashHeadID Then
        If (gCurrUser.UserPermissions And perOnlyWaves) = 0 Then Exit Sub
    End If
    LoadGridFromFrm (RowNum)
ElseIf TabIndex = 1 Then
    'If the Clicked Details is of Bank Heads Then Do Not Update it
    
    LoadDataToUpload (RowNum)
End If

End Sub

' This will return adodb.recordset if transid has mulitple credit and debit heads
' takes TransID as Long as Argument

Private Function GetMultiHeads() As ADODB.Recordset

On Error GoTo Hell:

Dim rstMultiHeads As ADODB.Recordset

Dim QryNametoFetch As String

Set GetMultiHeads = Nothing

' This will prepare queries
QryNametoFetch = CreateViewsforMultiHeads

' This means creation of qry is failed
If QryNametoFetch = "" Then Exit Function
    
gDbTrans.SqlStmt = " SELECT * FROM " & QryNametoFetch

If gDbTrans.Fetch(rstMultiHeads, adOpenStatic) = -1 Then Exit Function

Set GetMultiHeads = rstMultiHeads

Set rstMultiHeads = Nothing

Exit Function

Hell:
    
    Set rstMultiHeads = Nothing

End Function


Private Function CreateViewsforShowLedger(ByVal DateToConsider As Boolean, _
                                          ByRef rstLedger As ADODB.Recordset, _
                                          ByVal fromDate As String, _
                                          ByVal toDate As String, _
                                          ByVal headID As Long) As Long
' Set the Sql Statement

If DateToConsider Then
        
        
    gDbTrans.SqlStmt = " SELECT a.Headid, Debit, Credit, " & _
                            " TransID, VoucherType, Headname, TransDate " & _
                      " FROM qryAccTransMerge AS a, Heads AS b " & _
                      " WHERE a.HeadID = " & headID & _
                      " AND a.HeadID = b.HeadID " & _
                      " AND a.Debit=0" & _
                      " AND TransDate BETWEEN #" & GetSysFormatDate(fromDate) & "#" & _
                      " AND #" & GetSysFormatDate(toDate) & "#"
        
    gDbTrans.CreateView ("QryCrHeads")
    
    gDbTrans.SqlStmt = " SELECT a.Headid, Debit, Credit, " & _
                            " TransID, VoucherType, Headname, TransDate " & _
                      " FROM qryAccTransMerge AS a, Heads AS b " & _
                      " WHERE a.HeadID = " & headID & _
                      " AND a.HeadID = b.HeadID " & _
                      " AND a.Credit=0" & _
                      " AND TransDate BETWEEN #" & GetSysFormatDate(fromDate) & "#" & _
                      " AND #" & GetSysFormatDate(toDate) & "#"
        
    gDbTrans.CreateView ("QryDrHeads")
    
End If

If Not DateToConsider Then
        
    gDbTrans.SqlStmt = " SELECT CrHeadid AS HeadID, 0 AS Debit, Amount AS Credit, " & _
                            " TransID, VoucherType, Headname, TransDate " & _
                      " FROM AccTrans AS a, Heads AS b " & _
                      " WHERE DrHeadID = " & headID & _
                      " AND a.CrHeadID = b.HeadID "
        
    gDbTrans.CreateView ("QryCrHeads")
    
    
    gDbTrans.SqlStmt = " SELECT DrHeadid AS HeadID, Amount AS Debit, 0 AS Credit, " & _
                            " TransID, VoucherType, Headname, TransDate " & _
                      " FROM AccTrans AS a, Heads AS b " & _
                      " WHERE CrHeadID = " & headID & _
                      " AND a.DrHeadID = b.HeadID "

    gDbTrans.CreateView ("QryDrHeads")
    
End If
                        
gDbTrans.SqlStmt = " SELECT * FROM qryCrHeads " & _
                   " UNION ALL " & _
                   " SELECT * FROM qryDrHeads "
                   
If Not gDbTrans.CreateProcedures("QryLedger") Then Exit Function
    

gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM QryLedger" & _
                   " ORDER BY TransDate,TransID"
                
'Fetch the Data to Recordset based on the Condition
CreateViewsforShowLedger = gDbTrans.Fetch(rstLedger, adOpenStatic)

End Function

Public Sub frmHeadClicked(ByVal headID As Long)

' Handle Errors
On Error GoTo Hell:

' Declare Variables
Dim Balance As Currency

' Validate the Inputs

If headID = 0 Then Exit Sub
If m_frmAccTrans Is Nothing Then Exit Sub

If Not DateValidate(m_frmAccTrans.txtTab0CurrentDate.Text, "/", True) Then Exit Sub

With m_frmAccTrans
    'this will fetch the balance from the base table
    Balance = GetBalance(headID, GetSysFormatDate(.txtTab0CurrentDate.Text))
    .lblTab0Balance(0).Caption = FormatCurrency(Balance)
    'If The TransCtion Is of sales then allow him to enter the Closing Balance
    .cmdCBStock.Visible = IIf(GetParentID(headID) = parSales, True, False)
    
End With

Exit Sub

Hell:
    
End Sub


Public Function frmOKClicked() As wis_FunctionReturned

' Handle Errors
On Error GoTo Hell:

' Declare Variable
Dim TransDate As Date
Dim VoucherType As Wis_VoucherTypes
Dim Particulars As String
Dim Amount As Currency
Dim BlUpdate As Boolean
Dim StockValue As Currency
Dim StockId As Long
Dim rstTemp As Recordset


If m_frmAccTrans.DBOperation = Update Then BlUpdate = True Else BlUpdate = False
If m_GrdFunctions.fGrd Is Nothing Then Set m_GrdFunctions.fGrd = m_frmAccTrans.grdLedger(TransGrid)
If m_frmAccTrans Is Nothing Then Exit Function

frmOKClicked = Failure

Dim WavesUser As Boolean
WavesUser = (gCurrUser.UserPermissions And perOnlyWaves)

If IsCashTransaction And Not WavesUser Then
    Amount = m_frmAccTrans.txtTab0Amount
    'Check for the Stock Value
    With m_frmAccTrans
        If .lblTab0Balance(0).Tag = "1" Then
            .lblTab0Balance(0).Tag = ""
            StockValue = Val(.lblTab0Balance(0).Caption)
            StockId = .cmbTab0Ledger(LedgerCombo).ItemData(.cmbTab0Ledger(LedgerCombo).ListIndex)
        End If
    End With
    
    If BlUpdate Then
        If UpdateProcedures <> Success Then GoTo Hell
        'Now load the Update Cash Transaction Details
        With m_frmAccTrans
            With .grdLedger(TransGrid)
                .Row = IIf(.Row = 1, 2, 1)
                'Now Load the Cash TransCtion Details
                LoadGridFromFrm (.Row)
            End With
            .txtTab0Amount = Amount
        End With
    Else
        If AddProcedures <> Success Then GoTo Hell
    End If

    'Now Update the Cash Head Id
    If BlUpdate Then UpdateProcedures Else AddProcedures
End If

With m_frmAccTrans
    TransDate = GetSysFormatDate(.txtTab0CurrentDate.Text)
    VoucherType = .cmbTab0Ledger(VoucherCombo).ItemData(.cmbTab0Ledger(VoucherCombo).ListIndex)
    Particulars = .txtParticulars.Text
    .grdLedger(TransGrid).Enabled = True
End With

If (VoucherType = Receipt Or VoucherType = payment) And IsCashTransaction Then
    'Check Whether there is a transaction Between the
    'Bank Accounts and cash
    
End If

' Check the Incomplete Transactions
If DrCrTotalsSame(False) > 0 Then Exit Function

If AllTransVoucherCheck(VoucherType) <> Success Then Exit Function


' Save Vocuhers
gDbTrans.BeginTrans

'First Delete the TransID in case of updation
If BlUpdate Then DeleteTransID
If SaveVouchers(VoucherType, TransDate, Particulars) <> Success Then _
        gDbTrans.RollBacknRaiseError

'If the Transction is of sales then insert the Closing Stock
EnterStock:
If StockValue > 0 Then

    Dim InTrans As Boolean
'    InTrans = gDbTrans.BeginTrans
    gDbTrans.SqlStmt = "select opId From CloseStock " & _
                " WHere TransDate = #" & TransDate & "#" & _
                " ANd HeadID = " & StockId
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) Then
        gDbTrans.SqlStmt = "Update CloseStock " & _
                " Set StockValue = " & StockValue & _
                " WHere opID = " & rstTemp(0)
    
    Else
        gDbTrans.SqlStmt = "Insert Into CloseStock " & _
            " (TransDate,HeadID,StockValue)  VALUES " & _
            " ( #" & TransDate & "#," & StockId & "," & StockValue & ")"
    End If
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
'    If InTrans Then gDbTrans.CommitTrans: Exit Function
End If
gDbTrans.CommitTrans

'Checking prg can be deleted - pradeep
CheckTransId

If IsCashTransaction <> Success Then
    'If IsCashTransaction <> Success Then ClearCollections
    ClearCollections
    ClearAccTransControls
    ' With the Collections
    headID = 0
    SlNo = 0
    Debit = 0
    Credit = 0
    headName = ""
    TransID = 0
    UpdatingTrans = False
Else
    UpdatingTrans = True
End If

' Send Message
MsgBox "Saved the Transactions"

frmOKClicked = Success

Exit Function

Hell:
    MsgBox "Could Not Update Transactions"
    frmOKClicked = FatalError
    
End Function


Private Sub LoadDataToUpload(ByVal RowNum As Integer)

' Handle Error
On Error GoTo Hell:

' Declarations
Dim rstTrans As ADODB.Recordset

Dim VoucherType As Wis_VoucherTypes

Dim Item As Integer
Dim count As Integer

Dim lngTransID As Long
Dim Particulars As String

Dim LedgerClass As clsLedger

If RowNum = 0 Then Exit Sub
   
With m_frmAccTrans.grdLedger(LedgerGrid)
    .Row = RowNum
    ' Get the Transaction Id of the Selected trans from the Row
    lngTransID = .RowData(RowNum)
End With

If lngTransID = 0 Then Exit Sub

'Before Loading the Transction Details To the TransGrid
'Clear the All account which have half transction
'Or loaded any Previous transction
Call m_frmAccTrans.RefreshForTab0
Call ClearCollections

' if data is from Inventory then it can not be edited
If IsTransIDFromInventory(lngTransID) Then Exit Sub
'Check For the bank head transction
If IsBankHeadTransaction(lngTransID) Then
    If (gCurrUser.UserPermissions And perOnlyWaves) = 0 Then Exit Sub
End If
' Get the Records to Recordset
gDbTrans.SqlStmt = " SELECT * FROM AccTrans " & _
                " WHERE TransID = " & lngTransID

If gDbTrans.Fetch(rstTrans, adOpenDynamic) < 1 Then Exit Sub

With m_frmAccTrans
    .tabTrans.Tabs(1).Selected = True
    Set m_GrdFunctions.fGrd = m_frmAccTrans.grdLedger(0)
    
    .txtTab0CurrentDate.Text = FormatField(rstTrans.Fields("TransDate"))
    VoucherType = FormatField(rstTrans.Fields("VoucherType"))
    With .cmbTab0Ledger(VoucherCombo)
        count = .ListCount - 1
        For Item = 0 To count
            If .ItemData(Item) = VoucherType Then
                .ListIndex = Item
                Exit For
            End If
        Next Item
    End With
End With

SlNo = 0

Dim l_ParentID As Long

' Start the Loop
Do While Not rstTrans.EOF
    ' Feed data to Controls
    headID = FormatField(rstTrans.Fields("HeadID"))
    
    Debit = FormatField(rstTrans.Fields("Debit"))
    Credit = FormatField(rstTrans.Fields("Credit"))
    TransDate = FormatField(rstTrans.Fields("TransDate"))
    
    headName = GetHeadName(headID)
    Balance = GetBalance(headID, rstTrans.Fields("TransDate"), rstTrans.Fields("TransID"))
    
    ' Update the Grid
    If UpDateGrids <> Success Then Exit Sub
    'Update the Collections
    If UpdateCollections <> Success Then Exit Sub
    
    rstTrans.MoveNext
Loop

' Get the Particualrs
Set LedgerClass = New clsLedger
Particulars = LedgerClass.GetTransIDParticulars(lngTransID)
Set LedgerClass = Nothing

With m_frmAccTrans
    If Len(Particulars) > 0 Then .txtParticulars.Text = Particulars
    .txtTab0Amount.Value = DrCrTotalsSame(False)
    
    If .txtTab0Amount.Value = 0 Then
        If AllTrans.count > 0 Then
            .cmdTab0AddtoGrid(AddToGrid).Enabled = False
            .cmdOk.Enabled = True
            .cmbTab0Ledger(LedgerCombo).Clear
        End If
    End If
    
    .DBOperation = Update
    UpdatingTrans = True
End With

TransID = lngTransID

Set rstTrans = Nothing

' Exit
Exit Sub

Hell:
    
    Set rstTrans = Nothing
    Debug.Assert Err.Number <> 0
    'Resume

End Sub

Public Function frmUpdateClicked() As wis_FunctionReturned

If m_frmAccTrans Is Nothing Then Exit Function

frmUpdateClicked = Failure
'If HeadID = wis_CashHeadID Then
If IsCashTransaction Then
    frmUpdateClicked = Success
Else
    If UpdateProcedures = Success Then frmUpdateClicked = Success
End If

End Function


Public Sub frmVoucherClicked()

If m_frmAccTrans Is Nothing Then Exit Sub

Call HandleVouchers

End Sub

'This prepare the database to return the recordset of multiple heads
' it will create the conserned quiries
Private Function CreateViewsforMultiHeads() As String

On Error GoTo Hell:

CreateViewsforMultiHeads = ""

gDbTrans.SqlStmt = " SELECT HeadID as CrHeadID, TransID " & _
                 " FROM AccTrans " & _
                 " WHERE Credit > 0 " & _
                 " ORDER BY TRANSID"
                 
If Not gDbTrans.CreateView("QryCreditHeads") Then Exit Function

gDbTrans.SqlStmt = " SELECT HeadID as DrHeadID, TransID " & _
                 " FROM AccTrans " & _
                 " WHERE Debit > 0 " & _
                 " ORDER BY TRANSID"
                 
If Not gDbTrans.CreateView("QryDebitHeads") Then Exit Function

gDbTrans.SqlStmt = " SELECT DISTINCT * " & _
                   " FROM  QryCreditHeads A ,QryDebitHeads B " & _
                   " WHERE A.TransID = B.TransID " '& _
                   " AND A.CrHeadID=B.DrHeadID "

If Not gDbTrans.CreateView("QryMultiHeads") Then Exit Function

CreateViewsforMultiHeads = "QryMultiHeads"

Exit Function

Hell:
    
End Function



