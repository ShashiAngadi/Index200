VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransferNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Constants
Private Const TRANSFER_TRANSACTION As Integer = 0
Private Const TRANSFER_SHOW As Integer = 1
'set this to 0 to disable debug code in this class
#Const DebugMode = 0


#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    Private mlClassDebugID As Long
#End If

Private WithEvents m_frmTransfer As frmTransferNew
Attribute m_frmTransfer.VB_VarHelpID = -1

Private m_colTransfers As colTransfers
Private m_GrdFunctions As clsGrdFunctions

'local variable(s) to hold property value(s)
Private m_GodownID As Byte 'local copy
Private m_HeadID As Long 'local copy
Private m_HeadName As String 'local copy
Private m_RelationID As Long 'local copy
Private m_FromQuantity As Double 'local copy
Private m_ToQuantity As Double 'local copy
Private m_GodownName As String 'local copy
Private m_Balance As Double
Private m_TransID As Long 'local copy

Private m_TransDate As String
Private m_ToQtyBeforeUpdation As Double


Public Property Get ObjectHandle()
    If Not m_frmTransfer Is Nothing Then _
            ObjectHandle = m_frmTransfer.hwnd
End Property

Public Property Get ObjectName()
    ObjectName = "TransferClass"
End Property


Private Function PrepareRstForShowTransferDetails(ByRef rstTransfer As ADODB.Recordset, ByVal FromDateUS As String, ByVal ToDateUS As String) As Boolean
'Declare the variables
Dim StockInVoucher As Wis_VoucherTypes
Dim StockOutVoucher As Wis_VoucherTypes

StockInVoucher = StockIn
StockOutVoucher = StockOut

gDbTrans.SqlStmt = " SELECT RelationID,TransID,GodownID,Quantity,VoucherType,TransDate" & _
                   " FROM Stock" & _
                   " WHERE TransDate BETWEEN #" & FromDateUS & "#" & _
                   " AND #" & ToDateUS & "#" & _
                   " AND ( VoucherType =" & StockInVoucher & _
                   " OR VoucherType = " & StockOutVoucher & " )" & _
                   " ORDER BY TransDate,TransID ASC,VoucherType DESC"
                   
                   
gDbTrans.CreateView ("QryTransferDetails")

gDbTrans.SqlStmt = " SELECT *" & _
                  " FROM QryTransferDetails"
If gDbTrans.Fetch(rstTransfer, adOpenForwardOnly) < 1 Then Exit Function

PrepareRstForShowTransferDetails = True

End Function

Public Sub Show()
    
    On Error GoTo ShowErr
    
    If m_frmTransfer Is Nothing Then Set m_frmTransfer = New frmTransferNew
    If m_GrdFunctions Is Nothing Then Set m_GrdFunctions = New clsGrdFunctions
 
    'your code goes here...
    If m_GrdFunctions Is Nothing Then Exit Sub
    
    gWindowHandle = m_frmTransfer.hwnd
    m_frmTransfer.Show 'vbModal

    Exit Sub


ShowErr:
    Call RaiseError(MyUnhandledError, "clsTransferNew:Show Method")
End Sub

Private Sub LoadDetailsToTransactionGrid()
'Declare the Variables
Dim rstTransaction As ADODB.Recordset
Dim rstHeadName As ADODB.Recordset
Dim MaterialClass As clsMaterial

Dim SelectedRow As Integer

Dim VoucherType As Wis_VoucherTypes

Dim StockBalance As Double
Dim ProductName As String

'Setup an error handler...
On Error GoTo ErrLine

With m_frmTransfer.grd(TRANSFER_SHOW)
    SelectedRow = .Row
    m_TransID = .RowData(SelectedRow)
End With

If m_TransID = 0 Then Exit Sub
With m_frmTransfer
    .tabTransfer.Tabs(1).Selected = True
    .InitGridForTab0
    .cmbBranch.ListIndex = -1
    .cmbGroup.ListIndex = -1
    .cmbUnit.ListIndex = -1
End With




Call CreateTransactionView("QryTransaction", m_TransID)

gDbTrans.SqlStmt = " SELECT * FROM QryTransaction"

If gDbTrans.Fetch(rstTransaction, adOpenForwardOnly) < 1 Then Exit Sub
    
    
gDbTrans.SqlStmt = " SELECT DISTINCT A.RelationID,A.GodownID,GodownName,ProductName,UnitName" & _
                   " FROM RelationMaster A,QryTransaction B," & _
                        " GodownDet C,Products D,Units E" & _
                   " WHERE A.RelationID=B.RelationID" & _
                   " AND A.GodownID=B.GodownID" & _
                   " AND A.GodownID=C.GodownID" & _
                   " AND A.ProductID=D.ProductID" & _
                   " AND A.UnitID=E.UnitID" & _
                   " AND A.GroupID=D.GroupID"
                   
                   
Call gDbTrans.Fetch(rstHeadName, adOpenForwardOnly)

Set MaterialClass = New clsMaterial

Set m_colTransfers = New colTransfers

m_TransDate = ""

With m_frmTransfer.grd(TRANSFER_TRANSACTION)
    Do While Not rstTransaction.EOF
        
        m_RelationID = rstTransaction.Fields("RelationID")
        VoucherType = rstTransaction.Fields("VoucherType")
        m_TransDate = FormatField(rstTransaction.Fields("TransDate"))
        
        If gDbTrans.FindRecord(rstHeadName, "RelationID=" & RelationID) Then
            m_GodownID = rstHeadName.Fields("GodownID")
            m_GodownName = rstHeadName.Fields("GodownName")
            ProductName = rstHeadName.Fields("ProductName") & " " & rstHeadName.Fields("UnitName")
            StockBalance = MaterialClass.GetItemOnDateClosingStock(RelationID, m_TransDate, GodownID)
        End If
        
        Select Case VoucherType
            Case StockIn
                m_ToQuantity = FormatField(rstTransaction.Fields("Quantity"))
                If .Rows <= .Row + 1 Then .Rows = .Rows + 1
                .Row = .Row + 1
                .Col = 0: .Text = "To " & GodownName
                .Col = 1: .Text = ProductName: .RowData(.Row) = RelationID
                .Col = 2: .Text = m_ToQuantity
                .Col = 3: .Text = StockBalance
                Call m_colTransfers.Add(m_GodownID, m_GodownName, m_HeadID, m_HeadName, m_RelationID, m_TransID, 0, m_ToQuantity, StockBalance, "TRANSFER" & m_RelationID)
            Case StockOut
                
                m_FromQuantity = FormatField(rstTransaction.Fields("Quantity"))
                If .Rows <= .Row + 1 Then .Rows = .Rows + 1
                .Row = .Row + 1
                .Col = 0: .Text = "From " & GodownName
                .Col = 1: .Text = ProductName: .RowData(.Row) = RelationID
                .Col = 2: .Text = m_FromQuantity
                .Col = 3: .Text = StockBalance
                Call m_colTransfers.Add(m_GodownID, m_GodownName, m_HeadID, m_HeadName, m_RelationID, m_TransID, m_FromQuantity, 0, StockBalance, "TRANSFER" & m_RelationID)
        End Select
        
        rstTransaction.MoveNext
        
    Loop

m_frmTransfer.txtTransferDate.Text = m_TransDate

End With



Set MaterialClass = Nothing
Set rstTransaction = Nothing
Set rstHeadName = Nothing

Exit Sub

ErrLine:
    MsgBox "LoadTransactionsToGrid: " & Err.Description, vbCritical
    

End Sub

Private Sub CreateTransactionView(ByVal ViewName As String, ByVal TransID As Long)

gDbTrans.SqlStmt = " SELECT RelationID,GodownID,Quantity,VoucherType,TransDate" & _
                   " FROM Stock" & _
                   " WHERE TransID=" & TransID & _
                   " ORDER BY VoucherType DESC"

Call gDbTrans.CreateView(ViewName)

End Sub






Private Sub LoadGridSelectedDataToCombo()
'Declare the variables
Dim RelationID As Long
Dim GodownID As Byte
Dim headID As Long
Dim GroupID As Long
Dim UnitID As Long
Dim VoucherType As Wis_VoucherTypes
Dim MaxCount As Integer
Dim lpCount As Integer
Dim Quantity As Double
Dim SelectedRow As Integer

Dim rstRelation As ADODB.Recordset
Dim rstVoucher As ADODB.Recordset

'Setup an error handler...
On Error GoTo ErrLine

With m_frmTransfer.grd(TRANSFER_TRANSACTION)
    SelectedRow = .Row
    RelationID = .RowData(SelectedRow)
    .Col = 1: If .Text = "" Then Exit Sub
    .Col = 2:   Quantity = Val(.Text)
End With



If RelationID = 0 Then Exit Sub
If m_TransID = 0 Then Exit Sub

gDbTrans.SqlStmt = " SELECT GodownID,GroupID,UnitID" & _
                   " FROM RelationMaster" & _
                   " WHERE RelationID=" & RelationID
                   
If gDbTrans.Fetch(rstRelation, adOpenForwardOnly) < 1 Then Exit Sub

GodownID = rstRelation.Fields("GodownID")
'HeadID = rstRelation.Fields("HeadID")
GroupID = rstRelation.Fields("GroupID")
UnitID = rstRelation.Fields("UnitID")

Set rstRelation = Nothing

gDbTrans.SqlStmt = " SELECT VoucherType" & _
                   " FROM Stock" & _
                   " WHERE RelationID=" & RelationID & _
                   " AND TransID=" & m_TransID

If gDbTrans.Fetch(rstVoucher, adOpenForwardOnly) < 1 Then Exit Sub

'Enable the command button
With m_frmTransfer
    .cmdAdd.Enabled = True
    .cmdAdd.Caption = "&Update"
    .DBOperation = Update
'    .optFrom.Enabled = True
'    .optTo.Enabled = True
End With

VoucherType = rstVoucher.Fields("VoucherType")

Set rstVoucher = Nothing

If VoucherType = StockIn Then
    m_frmTransfer.optTo.Value = True
    m_ToQtyBeforeUpdation = Quantity
End If
If VoucherType = StockOut Then m_frmTransfer.optFrom.Value = True


MaxCount = m_frmTransfer.cmbBranch.ListCount - 1

With m_frmTransfer.cmbBranch
    For lpCount = 0 To MaxCount
        If GodownID = .ItemData(lpCount) Then
            .ListIndex = lpCount
            Exit For
        End If
    Next lpCount
End With
lpCount = 0: MaxCount = 0

MaxCount = m_frmTransfer.cmbGroup.ListCount - 1
With m_frmTransfer.cmbGroup
    For lpCount = 0 To MaxCount
        If GroupID = .ItemData(lpCount) Then
            .ListIndex = lpCount
            Exit For
        End If
    Next lpCount
End With

lpCount = 0: MaxCount = 0

MaxCount = m_frmTransfer.cmbUnit.ListCount - 1
With m_frmTransfer.cmbUnit
    For lpCount = 0 To MaxCount
        If UnitID = .ItemData(lpCount) Then
            .ListIndex = lpCount
            Exit For
        End If
    Next lpCount
End With

lpCount = 0: MaxCount = 0

MaxCount = m_frmTransfer.cmbProductName.ListCount - 1
With m_frmTransfer.cmbProductName
    For lpCount = 0 To MaxCount
        If RelationID = .ItemData(lpCount) Then
            .ListIndex = lpCount
            Exit For
        End If
    Next lpCount
End With

m_frmTransfer.txtQuantity.Text = Quantity

RemoveDataFromCollection (RelationID * GodownID)
m_frmTransfer.InitGridForTab0
AddToGrid
m_frmTransfer.grd(TRANSFER_TRANSACTION).Enabled = False
Exit Sub

ErrLine:
    MsgBox "LoadGridSelectedDataToCombo() " & Err.Description
    

End Sub


Private Function ShowTransferDetails(ByVal FromDateUS As String, ByVal ToDateUS As String) As Boolean
'Declare the variables
Dim rstTransfer As ADODB.Recordset
Dim rstHeadName As ADODB.Recordset
Dim MaterialClass As clsMaterial

Dim SlNo As Integer

Dim RelationID As Long
Dim TransID As Long

Dim TransDate As String
Dim ProductName As String
Dim UnitName As String

Dim FromQuantity As Double
Dim ToQuantity As Double
Dim StockBalance As Double

Dim GodownID As Byte
Dim VoucherType As Wis_VoucherTypes

'setup an error handler...
On Error GoTo ErrLine:

'Fetch the data
If Not PrepareRstForShowTransferDetails(rstTransfer, FromDateUS, ToDateUS) Then Exit Function

gDbTrans.SqlStmt = " SELECT DISTINCT A.GodownID,A.RelationID,GodownName,ProductName,UnitName" & _
                   " FROM RelationMaster A,QryTransferDetails B," & _
                        " GodownDet C,Products D,Units E" & _
                   " WHERE A.RelationID=B.RelationID" & _
                   " AND A.GodownID=B.GodownID" & _
                   " AND A.GodownID=C.GodownID" & _
                   " AND A.ProductID=D.ProductID" & _
                   " AND A.UnitID=E.UnitID" & _
                   " AND A.GroupID=D.GroupID"

If gDbTrans.Fetch(rstHeadName, adOpenForwardOnly) < 1 Then Exit Function

Set MaterialClass = New clsMaterial
Set m_GrdFunctions.fGrd = m_frmTransfer.grd(TRANSFER_SHOW)
SlNo = 0

With m_frmTransfer.grd(TRANSFER_SHOW)
    '.Visible = False
    .Row = .FixedRows - 1
    Do While Not rstTransfer.EOF
        RelationID = FormatField(rstTransfer.Fields("RelationID"))
        TransDate = FormatField(rstTransfer.Fields("TransDate"))
        GodownID = FormatField(rstTransfer.Fields("GodownID"))
        VoucherType = FormatField(rstTransfer.Fields("VoucherType"))
        TransID = FormatField(rstTransfer.Fields("TransID"))
        
        If gDbTrans.FindRecord(rstHeadName, "RelationID=" & RelationID) Then
            GodownName = FormatField(rstHeadName.Fields("GodownName"))
            ProductName = FormatField(rstHeadName.Fields("ProductName")) & " " & FormatField(rstHeadName.Fields("UnitName"))
            
            StockBalance = MaterialClass.GetItemOnDateClosingStock(RelationID, TransDate, GodownID)
        End If
        
        Select Case VoucherType
            Case StockIn
                FromQuantity = FormatField(rstTransfer.Fields("Quantity"))
                m_GrdFunctions.fMoreRows (1)
                .Row = .Row + 1
                .Col = 0: .Text = Format(SlNo, "00")
                .Col = 1: .Text = TransDate
                .Col = 2: .Text = "To " & GodownName: .RowData(.Row) = TransID
                .Col = 3: .Text = ProductName
                .Col = 4: .Text = FromQuantity
                .Col = 5: .Text = StockBalance
            
            Case StockOut
                SlNo = SlNo + 1
                ToQuantity = FormatField(rstTransfer.Fields("Quantity"))
                m_GrdFunctions.fMoreRows (1)
                .Row = .Row + 1
                .Col = 0: .Text = Format(SlNo, "00")
                .Col = 1: .Text = TransDate
                .Col = 2: .Text = "From " & GodownName: .RowData(.Row) = TransID
                .Col = 3: .Text = ProductName
                .Col = 4: .Text = ToQuantity
                .Col = 5: .Text = StockBalance
        End Select
        
        rstTransfer.MoveNext
    Loop
    
    '.Visible = True
End With

Set MaterialClass = Nothing
Set rstTransfer = Nothing
Set rstHeadName = Nothing

ShowTransferDetails = True

Exit Function
    
ErrLine:
    MsgBox "ShowTransferDetails: " & Err.Description, vbCritical
End Function

Public Property Let ToQuantity(ByVal vData As Double)
    On Error GoTo ToQuantityLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToQuantity = 5
    m_ToQuantity = vData
    Exit Property


ToQuantityLetErr:


    Call RaiseError(MyUnhandledError, "clsTransferNew:ToQuantity Property Let")
End Property


Public Property Get ToQuantity() As Double
    On Error GoTo ToQuantityGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToQuantity
    ToQuantity = m_ToQuantity
    Exit Property


ToQuantityGetErr:


    Call RaiseError(MyUnhandledError, "clsTransferNew:ToQuantity Property Get")
End Property



Public Property Let FromQuantity(ByVal vData As Double)
    On Error GoTo FromQuantityLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FromQuantity = 5
    m_FromQuantity = vData
    Exit Property


FromQuantityLetErr:


    Call RaiseError(MyUnhandledError, "clsTransferNew:FromQuantity Property Let")
End Property


Public Property Get FromQuantity() As Double
    On Error GoTo FromQuantityGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FromQuantity
    FromQuantity = m_FromQuantity
    Exit Property


FromQuantityGetErr:


    Call RaiseError(MyUnhandledError, "clsTransferNew:FromQuantity Property Get")
End Property



Public Property Let RelationID(ByVal vData As Long)
    On Error GoTo RelationIDLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RelationID = 5
    m_RelationID = vData
    Exit Property


RelationIDLetErr:


    Call RaiseError(MyUnhandledError, "clsTransferNew:RelationID Property Let")
End Property


Public Property Get RelationID() As Long
    On Error GoTo RelationIDGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RelationID
    RelationID = m_RelationID
    Exit Property


RelationIDGetErr:


    Call RaiseError(MyUnhandledError, "clsTransferNew:RelationID Property Get")
End Property



Public Property Let headName(ByVal vData As String)
    On Error GoTo HeadNameLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HeadName = 5
    m_HeadName = vData
    Exit Property


HeadNameLetErr:


    Call RaiseError(MyUnhandledError, "clsTransferNew:HeadName Property Let")
End Property


Public Property Get headName() As String
    On Error GoTo HeadNameGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HeadName
    headName = m_HeadName
    Exit Property


HeadNameGetErr:


    Call RaiseError(MyUnhandledError, "clsTransferNew:HeadName Property Get")
End Property



Public Property Let headID(ByVal vData As Long)
    On Error GoTo HeadIDLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HeadID = 5
    m_HeadID = vData
    Exit Property


HeadIDLetErr:


    Call RaiseError(MyUnhandledError, "clsTransferNew:HeadID Property Let")
End Property


Public Property Get headID() As Long
    
    On Error GoTo HeadIDGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HeadID
    headID = m_HeadID
    Exit Property


HeadIDGetErr:
    Call RaiseError(MyUnhandledError, "clsTransferNew:HeadID Property Get")
    
End Property





Public Property Let GodownID(ByVal vData As Byte)
    On Error GoTo GodownIDLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GodownID = 5
    m_GodownID = vData
    Exit Property


GodownIDLetErr:


    Call RaiseError(MyUnhandledError, "clsTransferNew:GodownID Property Let")
End Property


Public Property Get GodownID() As Byte
    On Error GoTo GodownIDGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GodownID
    GodownID = m_GodownID
    Exit Property


GodownIDGetErr:


    Call RaiseError(MyUnhandledError, "clsTransferNew:GodownID Property Get")
End Property

Private Sub Class_Initialize()
   
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & mlClassDebugID & " created"
    #End If
End Sub


Private Sub Class_Terminate()
    'the class is being destroyed
    #If DebugMode Then
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub


#If DebugMode Then
    Public Property Get ClassDebugID()
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

Private Sub m_frmTransfer_AddClicked()
If m_frmTransfer Is Nothing Then Exit Sub

With m_frmTransfer
    .cmbGroup.Locked = False
    If .optFrom.Value Then .cmbGroup.Locked = True
End With
'Check for validation
If Validated <> Success Then Exit Sub

'Add To the collection
If AddToCollection() <> Success Then Exit Sub

'add to grid
If AddToGrid <> Success Then Exit Sub
Dim I As Integer

With m_frmTransfer
    .cmdSave.Enabled = True
    I = .cmbGroup.ListIndex
    With .optTo
        If Not .Enabled Then
            .Enabled = True
            .Value = True
            On Error Resume Next
            .SetFocus
        End If
    End With
    If .optTo And .cmbBranch.ListIndex = 1 Then .cmbGroup.ListIndex = I
    .cmdUndo.Enabled = True
    .lblBalance.Caption = ""
    .lblPerUnit.Caption = ""
    .txtQuantity = ""
    .cmbProductName.ListIndex = -1
    .txtTransferDate.Locked = True
    .grd(TRANSFER_TRANSACTION).Enabled = True
End With



End Sub


Private Function AddToCollection() As wis_FunctionReturned
'Declare the variables
Dim MaterialClass As clsMaterial
Dim ClStock As Double
Dim OpStock As Double


'Setup an error handler...
On Error GoTo ErrLine

'Initialse the function
AddToCollection = Failure

With m_frmTransfer
    If .cmbBranch.ListIndex = -1 Then Exit Function
    If .cmbGroup.ListIndex = -1 Then Exit Function
    If .cmbUnit.ListIndex = -1 Then Exit Function
    
    m_GodownID = .cmbBranch.ItemData(.cmbBranch.ListIndex)
    m_GodownName = .cmbBranch.Text
'    m_HeadID = .cmbManufacturer.ItemData(.cmbManufacturer.ListIndex)
'    m_HeadName = Trim$(.cmbManufacturer.Text)
    m_RelationID = .cmbProductName.ItemData(.cmbProductName.ListIndex)
    m_Balance = 0
    m_Balance = CCur(.lblBalance.Caption)
    OpStock = m_Balance
    
    Set MaterialClass = New clsMaterial
    ClStock = MaterialClass.GetItemClStockBetweenDates(m_RelationID, GetSysFormatDate(.txtTransferDate), GetSysFormatDate(FinIndianEndDate), OpStock)
    
    If .optFrom.Value Then
        If .DBOperation = Update Then
            ClStock = ClStock + m_FromQuantity
            m_FromQuantity = 0
        Else
            m_FromQuantity = 0
            m_TransID = MaterialClass.GetMaxTransID() + 1
        End If
        
        m_FromQuantity = Val(.txtQuantity)
        m_Balance = ClStock - m_FromQuantity
        
        If m_Balance < 0 Then
            MsgBox "Insufficient Stock", vbInformation
            On Error Resume Next
            .cmbProductName.SetFocus
            Exit Function
        End If
        m_ToQuantity = 0
    
    Else
        
        m_FromQuantity = 0
        ClStock = ClStock - m_ToQtyBeforeUpdation
        m_ToQuantity = 0
        m_ToQuantity = Val(.txtQuantity)
        m_Balance = ClStock + m_ToQuantity
    
    End If
    
    If .cmbProductName.ListIndex = -1 Then Exit Function

End With

If m_colTransfers Is Nothing Then Set m_colTransfers = New colTransfers

Dim sKey As String
sKey = "TRANSFER" & m_RelationID
Call m_colTransfers.Add(m_GodownID, m_GodownName, m_HeadID, m_HeadName, m_RelationID, m_TransID, m_FromQuantity, m_ToQuantity, m_Balance, sKey)


AddToCollection = Success

Set MaterialClass = Nothing

Exit Function

ErrLine:
    MsgBox "AddToCollection(): " & vbCrLf & Err.Description
    AddToCollection = FatalError
    
End Function
Private Function AddToGrid() As wis_FunctionReturned
'Declare the variables
Dim MaxCount As Integer
Dim ItemCount As Integer
Dim lKeys
Dim TransferClass As clsTransferNew
Dim Qty As Double
Dim QtyBalance As Double
Dim IsOptFrom As Boolean
Dim FromPrinted As Boolean

'Initialse the function
AddToGrid = Failure


'.Col = 1: .Text = "Manufacturer": .CellFontBold = True
'.Col = 2: .Text = "Description": .CellFontBold = True
'.Col = 3: .Text = "Quantity": .CellFontBold = True
'.Col = 4: .Text = "TrQuantity": .CellFontBold = True

If m_GrdFunctions.fGrd Is Nothing Then Set m_GrdFunctions.fGrd = m_frmTransfer.grd(TRANSFER_TRANSACTION)

MaxCount = m_colTransfers.count

lKeys = m_colTransfers.GetKeys

'Add data to the grid
ItemCount = 0
IsOptFrom = False

For ItemCount = 0 To MaxCount - 1
    Set TransferClass = m_colTransfers.Item(lKeys(ItemCount))
    With m_frmTransfer.grd(TRANSFER_TRANSACTION)
        m_GrdFunctions.fMoreRows (2)
        .Row = ItemCount + 1
        
        .Col = 1: .Text = Trim$(m_frmTransfer.cmbProductName.Text): .RowData(.Row) = TransferClass.RelationID
        
        Qty = 0
        Qty = TransferClass.FromQuantity
        
        If MaxCount > 0 Then
            If Qty > 0 Then
                .Col = 0: .Text = "From " & TransferClass.GodownName
                IsOptFrom = True
            Else
                .Col = 0: .Text = "To " & TransferClass.GodownName
            End If
        End If
        
        If Qty > 0 Then .Col = 2: .Text = Qty
        
        Qty = 0
        Qty = TransferClass.ToQuantity
        
        If Qty > 0 Then .Col = 2: .Text = Qty
        
        .Col = 3: .Text = TransferClass.Balance
    End With
Next ItemCount

If m_frmTransfer.DBOperation = Insert Then If IsOptFrom Then m_frmTransfer.optFrom.Value = True
If MaxCount = 0 Then m_frmTransfer.optFrom.Value = True

With m_frmTransfer
    .cmdAdd.Enabled = True
    '.fra(TRANSFER_TRANSACTION).Enabled = True
    If MaxCount = 2 Then .cmdAdd.Enabled = False
    '   .fra(TRANSFER_TRANSACTION).Enabled = False
End With
'Assign Return to the function
AddToGrid = Success

Exit Function

ErrLine:
    
    MsgBox "AddToGrid: " & vbCrLf & Err.Description
    AddToGrid = FatalError
    
End Function


Private Function Validated() As wis_FunctionReturned
'Validate the controls
Dim StockBalance As Double
Dim Qty As Double
Validated = Failure

With m_frmTransfer
    If Not TextBoxDateValidate(.txtTransferDate, "/", True, True, True) Then Exit Function
    
    If .cmbBranch.ListIndex = -1 Then Exit Function
'    If .cmbManufacturer.ListIndex = -1 Then Exit Function
    If .cmbGroup.ListIndex = -1 Then Exit Function
    If .cmbUnit.ListIndex = -1 Then Exit Function
    If .cmbProductName.ListIndex = -1 Then Exit Function
    
 
    'Check Stock balance available
    StockBalance = Val(.lblBalance.Caption)
    Qty = Val(.txtQuantity)
        
    If Qty <= 0 Then
        MsgBox "Invalid Quantity Specified", vbInformation
        Exit Function
    End If
    
    Validated = Success
    If .DBOperation = Update Then Exit Function
    
    Validated = Failure
    
    If .optFrom Then
        If Qty > StockBalance Then
            MsgBox "Insuffient Stock ", vbInformation
            On Error Resume Next
            ActivateTextBox m_frmTransfer.txtQuantity
            Exit Function
        End If
    End If
    
    
End With

Validated = Success

End Function


Private Sub ClearCollections()

On Error Resume Next

Set m_colTransfers = Nothing

End Sub



Private Function LoadBranchRelationID(ByVal TheCombo As ComboBox, ByVal ToGodownID As Byte) As Long
'Declare the variables
Dim rst As ADODB.Recordset
Dim ProductID As Long
Dim headID As Long
Dim GroupID As Integer
Dim UnitID As Long

Dim fldProductName As ADODB.Field
Dim fldUnitName As ADODB.Field
Dim fldRelationID As ADODB.Field
   

gDbTrans.SqlStmt = "SELECT GroupID,ProductID FROM RelationMaster " & _
                " WHERE RelationID= " & m_RelationID

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 0 Then Exit Function

ProductID = FormatField(rst.Fields("ProductID"))
'HeadID = FormatField(Rst.Fields("HeadID"))
GroupID = FormatField(rst.Fields("GroupID"))

gDbTrans.SqlStmt = " SELECT ProductName,UnitName,RelationID FROM RelationMaster A," & _
                " Products B,ProductGroup C ,Units D " & _
                " WHERE GodownID = " & ToGodownID & _
                " " & _
                " AND A.GroupID = " & GroupID & _
                " AND A.ProductID = " & ProductID & _
                " AND A.ProductID = b.ProductID  " & _
                " AND A.UnitID = D.UnitID " & _
                " AND B.GroupID= C.GroupID " & _
                " AND A.GroupID = b.GroupID  " ' &
'                " AND HeadID = " & HeadID


TheCombo.Clear

If gDbTrans.Fetch(rst, adOpenForwardOnly) < 0 Then Exit Function

Set fldProductName = rst.Fields("ProductName")
Set fldUnitName = rst.Fields("UnitName")
Set fldRelationID = rst.Fields("RelationID")
 
 
Do While Not rst.EOF
    TheCombo.AddItem fldProductName.Value & "  " & fldUnitName.Value
    TheCombo.ItemData(TheCombo.newIndex) = fldRelationID.Value
    
    'Move the record set
    rst.MoveNext
Loop


End Function


Private Sub ClearTransferControls()
Dim MaterialClass As clsMaterial
'Dim eManufacture As wis_CompanyType

Set MaterialClass = New clsMaterial

'eManufacture = Enum_Manufacturer

With m_frmTransfer
    
    'Load the branches
    Call MaterialClass.LoadAllBranches(.cmbBranch)
    
    ' if the company type is Stockist  the load manufacturers
    'Call MaterialClass.LoadCompaniesToCombo(eManufacture,  .cmbManufacturer)
    
    'Load Group Combo box
    Call MaterialClass.LoadProductGroups(.cmbGroup)
    
    'Load Unit combo box
    .cmbUnit.Clear
    Call MaterialClass.LoadUnits(.cmbUnit)

    'Initalise the grid.
    .InitGridForTab0
    
    .InitGridForTab1
    
    .cmdUndo.Enabled = False
    .cmdSave.Enabled = False
    .cmdAdd.Enabled = True
    .cmdAdd.Caption = GetResourceString(10) '"&Add"
    .cmbBranch.Locked = False
        
    .txtQuantity.Text = ""
    
    .DBOperation = Insert
    
    .optFrom.Value = True
    .cmbGroup.Locked = False
    .fra(TRANSFER_TRANSACTION).Enabled = True
End With

End Sub



Private Sub m_frmTransfer_ClearClicked()
'Clear the collection
ClearCollections

'Clear module level variables
m_GodownID = 0
m_HeadID = 0
m_HeadName = 0
m_RelationID = 0
m_FromQuantity = 0
m_ToQuantity = 0
m_GodownName = 0

'refresh the cntrols
ClearTransferControls

End Sub

Private Function RemoveDataFromCollection(ByVal lngRelationID As Long) As Boolean

'Trap an error
On Error GoTo ErrLine

'Declare the variables

'remove the item from the collection
Call m_colTransfers.Remove("TRANSFER" & lngRelationID)


RemoveDataFromCollection = True

Exit Function

ErrLine:
    
End Function


 
Private Sub m_frmTransfer_DeleteClicked()

DeleteTransaction
End Sub

Private Sub m_frmTransfer_GridShowTransferClicked()

LoadDetailsToTransactionGrid

End Sub

Private Sub m_frmTransfer_GridTransactionClicked()
LoadGridSelectedDataToCombo

End Sub


'This function will check the Debit side total and Credit side total
Private Function IsFromToQuantityExists() As Boolean

' Setup Error Handler...
On Error GoTo ErrLine:

' Declare Variables
Dim FromQty As Double
Dim ToQty As Double
Dim Item As Long
Dim MaxCount As Long

Dim ObjKeys As Variant

If m_colTransfers Is Nothing Then Exit Function


MaxCount = m_colTransfers.count - 1
ObjKeys = m_colTransfers.GetKeys
For Item = 0 To MaxCount
    FromQty = FromQty + m_colTransfers(ObjKeys(Item)).FromQuantity
    ToQty = ToQty + m_colTransfers(ObjKeys(Item)).ToQuantity
Next Item

If FromQty = 0 Then Exit Function
If ToQty = 0 Then Exit Function


IsFromToQuantityExists = True

Exit Function


ErrLine:

    MsgBox "IsFromToQuantityExists:" & vbCrLf & Err.Description
    
    
End Function
Private Sub m_frmTransfer_OptFromClicked()
'Declare the variables
Dim MaterialClass As clsMaterial

'Setup an error handler...
On Error GoTo ErrLine

Set MaterialClass = New clsMaterial

With m_frmTransfer
    'Load the branches
    Call MaterialClass.LoadBranchesExceptSpecified(.cmbBranch, 0)
        
    ' if the company type is Stockist  the load manufacturers
    'Call MaterialClass.LoadCompaniesToCombo(Enum_Manufacturer, .cmbManufacturer)
        
    'Load Group Combo box
    Call MaterialClass.LoadProductGroups(.cmbGroup)
        
    'Load Unit combo box
    Call MaterialClass.LoadUnits(.cmbUnit)

End With

Set MaterialClass = Nothing

Exit Sub

ErrLine:
    MsgBox "OptFromClicked: " & Err.Description, vbCritical
    


End Sub

Private Sub m_frmTransfer_OptToClicked()
'Declare the variables
Dim MaterialClass As clsMaterial
Dim GodownID As Byte

'Setup an error handler...
On Error GoTo ErrLine

Set MaterialClass = New clsMaterial
If m_frmTransfer Is Nothing Then Exit Sub

With m_frmTransfer

    If .DBOperation = Insert Then Call MaterialClass.LoadBranchesExceptSpecified(.cmbBranch, m_GodownID)
    If .DBOperation = Update Then Call MaterialClass.LoadBranchesExceptSpecified(.cmbBranch, 0)
    Set MaterialClass = Nothing
    
    .cmbProductName.Clear
    
    If .cmbBranch.ListIndex = -1 Then Exit Sub
    
    GodownID = .cmbBranch.ItemData(.cmbBranch.ListIndex)
    
    
    Call LoadBranchRelationID(.cmbProductName, GodownID)
    
End With

Exit Sub

ErrLine:
    MsgBox "OptToClicked: " & Err.Description, vbCritical
    
End Sub

Private Sub m_frmTransfer_SaveClicked()

If Not IsFromToQuantityExists Then
    MsgBox "Unable to Transfer", vbInformation
    Exit Sub
End If

If m_frmTransfer.DBOperation = Insert Then
    If Not SaveDetails() Then Exit Sub
Else
    If Not UpdateDetails() Then Exit Sub
End If

Set m_colTransfers = Nothing



End Sub


Private Function SaveDetails() As Boolean
'Trtap an error
On Error GoTo ErrLine

'Declare the variables
Dim GodownID As Byte
Dim RelationID As Long
Dim TransID As Long
Dim FromQuantity As Double
Dim ToQuantity As Double
Dim PurchasePrice As Currency
Dim TransDate As String
Dim TransDateUS As String

Dim eVoucherType As Wis_VoucherTypes

Dim lpCount As Integer
Dim MaxCount As Integer
Dim lKeys As Variant

Dim TransferClass As clsTransferNew
Dim MaterialClass As clsMaterial


With m_frmTransfer
    TransDate = .txtTransferDate
End With

MaxCount = m_colTransfers.count - 1
lKeys = m_colTransfers.GetKeys

Set TransferClass = New clsTransferNew
Set MaterialClass = New clsMaterial

TransDateUS = GetSysFormatDate(TransDate)

gDbTrans.BeginTrans

For lpCount = 0 To MaxCount
    Set TransferClass = m_colTransfers.Item(lKeys(lpCount))
    
    FromQuantity = TransferClass.FromQuantity
    ToQuantity = TransferClass.ToQuantity
    TransID = TransferClass.TransID
    If FromQuantity > 0 Then
        RelationID = TransferClass.RelationID
        GodownID = TransferClass.GodownID
        PurchasePrice = MaterialClass.GetPurchasePrice(RelationID, GodownID, FinIndianFromDate, TransDate)
        
        eVoucherType = StockOut
        
        gDbTrans.SqlStmt = " INSERT INTO Stock (RelationID,TransID,GodownID, " & _
                       " Quantity,UnitPrice,VoucherType,PurOrSaleID,TransDate )" & _
                       " VALUES ( " & _
                       RelationID & "," & _
                       TransID & "," & _
                       GodownID & "," & _
                       FromQuantity & "," & _
                       PurchasePrice & "," & _
                       eVoucherType & "," & _
                       0 & "," & _
                       "#" & TransDateUS & "# ) "
                       
        If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    End If
    
    If ToQuantity > 0 Then
        RelationID = TransferClass.RelationID
        GodownID = TransferClass.GodownID
        PurchasePrice = MaterialClass.GetPurchasePrice(RelationID, GodownID, FinIndianFromDate, TransDate)
        
        eVoucherType = StockIn
        
        gDbTrans.SqlStmt = " INSERT INTO Stock (RelationID,TransID,GodownID, " & _
                       " Quantity,UnitPrice,VoucherType,PurOrSaleID,TransDate )" & _
                       " VALUES ( " & _
                       RelationID & "," & _
                       TransID & "," & _
                       GodownID & "," & _
                       ToQuantity & "," & _
                       PurchasePrice & "," & _
                       eVoucherType & "," & _
                       0 & "," & _
                       "#" & TransDateUS & "# ) "
                       
        If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    End If
    
Next lpCount

gDbTrans.CommitTrans

SaveDetails = True

MsgBox "Stock Transfered ", vbInformation


Exit Function

ErrLine:
        MsgBox "SaveDetails(): " & vbCrLf & Err.Description, vbCritical

End Function
Private Function UpdateDetails() As Boolean
'Trtap an error
On Error GoTo ErrLine

'Declare the variables
Dim GodownID As Byte
Dim RelationID As Long
Dim TransID As Long
Dim FromQuantity As Double
Dim ToQuantity As Double
Dim PurchasePrice As Currency
Dim TransDate As String
Dim TransDateUS As String

Dim eVoucherType As Wis_VoucherTypes

Dim lpCount As Integer
Dim MaxCount As Integer
Dim lKeys As Variant

Dim TransferClass As clsTransferNew
Dim MaterialClass As clsMaterial


With m_frmTransfer
    TransDate = .txtTransferDate
End With

MaxCount = m_colTransfers.count - 1
lKeys = m_colTransfers.GetKeys

Set TransferClass = New clsTransferNew
Set MaterialClass = New clsMaterial

TransDateUS = GetSysFormatDate(TransDate)

gDbTrans.BeginTrans

For lpCount = 0 To MaxCount
    Set TransferClass = m_colTransfers.Item(lKeys(lpCount))
    
    FromQuantity = TransferClass.FromQuantity
    ToQuantity = TransferClass.ToQuantity
    
    If FromQuantity > 0 Then
        RelationID = TransferClass.RelationID
        GodownID = TransferClass.GodownID
        eVoucherType = StockOut
        
        gDbTrans.SqlStmt = " UPDATE Stock" & _
                           " SET Quantity=" & FromQuantity & "," & _
                           " TransDate=#" & TransDateUS & "#" & _
                           " WHERE TransID=" & m_TransID & _
                           " AND RelationID=" & RelationID & _
                           " AND GodownID=" & GodownID & _
                           " AND VoucherType=" & eVoucherType
                           
        
        If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    End If
    
    If ToQuantity > 0 Then
        RelationID = TransferClass.RelationID
        GodownID = TransferClass.GodownID
        PurchasePrice = MaterialClass.GetPurchasePrice(RelationID, GodownID, FinIndianFromDate, TransDate)
        
        eVoucherType = StockIn
        
        gDbTrans.SqlStmt = " UPDATE Stock" & _
                           " SET Quantity=" & ToQuantity & "," & _
                           " TransDate=#" & TransDateUS & "#" & _
                           " WHERE TransID=" & m_TransID & _
                           " AND RelationID=" & RelationID & _
                           " AND GodownID=" & GodownID & _
                           " AND VoucherType=" & eVoucherType
        
                       
        If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
    End If
    
Next lpCount

gDbTrans.CommitTrans

UpdateDetails = True

MsgBox "Stock Transfer Updated", vbInformation


Exit Function

ErrLine:
        MsgBox "UpdateDetails(): " & vbCrLf & Err.Description, vbCritical

End Function
Private Function DeleteTransaction() As Boolean
'Declare the variables

'Setup an error handler
On Error GoTo ErrLine


gDbTrans.SqlStmt = " DELETE *" & _
                  " FROM Stock" & _
                  " WHERE TransID=" & m_TransID
gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
gDbTrans.CommitTrans

DeleteTransaction = True

MsgBox "Data Deleted!", vbInformation


Exit Function

ErrLine:
    MsgBox "DeleteTransaction: " & Err.Description, vbCritical
    

End Function

Public Property Let GodownName(ByVal vData As String)
    On Error GoTo HeadNameLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HeadName = 5
    m_GodownName = vData
    Exit Property


HeadNameLetErr:


    Call RaiseError(MyUnhandledError, "clsTransferNew:HeadName Property Let")
End Property


Public Property Get GodownName() As String
    On Error GoTo HeadNameGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HeadName
    GodownName = m_GodownName
    Exit Property


HeadNameGetErr:


    Call RaiseError(MyUnhandledError, "clsTransferNew:HeadName Property Get")
End Property

Private Sub m_frmTransfer_ShowClicked(ByVal FromDateUS As String, ByVal ToDateUS As String)
Call ShowTransferDetails(FromDateUS, ToDateUS)

End Sub


Private Sub m_frmTransfer_UndoClicked()
'Declare the variables
Dim SelectedRow As Integer
Dim lngRelationID As Long



'Get the Selected row
With m_frmTransfer.grd(TRANSFER_TRANSACTION)
    SelectedRow = .Row
    .Col = 2
    'Get the RelationID from the selected row
    lngRelationID = .RowData(SelectedRow)
End With


If RemoveDataFromCollection(lngRelationID) Then
    m_frmTransfer.InitGridForTab0
    AddToGrid
End If


End Sub


Public Property Let Balance(ByVal vData As Double)

On Error GoTo FromQuantityLetErr

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FromQuantity = 5
m_Balance = vData

Exit Property

FromQuantityLetErr:
    Call RaiseError(MyUnhandledError, "clsTransferNew:FromQuantity Property Let")
End Property


Public Property Get Balance() As Double

On Error GoTo FromQuantityGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FromQuantity
Balance = m_Balance
Exit Property

FromQuantityGetErr:
    Call RaiseError(MyUnhandledError, "clsTransferNew:FromQuantity Property Get")
End Property
Public Property Let TransID(ByVal vData As Long)
    On Error GoTo HeadIDLetErr


'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HeadID = 5
    m_TransID = vData
    Exit Property


HeadIDLetErr:


    Call RaiseError(MyUnhandledError, "clsTransferNew:HeadID Property Let")
End Property


Public Property Get TransID() As Long
    
    On Error GoTo HeadIDGetErr


'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HeadID
    TransID = m_TransID
    Exit Property


HeadIDGetErr:
    Call RaiseError(MyUnhandledError, "clsTransferNew:HeadID Property Get")
    
End Property

Private Sub m_frmTransfer_WindowClosed()
gWindowHandle = 0
Set m_frmTransfer = Nothing
End Sub


