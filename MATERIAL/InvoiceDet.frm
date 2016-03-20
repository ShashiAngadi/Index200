VERSION 5.00
Begin VB.Form frmInvoiceDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Details"
   ClientHeight    =   5355
   ClientLeft      =   2640
   ClientTop       =   1875
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtInvoiceNo 
      Height          =   315
      Left            =   60
      TabIndex        =   15
      Top             =   4110
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5340
      TabIndex        =   21
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   400
      Left            =   4020
      TabIndex        =   20
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtInvoiceAmount 
      Height          =   315
      Left            =   4560
      TabIndex        =   19
      Top             =   4110
      Width           =   2085
   End
   Begin VB.TextBox txtInvoiceDate 
      Height          =   315
      Left            =   2430
      TabIndex        =   17
      Top             =   4110
      Width           =   1665
   End
   Begin VB.ComboBox cmbInvoiceType 
      Height          =   315
      Left            =   2430
      TabIndex        =   13
      Top             =   3120
      Width           =   4125
   End
   Begin VB.ComboBox cmbRONo 
      Height          =   315
      Left            =   60
      TabIndex        =   7
      Top             =   2160
      Width           =   2055
   End
   Begin VB.OptionButton optSTANo 
      Caption         =   "STA No"
      Height          =   345
      Left            =   3750
      TabIndex        =   1
      Top             =   210
      Width           =   2775
   End
   Begin VB.ComboBox cmbCompany 
      Height          =   315
      Left            =   2070
      TabIndex        =   5
      Top             =   1170
      Width           =   4605
   End
   Begin VB.OptionButton optROno 
      Caption         =   "Release Order No"
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   2985
   End
   Begin VB.ComboBox cmbOpMode 
      Height          =   315
      Left            =   2070
      TabIndex        =   3
      Top             =   690
      Width           =   4605
   End
   Begin VB.Label lblOpMode 
      Caption         =   "Select Operation"
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblInvoiceNo 
      Caption         =   "Invoice No"
      Height          =   315
      Left            =   60
      TabIndex        =   14
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblCompany 
      Caption         =   "Select Company"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblInvoiceDate 
      Caption         =   "Invoice Date"
      Height          =   315
      Left            =   2430
      TabIndex        =   16
      Top             =   3600
      Width           =   1425
   End
   Begin VB.Label lblInvoiceAmount 
      Caption         =   "Invoice Amount"
      Height          =   315
      Left            =   4590
      TabIndex        =   18
      Top             =   3600
      Width           =   2085
   End
   Begin VB.Label lblInvoiceType 
      Caption         =   "Select Invoice Type"
      Height          =   315
      Left            =   60
      TabIndex        =   12
      Top             =   3150
      Width           =   2295
   End
   Begin VB.Label lblRODateL 
      Caption         =   "R O Date"
      Height          =   315
      Left            =   2550
      TabIndex        =   8
      Top             =   1770
      Width           =   1905
   End
   Begin VB.Label lblROAmountL 
      Caption         =   "R O Amount"
      Height          =   315
      Left            =   4830
      TabIndex        =   10
      Top             =   1800
      Width           =   2085
   End
   Begin VB.Label lblSelRo 
      Caption         =   "Select R O No"
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   1770
      Width           =   2295
   End
   Begin VB.Label lblROAmount 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4590
      TabIndex        =   11
      Top             =   2160
      Width           =   2085
   End
   Begin VB.Label lblRODate 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   2430
      TabIndex        =   9
      Top             =   2160
      Width           =   1425
   End
End
Attribute VB_Name = "frmInvoiceDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OkClicked()
Public Event UpdateClicked()
Public Event CancelClicked()


Private m_dbOperation As wis_DBOperation


Private Sub LoadCompanies()
'Trap an error
On Error GoTo ErrLine
'Declare the variables
Dim rst As ADODB.Recordset
Dim headID As ADODB.Field
Dim CompanyName As ADODB.Field

'Dim eManufacturer As wis_CompanyType
Dim eStockist As wis_CompanyType

'eManufacturer = Enum_Manufacturer
eStockist = Enum_Stockist


'Fire the query
gDbTrans.SqlStmt = " SELECT HeadID,CompanyName FROM CompanyCreation " & _
                 " WHERE CompanyType = " & eStockist & " " & _
                 " ORDER BY HeadID"
                 
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub


Set headID = rst.Fields("HeadID")
Set CompanyName = rst.Fields("CompanyName")

'Clear the combobox
cmbCompany.Clear

'Load the  Comapnies to the combo box
Do While Not rst.EOF
    cmbCompany.AddItem CompanyName.Value
    cmbCompany.ItemData(cmbCompany.newIndex) = headID.Value
    
    rst.MoveNext
Loop

Exit Sub

ErrLine:
End Sub

'This function will load the RO No or STANO, or INvoice NO As Specified
'to the soecified combobox
'
Private Sub LoadInvoiceNumbersForInsert(ByVal TheCombo As ComboBox)
'Trap an error
On Error GoTo ErrLine

'declare the variables
Dim rst As ADODB.Recordset
Dim PurchaseId As ADODB.Field
Dim InvoiceNo As ADODB.Field
Dim headID As Long

If cmbCompany.ListIndex = -1 Then Exit Sub

headID = cmbCompany.ItemData(cmbCompany.ListIndex)

'InvoiceType=1 for invoices

gDbTrans.SqlStmt = " SELECT PurchaseID,InvoiceType" & _
                   " FROM InvoiceDetails " & _
                   " WHERE InvoiceType=" & 1 & _
                   " AND HeadID = " & headID
        
Call gDbTrans.CreateView("QryInvoice")

gDbTrans.SqlStmt = " SELECT TransID,InvoiceType,InvoiceNo " & _
                   " FROM Purchase  " & _
                   " WHERE headid= " & headID

Call gDbTrans.CreateView("QryInvoicePurchase")

                   
gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM QryInvoicePurchase AS A " & _
                   " LEFT JOIN QryInvoice AS B ON A.Transid=B.PurchaseID"


'If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Sub

rst.Filter = adFilterNone

rst.Filter = "PurchaseID = NULL"

Set PurchaseId = rst.Fields("TransID")
Set InvoiceNo = rst.Fields("InvoiceNO")

TheCombo.Clear
Do While Not rst.EOF
    TheCombo.AddItem InvoiceNo.Value
    TheCombo.ItemData(TheCombo.newIndex) = PurchaseId.Value
    
    rst.MoveNext
Loop


Exit Sub

ErrLine:
    If Err Then
        MsgBox "LoadInvoiceNo(): " & Err.Description, vbCritical
        Exit Sub
    End If
End Sub

'This function will load the RO No or STANO, or INvoice NO As Specified
'to the soecified combobox
'
Private Sub LoadInvoiceNumbersForUpdate(ByVal TheCombo As ComboBox)
'Trap an error
On Error GoTo ErrLine

'declare the variables
Dim rst As ADODB.Recordset
Dim PurchaseId As ADODB.Field
Dim InvoiceNo As ADODB.Field
Dim headID As Long

If cmbCompany.ListIndex = -1 Then Exit Sub

headID = cmbCompany.ItemData(cmbCompany.ListIndex)

'InvoiceType=1 for invoices

gDbTrans.SqlStmt = " SELECT PurchaseID,InvoiceType" & _
                   " FROM InvoiceDetails " & _
                   " WHERE InvoiceType=" & 1 & _
                   " AND HeadID = " & headID
        
Call gDbTrans.CreateView("QryInvoice")

gDbTrans.SqlStmt = " SELECT TransID,InvoiceType,InvoiceNo " & _
                   " FROM Purchase  " & _
                   " WHERE headid= " & headID

Call gDbTrans.CreateView("QryInvoicePurchase")

                   
gDbTrans.SqlStmt = " SELECT * " & _
                   " FROM QryInvoicePurchase AS A " & _
                   " LEFT JOIN QryInvoice AS B ON A.Transid=B.PurchaseID"


If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub


rst.Filter = "PurchaseID <> NULL"

Set PurchaseId = rst.Fields("TransID")
Set InvoiceNo = rst.Fields("InvoiceNO")

TheCombo.Clear
Do While Not rst.EOF
    TheCombo.AddItem InvoiceNo.Value
    TheCombo.ItemData(TheCombo.newIndex) = PurchaseId.Value
    
    rst.MoveNext
Loop


Exit Sub

ErrLine:
    If Err Then
        MsgBox "LoadInvoiceNo(): " & Err.Description, vbCritical
        Exit Sub
    End If
End Sub


Private Function Validated() As Boolean

Validated = False

If cmbOpMode.ListIndex = -1 Then
    MsgBox GetResourceString(681), vbInformation, wis_MESSAGE_TITLE '"Please Select Operation mode", vbInformation
    cmbOpMode.SetFocus
    Exit Function
End If
If cmbCompany.ListIndex = -1 Then
    'MsgBox "Please Select Company", vbInformation
    MsgBox GetResourceString(682), vbInformation, wis_MESSAGE_TITLE
    cmbCompany.SetFocus
    Exit Function
End If
If cmbRONo.ListIndex = -1 Then
   ' MsgBox "Please Select RO/STA Number", vbInformation
    MsgBox GetResourceString(683), vbInformation, wis_MESSAGE_TITLE
    cmbRONo.SetFocus
    Exit Function
End If

If cmbInvoiceType.ListIndex = -1 Then
    'MsgBox "Please Select Invoice type", vbInformation
    MsgBox GetResourceString(684), vbInformation, wis_MESSAGE_TITLE
    cmbInvoiceType.SetFocus
    Exit Function
End If

If Not DateValidate(txtInvoiceDate.Text, "/", True) Then
    'MsgBox "Invalid Invoice date Specifed!", vbInformation
    MsgBox GetResourceString(509), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtInvoiceDate
    Exit Function
End If
If Not CurrencyValidate(txtInvoiceAmount.Text, True) Then
    'MsgBox "Invalid Invoice amount Specifed!", vbInformation
    MsgBox GetResourceString(513), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtInvoiceAmount
    Exit Function
End If

Validated = True
End Function

Private Sub cmbCompany_Click()
'Declare the variables
Dim eInvoiceType As Wis_InvoiceType

If cmbCompany.ListIndex = -1 Then Exit Sub
If cmbOpMode.ListIndex = -1 Then Exit Sub

If optRONo.Value Then
    eInvoiceType = RONumber
    With cmbInvoiceType
        .Clear
        .AddItem GetResourceString(172, 60) '"Invoice No"
        .ItemData(.newIndex) = 1
        
        .AddItem GetResourceString(226, 60) '"STA No"
        .ItemData(.newIndex) = 3
    End With
ElseIf optSTANo.Value Then
    eInvoiceType = STANumber
    With cmbInvoiceType
        .Clear
        .AddItem GetResourceString(172, 60) '"Invoice No"
        .ItemData(.newIndex) = 1
        
        .AddItem GetResourceString(225, 60) '"RO No"
        .ItemData(.newIndex) = 2
    End With
End If

m_dbOperation = cmbOpMode.ItemData(cmbOpMode.ListIndex)

'Load invoice no to combo
If m_dbOperation = Insert Then
    Call LoadInvoiceNumbersForInsert(cmbRONo)
    cmdOk.Caption = GetResourceString(1)
    'fSet + 1) '"&Ok"
ElseIf m_dbOperation = Update Then
    Call LoadInvoiceNumbersForUpdate(cmbRONo)
    cmdOk.Caption = GetResourceString(171) '"&Update"
End If


End Sub


Private Sub cmbInvoiceType_Click()
Dim PurchaseId As Long
Dim headID As Long
Dim rst As ADODB.Recordset
Dim eInvoiceType As Wis_InvoiceType


If cmbCompany.ListIndex = -1 Then Exit Sub
If cmbInvoiceType.ListIndex = -1 Then Exit Sub
If cmbRONo.ListIndex = -1 Then Exit Sub

eInvoiceType = cmbInvoiceType.ItemData(cmbInvoiceType.ListIndex)
Select Case eInvoiceType
    Case InvoiceNumber
        lblInvoiceNo.Caption = "Invoice No"
        lblInvoiceDate.Caption = "Invoice Date"
        lblInvoiceAmount.Caption = "Invoice Amount"
        
    Case RONumber
        lblInvoiceNo.Caption = "R O No"
        lblInvoiceDate.Caption = "R O Date"
        lblInvoiceAmount.Caption = "R O Amount"
    Case STANumber
        lblInvoiceNo.Caption = "STA No"
        lblInvoiceDate.Caption = "STA Date"
        lblInvoiceAmount.Caption = "STA Amount"
End Select

If m_dbOperation = Insert Then Exit Sub



headID = cmbCompany.ItemData(cmbCompany.ListIndex)
PurchaseId = cmbRONo.ItemData(cmbRONo.ListIndex)

txtInvoiceDate.Text = ""
txtInvoiceAmount.Text = ""
txtInvoiceNo.Text = ""
gDbTrans.SqlStmt = " SELECT InvoiceNo,InvoiceDate,InvoiceAmount " & _
                   " FROM InvoiceDetails " & _
                   " WHERE HeadID = " & headID & _
                   " AND PurchaseID = " & PurchaseId & _
                   " AND InvoiceType = " & eInvoiceType
                   
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

txtInvoiceDate.Text = FormatField(rst.Fields("InvoiceDate"))
txtInvoiceAmount.Text = FormatField(rst.Fields("InvoiceAmount"))
txtInvoiceNo.Text = FormatField(rst.Fields("InvoiceNo"))

End Sub


Private Sub cmbOpMode_Click()
If cmbOpMode.ListIndex = -1 Then Exit Sub
m_dbOperation = cmbOpMode.ItemData(cmbOpMode.ListIndex)

If m_dbOperation = Insert Then
    cmdOk.Caption = GetResourceString(1) '"&Ok"
ElseIf m_dbOperation = Update Then
    cmdOk.Caption = GetResourceString(171) '"&Update"
End If

cmbRONo.Clear

End Sub


Private Sub cmbRONo_Click()
Dim PurchaseId As Long
Dim headID As Long
Dim rst As ADODB.Recordset
Dim eInvoiceType As Wis_InvoiceType

If cmbCompany.ListIndex = -1 Then Exit Sub
If cmbRONo.ListIndex = -1 Then Exit Sub

eInvoiceType = RONumber
If optSTANo.Value Then eInvoiceType = STANumber


headID = cmbCompany.ItemData(cmbCompany.ListIndex)
PurchaseId = cmbRONo.ItemData(cmbRONo.ListIndex)

gDbTrans.SqlStmt = " SELECT InvoiceDate,InvoiceAmount " & _
                   " FROM InvoiceDetails " & _
                   " WHERE HeadID = " & headID & _
                   " AND PurchaseID = " & PurchaseId & _
                   " AND InvoiceType = " & eInvoiceType
                   
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub

lblRODate.Caption = FormatField(rst.Fields("InvoiceDate"))
lblROAmount.Caption = FormatField(rst.Fields("InvoiceAmount"))

End Sub


Private Sub cmdCancel_Click()

RaiseEvent CancelClicked

End Sub


Private Sub cmdOk_Click()

If Not Validated Then Exit Sub

If m_dbOperation = Insert Then
    RaiseEvent OkClicked
ElseIf m_dbOperation = Update Then
    RaiseEvent UpdateClicked
End If

End Sub

Private Sub Form_Load()
'Declare the variables
 
'Center the form
CenterMe Me

'Set Kannada caption
Call SetKannadaCaption
Call optRONo_Click
'Load company to combo
LoadCompanies

m_dbOperation = Insert
cmdOk.Caption = GetResourceString(1) '"&Ok"

'Load operation mode combo box
With cmbOpMode
    .AddItem GetResourceString(185)  '"New Entry "
    .ItemData(.newIndex) = 1
    .AddItem GetResourceString(171) '"Update "
    .ItemData(.newIndex) = 2
    .ListIndex = 0
End With

End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

optRONo.Caption = GetResourceString(225, 60)
optSTANo.Caption = GetResourceString(226, 60)
lblOpMode.Caption = GetResourceString(189, 27)
lblCompany.Caption = GetResourceString(138, 27)
lblSelRo.Caption = GetResourceString(201)
lblRODate.Caption = GetResourceString(225, 37)
lblROAmount.Caption = GetResourceString(225, 40)
lblInvoiceType.Caption = GetResourceString(172, 98, 27)
lblInvoiceNo.Caption = GetResourceString(172, 60)
lblInvoiceDate.Caption = GetResourceString(172, 37)
lblInvoiceAmount.Caption = GetResourceString(172, 40)

cmdOk.Caption = GetResourceString(1)
cmdCancel.Caption = GetResourceString(2)
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmInvoiceDet = Nothing

End Sub

Private Sub optRONo_Click()
Call cmbCompany_Click
lblSelRo.Caption = GetResourceString(201) '"Select R O No"
lblRODateL.Caption = GetResourceString(225, 37) '"R O Date"
lblROAmountL.Caption = GetResourceString(225, 40) '"R O Amount"
End Sub


Private Sub optSTANo_Click()
Call cmbCompany_Click

lblSelRo.Caption = GetResourceString(226, 60) '"Select STA No"
lblRODateL.Caption = GetResourceString(226) & " " & _
                     GetResourceString(37) '"STA Date"
lblROAmountL.Caption = GetResourceString(226) & " " & _
                       GetResourceString(40) '"STA Amount"

End Sub


