VERSION 5.00
Begin VB.Form frmInvPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Printing Properties for Sales Invoice"
   ClientHeight    =   2865
   ClientLeft      =   975
   ClientTop       =   3135
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6825
   Begin VB.TextBox txtLRNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   395
      Left            =   4830
      TabIndex        =   15
      Top             =   1530
      Width           =   1785
   End
   Begin VB.CommandButton cmdTransPort 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   1050
      Width           =   345
   End
   Begin VB.CheckBox chkTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Do you want Company Title Printed?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3210
      TabIndex        =   2
      Top             =   120
      Width           =   3585
   End
   Begin VB.ComboBox cmbTransPortMode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4830
      TabIndex        =   10
      Top             =   1050
      Width           =   1605
   End
   Begin VB.ComboBox cmbPaymentTerm 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   13
      Top             =   1530
      Width           =   1605
   End
   Begin VB.TextBox txtInvoiceNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   395
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4860
      TabIndex        =   16
      Top             =   2295
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5820
      TabIndex        =   17
      Top             =   2295
      Width           =   885
   End
   Begin VB.TextBox txtQtnDate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   395
      Left            =   4830
      TabIndex        =   6
      Top             =   570
      Width           =   1785
   End
   Begin VB.TextBox txtDCNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   395
      Left            =   1560
      TabIndex        =   8
      Top             =   1050
      Width           =   1575
   End
   Begin VB.TextBox txtQtn 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   395
      Left            =   1560
      TabIndex        =   4
      Top             =   570
      Width           =   1575
   End
   Begin VB.Label lblLRNo 
      Caption         =   "L.R.No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3210
      TabIndex        =   14
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblInvoiceNo 
      Caption         =   "Invoice No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   150
      Width           =   1545
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6750
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblTerms 
      Caption         =   "Payment Terms"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   12
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label lblTransport 
      Caption         =   "Mode of Transport"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3210
      TabIndex        =   9
      Top             =   1080
      Width           =   1755
   End
   Begin VB.Label lblQtnDate 
      Caption         =   "Quotation Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3210
      TabIndex        =   5
      Top             =   570
      Width           =   1455
   End
   Begin VB.Label lblDCNo 
      Caption         =   "Deliver Channel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   7
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label lblQtn 
      Caption         =   "Quotation No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   3
      Top             =   600
      Width           =   1545
   End
End
Attribute VB_Name = "frmInvPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_SaleTransID As Long
Private m_IsCashSales As Boolean

Private m_DBOperation As wis_DBOperation

Private Function GetDBOperationStatus() As wis_DBOperation
'Declare the variables
Dim Rst As ADODB.Recordset


GetDBOperationStatus = Insert


gDbTrans.SqlStmt = " SELECT SaleTransID" & _
                 " FROM InvoicePrint" & _
                 " WHERE SaleTransID=" & m_SaleTransID
                 
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Function

Set Rst = Nothing

GetDBOperationStatus = Update

End Function

Private Sub LoadPaymentTerms()

With cmbPaymentTerm
    .AddItem "Cash"
    .ItemData(.NewIndex) = 1
    .AddItem "Credit"
    .ItemData(.NewIndex) = 2
    .AddItem "Cheque"
    .ItemData(.NewIndex) = 3
    .AddItem "Demand Draft"
    .ItemData(.NewIndex) = 4
    
    If m_IsCashSales Then
        .ListIndex = 0
    Else
        .ListIndex = 1
    End If

End With


End Sub

Private Sub LoadDetails()
Dim Rst As ADODB.Recordset
Dim myTitle  As Byte
Dim PaymentTerms As wis_PaymentTerm
Dim ItemCount As Integer
Dim MaxCount As Integer
Dim TransPortMode As Byte

gDbTrans.SqlStmt = " SELECT * FROM" & _
                  " InvoicePrint" & _
                  " WHERE SaleTransID=" & m_SaleTransID
                  
                  
If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub


txtQtn.Text = FormatField(Rst.Fields("QuotationNo"))
txtQtnDate.Text = FormatField(Rst.Fields("QuotationDate"))
txtDCNo.Text = FormatField(Rst.Fields("DCNo"))
TransPortMode = FormatField(Rst.Fields("TransportMode"))
txtLRNo.Text = FormatField(Rst.Fields("LRNo"))

myTitle = FormatField(Rst.Fields("TitlePrint"))
chkTitle.Value = vbUnchecked
If myTitle = 1 Then chkTitle.Value = vbChecked

PaymentTerms = FormatField(Rst.Fields("PaymentTerm"))

MaxCount = cmbPaymentTerm.ListCount - 1

For ItemCount = 0 To MaxCount
    If PaymentTerms = cmbPaymentTerm.ItemData(ItemCount) Then
        cmbPaymentTerm.ListIndex = ItemCount
        Exit For
    End If
Next ItemCount

If TransPortMode = 0 Then Exit Sub
ItemCount = 0
MaxCount = cmbTransPortMode.ListCount - 1
For ItemCount = 0 To MaxCount
    If TransPortMode = cmbTransPortMode.ItemData(ItemCount) Then
        cmbTransPortMode.ListIndex = ItemCount
        Exit For
    End If
Next ItemCount

End Sub

Private Sub LoadTransportModes()
'Declare the vaiables
Dim Rst As ADODB.Recordset
Dim fldTransModeId As ADODB.Field
Dim fldTransModeName As ADODB.Field


'Setup an error handler...
On Error GoTo ErrLine

gDbTrans.SqlStmt = " SELECT *" & _
                  " FROM TransPortMode" & _
                  " ORDER BY TransModeName"

If gDbTrans.Fetch(Rst, adOpenForwardOnly) < 1 Then Exit Sub

cmbTransPortMode.Clear

Set fldTransModeId = Rst.Fields("TransModeID")
Set fldTransModeName = Rst.Fields("TransModeName")

Do While Not Rst.EOF
    cmbTransPortMode.AddItem fldTransModeName.Value
    cmbTransPortMode.ItemData(cmbTransPortMode.NewIndex) = fldTransModeId.Value
    
    'Move the reocrdset
    Rst.MoveNext
    
Loop

Exit Sub

ErrLine:
    MsgBox "LoadTransModes: " & vbCrLf & Err.Description, vbCritical
    


End Sub

Private Function SaveDetails() As Boolean
'Declare the variables
Dim TitlePrint As Byte
Dim PaymentTerm As wis_PaymentTerm
Dim TransPortMode As Byte

Dim Rst As ADODB.Recordset

'Setup an error handler...
On Error GoTo ErrLine

If m_SaleTransID = 0 Then Exit Function
If cmbPaymentTerm.ListIndex = -1 Then Exit Function

TransPortMode = 0
If cmbTransPortMode.ListIndex <> -1 Then TransPortMode = cmbTransPortMode.ItemData(cmbTransPortMode.ListIndex)

PaymentTerm = cmbPaymentTerm.ItemData(cmbPaymentTerm.ListIndex)

TitlePrint = 2
If chkTitle.Value = vbChecked Then TitlePrint = 1

gDbTrans.SqlStmt = " INSERT INTO InvoicePrint ( SaleTransID,DCNo," & _
                  " PaymentTerm,QuotationNo," & _
                  " TitlePrint,TransPortMode,LRNo ) " & _
                  " VALUES ( " & _
                  m_SaleTransID & "," & _
                  AddQuotes(txtDCNo.Text) & "," & _
                  PaymentTerm & "," & _
                  AddQuotes(txtQtn.Text) & "," & _
                  TitlePrint & "," & _
                  TransPortMode & "," & _
                  AddQuotes(txtLRNo.Text) & " )"
                  
If txtQtnDate.Text <> "" Then
    gDbTrans.SqlStmt = " INSERT INTO InvoicePrint ( SaleTransID,DCNo," & _
                    " PaymentTerm,QuotationNo,QuotationDate," & _
                    " TitlePrint,TransPortMode,LRNo ) " & _
                    " VALUES ( " & _
                    m_SaleTransID & "," & _
                    AddQuotes(txtDCNo.Text) & "," & _
                    PaymentTerm & "," & _
                    AddQuotes(txtQtn.Text) & "," & _
                    "#" & FormatDate(txtQtnDate.Text) & "#," & _
                    TitlePrint & "," & _
                   TransPortMode & "," & _
                   AddQuotes(txtLRNo.Text) & " )"
                    
End If
                  
gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

gDbTrans.CommitTrans

SaveDetails = True

Exit Function

ErrLine:
    MsgBox "SaveDetails: " & Err.Description, vbCritical
    
End Function
Private Function UpdateDetails() As Boolean
'Declare the variables
Dim TitlePrint As Byte
Dim PaymentTerm As wis_PaymentTerm
Dim TransPortMode As Byte
           
Dim Rst As ADODB.Recordset

'Setup an error handler...
On Error GoTo ErrLine

If m_SaleTransID = 0 Then Exit Function
If cmbPaymentTerm.ListIndex = -1 Then Exit Function
TransPortMode = 0

If cmbTransPortMode.ListIndex <> -1 Then TransPortMode = cmbTransPortMode.ItemData(cmbTransPortMode.ListIndex)

PaymentTerm = cmbPaymentTerm.ItemData(cmbPaymentTerm.ListIndex)

TitlePrint = 2

If chkTitle.Value = vbChecked Then TitlePrint = 1

gDbTrans.SqlStmt = " UPDATE InvoicePrint" & _
                   " SET DCNo=" & AddQuotes(txtDCNo.Text) & "," & _
                   " PaymentTerm=" & PaymentTerm & "," & _
                   " QuotationNo=" & AddQuotes(txtQtn.Text) & "," & _
                   " QuotationDate=Null" & "," & _
                   " TitlePrint=" & TitlePrint & "," & _
                   " TransPortMode=" & TransPortMode & "," & _
                   " LRNo= " & AddQuotes(txtLRNo.Text) & _
                   " WHERE SaleTransID=" & m_SaleTransID

If txtQtnDate.Text <> "" Then
    gDbTrans.SqlStmt = " UPDATE InvoicePrint" & _
                       " SET DCNo=" & AddQuotes(txtDCNo.Text) & "," & _
                       " PaymentTerm=" & PaymentTerm & "," & _
                       " QuotationNo=" & AddQuotes(txtQtn.Text) & "," & _
                       " QuotationDate=#" & FormatDate(txtQtnDate.Text) & "#," & _
                       " TitlePrint=" & TitlePrint & "," & _
                       " TransPortMode=" & TransPortMode & "," & _
                       " LRNo=" & AddQuotes(txtLRNo.Text) & _
                       " WHERE SaleTransID=" & m_SaleTransID

End If
                  
gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

gDbTrans.CommitTrans

UpdateDetails = True

Exit Function

ErrLine:
    MsgBox "UpdateDetails: " & Err.Description, vbCritical

End Function
Private Function Validated() As Boolean
'Setup an error handler...
On Error GoTo ErrLine

If txtQtnDate.Text <> "" Then If Not TextBoxDateValidate(txtQtnDate, "/", True, True) Then Exit Function

'If cmbPaymentTerm.ListIndex = -1 Then Exit Function


Validated = True

Exit Function

ErrLine:
    MsgBox "Validated: " & Err.Description, vbCritical
    
End Function

Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdOk_Click()

If Not Validated Then Exit Sub

m_DBOperation = GetDBOperationStatus

If m_DBOperation = Insert Then Call SaveDetails
If m_DBOperation = Update Then Call UpdateDetails

Unload Me
End Sub

Private Sub cmdTransPort_Click()
frmTransPort.Show vbModal

'Load the transport modes from the database
LoadTransportModes

End Sub

Private Sub Form_Load()
'Center the form
CenterMe Me

'Set the Icon for the form
Me.Icon = LoadResPicture(147, vbResIcon)

chkTitle.Value = vbChecked
'Load Combo box
Call LoadPaymentTerms

'Load Transport Modes
LoadTransportModes

'Load Previous Details
Call LoadDetails


End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmInvPrint = Nothing

End Sub



Public Property Get SaleTransID() As Long
SaleTransID = m_SaleTransID

End Property

Public Property Let SaleTransID(ByVal vNewValue As Long)
m_SaleTransID = vNewValue
End Property

Public Property Get IsCashSales() As Boolean
IsCashSales = m_IsCashSales
End Property

Public Property Let IsCashSales(ByVal vNewValue As Boolean)
m_IsCashSales = vNewValue
End Property
