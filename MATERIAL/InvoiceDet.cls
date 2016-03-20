VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInvoiceDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents m_frmInvoice As frmInvoiceDet
Attribute m_frmInvoice.VB_VarHelpID = -1


Public Property Get ObjectName()
    ObjectName = "InvoiceClass"
End Property


Private Function SaveDetails() As Boolean
'trap an error
On Error GoTo ErrLine
'Declare the variables
Dim HeadID As Long
Dim PurchaseId As Long
Dim eInvoiceType As Wis_InvoiceType
Dim USInvoiceDate As String
Dim InvoiceAmount As Currency
Dim InvoiceNo As String

SaveDetails = False

With m_frmInvoice
    If .cmbCompany.ListIndex = -1 Then Exit Function
    If .cmbRONo.ListIndex = -1 Then Exit Function
    If .cmbInvoiceType.ListIndex = -1 Then Exit Function
    
    HeadID = .cmbCompany.ItemData(.cmbCompany.ListIndex)
    PurchaseId = .cmbRONo.ItemData(.cmbRONo.ListIndex)
    eInvoiceType = .cmbInvoiceType.ItemData(.cmbInvoiceType.ListIndex)
    USInvoiceDate = GetSysFormatDate(.txtInvoiceDate.Text)
    InvoiceAmount = Val(.txtInvoiceAmount.Text)
    InvoiceNo = Trim$(.txtInvoiceNo.Text)
    
End With


'Now Insert the above data into the data base


gDbTrans.SQLStmt = " INSERT INTO InvoiceDetails " & _
                   " (HeadID,PurchaseID,InvoiceType,InvoiceNo,InvoiceDate,InvoiceAmount )" & _
                   " VALUES ( " & _
                   HeadID & "," & _
                   PurchaseId & "," & _
                   eInvoiceType & "," & _
                   "'" & InvoiceNo & "'," & _
                   "#" & USInvoiceDate & "#," & _
                   InvoiceAmount & " ) "
                   

gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
gDbTrans.CommitTrans

SaveDetails = True

MsgBox "Saved the Details ", vbInformation

Exit Function

ErrLine:
    MsgBox "SaveDetails()"
    

End Function


Private Function UpdateDetails() As Boolean
'trap an error
On Error GoTo ErrLine
'Declare the variables
Dim HeadID As Long
Dim PurchaseId As Long
Dim eInvoiceType As Wis_InvoiceType
Dim USInvoiceDate As String
Dim InvoiceAmount As Currency
Dim InvoiceNo As String

UpdateDetails = False

With m_frmInvoice
    If .cmbCompany.ListIndex = -1 Then Exit Function
    If .cmbRONo.ListIndex = -1 Then Exit Function
    If .cmbInvoiceType.ListIndex = -1 Then Exit Function
    
    HeadID = .cmbCompany.ItemData(.cmbCompany.ListIndex)
    PurchaseId = .cmbRONo.ItemData(.cmbRONo.ListIndex)
    eInvoiceType = .cmbInvoiceType.ItemData(.cmbInvoiceType.ListIndex)
    USInvoiceDate = GetSysFormatDate(.txtInvoiceDate.Text)
    InvoiceAmount = Val(.txtInvoiceAmount.Text)
    InvoiceNo = Trim$(.txtInvoiceNo.Text)
    
End With


'Now Insert the above data into the data base


gDbTrans.SQLStmt = " UPDATE InvoiceDetails " & _
                   " SET InvoiceNo = " & AddQuotes(InvoiceNo) & "," & _
                   " InvoiceDate = #" & USInvoiceDate & "#," & _
                   " InvoiceAmount = " & InvoiceAmount & _
                   " WHERE HeadID = " & HeadID & _
                   " AND PurchaseID = " & PurchaseId & _
                   " AND InvoiceType = " & eInvoiceType

                   

gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
gDbTrans.CommitTrans

UpdateDetails = True

MsgBox "Details Updated", vbInformation

Exit Function

ErrLine:
    MsgBox "m_frmInvoice_OKClicked()"
    

End Function

Public Sub Show()
If m_frmInvoice Is Nothing Then Set m_frmInvoice = New frmInvoiceDet

m_frmInvoice.Show vbModal


End Sub


Private Sub m_frmInvoice_CancelClicked()
Unload m_frmInvoice

Set m_frmInvoice = Nothing

End Sub

Public Property Get ObjectHandle() As Long
    If Not m_frmInvoice Is Nothing Then _
                ObjectHandle = m_frmInvoice.hwnd
End Property

Private Sub m_frmInvoice_OKClicked()
If Not SaveDetails Then
    MsgBox "Unable to Save the data ", vbInformation
    Exit Sub
End If

Call ClearControls



End Sub


Private Sub m_frmInvoice_UpdateClicked()
If Not UpdateDetails Then
    MsgBox "Unable to update the data ", vbInformation
    Exit Sub
End If
Call ClearControls
End Sub


Private Sub ClearControls()
With m_frmInvoice
    .cmbOpMode.ListIndex = 0
    .cmbCompany.ListIndex = -1
    .cmbRONo.Clear
    .lblRODate.Caption = ""
    .lblROAmount.Caption = ""
    .cmbInvoiceType.Clear
    .txtInvoiceNo.Text = ""
    .txtInvoiceDate.Text = ""
    .txtInvoiceAmount.Text = ""
End With
End Sub


