VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClearing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Implement this module as an ActiveX DLL.
' Plugging this module into INDX2000 will be dammmmnnn easy.
Private WithEvents m_frmClearing  As frmClearing
Attribute m_frmClearing.VB_VarHelpID = -1
Private WithEvents m_frmClrReport  As frmCLReport
Attribute m_frmClrReport.VB_VarHelpID = -1


Public Property Get ObjectName()
    ObjectName = "ClearingClass"
End Property


Public Property Get ObjectHandle() As Long
    If Not m_frmClearing Is Nothing Then _
                ObjectHandle = m_frmClearing.hwnd
End Property

Public Sub Show()
    If m_frmClearing Is Nothing Then Set m_frmClearing = New frmClearing
    gWindowHandle = m_frmClearing.hwnd
    m_frmClearing.Show
End Sub

Public Function UndoTransaction(AccId As Long, TransID As Long)
Dim AccHeadID As Long

Dim lastTransID As Long

MsgBox "Incomplet code"
'Get the Last TransID
AccHeadID = GetIndexHeadID(GetResourceString(168) & " " & _
        GetResourceString(36))
gDbTrans.SqlStmt = "select Top 1 * from ClearingTrans WHere ChequeId = " & _
        "(Select "

End Function

Private Sub m_frmClearing_ShowReport(ReportType As wis_ClearingReports, fromDate As String, toDate As String)
    If Not m_frmClrReport Is Nothing Then Unload m_frmClrReport
    Set m_frmClrReport = New frmCLReport
    With m_frmClrReport
        .fromDate = fromDate
        .toDate = toDate
        .ReportType = ReportType
    End With
    Load m_frmClrReport
    m_frmClrReport.Show
    
End Sub

Private Sub m_frmClearing_WindowClosed()
    Set m_frmClearing = Nothing
End Sub


Private Sub m_frmClrReport_WindowClosed()
    Unload m_frmClrReport
End Sub
