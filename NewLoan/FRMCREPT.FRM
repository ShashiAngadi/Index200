VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.1#0"; "CRVIEWER.DLL"
Begin VB.Form frmCRept 
   Caption         =   "INDEX2000  -  Report Wizard..."
   ClientHeight    =   3825
   ClientLeft      =   2085
   ClientTop       =   1530
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4230
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   3525
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   3720
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControl=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertControl=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
   End
End
Attribute VB_Name = "frmCRept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowLoanHolders(rs As Recordset, stDate As String, endDate As String)
Dim crRept As New crLoansIssued

' Set the properties...
With crRept
    .Database.SetDataSource rs
    .stDate = stDate
    .endDate = endDate
End With

With CRViewer1
    .ReportSource = crRept
    .ViewReport
End With
Me.Show vbModal
Set crRept = Nothing

End Sub
Private Sub Form_Resize()
With CRViewer1
    .Left = 0
    .Top = 0
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight
End With
End Sub


