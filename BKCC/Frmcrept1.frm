VERSION 5.00
Begin VB.Form frmCRept 
   Caption         =   "INDEX2000  -  Report Wizard..."
   ClientHeight    =   3825
   ClientLeft      =   2085
   ClientTop       =   1530
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4230
   Begin VB.PictureBox CRViewer1 
      Height          =   3525
      Left            =   195
      ScaleHeight     =   3465
      ScaleWidth      =   3660
      TabIndex        =   0
      Top             =   105
      Width           =   3720
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


