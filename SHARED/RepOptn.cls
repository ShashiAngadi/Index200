VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRepOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Form varaible
Private WithEvents m_frm As frmReportOption
Attribute m_frm.VB_VarHelpID = -1

'Varaibles tostrote the report options
Private m_Gender As wis_Gender
Private m_Caste As String
Private m_Place As String
Private m_FromAmount As Double
Private m_ToAmount As Double
Private m_AccGroupID As Integer
Private m_EnableAmount As Boolean
Private m_EnableControls As Boolean
Private Sub Class_Initialize()
    m_Gender = wisNoGender
    m_Place = ""
    m_Caste = ""
    m_FromAmount = 0
    m_ToAmount = 0
    m_AccGroupID = 1
    
End Sub



Public Property Let EnableCasteControls(NewValue As Boolean)
    m_EnableControls = NewValue
End Property

Public Property Let EnableAmountRange(NewValue As Boolean)
    m_EnableAmount = NewValue
End Property
 
Public Sub ShowDialog()
    If m_frm Is Nothing Then _
        Set m_frm = New frmReportOption
    'End If
    Load m_frm
    m_frm.EnableAmountRange = m_EnableAmount
    m_frm.EnableControls = m_EnableControls
    m_frm.Show vbModal
    
End Sub

Public Property Get AccountGroupID() As Integer
    AccountGroupID = m_AccGroupID
End Property

Public Property Get Caste() As String
    Caste = m_Caste
End Property

Public Property Get FromAmount() As Double
    FromAmount = m_FromAmount
End Property

Public Property Get ToAmount() As Double
    ToAmount = m_ToAmount
End Property

Public Property Get Gender() As wis_Gender
    Gender = m_Gender
End Property

Public Property Get Place() As String
    Place = m_Place
End Property

Private Sub Class_Terminate()
    If Not m_frm Is Nothing Then _
            Set m_frm = Nothing
        
End Sub

Private Sub m_frm_WindowClosed()

With m_frm
    m_AccGroupID = 1
    With .cmbAccGroup
        If .ListIndex > 0 Then m_AccGroupID = .ItemData(.ListIndex)
    End With

    'Gender
    m_Gender = wisNoGender
    With .cmbGender
        If .Enabled Then
            If .ListIndex >= 0 Then m_Gender = .ItemData(.ListIndex)
        End If
    End With

    'Place
    m_Place = ""
    If .cmbPlaces.Enabled Then m_Place = .cmbPlaces.Text
    
    'Caste
    m_Caste = ""
    If .cmbCastes.Enabled Then m_Caste = .cmbCastes.Text

    'From Amount
    m_FromAmount = 0
    If .txtStartAmt.Enabled Then m_FromAmount = .txtStartAmt.Value
    
    'To Amount
    m_ToAmount = 0
    If .txtEndAmt.Enabled Then m_ToAmount = .txtEndAmt.Value
    
End With

End Sub
