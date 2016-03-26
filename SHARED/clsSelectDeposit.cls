VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSelectDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_frmSelectDeposit As frmSelectFD
Attribute m_frmSelectDeposit.VB_VarHelpID = -1
Private m_DepositType As Integer
Private m_MultiDeposit As Boolean
Private m_Cancel As Boolean

Public Function SelectDeposit(ByVal ModuleID As wisModules, ByVal GroupType As wis_GroupType, ByRef multiDeposit As Boolean, ByRef cancel As Boolean) As Integer
    Dim deposityType As Integer
    
    'Set Some Defualt value to DepositType
    m_DepositType = -1
    Set m_frmSelectDeposit = Nothing
    Set m_frmSelectDeposit = New frmSelectFD
    'Set the Properties to show the proper Deposit/MemberTYpe
    m_frmSelectDeposit.ModuleID = ModuleID
    m_frmSelectDeposit.GroupType = GroupType
    Load m_frmSelectDeposit
    ''IF user has only one deposit/member TYpe
    '' then need not to show the dailogue. that single deposit will be selected automatically
    ''If slected deposit type is no -1 then, show the dailogue
    If m_DepositType = -1 Then m_frmSelectDeposit.Show 1
    Unload m_frmSelectDeposit
    'm_frmFDAcc.DepositType = m_DepositType
    
    SelectDeposit = m_DepositType
    cancel = m_Cancel
    multiDeposit = m_MultiDeposit
    
End Function

Private Sub m_frmSelectDeposit_CancelClick()
    m_Cancel = True
End Sub

Private Sub m_frmSelectDeposit_DepositCount(ByVal NumOfDeposits As Integer)
    m_MultiDeposit = (NumOfDeposits > 1)
End Sub

Private Sub m_frmSelectDeposit_OKClick(ByVal DepositType As Integer)
    m_DepositType = DepositType Mod 100
End Sub
