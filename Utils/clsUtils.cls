VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_frmUtils As frmUtils
Attribute m_frmUtils.VB_VarHelpID = -1
Public Property Get ObjectName()
    ObjectName = "UtilClass"
End Property

Public Property Get ObjectHandle() As Long
    If Not m_frmUtils Is Nothing Then ObjectHandle = m_frmUtils.hwnd
End Property


Public Sub Show()
    If m_frmUtils Is Nothing Then Set m_frmUtils = New frmUtils
    m_frmUtils.Show
    gWindowHandle = m_frmUtils.hwnd

End Sub

Private Sub m_frmUtils_WindowClosed()
Set m_frmUtils = Nothing
End Sub


