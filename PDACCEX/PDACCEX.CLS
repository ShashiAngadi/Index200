VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPDAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_User As clsUsers
Attribute m_User.VB_VarHelpID = -1




Public Function ShowExp() As Boolean
frmMain.Show
'End
End Function


Private Sub Class_Initialize()
m_User.ShowLoginDialog
End Sub


