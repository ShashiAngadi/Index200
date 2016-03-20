VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTimeIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m_StartIT As Double
Private m_EndIT As Double

Public Function RunResult() As String

RunResult = Format$(EndIT - StartIT, "00.00") & " Secs"

Debug.Print RunResult

StartIT = 0
EndIT = 0

End Function

Public Property Get StartIT() As Double

StartIT = m_StartIT

End Property

Public Property Let StartIT(ByVal NewValue As Double)

m_StartIT = NewValue

End Property

Public Property Get EndIT() As Double

EndIT = m_EndIT

End Property

Public Property Let EndIT(ByVal NewValue As Double)

m_EndIT = NewValue

End Property

Private Sub Class_Initialize()
StartIT = 0
EndIT = 0
End Sub


