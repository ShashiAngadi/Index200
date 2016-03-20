Attribute VB_Name = "basHelp"
Option Explicit

'this is keeps the track of the present active control
Public gPresentControl As String
''------------------------------------------------------------
''this sub displays the passed in message in the status
''bar on the bottom of the Main form
''------------------------------------------------------------
'Sub MsgBar(strMsg As String, isPause As Integer)
'
'Screen.MousePointer = vbDefault
'MDIMain.stsStatusBar.Panels(1).Text = "Ready "
'
'If Len(strMsg) > 0 Then MDIMain.stsStatusBar.Panels(1).Text = strMsg
'
'If isPause Then MDIMain.stsStatusBar.Panels(1).Text = strMsg & " Please wait..."
'
'MDIMain.stsStatusBar.Refresh
'
'End Sub
'
'
' this sub routine will fetch the tag of the active control
' of the active form
' takes no arguments
' returns the string which might have got from the tag
'
Public Function DisplayHelp() As String

On Error Resume Next

With Screen.ActiveControl
    
    ' place the activecontrol name in the variable
    
    gPresentControl = .Name
    
    'ControlTag = Val(Screen.ActiveControl.tag)
    DisplayHelp = "Help for this is Not Available"
    
    If .Tag <> "" Then DisplayHelp = Screen.ActiveControl.Tag
    
End With
End Function


