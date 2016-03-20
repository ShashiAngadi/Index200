Attribute VB_Name = "basMain"
Option Explicit

Public Const gAppName = "NOTE Demo App"
Dim gAppPath As String
Public gDBTrans As clsTransact

Private Sub Initialize()
'Initialize the global variables
    gAppPath = App.Path
    
    If gDBTrans Is Nothing Then
        Set gDBTrans = New clsTransact
    End If

'Open the data base
    If Not gDBTrans.OpenDB(gAppPath & "\Notes.MDB", "") Then
        If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
            
            End
        End If
        If Not gDBTrans.CreateDB(gAppPath & "\NOTEs.INI", "") Then
            MsgBox "Unable to create new database !", vbCritical, gAppName & " - Error"
            On Error Resume Next
            Kill gAppPath & "\SBAcc.MDB"
            End
        End If
    End If



End Sub


Sub Main()
    Call Initialize
    frmMain.Show vbModal

End Sub


