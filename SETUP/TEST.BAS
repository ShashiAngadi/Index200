Attribute VB_Name = "basTest"
Option Explicit

Public gDBTrans As New clsTransact
Public gAppPath As String
Public Const gAppName = "Index 2000 - Setup Module"
Public Sub Initialize()
'Initialize the global variables
    gAppPath = App.Path
    
    If gDBTrans Is Nothing Then
        Set gDBTrans = New clsTransact
    End If

'Open the data base
    If Not gDBTrans.OpenDB(gAppPath & "\SETUPDB.MDB", "") Then
        If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
            
            End
        End If
        If Not gDBTrans.CreateDB(gAppPath & "\SETUPDB.INI", "") Then
            MsgBox "Unable to create new database !", vbCritical, gAppName & " - Error"
            On Error Resume Next
            Kill gAppPath & "\SBAcc.MDB"
            End
        End If
    End If

End Sub


Public Sub Main()
Call Initialize
frmTest.Show vbModal
End Sub


