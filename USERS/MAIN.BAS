Attribute VB_Name = "basMain"
Option Explicit

Public gAppPath As String
Public Const gAppName = "INDEX2000 - SB Acounts"
Public gDBTrans As clsTransact

Public gtranstype As wisTransactionTypes

Public Sub Initialize()

'Initialize the global variables
    gAppPath = App.Path
    
    If gDBTrans Is Nothing Then
        Set gDBTrans = New clsTransact
    End If

'Open the data base
    If Not gDBTrans.OpenDB(gAppPath & "\Users.MDB", "") Then
        If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
        End If
        If Not gDBTrans.CreateDB(gAppPath & "\Users.TAB", "") Then
            MsgBox "Unable to create new database !", vbCritical, gAppName & " - Error"
            On Error Resume Next
            Kill gAppPath & "\SBAcc.MDB"
            
        End If
    End If



End Sub

Public Sub Main()
Call Initialize
frmMain.Show
End Sub


