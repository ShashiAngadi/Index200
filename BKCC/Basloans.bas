Attribute VB_Name = "basLoans"
Option Explicit

Public Sub Initialize()

'Initialize the global variables
    gAppPath = App.Path
    
    If gDBTrans Is Nothing Then
        Set gDBTrans = New clsTransact
    End If

'Open the data base
    If Not gDBTrans.OpenDB(gAppPath & "\loandb.MDB", "") Then
        If MsgBox("Unable to open the database !" & vbCrLf & vbCrLf & " Creating New Database", vbQuestion + vbOKCancel, gAppName & " - Confirmation") = vbCancel Then
            End
        End If
        If Not gDBTrans.CreateDB(gAppPath & "\loanDB.INI", "") Then
            MsgBox "Unable to create new database !", vbCritical, gAppName & " - Error"
            On Error Resume Next
            Kill gAppPath & "\loanDB.MDB"
            End
        End If
    End If

End Sub

Public Sub main()
Call Initialize
frmMain.Show
End Sub

