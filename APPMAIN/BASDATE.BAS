Attribute VB_Name = "basDate"
Option Explicit

Public Sub Main()
Dim FileNo As Integer
FileNo = FreeFile
Dim StrDate As String
StrDate = Format(Now, "mm/dd/yyyy")
Open App.Path & "\DateFile.Dat" For Output As #FileNo
Write #FileNo, StrDate
Close #FileNo
End Sub


