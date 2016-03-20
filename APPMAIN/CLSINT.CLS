VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInterest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_ModuleID As Long


Public Function InterestRate(SchemeId As Long, Optional IndianDate As String) As Single
'This Function Gives The Interest Rate And




If IsMissing(IndianDate) Then
    IndianDate = CStr(FormatDate(Date)) 'put todays date
End If
gDBTrans.SQLStmt = "Select InterestRate From InterestTab Where TransID = " & _
                  "(SELECT MAX(TransId) from InterestTab WHERE SchemeId =" & SchemeId & _
                   " AND Startdate = " & FormatDate(IndianDate) & ")"

If gDBTrans.SQLFetch < 1 Then GoTo ErrLine
InterestRate = CSng(FormatField(gDBTrans.Rst("InterestRate")))
Exit Function

ErrLine:
InterestRate = 0#

End Function


Public Property Get ModuleID() As Variant
ModuleID = m_ModuleID
End Property

Public Property Let ModuleID(ByVal ModuleID As Variant)
m_ModuleID = ModuleID
End Property
