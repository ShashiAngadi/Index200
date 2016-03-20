VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_frmShg As frmShg
Attribute m_frmShg.VB_VarHelpID = -1
Private WithEvents m_frmShgReport As frmShgReport
Attribute m_frmShgReport.VB_VarHelpID = -1
Private m_CustReg As clsCustReg
Attribute m_CustReg.VB_VarHelpID = -1
Private m_clsRepOption As clsRepOption

Private m_dbOperation As wis_DBOperation
Private m_LoanSchemeId As Long
Private m_TransID As Integer

'Private m_CustomerID As Long
'Private m_SbId As Integer
'Private m_LoanID As Integer

Private m_SBNum As String
Private m_LoanNum As String

Private m_AccID As Long

' Returns the text value from a control array
' bound the field "FieldName".
Private Function GetVal(FieldName As String) As String
Dim I As Integer
Dim strTxt As String
With m_frmShg
    For I = 0 To .txtData.count - 1
        strTxt = ExtractToken(.txtPrompt(I).Tag, "DataSource")
        If StrComp(strTxt, FieldName, vbTextCompare) = 0 Then
            GetVal = .txtData(I).Text
            Exit For
        End If
    Next
End With
End Function


Private Function AddTrainingRecord() As Boolean

If m_AccID = 0 Then
    MsgBox GetResourceString(662), vbInformation, wis_MESSAGE_TITLE
    Exit Function
End If

'First Validate the Records
With m_frmShg
    If Not DateValidate(.txtTrainFrom, "/", True) Then
        MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox .txtTrainFrom
        Exit Function
    End If
    Dim fromDate As Date
    Dim toDate As Date
    fromDate = GetSysFormatDate(.txtTrainFrom)
    toDate = fromDate
    If Len(Trim$(.txtTrainTo)) Then
        If Not DateValidate(.txtTrainTo, "/", True) Then
            MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox .txtTrainTo
            Exit Function
        End If
        toDate = GetSysFormatDate(.txtTrainTo)
        If DateDiff("d", fromDate, toDate) < 0 Then
            MsgBox GetResourceString(563), vbInformation, wis_MESSAGE_TITLE
            ActivateTextBox .txtTrainTo
            Exit Function
        End If
    End If
    Dim strTraining As String
    If Len(Trim$(.txtTraining)) = 0 Then
        MsgBox GetResourceString(621), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox .txtTrainTo
        Exit Function
    End If
    strTraining = Trim$(.txtTraining)
    
    Dim strPLace As String
    strPLace = Trim$(.txtTrainingPlace)
    Dim intMembers As Integer
    intMembers = Val(.txtTrainMembers)
    
End With

If m_TransID = 0 Then
    Dim Rst As Recordset
    gDbTrans.SqlStmt = "Select Max(TransId) From ShgTrans"
    If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then m_TransID = FormatField(Rst(0))
    m_TransID = m_TransID + 1
    gDbTrans.SqlStmt = "Insert Into ShgTrans" & _
                    " (TransID,AccID,FromDate,ToDate," & _
                    " TrainingDetail,Place,Members)" & _
                    " Values ( " & m_TransID & _
                    "," & m_AccID & ", #" & fromDate & "#, #" & toDate & "#" & _
                    "," & AddQuotes(strTraining) & "," & AddQuotes(strPLace) & _
                    "," & intMembers & ")"
                    
Else
    gDbTrans.SqlStmt = "Update ShgTrans " & _
                " Set FromDate = #" & fromDate & "#," & _
                " TODate = #" & toDate & "#," & _
                " TrainingDetail = " & AddQuotes(strTraining) & "," & _
                " Place = " & AddQuotes(strPLace) & "," & _
                " Members = " & intMembers & _
                " Where TransID = " & m_TransID
End If

gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Function
End If
gDbTrans.CommitTrans
AddTrainingRecord = True


End Function

Public Property Get ObjectName()
    ObjectName = "ShgClass"
End Property


Private Sub ClearControls()
If m_frmShg Is Nothing Then Exit Sub

With m_frmShg
    
    .cmdSave.Caption = GetResourceString(7)
    .cmdDelete.Enabled = False
    
    .txtData(GetIndex("AccID")) = ""
    .txtData(GetIndex("AccID")).Locked = False
    .txtData(GetIndex("AccName")) = ""
    .CustomerID = 0
    If Not m_CustReg Is Nothing Then m_CustReg.NewCustomer
    
    .txtData(GetIndex("CreateDate")) = ""
    .txtData(GetIndex("Contact")) = ""
    .cmb(Val(ExtractToken(.txtPrompt(GetIndex("MeetDay")).Tag, "TextIndex"))).ListIndex = -1
    .cmb(Val(ExtractToken(.txtPrompt(GetIndex("Gender")).Tag, "TextIndex"))).ListIndex = -1
    .cmb(Val(ExtractToken(.txtPrompt(GetIndex("Place")).Tag, "TextIndex"))).ListIndex = 0
    .cmb(Val(ExtractToken(.txtPrompt(GetIndex("Caste")).Tag, "TextIndex"))).ListIndex = 0
    
    .txtData(GetIndex("MeetPlace")) = ""
    .txtData(GetIndex("TotalMem")) = ""
    .txtData(GetIndex("FemaleMem")) = ""
    .txtData(GetIndex("FemaleScStMem")) = ""
    .txtData(GetIndex("LoanAccNum")) = ""
    .txtData(GetIndex("SBAccNum")) = ""
    .txtData(GetIndex("SCSTMem")) = ""
    .txtData(GetIndex("Remarks")) = ""
    
    .cmdAdd.Enabled = False
    .cmdDelete.Enabled = False
    
    m_dbOperation = Insert
    m_AccID = 0
    .SHGID = 0
    Call InitGrid
End With

End Sub


Private Sub ClearControls1()
#If junk Then
If m_frmShg Is Nothing Then Exit Sub

With m_frmShg
    .cmdCreate.Caption = GetResourceString(15)
    .cmdLoad.Caption = GetResourceString(3)
    
    .txtShgNum = ""
    .txtShgNum.Locked = False
    .txtCustNAme = ""
    m_CustReg.NewCustomer
    
    .txtDOB = ""
    .txtContactPerson = ""
    .cmbCaste.ListIndex = 0
    .cmbDay.ListIndex = -1
    .cmbGender.ListIndex = 0
    .cmbPlace.ListIndex = 0
    .txtMeetingPlace = ""
    .txtFemaleMem = ""
    .txtFemSc = ""
    .txtLoan = ""
    .txtMemTotal = ""
    .txtRemark = ""
    .txtSb = ""
    .txtScMem = ""
    
    .cmdAdd.Enabled = False
    .cmdDelete.Enabled = False
    
    m_dbOperation = Insert
    m_AccID = 0
    .SHGID = 0
    Call InitGrid
End With
#End If
End Sub

Private Sub InitGrid()

With m_frmShg
    .txtTrainFrom = ""
    .txtTraining = ""
    .txtTrainingPlace = ""
    .txtTrainMembers = ""
    .txtTrainTo = ""
    .cmdAdd.Caption = GetResourceString(10)
    m_TransID = 0
    .cmdTrainingDelete.Enabled = False
End With

With m_frmShg.grd
    .Clear
    .Rows = 1
    .Cols = 1
    .Cols = 6
    .Rows = 5
    .FixedCols = 1
    .FixedRows = 1
    .AllowUserResizing = flexResizeBoth
    .Row = 0
    .Col = 0: .ColWidth(0) = 400
    .Text = GetResourceString(33)
    .Col = 1: .ColWidth(1) = 1000
    .Text = GetResourceString(107)
    .Col = 2: .ColWidth(2) = 1000
    .Text = GetResourceString(108)
    .Col = 3: .ColWidth(3) = 2400
    .Text = GetResourceString(295)
    .Col = 4: .ColWidth(4) = 1200
    .Text = GetResourceString(94)
    .Col = 5: .ColWidth(5) = 600
    .Text = GetResourceString(112)
End With

End Sub

Private Sub LoadSHGDetail1()
#If junk Then
Dim Rst As Recordset
Dim AccNum As String
AccNum = Trim(m_frmShg.txtShgNum)
'Now Fetch the Shg detial
gDbTrans.SqlStmt = "Select * From ShgMaster " & _
        " Where AccNum = " & AddQuotes(AccNum)

If gDbTrans.Fetch(Rst, adOpenDynamic) < 1 Then Exit Sub

Dim I As Integer
Dim MaxI As Integer
Dim Str As String
Dim SbId As Long
Dim LoanID As Long
Dim CustomerID As Long

If m_AccID = FormatField(Rst("AccID")) Then Exit Sub

With m_frmShg
    .txtShgNum = FormatField(Rst("AccNum"))
    .txtShgNum.Locked = True
    
    m_AccID = FormatField(Rst("AccID"))
    .SHGID = m_AccID
    
    SbId = FormatField(Rst("SbAccid"))
    LoanID = FormatField(Rst("Loanid"))
    CustomerID = FormatField(Rst("CustomerID"))
    
    m_dbOperation = Update
    .cmdCreate.Caption = GetResourceString(171)
    .cmdLoad.Caption = GetResourceString(8)
    
    If m_CustReg Is Nothing Then Set m_CustReg = New clsCustReg
    Call m_CustReg.LoadCustomerInfo(FormatField(Rst("CustomerID")))
    .txtCustNAme = m_CustReg.FullName
    
    .txtContactPerson = FormatField(Rst("ContactPerson"))
    .txtMeetingPlace = FormatField(Rst("MeetingPlace"))
    .txtDOB = FormatField(Rst("CreateDate"))
    
    .txtFemaleMem = FormatField(Rst("FemaleMembers"))
    .txtMemTotal = FormatField(Rst("TotalMembers"))
    .txtScMem = FormatField(Rst("ScStMembers"))
    .txtFemSc = FormatField(Rst("FemaleScStMembers"))
    
    .cmbGender.ListIndex = FormatField(Rst("Gender"))
    .txtRemark = FormatField(Rst("Remarks"))
    
    'Caste
    Str = FormatField(Rst("Caste"))
    With .cmbCaste
        MaxI = .ListCount - 1
        For I = 0 To MaxI
            If .List(I) = Str Then _
                .ListIndex = I: Exit For
        Next
    End With
    'Place
    Str = FormatField(Rst("Place"))
    With .cmbPlace
        MaxI = .ListCount - 1
        For I = 0 To MaxI
            If .List(I) = Str Then _
                .ListIndex = I: Exit For
        Next
    End With
    'Meeting Day
    Str = FormatField(Rst("MeetingDay"))
    With .cmbDay
        MaxI = .ListCount - 1
        For I = 0 To MaxI
            If .List(I) = Str Then _
                    .ListIndex = I: Exit For
        Next
    End With
    
    gDbTrans.SqlStmt = ""
    
'get the Savings account details
    If SbId > 0 Then
        gDbTrans.SqlStmt = "select AccNum,AccID From SbMaster " & _
                            " Where AccId = " & SbId
    Else
        gDbTrans.SqlStmt = "select AccNum,AccID From SbMaster " & _
                            " Where CustomerID = " & CustomerID
    End If
    
    If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then
        .txtSb.Text = FormatField(Rst(0))
 '       SbId = FormatField(Rst(1))
        gDbTrans.SqlStmt = "Select Top 1 Balance From SbTrans " & _
                        " Where AccId = " & SbId & " Order By TransID Desc;"
        If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then _
                .txtSb.Tag = "Balance Rs." & FormatField(Rst(0)): .lblSb.Tag = .txtSb.Tag
    End If
    
    'get the Loan details
    If LoanID Then
        gDbTrans.SqlStmt = "select AccNum,LoanID From LoanMaster " & _
                " Where LoanId = " & LoanID
    ElseIf m_LoanSchemeId Then
        gDbTrans.SqlStmt = "Select AccNum,LoanID From LoanMaster " & _
                    " Where SchemeID = " & m_LoanSchemeId & _
                    " ANd CustomerId = " & CustomerID
    Else
        gDbTrans.SqlStmt = "Select AccNum,LoanID From LoanMaster " & _
                    " Where CustomerId = " & CustomerID
        
    End If
    If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then
        .txtLoan = FormatField(Rst(0))
'        LoanID = FormatField(Rst(1))
        gDbTrans.SqlStmt = "Select Top 1 Balance From LoanTrans " & _
                    " Where LoanId = " & LoanID & " Order By TransID Desc;"
        If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then _
                .txtLoan.Tag = "Balance Rs." & FormatField(Rst(0)): .lblLoan.Tag = .txtLoan.Tag
    End If
    
    
.cmdAdd.Enabled = True

End With
#End If
End Sub

Private Sub LoadSHGDetail()

 
Dim Rst As Recordset
Dim AccNum As String
Dim txtIndex As Byte
Dim cmbIndex As Byte

txtIndex = GetIndex("AccID")
AccNum = Trim(m_frmShg.txtData(txtIndex))
'Now Fetch the Shg detail
gDbTrans.SqlStmt = "Select * From ShgMaster " & _
        " Where AccNum = " & AddQuotes(AccNum)

If gDbTrans.Fetch(Rst, adOpenDynamic) < 1 Then Exit Sub

Dim I As Integer
Dim MaxI As Integer
Dim Str As String
Dim SbId As Long
Dim LoanID As Long
Dim CustomerID As Long

If m_AccID = FormatField(Rst("AccID")) Then Exit Sub

With m_frmShg

    txtIndex = GetIndex("AccID")
    .txtData(txtIndex) = FormatField(Rst("AccNum"))
    .txtData(txtIndex).Locked = True
    
    m_AccID = FormatField(Rst("AccID"))
    .SHGID = m_AccID
    
    SbId = FormatField(Rst("SbAccid"))
    LoanID = FormatField(Rst("Loanid"))
    CustomerID = FormatField(Rst("CustomerID"))
    .CustomerID = CustomerID
    
    m_dbOperation = Update
    txtIndex = GetIndex("CreateDate")
    .cmdSave.Caption = GetResourceString(171)
    .cmdDelete.Enabled = True
    '.cmd(0).Caption = GetResourceString(8)
    
    If m_CustReg Is Nothing Then Set m_CustReg = New clsCustReg
    Call m_CustReg.LoadCustomerInfo(FormatField(Rst("CustomerID")))
    txtIndex = GetIndex("AccName")
    .txtData(txtIndex) = m_CustReg.FullName
    
    .txtData(GetIndex("Contact")) = FormatField(Rst("ContactPerson"))
    .txtData(GetIndex("MeetPlace")) = FormatField(Rst("MeetingPlace"))
    .txtData(GetIndex("CreateDate")) = FormatField(Rst("CreateDate"))
    
    .txtData(GetIndex("FemaleMem")) = FormatField(Rst("FemaleMembers"))
    .txtData(GetIndex("TotalMem")) = FormatField(Rst("TotalMembers"))
    .txtData(GetIndex("SCSTMem")) = FormatField(Rst("ScStMembers"))
    .txtData(GetIndex("FemaleSCSTMem")) = FormatField(Rst("FemaleScStMembers"))
    
    txtIndex = (GetIndex("Gender"))
    cmbIndex = Val(ExtractToken(.txtPrompt(txtIndex).Tag, "TextIndex"))
    .cmb(cmbIndex).ListIndex = FormatField(Rst("Gender"))
    .txtData(txtIndex).Text = .cmb(cmbIndex).Text
    
    .txtData(GetIndex("Remarks")) = FormatField(Rst("Remarks"))
    
    'Caste
    Str = FormatField(Rst("Caste"))
    txtIndex = GetIndex("Caste")
    .txtData(txtIndex) = Str
    cmbIndex = ExtractToken(.txtPrompt(txtIndex).Tag, "TextIndex")
    With .cmb(cmbIndex)
        MaxI = .ListCount - 1
        For I = 0 To MaxI
            If .List(I) = Str Then _
                .ListIndex = I: Exit For
        Next
    End With
    'Place
    Str = FormatField(Rst("Place"))
    txtIndex = GetIndex("Place")
    .txtData(txtIndex) = Str
    cmbIndex = ExtractToken(.txtPrompt(txtIndex).Tag, "TextIndex")
    With .cmb(cmbIndex)
        MaxI = .ListCount - 1
        For I = 0 To MaxI
            If .List(I) = Str Then _
                .ListIndex = I: Exit For
        Next
    End With
    'Meeting Day
    Str = FormatField(Rst("MeetingDay"))
    txtIndex = GetIndex("MeetDay")
    .txtData(txtIndex) = Str
    cmbIndex = ExtractToken(.txtPrompt(txtIndex).Tag, "TextIndex")
    With .cmb(cmbIndex)
        MaxI = .ListCount - 1
        For I = 0 To MaxI
            If .List(I) = Str Then _
                    .ListIndex = I: Exit For
        Next
    End With
    
    gDbTrans.SqlStmt = ""
    
'get the Savings account details
    If SbId > 0 Then
        gDbTrans.SqlStmt = "Select AccNum,AccID From SbMaster " & _
                            " Where AccId = " & SbId
    Else
        gDbTrans.SqlStmt = "Select AccNum,AccID From SbMaster " & _
                            " Where CustomerID = " & CustomerID
    End If
    
    If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then
        txtIndex = GetIndex("SBAccNum")
        .txtData(txtIndex).Text = FormatField(Rst(0))
 '       SbId = FormatField(Rst(1))
        gDbTrans.SqlStmt = "Select Top 1 Balance From SbTrans " & _
                        " Where AccId = " & SbId & " Order By TransID Desc;"
        If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then _
                Call PutToken(.txtData(txtIndex).Tag, "Balance", FormatField(Rst(0)))
                '.txtData(txtIndex).Tag = "Balance Rs." & FormatField(Rst(0)): .lblSb.Tag = .txtSb.Tag
                
    End If
    
    'get the Loan details
    If LoanID Then
        gDbTrans.SqlStmt = "select AccNum,LoanID From LoanMaster " & _
                " Where LoanId = " & LoanID
    ElseIf m_LoanSchemeId Then
        gDbTrans.SqlStmt = "Select AccNum,LoanID From LoanMaster " & _
                    " Where SchemeID = " & m_LoanSchemeId & _
                    " And CustomerId = " & CustomerID
    Else
        gDbTrans.SqlStmt = "Select AccNum,LoanID From LoanMaster " & _
                    " Where CustomerId = " & CustomerID
        
    End If
    
    If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then
        txtIndex = GetIndex("LoanAccNum")
        .txtData(txtIndex) = FormatField(Rst(0))
'        LoanID = FormatField(Rst(1))
        gDbTrans.SqlStmt = "Select Top 1 Balance From LoanTrans " & _
                    " Where LoanId = " & LoanID & " Order By TransID Desc;"
        If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then _
                Call PutToken(.txtData(txtIndex).Tag, "Balance", FormatField(Rst(0)))
                '.txtLoan.Tag = "Balance Rs." & FormatField(Rst(0)): .lblLoan.Tag = .txtLoan.Tag
    End If
    
    
.cmdAdd.Enabled = True

End With

End Sub


Private Sub LoadTrainingDetail()

Call InitGrid

Dim Rst As Recordset

gDbTrans.SqlStmt = "Select * From shgTrans " & _
                " Where AccID = " & m_AccID & _
                " Order By FromDate"
If gDbTrans.Fetch(Rst, adOpenDynamic) < 1 Then Exit Sub

With m_frmShg.grd
    .Row = 0
    While Not Rst.EOF
        If .Rows = .Row + 1 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .RowData(.Row) = FormatField(Rst("TransID"))
        .Col = 0: .Text = .Row
        .Col = 1: .Text = FormatField(Rst("FromDate"))
        .Col = 2: .Text = FormatField(Rst("TODate"))
        .Col = 3: .Text = FormatField(Rst("TrainingDetail"))
        .Col = 4: .Text = FormatField(Rst("Members"))
        .Col = 5: .Text = FormatField(Rst("Place"))
        Rst.MoveNext
    Wend
End With

End Sub


Private Function SaveDetail1()
#If junk Then
Dim strCreateDate As String
Dim intGender As Byte
Dim Str As String
Dim custId As Long

Dim SbId As Integer
Dim LoanID As Integer

Dim SqlStr As String
Dim Rst As Recordset

With m_frmShg
    If DateValidate(.txtDOB, "/", True) Then _
                strCreateDate = " #" & GetSysFormatDate(.txtDOB) & "#,"
    intGender = .cmbGender.ListIndex
End With

custId = m_CustReg.CustomerID
    
'Get the Savings Account Number if any
Str = Trim(m_frmShg.txtSb)
If Len(Str) Then
    gDbTrans.SqlStmt = "Select AccID From SbMaster " & _
                " Where AccNum = " & AddQuotes(Str)
Else
    gDbTrans.SqlStmt = "Select AccID From SbMaster " & _
                " Where CustomerID = " & custId
End If
If gDbTrans.Fetch(Rst, adOpenDynamic) = 1 Then SbId = FormatField(Rst(0))

'Get the Loan account Number if any
Str = Trim(m_frmShg.txtLoan)
If Len(Str) Then
    If m_LoanSchemeId Then
        gDbTrans.SqlStmt = "Select LoanID From LoanMaster " & _
                    " Where SchemeId = " & m_LoanSchemeId & _
                    " And AccNum = " & AddQuotes(Str)
    Else
        gDbTrans.SqlStmt = "Select LoanID From LoanMaster " & _
                    " Where AccNum = " & AddQuotes(Str) & _
                    " And CustomerID = " & custId
    End If
Else
    gDbTrans.SqlStmt = "Select LoanID From LoanMaster " & _
                " Where CustomerID = " & custId
End If
If gDbTrans.Fetch(Rst, adOpenDynamic) = 1 Then LoanID = FormatField(Rst(0))

If m_dbOperation = Insert Then
    m_AccID = 0
    gDbTrans.SqlStmt = "Select Max(AccId) From ShgMaster;"
    If gDbTrans.Fetch(Rst, adOpenDynamic) Then m_AccID = FormatField(Rst(0)) + 1
    If m_AccID <= 0 Then m_AccID = 1
    
    With m_frmShg
        SqlStr = "Insert Into ShgMaster (AccNum,AccId," & _
            "CustomerID,ContactPerson," & _
            "MeetingDay,MeetingPlace," & _
            "Gender,Place,Caste," & _
            IIf(Len(strCreateDate), "CreateDate,", "") & " TotalMembers," & _
            "FemaleMembers,ScStMembers,FemaleScStMembers," & _
            "SbAccId,LoanId ) " & _
            " VALUES  (" & _
            AddQuotes(.txtShgNum) & "," & m_AccID & "," & _
            custId & "," & AddQuotes(.txtContactPerson) & "," & _
            AddQuotes(.cmbDay.Text) & "," & AddQuotes(.txtMeetingPlace) & "," & _
            intGender & "," & AddQuotes(.cmbPlace.Text) & "," & AddQuotes(.cmbCaste.Text) & "," & _
            strCreateDate & Val(.txtMemTotal) & "," & _
            Val(.txtFemaleMem) & "," & Val(.txtScMem) & "," & Val(.txtFemSc) & "," & _
            SbId & "," & LoanID & ")"
    End With

Else
    
    With m_frmShg
        SqlStr = "UPDate ShgMaster Set" & _
            " ContactPerson = " & AddQuotes(.txtContactPerson) & "," & _
            " MeetingDay = " & AddQuotes(.cmbDay.Text) & "," & _
            " MeetingPlace = " & AddQuotes(.txtMeetingPlace) & "," & _
            " Gender = " & intGender & "," & _
            " Place = " & AddQuotes(.cmbPlace.Text) & "," & _
            " Caste = " & AddQuotes(.cmbCaste.Text) & "," & _
            IIf(Len(strCreateDate), "CreateDate = " & strCreateDate, "") & _
            " TotalMembers = " & Val(.txtMemTotal) & "," & _
            " FemaleMembers = " & Val(.txtFemaleMem) & "," & _
            " ScStMembers = " & Val(.txtScMem) & "," & _
            " FemaleScStMembers = " & Val(.txtFemSc) & "," & _
            " SbAccId = " & SbId & ", LoanId = " & LoanID & _
            " WHERE AccId = " & m_AccID
            
    End With
End If

gDbTrans.BeginTrans
m_CustReg.ModuleID = wis_Users
If Not m_CustReg.SaveCustomer Then gDbTrans.RollBacknRaiseError
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

gDbTrans.CommitTrans

MsgBox GetResourceString(IIf(m_dbOperation = Insert, 528, 707)), vbInformation, wis_MESSAGE_TITLE

If m_dbOperation = Insert Then m_AccID = 0: Call LoadSHGDetail

Err_line:
#End If
End Function

Private Function SaveDetail()

Dim strCreateDate As String
Dim intGender As Byte
Dim Str As String
Dim custId As Long

Dim SbId As Integer
Dim LoanID As Integer

Dim SqlStr As String
Dim Rst As Recordset
Dim txtIndex As Byte

With m_frmShg
    txtIndex = GetIndex("CreateDate")
    If DateValidate(.txtData(txtIndex).Text, "/", True) Then _
                strCreateDate = " #" & GetSysFormatDate(.txtData(txtIndex).Text) & "#,"
    'Get the Combo ondex
    txtIndex = GetIndex("Gender")
    txtIndex = ExtractToken(.txtPrompt(txtIndex).Tag, "textindex")
    intGender = .cmb(txtIndex).ListIndex
End With

custId = m_CustReg.CustomerID
    
'Get the Savings Account Number if any
txtIndex = GetIndex("SbAccNum")
Str = Trim(m_frmShg.txtData(txtIndex))
If Len(Str) Then
    gDbTrans.SqlStmt = "Select AccID From SbMaster " & _
                " Where AccNum = " & AddQuotes(Str)
Else
    gDbTrans.SqlStmt = "Select AccID From SbMaster " & _
                " Where CustomerID = " & custId
End If
If gDbTrans.Fetch(Rst, adOpenDynamic) = 1 Then SbId = FormatField(Rst(0))

'Get the Loan account Number if any
txtIndex = GetIndex("LoanAccNum")
Str = Trim$(m_frmShg.txtData(txtIndex).Text)
If Len(Str) Then
    If m_LoanSchemeId Then
        gDbTrans.SqlStmt = "Select LoanID From LoanMaster " & _
                    " Where SchemeId = " & m_LoanSchemeId & _
                    " And AccNum = " & AddQuotes(Str)
    Else
        gDbTrans.SqlStmt = "Select LoanID From LoanMaster " & _
                    " Where AccNum = " & AddQuotes(Str) & _
                    " And CustomerID = " & custId
    End If
Else
    gDbTrans.SqlStmt = "Select LoanID From LoanMaster " & _
                " Where CustomerID = " & custId
End If
If gDbTrans.Fetch(Rst, adOpenDynamic) = 1 Then LoanID = FormatField(Rst(0))

If m_dbOperation = Insert Then
    m_AccID = 0
    gDbTrans.SqlStmt = "Select Max(AccId) From ShgMaster;"
    If gDbTrans.Fetch(Rst, adOpenDynamic) Then m_AccID = FormatField(Rst(0)) + 1
    If m_AccID <= 0 Then m_AccID = 1
    
    With m_frmShg
        SqlStr = "Insert Into ShgMaster (AccNum,AccId," & _
            "CustomerID,ContactPerson," & _
            "MeetingDay,MeetingPlace," & _
            "Gender,Place,Caste," & _
            IIf(Len(strCreateDate), "CreateDate,", "") & " TotalMembers," & _
            "FemaleMembers,ScStMembers,FemaleScStMembers," & _
            "SbAccId,LoanId ) " & _
            " VALUES  (" & _
            AddQuotes(GetVal("AccID")) & "," & m_AccID & "," & _
            custId & "," & AddQuotes(GetVal("Contact")) & "," & _
            AddQuotes(GetVal("MeetDay")) & "," & AddQuotes(GetVal("MeetPlace")) & "," & _
            intGender & "," & AddQuotes(GetVal("Place")) & "," & AddQuotes(GetVal("Caste")) & "," & _
            strCreateDate & Val(GetVal("TotalMem")) & "," & _
            Val(GetVal("FemaleMem")) & "," & Val(GetVal("SCSTMem")) & "," & Val(GetVal("FemaleSCSTMem")) & "," & _
            SbId & "," & LoanID & ")"
    End With

Else
    
    With m_frmShg
        SqlStr = "UPDate ShgMaster Set" & _
            " ContactPerson = " & AddQuotes(GetVal("Contact")) & "," & _
            " MeetingDay = " & AddQuotes(GetVal("MeetDay")) & "," & _
            " MeetingPlace = " & AddQuotes(GetVal("MeetPlace")) & "," & _
            " Gender = " & intGender & "," & _
            " Place = " & AddQuotes(GetVal("Place")) & "," & _
            " Caste = " & AddQuotes(GetVal("Caste")) & "," & _
            IIf(Len(strCreateDate), "CreateDate = " & strCreateDate, "") & _
            " TotalMembers = " & Val(GetVal("TotalMem")) & "," & _
            " FemaleMembers = " & Val(GetVal("FemaleMem")) & "," & _
            " ScStMembers = " & Val(GetVal("SCSTMem")) & "," & _
            " FemaleScStMembers = " & Val(GetVal("FemaleSCSTMem")) & "," & _
            " SbAccId = " & SbId & ", LoanId = " & LoanID & _
            " WHERE AccId = " & m_AccID
    End With
End If

gDbTrans.BeginTrans
m_CustReg.ModuleID = wis_Users
If Not m_CustReg.SaveCustomer Then gDbTrans.RollBacknRaiseError
gDbTrans.SqlStmt = SqlStr
If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

gDbTrans.CommitTrans

MsgBox GetResourceString(IIf(m_dbOperation = Insert, 528, 707)), vbInformation, wis_MESSAGE_TITLE

If m_dbOperation = Insert Then m_AccID = 0: Call LoadSHGDetail

Err_line:


End Function


' Returns the index of the control bound to "strDatasrc".
Private Function GetIndex(strDataSrc As String) As Integer
GetIndex = -1
Dim strTmp As String
Dim I As Integer
With m_frmShg
    For I = 0 To .txtPrompt.count - 1
        ' Get the data source for this control.
        strTmp = ExtractToken(.txtPrompt(I).Tag, "DataSource")
        If StrComp(strDataSrc, strTmp, vbTextCompare) = 0 Then
            GetIndex = I
            Exit For
        End If
    Next
End With

End Function

Public Sub Show()
'Get the Loan SchemID
Dim Rst As Recordset
gDbTrans.SqlStmt = "Select SchemeID From LoanScheme " & _
        " where LoanType = " & wisSHGLoan
If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then _
                m_LoanSchemeId = FormatField(Rst(0))
Set Rst = Nothing

If m_frmShg Is Nothing Then Set m_frmShg = New frmShg
Load m_frmShg
Call InitGrid

gWindowHandle = m_frmShg.hwnd
'now Check the permission
Dim Perms As wis_Permissions
Perms = gCurrUser.UserPermissions

If Perms = perBankAdmin Or Perms = perOnlyWaves Or Perms = perModifyAccount Then
    With m_frmShg
        .cmdDelete.Enabled = True
        .cmdSave.Enabled = True
    End With
Else
    With m_frmShg
        '.cmdDelete.Enabled = False
        '.cmdSave.Enabled = False
    End With
    'm_frmShg.cmdSave.Enabled = False
End If
m_frmShg.Show

End Sub


Private Function Validate1() As Boolean
#If junk Then
Validate1 = False

On Error GoTo Exit_line

Dim strMsg As String
If m_CustReg Is Nothing Then
    'No Customer Detials specified
    strMsg = GetResourceString(662)
    GoTo Exit_line
End If

If m_CustReg.CustomerID = 0 Then
    'No Customer Detials specified
    strMsg = GetResourceString(662)
    GoTo Exit_line
End If

Dim Rst As Recordset

With m_frmShg
    If Len(Trim$(.txtShgNum)) = 0 Then
        strMsg = GetResourceString(500)
        ActivateTextBox .txtShgNum
        GoTo Exit_line
    End If
    
    gDbTrans.SqlStmt = "Select * From ShgMaster " & _
                " Where AccNum = " & AddQuotes(.txtShgNum)
    If m_dbOperation = Update Then _
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And AccID <> " & m_AccID
        
    If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then
        strMsg = GetResourceString(545)
        ActivateTextBox .txtShgNum
        GoTo Exit_line
    End If

    If Len(.txtDOB) And Not DateValidate(.txtDOB, "/", True) Then
        ActivateTextBox .txtDOB
        strMsg = GetResourceString(499)
        GoTo Exit_line
    End If
    
    strMsg = "Invalid Member datails"
    strMsg = GetResourceString(760)
    If Not CurrencyValidate(.txtMemTotal, True) Then
        ActivateTextBox .txtMemTotal
        GoTo Exit_line
    End If
    'Checkthe No of members
    If Not CurrencyValidate(.txtFemaleMem, True) Then
        ActivateTextBox .txtFemaleMem
        GoTo Exit_line
    End If
    If Not CurrencyValidate(.txtScMem, True) Then
        ActivateTextBox .txtScMem
        GoTo Exit_line
    End If
    If Not CurrencyValidate(.txtFemSc, True) Then
        ActivateTextBox .txtFemSc
        GoTo Exit_line
    End If
    
    If .cmbGender.ListIndex < 0 Then .cmbGender.ListIndex = 0
    strMsg = "Specify the Gender"
    If .cmbGender.ListIndex < 0 Then
        ActivateTextBox .cmbGender
        GoTo Exit_line
    End If
    strMsg = "Specify the Place"
    If .cmbPlace.ListIndex < 0 Then .cmbPlace.ListIndex = 0
    
    strMsg = "Specify the Caste"
    If .cmbCaste.ListIndex < 0 Then .cmbCaste.ListIndex = 0
    
    strMsg = "Specify the meeting day "
    If .cmbDay.ListIndex < 0 Then
        ActivateTextBox .cmbDay
        GoTo Exit_line
    End If
    
    strMsg = "Contact person name not entered"
    If Len(Trim$(.txtContactPerson)) = 0 Then
        ActivateTextBox .txtMemTotal
        GoTo Exit_line
    End If
    strMsg = "Place of the weekly Meeting Not specified"
    If Len(Trim$(.txtMeetingPlace)) = 0 Then
        ActivateTextBox .txtMeetingPlace
        GoTo Exit_line
    End If

End With


Validate1 = True
Exit Function

Exit_line:

MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE
#End If
End Function

Private Function Validate() As Boolean
Validate = False

On Error GoTo Exit_line

Dim strMsg As String
If m_CustReg Is Nothing Then
    'No Customer Detials specified
    strMsg = GetResourceString(662)
    GoTo Exit_line
End If

If m_CustReg.CustomerID = 0 Then
    'No Customer Detials specified
    strMsg = GetResourceString(662)
    GoTo Exit_line
End If

Dim Rst As Recordset
Dim txtIndex As Byte

With m_frmShg
    'Check for the Shg Number
    txtIndex = GetIndex("AccID")
    If Len(Trim$(.txtData(txtIndex))) = 0 Then
        strMsg = GetResourceString(500)
        ActivateTextBox .txtData(txtIndex)
        GoTo Exit_line
    End If
    
    gDbTrans.SqlStmt = "Select * From ShgMaster " & _
                " Where AccNum = " & AddQuotes(Trim$(.txtData(txtIndex)))
    If m_dbOperation = Update Then _
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And AccID <> " & m_AccID
        
    If gDbTrans.Fetch(Rst, adOpenDynamic) > 0 Then
        strMsg = GetResourceString(545)
        ActivateTextBox .txtData(txtIndex)
        GoTo Exit_line
    End If

    'Check for the Date Of Inception
    txtIndex = GetIndex("CreateDate")
    If Len(.txtData(txtIndex)) And Not DateValidate(.txtData(txtIndex), "/", True) Then
        ActivateTextBox .txtData(txtIndex)
        strMsg = GetResourceString(499)
        GoTo Exit_line
    End If
    
    'Checkthe No of members
    txtIndex = GetIndex("TotalMem")
    strMsg = "Invalid Member datails"
    strMsg = GetResourceString(760)
    If Not CurrencyValidate(.txtData(txtIndex), True) Then
        ActivateTextBox .txtData(txtIndex)
        GoTo Exit_line
    End If
    
    'Checkthe No of female members
    txtIndex = GetIndex("FemaleMem")
    If Not CurrencyValidate(.txtData(txtIndex), True) Then
        ActivateTextBox .txtData(txtIndex)
        GoTo Exit_line
    End If
    
    'Checkthe No of Sc/St members
    txtIndex = GetIndex("SCSTMem")
    If Not CurrencyValidate(.txtData(txtIndex), True) Then
        ActivateTextBox .txtData(txtIndex)
        GoTo Exit_line
    End If
    
    txtIndex = GetIndex("FemaleSCSTMem")
    If Not CurrencyValidate(.txtData(txtIndex), True) Then
        ActivateTextBox .txtData(txtIndex)
        GoTo Exit_line
    End If
    
    txtIndex = GetIndex("Gender")
    txtIndex = Val(ExtractToken(.txtData(txtIndex).Tag, "textIndex"))
    If .cmb(txtIndex).ListIndex < 0 Then .cmb(txtIndex).ListIndex = 0
    strMsg = "Specify the Gender"
    If .cmb(txtIndex).ListIndex < 0 Then
        ActivateTextBox .cmb(txtIndex)
        GoTo Exit_line
    End If
    txtIndex = GetIndex("Place")
    txtIndex = Val(ExtractToken(.txtData(txtIndex).Tag, "textIndex"))
    If .cmb(txtIndex).ListIndex < 0 Then .cmb(txtIndex).ListIndex = 0
    
    txtIndex = GetIndex("Caste")
    txtIndex = Val(ExtractToken(.txtData(txtIndex).Tag, "textIndex"))
    If .cmb(txtIndex).ListIndex < 0 Then .cmb(txtIndex).ListIndex = 0
    
    strMsg = "Specify the meeting day "
    txtIndex = GetIndex("MeetDay")
    txtIndex = Val(ExtractToken(.txtData(txtIndex).Tag, "TextIndex"))
    If .cmb(txtIndex).ListIndex < 0 Then
        ActivateTextBox .cmb(txtIndex)
        GoTo Exit_line
    End If
    
    strMsg = "Contact person name not entered"
    txtIndex = GetIndex("Contact")
    If Len(Trim$(.txtData(txtIndex))) = 0 Then
        ActivateTextBox .txtData(txtIndex)
        GoTo Exit_line
    End If
    strMsg = "Place of the weekly Meeting Not specified"
    txtIndex = GetIndex("MeetPlace")
    If Len(Trim$(.txtData(txtIndex))) = 0 Then
        ActivateTextBox .txtData(txtIndex)
        GoTo Exit_line
    End If
    
End With

Validate = True
Exit Function

Exit_line:

MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE

End Function


Private Sub Class_Initialize()
m_dbOperation = Insert
End Sub

Private Sub m_frmShg_AddClick()
 
If AddTrainingRecord Then Call LoadTrainingDetail
 
End Sub

Private Sub m_frmShg_ClearClick()
    Call ClearControls
    Call ActivateTextBox(m_frmShg.txtData(GetIndex("AccID")))

End Sub

Private Sub m_frmShg_CloseClick()
gWindowHandle = 0
Set m_frmShg = Nothing

End Sub

Public Property Get ObjectHandle() As Long
    If Not m_frmShg Is Nothing Then _
            ObjectHandle = m_frmShg.hwnd: Exit Property
    If Not m_frmShgReport Is Nothing Then _
            Unload m_frmShgReport: Set m_frmShgReport = Nothing
End Property

Private Sub m_frmShg_CustClick()

If m_CustReg Is Nothing Then Set m_CustReg = New clsCustReg
Dim CustomerID As Long

With m_CustReg
    .ShowDialog
    If .CustomerLoaded Then
        'm_frmShg.txtCustName = .FullName
        m_frmShg.txtData(GetIndex("AccName")) = .FullName
        m_frmShg.CustomerID = .CustomerID
    End If
    CustomerID = .CustomerID
    
End With
If m_dbOperation = Update Then Exit Sub

'Check for the Existance of the Shg Group
Dim Rst As Recordset
Dim AccNum As String

gDbTrans.SqlStmt = "Select AccNum From shgMaster " & _
            " WHere CustomerID = " & CustomerID
If gDbTrans.Fetch(Rst, adOpenDynamic) < 1 Then Exit Sub

AccNum = FormatField(Rst("AccNum"))
'm_frmShg.txtShgNum = AccNum
m_frmShg.txtData(GetIndex("AccID")) = AccNum

'Now Load The Deatils
Call m_frmShg_LoadClick


End Sub


Private Sub m_frmShg_DeleteClick()

If m_AccID = 0 Then Exit Sub
If MsgBox(GetResourceString(539), vbQuestion + vbYesNo + vbDefaultButton2, _
            wis_MESSAGE_TITLE) = vbNo Then Exit Sub
            
gDbTrans.SqlStmt = "Delete A.* ,B.* " & _
            " From ShgMaster A,ShgTrans B " & _
            " Where A.AccID = B.AccID " & _
            " And A.AccID = " & m_AccID

gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Sub
End If
gDbTrans.CommitTrans

Call InitGrid
Call ClearControls

End Sub

Private Sub m_frmShg_DeleteTraining()

gDbTrans.SqlStmt = "Delete * " & _
            " From ShgTrans " & _
            " Where TransID = " & m_TransID

gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Sub
End If
gDbTrans.CommitTrans

End Sub

Private Sub m_frmShg_GridClick()

With m_frmShg.grd
    m_TransID = .RowData(.Row)
    If m_TransID = 0 Then Exit Sub
End With


Dim Rst As Recordset
gDbTrans.SqlStmt = "Select * From ShgTrans " & _
            " Where TransId = " & m_TransID
If gDbTrans.Fetch(Rst, adOpenDynamic) < 1 Then Exit Sub

With m_frmShg
    .txtTrainFrom = FormatField(Rst("FromDate"))
    .txtTrainTo = FormatField(Rst("ToDate"))
    .txtTraining = FormatField(Rst("TrainingDetail"))
    .txtTrainingPlace = FormatField(Rst("Place"))
    .txtTrainMembers = FormatField(Rst("Members"))
    
    .cmdAdd.Caption = GetResourceString(171)
    .cmdTrainingDelete.Enabled = True
End With


End Sub

Private Sub m_frmShg_LoadClick()

    Call LoadSHGDetail
    If m_AccID Then Call LoadTrainingDetail

End Sub


Private Sub m_frmShg_RepOptionClick()
    If m_clsRepOption Is Nothing Then _
        Set m_clsRepOption = New clsRepOption
    
    m_clsRepOption.EnableAmountRange = False
    
    With m_frmShg
        If .optReports(0) Or .optReports(1) Or .optReports(2) _
                Or .optReports(4) Or .optReports(5) Or .optReports(6) _
                Or .optReports(7) Or .optReports(8) Then
            m_clsRepOption.EnableCasteControls = True
        Else
            m_clsRepOption.EnableCasteControls = False
        End If
    End With
    
    m_clsRepOption.ShowDialog
    
End Sub

Private Sub m_frmShg_SaveClick()

If Not Validate Then Exit Sub

Call SaveDetail

End Sub


Private Sub m_frmShg_ShowClick()
Dim StrFrom As String
Dim strTo As String

Dim StrCaste As String
Dim strPLace As String
Dim intGender As Byte
Dim ReportType As wis_ShgReports
Dim ReportOrder As wis_ReportOrder

If m_clsRepOption Is Nothing Then _
    Set m_clsRepOption = New clsRepOption
    
With m_frmShg
    ReportOrder = IIf(.optAccId, wisByAccountNo, wisByName)
    StrCaste = m_clsRepOption.Caste
    strPLace = m_clsRepOption.Place
    intGender = m_clsRepOption.Gender
    
    If .txtDate1.Enabled Then StrFrom = .txtDate1.Text
    If .txtDate2.Enabled Then strTo = .txtDate2.Text
    
    If .optReports(0).Value Then ReportType = wisShgList
    If .optReports(1).Value Then ReportType = wisShgScStMembers
    If .optReports(2).Value Then ReportType = wisShgCreated
    If .optReports(3).Value Then ReportType = wisShgTraining
    If .optReports(4).Value Then ReportType = wisSHGSbBalance
    If .optReports(5).Value Then ReportType = wisShgLoanBalance
    If .optReports(6).Value Then ReportType = wisShgSbMonBalance
    If .optReports(7).Value Then ReportType = wisShgLoanMonBalnace
    If .optReports(8).Value Then ReportType = wisShgMonthlyStmt
    
End With

If Not m_frmShgReport Is Nothing Then Unload m_frmShgReport
Set m_frmShgReport = New frmShgReport

With m_frmShgReport
    .ReportOrder = ReportOrder
    .ReportType = ReportType
    .FromIndianDate = StrFrom
    .ToIndianDate = strTo
    .Caste = StrCaste
    .Gender = intGender
    .Place = strPLace
End With

Load m_frmShgReport

If gCancel <> 2 Then m_frmShgReport.Show 'vbModal
    

End Sub

