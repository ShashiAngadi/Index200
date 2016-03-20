VERSION 5.00
Begin VB.Form frmCreateCompany 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Creation"
   ClientHeight    =   4800
   ClientLeft      =   1260
   ClientTop       =   3045
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CreateCompany.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCompanyNameEnglish 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2430
      TabIndex        =   23
      Top             =   1200
      Width           =   4965
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   90
      TabIndex        =   22
      Top             =   4290
      Width           =   1215
   End
   Begin VB.CheckBox chkISOhterState 
      Caption         =   "Is Other state"
      Height          =   285
      Left            =   5490
      TabIndex        =   2
      Top             =   180
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6720
      TabIndex        =   21
      Top             =   4290
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1560
      TabIndex        =   20
      Top             =   4290
      Width           =   1215
   End
   Begin VB.CommandButton cmdCompany 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7620
      TabIndex        =   5
      Top             =   780
      Width           =   315
   End
   Begin VB.ComboBox cmbCompanyType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2460
      TabIndex        =   1
      Top             =   135
      Width           =   2925
   End
   Begin VB.TextBox txtEmail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5880
      TabIndex        =   19
      Top             =   3210
      Width           =   1815
   End
   Begin VB.TextBox txtMobileNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2460
      TabIndex        =   17
      Top             =   3330
      Width           =   1725
   End
   Begin VB.TextBox txtContactPerson 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2460
      TabIndex        =   15
      Top             =   2790
      Width           =   1725
   End
   Begin VB.TextBox txtPhoneNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2460
      TabIndex        =   11
      Top             =   2250
      Width           =   1725
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5910
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   2220
      Width           =   1815
   End
   Begin VB.TextBox txtCST 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2460
      TabIndex        =   7
      Top             =   1710
      Width           =   1725
   End
   Begin VB.TextBox txtKST 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5910
      TabIndex        =   9
      Top             =   1740
      Width           =   1815
   End
   Begin VB.TextBox txtCompanyName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2460
      TabIndex        =   4
      Top             =   780
      Width           =   4965
   End
   Begin VB.Label lblCompanyNameEnglish 
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   24
      Top             =   1230
      Width           =   2565
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   30
      X2              =   8040
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label lblEmail 
      Caption         =   "E-Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4470
      TabIndex        =   18
      Top             =   3240
      Width           =   1395
   End
   Begin VB.Label lblMobileNo 
      Caption         =   "Mobile Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   30
      TabIndex        =   16
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lblContactPerson 
      Caption         =   "Contact Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   30
      TabIndex        =   14
      Top             =   2820
      Width           =   2385
   End
   Begin VB.Label lblPhoneNo 
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   30
      TabIndex        =   10
      Top             =   2280
      Width           =   2475
   End
   Begin VB.Label lblAddress 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4500
      TabIndex        =   12
      Top             =   2250
      Width           =   1395
   End
   Begin VB.Label lblCST 
      Caption         =   "CST No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   30
      TabIndex        =   6
      Top             =   1740
      Width           =   2265
   End
   Begin VB.Label lblKST 
      Caption         =   "KST No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4500
      TabIndex        =   8
      Top             =   1770
      Width           =   1395
   End
   Begin VB.Label lblCompanyType 
      Caption         =   "Company Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Top             =   135
      Width           =   2325
   End
   Begin VB.Label lblCompanyName 
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   30
      TabIndex        =   3
      Top             =   810
      Width           =   2325
   End
End
Attribute VB_Name = "frmCreateCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_dbOperation As wis_DBOperation
Private m_HeadID As Long
Private m_CompanyType As wis_CompanyType

Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1


Private Function DeleteCompany() As Boolean


Dim eCompanyType As wis_CompanyType

On Error GoTo ErrLine

If cmbCompanyType.ListIndex = -1 Then Exit Function

DeleteCompany = False

If MsgBox("You are going to delete the company" & _
        vbCrLf & "Are you sure? ", vbQuestion + vbYesNo) = vbNo Then Exit Function

eCompanyType = cmbCompanyType.ItemData(cmbCompanyType.ListIndex)

If HasTranasction(m_HeadID) Then
    MsgBox "Has Transactions, will not be deleted.", vbInformation
    Call ClearControls
    Exit Function
End If

If eCompanyType <> Enum_Branch Then
    gDbTrans.SqlStmt = " DELETE * FROM CompanyCreation " & _
                       " WHERE HeadID = " & m_HeadID
    
Else
    gDbTrans.SqlStmt = " DELETE * FROM GodownDet " & _
                       " WHERE GodownID = " & m_HeadID
    
End If

gDbTrans.BeginTrans
'Delete the Company Details
If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

'Delete the Opening Balance
gDbTrans.SqlStmt = "DELETE * FROM OPBalance" & _
                            " WHERE HeadID = " & m_HeadID
If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
  
'now Delete the Head From The Heads table
gDbTrans.SqlStmt = "DELETE * FROM Heads " & _
                            " WHERE HeadID = " & m_HeadID

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
gDbTrans.CommitTrans

DeleteCompany = True

Exit Function

ErrLine:
    MsgBox "DeleteCompany" & vbCrLf & Err.Description, vbCritical
End Function

'set the Kannada option here.
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)
txtCompanyNameEnglish.Font.name = "MS Sans Serif"
If gLangOffSet = 0 Then txtCompanyNameEnglish.Enabled = False
'set the Kannada for all controls
lblCompanyType.Caption = GetResourceString(243)
lblCompanyName.Caption = GetResourceString(138)
lblCompanyNameEnglish.Caption = GetResourceString(138, 468)
lblCST.Caption = GetResourceString(242, 60)
lblKST.Caption = GetResourceString(241, 60)
lblPhoneNo.Caption = GetResourceString(135, 60)
lblAddress.Caption = GetResourceString(130)
lblContactPerson.Caption = GetResourceString(236)
lblMobileNo.Caption = GetResourceString(239, 60)
lblEmail.Caption = GetResourceString(136)

chkISOhterState.Caption = GetResourceString(128)
cmdOk.Caption = GetResourceString(1)
cmdCancel.Caption = GetResourceString(2)
cmdDelete.Caption = GetResourceString(14)

End Sub


Private Sub ClearControls()
txtCompanyName.Text = ""
txtCST.Text = ""
txtKST.Text = ""
txtPhoneNo.Text = ""
txtAddress.Text = ""
txtContactPerson.Text = ""
txtMobileNo.Text = ""
txtEmail.Text = ""
chkISOhterState.Value = vbUnchecked
cmbCompanyType.Locked = False

cmdOk.Caption = GetResourceString(1) '"&Ok"
m_dbOperation = Insert
cmdDelete.Enabled = False

On Error Resume Next
cmbCompanyType.SetFocus
End Sub

Private Sub SetFont()
On Error Resume Next

Dim Ctrl As Control

For Each Ctrl In Me
      Ctrl.FontName = "Arial"
      Ctrl.FONTSIZE = 10
Next Ctrl


End Sub


Private Sub LoadCompanyDetails(ByVal lngHeadID As Long)
Dim rstCompany As ADODB.Recordset
Dim TheState As wis_StatePosition


gDbTrans.SqlStmt = " SELECT * FROM CompanyCreation " & _
                   " WHERE HeadID = " & lngHeadID

If gDbTrans.Fetch(rstCompany, adOpenForwardOnly) < 1 Then Exit Sub

With Me
   .txtCompanyName = FormatField(rstCompany("CompanyName"))
   .txtCompanyNameEnglish = FormatField(rstCompany("CompanyNameEnglish"))
   .txtCST = FormatField(rstCompany("CST"))
   .txtCST.Enabled = True
   .txtKST = FormatField(rstCompany("KST"))
   .txtCST.Enabled = True
   .txtPhoneNo = FormatField(rstCompany("PhoneNo"))
   .txtAddress = FormatField(rstCompany("Address"))
   .txtContactPerson = FormatField(rstCompany("ContactPerson"))
   .txtMobileNo = FormatField(rstCompany("MobileNO"))
   .txtEmail = FormatField(rstCompany("Email"))
    
    TheState = FormatField(rstCompany("SameState"))
    chkISOhterState.Value = vbUnchecked
    If TheState = OtherState Then chkISOhterState.Value = vbChecked
End With

End Sub

Private Sub LoadGodownDetails(ByVal GodownID As Long)
Dim rstCompany As ADODB.Recordset
Dim TheState As wis_StatePosition


gDbTrans.SqlStmt = " SELECT * FROM GodownDet " & _
                   " WHERE GodownID = " & GodownID

If gDbTrans.Fetch(rstCompany, adOpenForwardOnly) < 1 Then Exit Sub

With Me
   .txtCompanyName = rstCompany.Fields("GodownName").Value
   .txtCompanyNameEnglish = rstCompany.Fields("GodownNameEnglish").Value
   .txtPhoneNo = rstCompany.Fields("PhoneNo").Value
   .txtAddress = rstCompany.Fields("Address").Value
   .txtContactPerson = rstCompany.Fields("ContactPerson").Value
   .txtMobileNo = rstCompany.Fields("MobileNO").Value
   .txtEmail = rstCompany.Fields("Email").Value
   .txtKST.Enabled = False
   .txtCST.Enabled = False
End With

End Sub
'
Private Function SaveDetails() As Boolean
Dim enumCompanyType As wis_CompanyType
Dim headID As Long
Dim OpAmount As Currency

Dim rst As ADODB.Recordset
Dim TheSameState As wis_StatePosition
Dim ParentID As Long

SaveDetails = False
If gLangOffSet = 0 Then txtCompanyNameEnglish.Text = txtCompanyName.Text

On Error GoTo ErrLine
If cmbCompanyType.ListIndex = -1 Then Exit Function
enumCompanyType = cmbCompanyType.ItemData(cmbCompanyType.ListIndex)
'opamount
'Insert the data into the database
Select Case enumCompanyType
    'Case Enum_Manufacturer
    '    ParentID = wis_CreditorsParentID
    '    gDbTrans.SQLStmt = " SELECT MAX(HeadID) FROM Heads " & _
                      " WHERE HeadID BETWEEN " & wis_CreditorsParentID & " AND " & (wis_CreditorsParentID + SUB_HEAD_OFFSET)
    Case Enum_Stockist
        ParentID = wis_CreditorsParentID
        gDbTrans.SqlStmt = " SELECT MAX(HeadID) FROM Heads " & _
                      " WHERE HeadID BETWEEN " & wis_CreditorsParentID & " AND " & (wis_CreditorsParentID + SUB_HEAD_OFFSET)
    Case Enum_Customers
        gDbTrans.SqlStmt = " SELECT MAX(HeadID) FROM Heads " & _
                        " WHERE HeadID BETWEEN " & (wis_DebitorsParentID + SUB_HEAD_OFFSET) & " AND " & (wis_DebitorsParentID + 2 * SUB_HEAD_OFFSET)
   
        ParentID = wis_DebitorsParentID
    Case Enum_Branch
        gDbTrans.SqlStmt = " SELECT MAX(GodownID) FROM GodownDet "

End Select
      
Call gDbTrans.Fetch(rst, adOpenForwardOnly)
   
headID = FormatField(rst(0)) + 1

If enumCompanyType <> Enum_Branch Then
    
    If headID < HEAD_OFFSET Then
        Select Case enumCompanyType
            'Case Enum_Manufacturer
            '    HeadID = HeadID + parPayAble  'wis_CreditorsParentID
            Case Enum_Stockist
                headID = headID + parPayAble 'wis_CreditorsParentID
            Case Enum_Customers
                headID = headID + parReceivable 'wis_DebitorsParentID
        End Select
    End If

    TheSameState = SameState
    
    If chkISOhterState.Value = vbChecked Then TheSameState = OtherState
    
    'Insert into the database
    gDbTrans.SqlStmt = "INSERT INTO CompanyCreation " & _
            " (HeadID,CompanyName,CompanyNameEnglish,CompanyType, " & _
            " KST,CST,Address,PhoneNo,ContactPerson," & _
            " MobileNo,Email,SameState ) " & _
            " VALUES ( " & _
            headID & "," & _
            AddQuotes(txtCompanyName.Text, True) & "," & _
            AddQuotes(txtCompanyNameEnglish.Text, True) & "," & _
            enumCompanyType & "," & _
            AddQuotes(txtKST.Text, True) & "," & _
            AddQuotes(txtCST.Text, True) & "," & _
            AddQuotes(txtAddress.Text, True) & "," & _
            AddQuotes(txtPhoneNo.Text, True) & "," & _
            AddQuotes(txtContactPerson.Text, True) & "," & _
            AddQuotes(txtMobileNo.Text, True) & "," & _
            AddQuotes(txtEmail.Text, True) & "," & _
            TheSameState & " ) "

Else 'If enumCompanyType = Branch Then
    
    gDbTrans.SqlStmt = "INSERT INTO GodownDet(GodownID,GodownName,GodownNameEnglish," & _
                " Address,PhoneNo,ContactPerson,MobileNo,Email) " & _
                " VALUES ( " & _
                headID & "," & _
                AddQuotes(txtCompanyName.Text, True) & "," & _
                AddQuotes(txtCompanyNameEnglish.Text, True) & "," & _
                AddQuotes(txtAddress.Text, True) & "," & _
                AddQuotes(txtPhoneNo.Text, True) & "," & _
                AddQuotes(txtContactPerson.Text, True) & "," & _
                AddQuotes(txtMobileNo.Text, True) & "," & _
                AddQuotes(txtEmail.Text, True) & " ) "

End If

gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   

If enumCompanyType <> Enum_Branch Then
    
    gDbTrans.SqlStmt = " INSERT INTO Heads (HeadID,HeadName,HeadNameEnglish,ParentID)" & _
                    " VALUES ( " & _
                    headID & "," & _
                    AddQuotes(txtCompanyName.Text, True) & "," & _
                    AddQuotes(txtCompanyNameEnglish.Text, True) & "," & _
                    ParentID & " )"
                       
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
        
    OpAmount = Val(InputBox("Enter Opening Balance", "Opening Balnace"))
    gDbTrans.SqlStmt = " INSERT INTO OpBalance (HeadID,OpDate,OpAmount)" & _
                       " VALUES ( " & _
                       headID & "," & _
                       "#" & GetSysFormatDate(FinIndianFromDate) & "#," & _
                       OpAmount & " )"
                       
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    
End If

gDbTrans.CommitTrans

SaveDetails = True
MsgBox "Saved the Details", vbInformation, wis_MESSAGE_TITLE

ClearControls

cmdOk.Caption = GetResourceString(1) '"&Ok"
m_dbOperation = Insert

Exit Function

ErrLine:
    MsgBox "SaveDetails" & vbCrLf & Err.Description, vbCritical
    'Resume
End Function
Private Function UpdateDetails() As Boolean
Dim enumCompanyType As wis_CompanyType
Dim TheSameState As wis_StatePosition
Dim ParentID As Long

On Error GoTo ErrLine

If gLangOffSet = 0 Then txtCompanyNameEnglish.Text = txtCompanyName.Text

UpdateDetails = False

If cmbCompanyType.ListIndex = -1 Then Exit Function

enumCompanyType = cmbCompanyType.ItemData(cmbCompanyType.ListIndex)

TheSameState = SameState

If chkISOhterState.Value = vbChecked Then TheSameState = OtherState

ParentID = wis_CreditorsParentID

If enumCompanyType = Enum_Customers Then ParentID = wis_DebitorsParentID
   
If enumCompanyType <> Enum_Branch Then
    gDbTrans.SqlStmt = " UPDATE CompanyCreation SET " & _
                    " CompanyName = " & AddQuotes(txtCompanyName.Text, True) & "," & _
                    " CompanyNameEnglish = " & AddQuotes(txtCompanyNameEnglish.Text, True) & "," & _
                    " CompanyType = " & enumCompanyType & "," & _
                    " KST =" & AddQuotes(txtKST.Text, True) & "," & _
                    " CST =" & AddQuotes(txtCST.Text, True) & "," & _
                    " Address = " & AddQuotes(txtAddress.Text, True) & "," & _
                    " PhoneNO = " & AddQuotes(txtPhoneNo.Text, True) & "," & _
                    " ContactPerson =" & AddQuotes(txtContactPerson.Text, True) & "," & _
                    " MobileNo = " & AddQuotes(txtMobileNo.Text, True) & "," & _
                    " SameState = " & TheSameState & "," & _
                    " Email = " & AddQuotes(txtEmail.Text, True) & _
                    " WHERE HeadID = " & m_HeadID

Else
    gDbTrans.SqlStmt = " UPDATE GodownDet SET " & _
                    " GodownName = " & AddQuotes(txtCompanyName.Text, True) & "," & _
                    " GodownNameEnglish = " & AddQuotes(txtCompanyNameEnglish.Text, True) & "," & _
                    " Address = " & AddQuotes(txtAddress.Text, True) & "," & _
                    " PhoneNO = " & AddQuotes(txtPhoneNo.Text, True) & "," & _
                    " ContactPerson =" & AddQuotes(txtContactPerson.Text, True) & "," & _
                    " MobileNo = " & AddQuotes(txtMobileNo.Text, True) & "," & _
                    " Email = " & AddQuotes(txtEmail.Text, True) & _
                    " WHERE GodownID = " & m_HeadID
End If

gDbTrans.BeginTrans

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
    

If enumCompanyType <> Enum_Branch Then
    gDbTrans.SqlStmt = " UPDATE Heads SET " & _
                       " HeadName = " & AddQuotes(txtCompanyName.Text) & "," & _
                       " ParentID =" & ParentID & _
                       " WHERE HeadID = " & m_HeadID
                   
    If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
        
End If

gDbTrans.CommitTrans

UpdateDetails = True

MsgBox "Details updated ", vbInformation, wis_MESSAGE_TITLE

'Clear the controls
ClearControls

cmdOk.Caption = GetResourceString(1) ' "&Ok"
m_dbOperation = Insert

Exit Function

ErrLine:
    MsgBox "UpdateDetails" & vbCrLf & Err.Description, vbCritical
    
End Function


Private Function Validated() As Boolean
Validated = False

If cmbCompanyType.ListIndex = -1 Then Exit Function
If txtCompanyName.Text = "" Then Exit Function

Validated = True

End Function

Private Sub cmbCompanyType_Click()
Dim eCompanyType As wis_CompanyType

If cmbCompanyType.ListIndex = -1 Then Exit Sub

eCompanyType = cmbCompanyType.ItemData(cmbCompanyType.ListIndex)

txtCST.Enabled = True
txtKST.Enabled = True
If eCompanyType = Enum_Branch Then
    txtCST.Enabled = False
    txtKST.Enabled = False
End If

End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCompany_Click()

Dim enumCompanyType As wis_CompanyType
Dim rst As ADODB.Recordset

If cmbCompanyType.ListIndex = -1 Then Exit Sub

enumCompanyType = cmbCompanyType.ItemData(cmbCompanyType.ListIndex)

If enumCompanyType <> Enum_Branch Then
    gDbTrans.SqlStmt = " SELECT HeadID,CompanyName,Address " & _
                         " FROM CompanyCreation " & _
                         " WHERE CompanyType = " & enumCompanyType & _
                         " ORDER BY CompanyName"

Else
    gDbTrans.SqlStmt = " SELECT GodownID,GodownName,Address " & _
                    " FROM GodownDet " & _
                    " WHERE GodownID > " & 1 & _
                    " ORDER BY GodownName"
End If


If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Sub
If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp

If enumCompanyType <> Enum_Branch Then
    'If Not FillView(m_frmLookUp.lvwReport, Rst, "HeadID", True) Then Exit Sub
    If Not FillView(m_frmLookUp.lvwReport, rst, True) Then Exit Sub
Else
    'If Not FillView(m_frmLookUp.lvwReport, Rst, "GodownID", True) Then Exit Sub
    If Not FillView(m_frmLookUp.lvwReport, rst, True) Then Exit Sub
End If

m_HeadID = 0
m_frmLookUp.Show vbModal

If m_HeadID > 0 Then
    If enumCompanyType = Enum_Branch Then
        Call LoadGodownDetails(m_HeadID)
    Else
        Call LoadCompanyDetails(m_HeadID)
    End If
    m_dbOperation = Update
    cmdOk.Caption = GetResourceString(171) '"&Update"
    cmbCompanyType.Locked = True
    cmdDelete.Enabled = True
End If
End Sub

Private Sub cmdDelete_Click()
If Not DeleteCompany Then
    MsgBox "Unable to Delete the company "
    Exit Sub
End If

Call ClearControls

End Sub


Private Function HasTranasction(ByVal headID As Long) As Boolean
Dim rst As ADODB.Recordset



gDbTrans.SqlStmt = " SELECT HeadID " & _
                  " FROM AccTrans " & _
                  " WHERE HeadID=" & headID
If gDbTrans.Fetch(rst, adOpenForwardOnly) < 1 Then Exit Function
                  
'HasTranasction = True

Set rst = Nothing

Exit Function


End Function



Private Sub cmdOk_Click()

If Not Validated Then Exit Sub
If m_dbOperation = Insert Then SaveDetails
If m_dbOperation = Update Then UpdateDetails
End Sub


Private Sub Form_Load()
'Center the form
CenterMe Me

'Set the icon
Me.Icon = LoadResPicture(147, vbResIcon)
'set kannada caption
Call SetKannadaCaption

With cmbCompanyType
'   .AddItem GetResourceString(174) '"Manufacturer"
'   .ItemData(.NewIndex) = 1
   .AddItem GetResourceString(204) '"Stockist"
   .ItemData(.newIndex) = 3
   .AddItem GetResourceString(205) '"Customer"
   .ItemData(.newIndex) = 2
   .AddItem GetResourceString(227) '"Branch"
   .ItemData(.newIndex) = 4
End With


m_dbOperation = Insert
cmdOk.Caption = GetResourceString(1) '"&Ok"
cmdDelete.Enabled = False

End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmCreateCompany = Nothing
Set m_frmLookUp = Nothing
End Sub


Private Sub m_frmCompanyType_OKClicked(intSelection As Integer)
m_CompanyType = intSelection
End Sub

Private Sub m_frmLookUp_SelectClick(strSelection As String)
m_HeadID = CLng(strSelection)
End Sub


Private Sub txtAddress_LostFocus()
'txtAddress.Text = ConvertToProperCase(txtAddress.Text)
End Sub


Private Sub txtCompanyName_LostFocus()
'txtCompanyName.Text = ConvertToProperCase(txtCompanyName)
End Sub


Private Sub txtCompanyNameEnglish_GotFocus()
    Call Translate(txtCompanyName, txtCompanyNameEnglish)
End Sub

Private Sub txtCompanyNameEnglish_LostFocus()
    Call ToggleWindowsKey(winScrlLock, True)
End Sub

Private Sub txtContactPerson_LostFocus()
'txtContactPerson.Text = ConvertToProperCase(txtContactPerson.Text)
End Sub


