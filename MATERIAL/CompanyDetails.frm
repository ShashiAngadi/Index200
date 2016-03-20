VERSION 5.00
Begin VB.Form frmCompanyDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Details"
   ClientHeight    =   4020
   ClientLeft      =   2910
   ClientTop       =   2895
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CompanyDetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7065
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
      Left            =   2160
      TabIndex        =   18
      Tag             =   "Shows the Name of the Company"
      Top             =   480
      Width           =   4155
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
      Left            =   5790
      TabIndex        =   17
      Top             =   3420
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
      Left            =   4320
      TabIndex        =   16
      Top             =   3420
      Width           =   1215
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
      Left            =   4980
      TabIndex        =   15
      Top             =   2460
      Width           =   1845
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
      Left            =   1470
      TabIndex        =   13
      Top             =   2460
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
      Left            =   1470
      TabIndex        =   11
      Top             =   2010
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
      Left            =   1470
      TabIndex        =   7
      Top             =   1470
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
      Left            =   4980
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1470
      Width           =   1845
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
      Left            =   1470
      TabIndex        =   3
      Top             =   960
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
      Left            =   4980
      TabIndex        =   5
      Top             =   960
      Width           =   1845
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
      Left            =   1470
      TabIndex        =   1
      Tag             =   "Shows the Name of the Company"
      Top             =   120
      Width           =   4875
   End
   Begin VB.Label lblCompanyEnglish 
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
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   510
      Width           =   1995
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   30
      X2              =   6930
      Y1              =   3180
      Y2              =   3180
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
      Height          =   315
      Left            =   3330
      TabIndex        =   14
      Top             =   2490
      Width           =   1275
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
      Height          =   315
      Left            =   30
      TabIndex        =   12
      Top             =   2490
      Width           =   1425
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
      Height          =   315
      Left            =   30
      TabIndex        =   10
      Top             =   2040
      Width           =   1515
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
      Height          =   315
      Left            =   30
      TabIndex        =   6
      Top             =   1500
      Width           =   1425
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
      Height          =   315
      Left            =   3330
      TabIndex        =   8
      Top             =   1500
      Width           =   1275
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
      Height          =   315
      Left            =   30
      TabIndex        =   2
      Top             =   990
      Width           =   1425
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
      Height          =   315
      Left            =   3360
      TabIndex        =   4
      Top             =   990
      Width           =   1275
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
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   150
      Width           =   1395
   End
End
Attribute VB_Name = "frmCompanyDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_dbOperation As wis_DBOperation
Private m_CompanyID As Long
Private m_CompanyName As String


Public Event OkClicked(strCompanyName As String)
Public Event CancelClicked()


'set the Kannada option here.
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

'set the Kannada for all controls
lblCompanyName.Caption = GetResourceString(138)
lblCST.Caption = GetResourceString(242, 60)
lblKST.Caption = GetResourceString(241, 60)
lblPhoneNo.Caption = GetResourceString(135, 60)
lblAddress.Caption = GetResourceString(130)
lblContactPerson.Caption = GetResourceString(236)
lblMobileNo.Caption = GetResourceString(239, 60)
lblEmail.Caption = GetResourceString(136)
cmdOk.Caption = GetResourceString(1)
cmdCancel.Caption = GetResourceString(2)

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
End Sub

Private Sub SetFont()
On Error Resume Next

Dim Ctrl As Control

For Each Ctrl In Me
      Ctrl.FontName = "Arial"
      Ctrl.FONTSIZE = 10
Next Ctrl


End Sub


Private Sub LoadCompanyDetails()
Dim rstCompany As ADODB.Recordset
Dim enumCompanyType As wis_CompanyType

enumCompanyType = Enum_Self

gDbTrans.SqlStmt = " SELECT * FROM CompanyCreation" & _
                   " WHERE CompanyType = " & enumCompanyType

If gDbTrans.Fetch(rstCompany, adOpenForwardOnly) < 1 Then Exit Sub

With Me
   .txtCompanyName = FormatField(rstCompany.Fields("CompanyName"))
   .txtCompanyNameEnglish = FormatField(rstCompany.Fields("CompanyNameEnglish"))
   .txtCST = FormatField(rstCompany.Fields("CST"))
   .txtKST = FormatField(rstCompany.Fields("KST"))
   .txtPhoneNo = FormatField(rstCompany.Fields("PhoneNo"))
   .txtAddress = FormatField(rstCompany.Fields("Address"))
   .txtContactPerson = FormatField(rstCompany.Fields("ContactPerson"))
   .txtMobileNo = FormatField(rstCompany.Fields("MobileNO"))
   .txtEmail = FormatField(rstCompany.Fields("Email"))
   
   .cmdOk.Caption = GetResourceString(171) '"&Update"
   m_dbOperation = Update
   
   m_CompanyID = FormatField(rstCompany.Fields("HeadID"))
End With

End Sub

Private Function UpdateDetails() As Boolean
Dim rst As ADODB.Recordset

On Error GoTo ErrLine

UpdateDetails = False

m_CompanyName = Trim$(txtCompanyName.Text)

gDbTrans.SqlStmt = " UPDATE CompanyCreation SET " & _
                " CompanyName = " & AddQuotes(m_CompanyName, True) & "," & _
                " CompanyNameEnglish = " & AddQuotes(txtCompanyNameEnglish.Text, True) & "," & _
                " CompanyType = " & 0 & "," & _
                " KST =" & AddQuotes(txtKST, True) & "," & _
                " CST =" & AddQuotes(txtCST, True) & "," & _
                " Address = " & AddQuotes(txtAddress, True) & "," & _
                " PhoneNO =" & AddQuotes(txtPhoneNo, True) & "," & _
                " ContactPerson =" & AddQuotes(txtContactPerson, True) & "," & _
                " MobileNo = " & AddQuotes(txtMobileNo, True) & "," & _
                " Email = " & AddQuotes(txtEmail, True) & _
                " WHERE HeadID = " & m_CompanyID


gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError

gDbTrans.SqlStmt = " UPDATE GodownDet SET " & _
                " GodownName = " & AddQuotes(m_CompanyName, True) & _
                " ,GodownNameEnglish = " & AddQuotes(txtCompanyNameEnglish.Text, True) & _
                " WHERE GodownID = " & m_CompanyID

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   

gDbTrans.CommitTrans

UpdateDetails = True
MsgBox GetResourceString(707), vbInformation, wis_MESSAGE_TITLE '"Details Updated"

Call ClearControls

cmdOk.Caption = GetResourceString(171) ' "&Update"
m_dbOperation = Update

Exit Function

ErrLine:
    MsgBox "UpdateDetails" & vbCrLf & Err.Description, vbCritical
End Function

Private Function SaveDetails() As Boolean
Dim rst As ADODB.Recordset

On Error GoTo ErrLine

SaveDetails = False
m_CompanyName = Trim$(txtCompanyName.Text)

'Insert the data into the database
gDbTrans.SqlStmt = " INSERT INTO CompanyCreation (HeadID,CompanyName,CompanyNameEnglish,CompanyType, " & _
            " KST,CST,Address,PhoneNo,ContactPerson,MobileNo,Email ) " & _
            " VALUES ( " & _
            1 & "," & _
            AddQuotes(m_CompanyName) & "," & _
            AddQuotes(txtCompanyNameEnglish.Text) & "," & _
            0 & "," & _
            AddQuotes(txtKST.Text) & "," & _
            AddQuotes(txtCST.Text) & "," & _
            AddQuotes(txtAddress.Text) & "," & _
            AddQuotes(txtPhoneNo.Text) & "," & _
            AddQuotes(txtContactPerson.Text) & "," & _
            AddQuotes(txtMobileNo.Text) & "," & _
            AddQuotes(txtEmail.Text) & " ) "

            
gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
'Now Insert these data into the GodownDet
gDbTrans.SqlStmt = " INSERT INTO GodownDet(GodownID,GodownName) " & _
                   " VALUES  ( " & _
                   1 & "," & _
                   AddQuotes(m_CompanyName) & " )"

If Not gDbTrans.SQLExecute Then gDbTrans.RollBacknRaiseError
   
If Not InsertCashInHand Then Exit Function

gDbTrans.CommitTrans

SaveDetails = True



MsgBox GetResourceString(528), vbInformation, wis_MESSAGE_TITLE '"Saved the Details"

Call ClearControls

cmdOk.Caption = GetResourceString(171) '"&Update"
m_dbOperation = Update


Exit Function

ErrLine:
    MsgBox "SaveDetails" & vbCrLf & Err.Description, vbCritical
    
End Function

Private Function InsertCashInHand() As Boolean
Dim LedgerClass As clsLedger
Dim headName As String
Dim headNameEnglish As String

Set LedgerClass = New clsLedger

headName = "Cash In Hand"
headName = GetResourceString(442)
headNameEnglish = LoadResString(442)

If LedgerClass.GetHeadIDCreated(wis_CashParentID, headName, 0, headNameEnglish) > 0 Then InsertCashInHand = True

End Function


Private Function Validated() As Boolean
Validated = False


If txtCompanyName.Text = "" Then Exit Function

Validated = True

End Function

Private Sub cmdCancel_Click()

Unload Me
Set frmCompanyDetails = Nothing
If m_dbOperation = Insert Then ShutDownInventory
End Sub

Private Sub cmdOk_Click()

If Not Validated Then Exit Sub

If m_dbOperation = Insert Then SaveDetails
If m_dbOperation = Update Then UpdateDetails


RaiseEvent OkClicked(m_CompanyName)
'Me.Hide
Unload Me
Set frmCompanyDetails = Nothing
End Sub


Private Sub Form_Load()
Debug.Print IIf(gDbTrans Is Nothing, "NOthing", "object")
'Center the form
CenterMe Me

'Set the icon
Me.Icon = LoadResPicture(147, vbResIcon)

'set kannada fonts
SetKannadaCaption
Call SkipFontToControls(lblCompanyName, lblCompanyEnglish, txtCompanyNameEnglish)
If gLangOffSet = 0 Then
    Call ReduceControlsTopPosition(lblCompanyEnglish.Top - lblCompanyEnglish.Top, _
        lblContactPerson, lblCST, lblKST, lblPhoneNo, lblAddress, lblContactPerson, lblMobileNo, lblEmail, _
        txtCST, txtKST, txtPhoneNo, txtAddress, txtContactPerson, txtMobileNo, txtEmail, cmdOk, cmdCancel, Line1)
    Me.Height = Me.Height - (lblCompanyEnglish.Top - lblCompanyEnglish.Top)
End If

m_dbOperation = Insert
cmdOk.Caption = GetResourceString(1) '"&Ok"

'Load CompanyDetails
LoadCompanyDetails
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Unload Me
Set frmCompanyDetails = Nothing

If m_dbOperation = Insert Then ShutDownInventory

End Sub


Private Sub txtAddress_LostFocus()
'txtAddress.Text = ConvertToProperCase(txtAddress.Text)
End Sub


Private Sub txtCompanyName_LostFocus()
Debug.Print IIf(gDbTrans Is Nothing, "NOthing", "object")
'txtCompanyName.Text = ConvertToProperCase(txtCompanyName)
End Sub


Private Sub txtCompanyNameEnglish_GotFocus()
    Call ToggleWindowsKey(winScrlLock, False)
End Sub

Private Sub txtCompanyNameEnglish_LostFocus()
    Call ToggleWindowsKey(winScrlLock, True)
End Sub

Private Sub txtContactPerson_LostFocus()
'txtContactPerson = ConvertToProperCase(txtContactPerson)
End Sub

