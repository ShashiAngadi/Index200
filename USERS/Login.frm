VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please identify yourself"
   ClientHeight    =   2715
   ClientLeft      =   2775
   ClientTop       =   2955
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5265
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   90
      TabIndex        =   8
      Top             =   60
      Width           =   5025
      Begin VB.ComboBox cmbFinancialYear 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   5
         Text            =   "cmbFinancialYear"
         ToolTipText     =   "Select the Finanicial Year You want to Explore"
         Top             =   1455
         Width           =   2385
      End
      Begin VB.TextBox txtUserPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   795
         Width           =   2385
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Width           =   2385
      End
      Begin VB.Image img 
         Height          =   480
         Left            =   150
         Picture         =   "Login.frx":0000
         Top             =   540
         Width           =   480
      End
      Begin VB.Label lblUserDate 
         Caption         =   "Finanacial Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   1455
         Width           =   1455
      End
      Begin VB.Label lblUserPassword 
         Caption         =   "User password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   795
         Width           =   1455
      End
      Begin VB.Label lblUserName 
         Caption         =   "User name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3840
      TabIndex        =   7
      Top             =   2130
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2550
      TabIndex        =   6
      Top             =   2130
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event LoginClicked(UserName As String, _
                            Userpassword As String, _
                            LoginDate As String, _
                            UnloadDialog As Boolean)

Public Event CancelClicked()
Public Event FinYearSelected(ByVal YearID As Integer)
Public Event FinYearChanged(ByVal YearID As Integer)

Private Sub SetKannadaCaption()
Call SetFontToControls(Me)

lblUserName.Caption = GetResourceString(151, 35)
lblUserPassword.Caption = GetResourceString(151, 153)
lblUserDate.Caption = GetResourceString(37)
cmdCancel.Caption = GetResourceString(2)
cmdLogin.Caption = GetResourceString(151)
Me.FontName = gFontName
End Sub
'
'This subroutine will add the Financial Year to the ComboBox
'From the External file FinYear.fin File
'WRITTEN By Lingappa Sindhanur
'DATED   "June 19, 2002
Private Sub GetFinancialYear()
Dim YearID As Long

Dim FinYearClass As clsFinChange

YearID = cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex)
   
Set FinYearClass = New clsFinChange

Call FinYearClass.LoadFinYearData(App.Path & "\" & constFINYEARFILE, YearID)
Dim rst As Recordset
If DayBeginDate = "" Then DayBeginDate = GetSysFormatDate(gStrDate)

'Now Initilise the Global varible
'Now get the Name of the Bank /Society from DataBase
gDbTrans.SqlStmt = "Select * from CompanyCreation"
If gDbTrans.Fetch(rst, adOpenStatic) > 0 Then _
            gCompanyName = FormatField(rst("CompanyName"))
    
Dim SetUp As New clsSetup
Dim retstr As String
retstr = SetUp.ReadSetupValue("General", "ONLINE", "False")
gOnLine = IIf(UCase(retstr) = "FALSE", False, True)
DateFormat = UCase(SetUp.ReadSetupValue("General", "DateFormat", "dd/mm/yyyy"))
 
If gOnLine Then
    retstr = SetUp.ReadSetupValue("General", "CashierWindow", "False")
    gCashier = IIf(UCase(retstr) = "FALSE", False, True)
End If

Set SetUp = Nothing

With wisMain.StatusBar1
    .Panels(1).Text = gCompanyName
    .Panels(2).Text = DayBeginDate
    .Panels(2).Key = "TransDate"
    .Panels(3).Text = GetWeekDayName(GetSysFormatDate(DayBeginDate))
    .Panels(3).Key = "TransDay"
    '.Panels(4).Style = sbrScrl
    .Panels(4).Text = gCurrUser.UserName
End With

Set FinYearClass = Nothing

End Sub
Private Sub LoadAdmin()
txtUserName.Text = "admin"
txtUserPassword.Text = "admin"
If cmbFinancialYear.ListIndex < 0 Then cmbFinancialYear.ListIndex = cmbFinancialYear.ListCount - 1

Call cmdLogin_Click

End Sub

Private Function Validated() As Boolean
Validated = False

If txtUserName.Text = "" Then
    MsgBox "Please enter the User Name"
    Exit Function
End If
If txtUserPassword.Text = "" Then
    MsgBox "Please enter the password"
    Exit Function
End If

If cmbFinancialYear.ListIndex = -1 Then
   MsgBox "Please Select the Current financial Year"
   Exit Function
End If

Validated = True
End Function

Private Sub cmdCancel_Click()

RaiseEvent CancelClicked

Unload Me
End Sub

Private Sub cmdLogin_Click()

Dim UnloadDialog As Boolean
Dim YearID As Integer
Dim DBPath As String
Dim FinYearClass As clsFinChange
Dim DbUtilClass As clsDBUtilities

If Not Validated Then Exit Sub

Me.MousePointer = vbHourglass

YearID = cmbFinancialYear.ItemData(cmbFinancialYear.ListIndex)

If txtUserName.Enabled = True Then
    RaiseEvent FinYearSelected(YearID)
    RaiseEvent LoginClicked(Trim$(txtUserName.Text), Trim$(txtUserPassword.Text), "", UnloadDialog)
    If UnloadDialog Then
        GetFinancialYear
        
        DBPath = GetRegistryValue(HKEY_LOCAL_MACHINE, constREGKEYNAME, "Server")
        
        If DBPath = "" Then
            Set FinYearClass = New clsFinChange
            
            DBPath = FilePath(FinYearClass.GetDBNameWithPath(App.Path & "\" & constFINYEARFILE, YearID))
            
            Set DbUtilClass = New clsDBUtilities
            
            DbUtilClass.MakeBackUp (DBPath & "\" & constDBName)
            
            Set DbUtilClass = Nothing
            Set FinYearClass = Nothing
            
        End If
        
        Me.MousePointer = vbDefault
        Unload Me
        
    End If
Else
    
    RaiseEvent FinYearChanged(YearID)
    
    Call GetFinancialYear

    Me.MousePointer = vbDefault
    Unload Me
    
End If

Me.MousePointer = vbDefault

End Sub


Private Sub Form_Load()

Call SetKannadaCaption

Dim FinYearClass As clsFinChange

Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

Me.Caption = "Please identify yourself"

'Set the Icon for the form
Me.Icon = LoadResPicture(147, vbResIcon)

Set FinYearClass = New clsFinChange

If Not FinYearClass.GetFinYearData(App.Path & "\" & constFINYEARFILE, cmbFinancialYear) Then End

Set FinYearClass = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmLogin = Nothing
End Sub

Private Sub lblUserDate_DblClick()
LoadAdmin
End Sub

Private Sub txtUserName_GotFocus()
ActivateTextBox txtUserName
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{tab}"
    If KeyAscii = 17 Then LoadAdmin
End Sub

Private Sub txtUserPassword_GotFocus()
ActivateTextBox txtUserPassword
End Sub

