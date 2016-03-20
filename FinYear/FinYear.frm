VERSION 5.00
Begin VB.Form frmFinYear 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Financial Year Entry"
   ClientHeight    =   1605
   ClientLeft      =   3030
   ClientTop       =   3675
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   855
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
      Left            =   2130
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtEndDate 
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
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   540
      Width           =   1455
   End
   Begin VB.TextBox txtStDate 
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
      Left            =   1500
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblEndDate 
      AutoSize        =   -1  'True
      Caption         =   "Year End Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   30
      TabIndex        =   5
      Top             =   570
      Width           =   1515
   End
   Begin VB.Label lblStDate 
      AutoSize        =   -1  'True
      Caption         =   "Year Start Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   30
      TabIndex        =   4
      Top             =   150
      Width           =   1350
   End
End
Attribute VB_Name = "frmFinYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Event OkClicked(strFinFromDate As String)
Public Event CancelClicked()




Private Sub SetKannadaCaption()
Call SetFontToControls(Me) 'cmdOK.Caption = GetResourceString(1)

cmdCancel.Caption = GetResourceString(2)
lblStDate.Caption = GetResourceString(109)
lblEndDate.Caption = GetResourceString(110)
End Sub



Private Function Validated() As Boolean
Dim MonthPart As Integer
Dim DayPart As Integer
Dim YearPart As Integer

Dim StartDate As String
Dim EndDate As String
Dim arrDate() As String

Validated = False

StartDate = txtStDate.Text
EndDate = txtEndDate.Text

'Check For Validate of Dates
If Not DateValidate(StartDate, "/", True) Then
    ActivateTextBox txtStDate
    Exit Function
End If

Call GetStringArray(StartDate, arrDate, "/")

DayPart = CInt(arrDate(0))
MonthPart = CInt(arrDate(1))
YearPart = CInt(arrDate(2))

If DayPart <> 1 Then
    MsgBox "Invalid Day Specified!", vbInformation
    ActivateTextBox txtStDate
    Exit Function
End If
If MonthPart <> 4 Then
    MsgBox "Invalid month Specified!", vbInformation
    ActivateTextBox txtStDate
    Exit Function
End If

txtEndDate.Text = "31/03/" & Format(YearPart + 1, "00")


Validated = True

End Function

        
Private Function DateValidate(DateText As String, Delimiter As String, Optional IsIndian As Boolean) As Boolean
DateValidate = False
On Error Resume Next
'Check For The Decimal point in the string.
'If there is any decimal point the cint will

If InStr(1, DateText, ".", vbTextCompare) Then
    Exit Function
End If
'Breakup the given string into array elements based on the delimiter.
Dim DateArray() As String
GetStringArray DateText, DateArray(), Delimiter

'Quit if ubound is < 3   - GIRISH 11/1/2000
    If UBound(DateArray) < 2 Then
        Exit Function
    End If

' Get the date, month and year parts.
Dim DayPart As Integer
Dim MonthPart As Integer
Dim YearPart As Integer
On Error GoTo ErrLine
If IsIndian Then
    DayPart = CInt(DateArray(0))
    MonthPart = CInt(DateArray(1))
Else
    DayPart = CInt(DateArray(1))
    MonthPart = CInt(DateArray(0))
End If

YearPart = CInt(DateArray(2))
On Error GoTo 0
' The day, month and year should not be 0.
If DayPart = 0 Then
    'MsgBox "Inavlid day value.", vbInformation
    Exit Function
End If
If MonthPart = 0 Then
    'MsgBox "Inavlid day value.", vbInformation
    Exit Function
End If
'Changed condition from = to < - Girish 11/1/2000
If YearPart < 0 Then
    'MsgBox "Inavlid year value.", vbInformation
    Exit Function
End If
'The yearpart should not exceed 4 digits.
If Len(CStr(YearPart)) > 4 Then
    'MsgBox "Year is too long.", vbInformation
    Exit Function
End If

' The month part should not exceed 12.
If MonthPart > 12 Then
    'MsgBox "Invalid month.", vbInformation
    Exit Function
End If

' If the year part is only 2 digits long,
' then prefix the century digits.
If Len(CStr(YearPart)) = 2 Then
    'YearPart = Left$(CStr(Year(gStrDate)), 2) & YearPart
    '5 lines added by Girish    11/1/2000
    If Val(YearPart) <= 30 Then
        YearPart = "20" & YearPart
    Else
        YearPart = "19" & YearPart
    End If
End If

' Check if it is a leap year.
Dim bLeapYear As Boolean


' Validations.
Select Case MonthPart
    Case 2  ' Check for February month.
        If bLeapYear Then
            If DayPart > 29 Then
                Exit Function
            End If
        Else
            If DayPart > 28 Then
                
                Exit Function
            End If
        End If
    
    Case 4, 6, 9, 11 ' Months having 30 days...
        If DayPart > 30 Then
            Exit Function
        End If
    Case Else
        If DayPart > 31 Then
            Exit Function
        End If
End Select

DateValidate = True

ErrLine:
    

End Function



Private Sub cmdCancel_Click()

RaiseEvent CancelClicked

End Sub


Private Sub cmdOk_Click()

RaiseEvent OkClicked(txtStDate.Text)
Unload Me
End Sub


Private Sub Form_Load()
'Center the form
CenterMe Me

'Set the icon
Me.Icon = LoadResPicture(147, vbResIcon)


cmdOk.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmFinYear = Nothing

End Sub


Private Sub txtEndDate_GotFocus()
txtEndDate.SelStart = 0
txtEndDate.SelLength = Len(txtEndDate)
End Sub


Private Sub txtStDate_GotFocus()
txtStDate.SelStart = 0
txtStDate.SelLength = Len(txtStDate)

End Sub


Private Sub txtStDate_LostFocus()
cmdOk.Enabled = False
If txtStDate.Text = "" Then Exit Sub
If Validated Then cmdOk.Enabled = True
End Sub


