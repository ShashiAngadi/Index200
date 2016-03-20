VERSION 5.00
Begin VB.Form frmProd 
   Caption         =   "Enter the Product Key"
   ClientHeight    =   2010
   ClientLeft      =   3285
   ClientTop       =   3195
   ClientWidth     =   4335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2010
   ScaleWidth      =   4335
   Begin VB.TextBox txtDays 
      Height          =   315
      Left            =   3390
      TabIndex        =   5
      Text            =   "30"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtProdID 
      Height          =   315
      Index           =   3
      Left            =   3390
      MaxLength       =   4
      TabIndex        =   4
      Top             =   630
      Width           =   855
   End
   Begin VB.TextBox txtProdID 
      Height          =   315
      Index           =   2
      Left            =   2250
      MaxLength       =   4
      TabIndex        =   3
      Top             =   630
      Width           =   855
   End
   Begin VB.TextBox txtProdID 
      Height          =   315
      Index           =   1
      Left            =   1140
      MaxLength       =   4
      TabIndex        =   2
      Top             =   630
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   2490
      TabIndex        =   6
      Top             =   1590
      Width           =   825
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3450
      TabIndex        =   7
      Top             =   1590
      Width           =   765
   End
   Begin VB.TextBox txtProdID 
      Height          =   315
      Index           =   0
      Left            =   60
      MaxLength       =   4
      TabIndex        =   1
      Top             =   630
      Width           =   855
   End
   Begin VB.Label lblValidity 
      Caption         =   "Enter Validity period in days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   11
      Top             =   1110
      Width           =   3225
   End
   Begin VB.Label lblHyphen 
      Alignment       =   2  'Center
      Caption         =   "__"
      Height          =   345
      Index           =   2
      Left            =   3120
      TabIndex        =   10
      Top             =   630
      Width           =   255
   End
   Begin VB.Label lblHyphen 
      Alignment       =   2  'Center
      Caption         =   "__"
      Height          =   345
      Index           =   1
      Left            =   2010
      TabIndex        =   9
      Top             =   630
      Width           =   225
   End
   Begin VB.Label lblHyphen 
      Alignment       =   2  'Center
      Caption         =   "__"
      Height          =   345
      Index           =   0
      Left            =   930
      TabIndex        =   8
      Top             =   630
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the product Key of ""INDEX 2000 V3""."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   4125
   End
End
Attribute VB_Name = "frmProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()

'Const wis_INDEX2000_KEY = "Software\Waves Information Systems\Index 2000V3"
Const constREGKEYNAME = "Software\Waves Information Systems\Index 2000V3\Settings"
Dim strwisReg As String

'''Check for the demo validity
Dim ProductID As String
Dim ProductVal As String
Dim InstallDate As String
Dim ExpiryDate As String
Dim strTemp As String

Dim MaxCount As Integer
Dim Count As Integer


MaxCount = 3
For Count = 0 To MaxCount
    If Len(txtProdID(Count).Text) <> 4 Then
        MsgBox "Invalid Product ID entered", vbInformation, "Index 2000"
        txtProdID(Count).SetFocus
        Exit Sub
    End If
    ProductID = ProductID & "-" & txtProdID(Count)
Next
ProductID = Mid(ProductID, 2)

strwisReg = "Software\Waves Information Systems\" & ProductID
If Len(GetRegistryValue(HKEY_CURRENT_USER, strwisReg, "License")) > 0 Then
    If Val(cmdOk.Tag) = 0 Then
        MsgBox "The product already registered", vbInformation, "INdex 2000"
        Exit Sub
    End If
End If

'Now Get the Product Value
ProductVal = ""
MaxCount = Len(ProductID)
For Count = 1 To MaxCount
    strTemp = Mid(ProductID, Count, 1)
    If Count Mod 5 Then strTemp = Right(CStr((Val(strTemp) + 1) * 3), 1)
    ProductVal = ProductVal & strTemp
Next

'Now Register the Product value
Call CreateRegistryKey(HKEY_CURRENT_USER, strwisReg)
Call SetRegistryValue(HKEY_CURRENT_USER, strwisReg, "License", ProductVal)

MsgBox "Product registered", , "Index 2000 ver 3.0"
If Val(txtDays) = 0 Then Exit Sub

Dim Days As Long

Days = Val(txtDays)
InstallDate = Format(Now, "dd/mm/yyyy")
'ExpiryDate = FormatDate(DateAdd("d", CLng(Days), FormatDate(InstallDate)))

If Len(GetRegistryValue(HKEY_CURRENT_USER, strwisReg, "InstallDate")) = 0 Then
    Call SetRegistryValue(HKEY_CURRENT_USER, strwisReg, "InstallDate", InstallDate)
End If
If Len(GetRegistryValue(HKEY_CURRENT_USER, strwisReg, "Validity")) = 0 Then
    Call SetRegistryValue(HKEY_CURRENT_USER, strwisReg, "Validity", Days)
End If


End Sub


' Formats the given date string according to DD/MM/YYYY.
' Currently, it assumes that the given date is in MM/DD/YYYY.
Private Function FormatDate(strDate As String) As String

On Error GoTo FormatDateError
' Swap the DD and MM portions of the given date string
Const Delimiter = "/"

Dim TempDelim As String
Dim YearPart As String
Dim strArray() As String

'First Check For the Space in the given string
'Because the Date & Time part will be seperated bt a space
strDate = Trim$(strDate)
Dim SpacePos As Integer

'check for the deimeter
TempDelim = IIf(InStr(1, strDate, "/"), "/", Delimiter)

SpacePos = InStr(1, strDate, " ")
If SpacePos Then strDate = Left(strDate, SpacePos - 1)

'Breakup the date string into array elements.
'GetStringArray strDate, strArray(), Delimiter
'GetStringArray strDate, strArray(), TempDelim
strArray = Split(strDate, TempDelim)
' Check if the year part contains 2 digits.
ReDim Preserve strArray(2)
YearPart = Left$(strArray(2), 4)
If Len(Trim$(strArray(2))) = 2 Then
    ' Check, if it is greater than 30, in which case,
    ' Add "20", else, add "19".
    If Val(strArray(2)) < 30 Then
        YearPart = "20" & Right$(Trim(YearPart), 2)
    Else
        YearPart = "19" & Right$(Trim(YearPart), 2)
    End If
End If

'Change the month and day portions and concatenate.
TempDelim = IIf(InStr(1, strDate, "/"), Delimiter, "/")

FormatDate = strArray(1) & TempDelim & strArray(0) & TempDelim & YearPart
'If gIsIndianDate Then FormatDate = strArray(0) & TempDelim & strArray(1) & TempDelim & YearPart

FormatDateError:

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift And vbShiftMask Then cmdOk.Tag = 1

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift And vbShiftMask Then cmdOk.Tag = 0
End Sub


Private Sub txtProdID_Change(Index As Integer)

If Index < 3 And Len(txtProdID(Index)) = 4 Then _
    txtProdID(Index + 1).SetFocus

End Sub


Private Sub txtProdID_GotFocus(Index As Integer)
On Error Resume Next
With txtProdID(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub


