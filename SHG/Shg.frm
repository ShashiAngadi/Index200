VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmShg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SHG Information"
   ClientHeight    =   7305
   ClientLeft      =   1680
   ClientTop       =   1260
   ClientWidth     =   7785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   7785
   Begin VB.CommandButton cmdClose 
      Caption         =   "Clos&e"
      Height          =   400
      Left            =   6480
      TabIndex        =   1
      Top             =   6870
      Width           =   1215
   End
   Begin VB.Frame fra 
      Height          =   6075
      Index           =   1
      Left            =   270
      TabIndex        =   39
      Top             =   570
      Width           =   7335
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0C0C0&
         Height          =   990
         Left            =   150
         ScaleHeight     =   930
         ScaleWidth      =   5685
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   255
         Width           =   5745
         Begin VB.Label lblHeading 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   54
            Top             =   30
            Width           =   135
         End
         Begin VB.Label lblDesc 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   555
            Left            =   840
            TabIndex        =   53
            Top             =   360
            Width           =   4710
         End
         Begin VB.Image imgNewAcc 
            Height          =   375
            Left            =   135
            Stretch         =   -1  'True
            Top             =   120
            Width           =   345
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   6000
         TabIndex        =   50
         Top             =   4875
         Width           =   1215
      End
      Begin VB.PictureBox picViewport 
         BackColor       =   &H00FFFFFF&
         Height          =   4380
         Left            =   150
         ScaleHeight     =   4320
         ScaleWidth      =   5685
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1305
         Width           =   5745
         Begin VB.PictureBox picSlider 
            Height          =   3645
            Left            =   -45
            ScaleHeight     =   3585
            ScaleWidth      =   5400
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   30
            Width           =   5460
            Begin VB.TextBox txtData 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   2610
               TabIndex        =   48
               Top             =   30
               Width           =   2760
            End
            Begin VB.CommandButton cmd 
               Caption         =   "..."
               Height          =   315
               Index           =   0
               Left            =   4860
               TabIndex        =   47
               Top             =   870
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.ComboBox cmb 
               Height          =   315
               Index           =   0
               Left            =   2340
               Style           =   2  'Dropdown List
               TabIndex        =   46
               Top             =   720
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.CheckBox chk 
               Alignment       =   1  'Right Justify
               Caption         =   "chk"
               Height          =   300
               Index           =   0
               Left            =   2820
               TabIndex        =   45
               Top             =   2070
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label txtPrompt 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Account Holder"
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   0
               Left            =   60
               TabIndex        =   49
               Top             =   30
               Width           =   2535
            End
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   4305
            Left            =   5460
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   30
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6000
         TabIndex        =   41
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   6000
         TabIndex        =   40
         Top             =   5310
         Width           =   1215
      End
      Begin VB.Label lblOperation 
         AutoSize        =   -1  'True
         Caption         =   "Operation Mode :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   52
         Top             =   5700
         Width           =   1545
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Reports"
      Height          =   6045
      Index           =   3
      Left            =   270
      TabIndex        =   3
      Top             =   600
      Width           =   7335
      Begin VB.OptionButton optReports 
         Caption         =   "Show Sb Monthly balance"
         Height          =   285
         Index           =   6
         Left            =   660
         TabIndex        =   38
         Top             =   1920
         Width           =   3195
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Show Loan monthly balance"
         Height          =   285
         Index           =   7
         Left            =   3870
         TabIndex        =   37
         Top             =   1920
         Width           =   3315
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Shg Monthly Stament"
         Height          =   225
         Index           =   8
         Left            =   660
         TabIndex        =   36
         Top             =   2520
         Width           =   3195
      End
      Begin VB.OptionButton optReports 
         Caption         =   "SHG Training"
         Height          =   315
         Index           =   3
         Left            =   3870
         TabIndex        =   35
         Top             =   870
         Width           =   3315
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Groups Created"
         Height          =   285
         Index           =   2
         Left            =   660
         TabIndex        =   12
         Top             =   860
         Width           =   3195
      End
      Begin VB.Frame fraOrder 
         Caption         =   " List Order"
         Height          =   1905
         Left            =   210
         TabIndex        =   9
         Top             =   3450
         Width           =   6855
         Begin VB.CommandButton cmdAdvance 
            Caption         =   "&Advanced"
            Height          =   400
            Left            =   5550
            TabIndex        =   19
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtDate1 
            Height          =   345
            Left            =   1620
            TabIndex        =   16
            Top             =   870
            Width           =   1245
         End
         Begin VB.TextBox txtDate2 
            Height          =   345
            Left            =   5130
            TabIndex        =   15
            Top             =   900
            Width           =   1245
         End
         Begin VB.CommandButton cmdDate1 
            Caption         =   "..."
            Height          =   315
            Left            =   2910
            TabIndex        =   14
            Top             =   900
            Width           =   315
         End
         Begin VB.CommandButton cmdDate2 
            Caption         =   "..."
            Height          =   315
            Left            =   6450
            TabIndex        =   13
            Top             =   900
            Width           =   315
         End
         Begin VB.OptionButton optAccId 
            Caption         =   "By Account No"
            Height          =   315
            Left            =   510
            TabIndex        =   11
            Top             =   270
            Width           =   1815
         End
         Begin VB.OptionButton optName 
            Caption         =   "By Name"
            Height          =   315
            Left            =   3690
            TabIndex        =   10
            Top             =   270
            Value           =   -1  'True
            Width           =   1905
         End
         Begin VB.Line Line1 
            X1              =   6930
            X2              =   0
            Y1              =   690
            Y2              =   690
         End
         Begin VB.Label lblDate1 
            Caption         =   "after (dd/mm/yyyy)"
            Height          =   225
            Left            =   90
            TabIndex        =   18
            Top             =   930
            Width           =   1545
         End
         Begin VB.Label lblDate2 
            Caption         =   "and before (dd/mm/yyyy)"
            Height          =   315
            Left            =   3300
            TabIndex        =   17
            Top             =   930
            Width           =   1815
         End
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Show Sc/St Members only"
         Height          =   315
         Index           =   1
         Left            =   3870
         TabIndex        =   8
         Top             =   330
         Width           =   3405
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Show Loan balance"
         Height          =   285
         Index           =   5
         Left            =   3870
         TabIndex        =   7
         Top             =   1410
         Width           =   3315
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Show Sb balance"
         Height          =   285
         Index           =   4
         Left            =   660
         TabIndex        =   6
         Top             =   1390
         Width           =   3195
      End
      Begin VB.OptionButton optReports 
         Caption         =   "Show Group List"
         Height          =   285
         Index           =   0
         Left            =   660
         TabIndex        =   5
         Top             =   330
         Width           =   3195
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Show"
         Default         =   -1  'True
         Height          =   400
         Left            =   5880
         TabIndex        =   4
         Top             =   5550
         Width           =   1215
      End
   End
   Begin VB.Frame fra 
      Height          =   6045
      Index           =   2
      Left            =   270
      TabIndex        =   2
      Top             =   600
      Width           =   7335
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   2925
         Left            =   240
         TabIndex        =   30
         Top             =   2970
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   5159
         _Version        =   393216
      End
      Begin VB.CommandButton cmdTrainingDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4410
         TabIndex        =   34
         Top             =   2490
         Width           =   1215
      End
      Begin VB.TextBox txtTrainMembers 
         Height          =   345
         Left            =   6240
         TabIndex        =   29
         Top             =   1800
         Width           =   675
      End
      Begin VB.CommandButton cmdTrainToDate 
         Caption         =   "..."
         Height          =   315
         Left            =   6930
         TabIndex        =   33
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton cmdTrainFromDate 
         Caption         =   "..."
         Height          =   315
         Left            =   3120
         TabIndex        =   32
         Top             =   330
         Width           =   315
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   5730
         TabIndex        =   31
         Top             =   2490
         Width           =   1215
      End
      Begin VB.TextBox txtTrainingPlace 
         Height          =   345
         Left            =   1860
         TabIndex        =   27
         Top             =   1800
         Width           =   1725
      End
      Begin VB.TextBox txtTraining 
         Height          =   735
         Left            =   1860
         TabIndex        =   25
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox txtTrainTo 
         Height          =   345
         Left            =   5670
         TabIndex        =   23
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox txtTrainFrom 
         Height          =   345
         Left            =   1860
         TabIndex        =   20
         Top             =   300
         Width           =   1215
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   60
         X2              =   7230
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Label lbltrainMembers 
         Caption         =   "No of Members Attended"
         Height          =   285
         Left            =   3930
         TabIndex        =   28
         Top             =   1830
         Width           =   1935
      End
      Begin VB.Label lblTrainingPlace 
         Caption         =   "Training Place :"
         Height          =   285
         Left            =   270
         TabIndex        =   26
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label lblTraining 
         Caption         =   "Training Detail"
         Height          =   315
         Left            =   270
         TabIndex        =   24
         Top             =   870
         Width           =   1455
      End
      Begin VB.Label lblTrainTo 
         Caption         =   "To Date:"
         Height          =   255
         Left            =   4170
         TabIndex        =   22
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label lblTrainFrom 
         Caption         =   "From date :"
         Height          =   255
         Left            =   330
         TabIndex        =   21
         Top             =   360
         Width           =   1515
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   6705
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11827
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "SHG Detail"
            Key             =   "shg"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Training"
            Key             =   "Training"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reports"
            Key             =   "Reports"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmShg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event CustClick()
Public Event SaveClick()
Public Event CloseClick()
Public Event ShowClick()
Public Event LoadClick()
Public Event ClearClick()
Public Event AddClick()
Public Event DeleteClick()
Public Event DeleteTraining()
Public Event GridClick()
Public Event RepOptionClick()

'Private m_clsRepOption As clsRepOption
Private m_AccID As Long
Private m_CustomerID As Long
Private m_dbOperation As wis_DBOperation

Private Sub ArrangePropSheet()
Const CTL_MARGIN = 15
Const BORDER_HEIGHT = 15
Dim NumItems As Integer
Dim NeedsScrollbar As Boolean

' Arrange the Slider panel.
With picSlider
    .BorderStyle = 0
    .Top = 0
    .Left = 0
    NumItems = VisibleCount()
    .Height = txtData(0).Height * NumItems + 1 _
            + BORDER_HEIGHT * (NumItems + 1)
    ' If the height is greater than viewport height,
    ' the scrollbar needs to be displayed.  So,
    ' reduce the width accordingly.
    If .Height > picViewport.ScaleHeight Then
        NeedsScrollbar = True
        .Width = picViewport.ScaleWidth - _
                VScroll1.Width
    Else
        .Width = picViewport.ScaleWidth
    End If

End With

' Set/Reset the properties of scrollbar.
With VScroll1
    .Height = picViewport.ScaleHeight
    .Min = 0
    .Max = picSlider.Height - picViewport.ScaleHeight
    If .Max < 0 Then .Max = 0
    .SmallChange = txtData(0).Height
    .LargeChange = picViewport.ScaleHeight / 2
End With

' Adjust the text controls on this panel.
Dim I As Integer
For I = 0 To txtData.count - 1
    txtData(I).Width = picSlider.ScaleWidth _
            - txtPrompt(I).Width - CTL_MARGIN
Next


If NeedsScrollbar Then
    VScroll1.Visible = True
End If

' Need to adjust the width of text boxes, due to
' change in width of the slider box.
Dim CtlIndex As Integer
For I = 0 To txtData.count - 1
    txtData(I).Width = picSlider.ScaleWidth - _
        (txtPrompt(I).Left + txtPrompt(I).Width) - CTL_MARGIN
Next

' Align all combo and command controls on this prop sheet.
For I = 0 To cmb.count - 1
    cmb(I).Width = txtData(I).Width
Next
For I = 0 To cmd.count - 1
    cmd(I).Left = txtData(I).Left + txtData(I).Width - cmd(I).Width
Next

End Sub

Public Property Let CustomerID(NewValue As Long)
    m_CustomerID = NewValue
End Property

Public Property Let SHGID(NewValue As Integer)
    m_AccID = NewValue
    m_dbOperation = Update
    If m_AccID = 0 Then m_dbOperation = Insert
End Property

' Returns the number of items that are visible for a control array.
' Looks in the control's tag for visible property, rather than
' depend upon the control's visible property for some obvious reasons.
Private Function VisibleCount() As Integer
On Error GoTo Err_line
Dim I As Integer
Dim strVisible As String
For I = 0 To txtPrompt.count - 1
    strVisible = ExtractToken(txtPrompt(I).Tag, "Visible")
    If StrComp(strVisible, "True") = 0 Then
        VisibleCount = VisibleCount + 1
    End If
Next
Err_line:
End Function



Private Function LoadPropSheet() As Boolean

'TabStrip.ZOrder 1
'TabStrip.Tabs(1).Selected = True
lblDesc.BorderStyle = 0
lblHeading.BorderStyle = 0
lblOperation.Caption = GetResourceString(54) '"Operation Mode : <INSERT>"
'
' Read the data from SBAcc.PRP and load the relevant data.
'
Const CTL_MARGIN = 15
' Check for the existence of the file.
Dim PropFile As String
PropFile = App.Path & "\SHG_" & gLangOffSet & ".PRP"

If Dir(PropFile, vbNormal) = "" Then
    If gLangOffSet <> wis_NoLangOffset Then
        PropFile = App.Path & "\SHGKan.PRP"
    Else
        PropFile = App.Path & "\SHG.PRP"
    End If
End If
If Dir(PropFile, vbNormal) = "" Then
    'MsgBox "Unable to locate the properties file '" _
            & PropFile & "' !", vbExclamation
    MsgBox GetResourceString(602) _
            & PropFile & "' !", vbExclamation
    Exit Function
End If

'Load the CLIP Icon
imgNewAcc.Picture = LoadResPicture(105, vbResIcon)

' Declare required variables...
Dim strTmp As String
Dim strPropType As String
Dim FirstImgCtl As Boolean
Dim FirstControl As Boolean
Dim I As Integer, CtlIndex As Integer
Dim strRet As String, imgCtlIndex As Integer
FirstControl = True
FirstImgCtl = True
Dim strTag As String

' Read all the prompts and load accordingly...
Do
    ' Read a line.
    strTag = ReadFromIniFile("Property Sheet", _
                "Prop" & I + 1, PropFile)
    If strTag = "" Then Exit Do

    ' Load a prompt and a data text.
    If FirstControl Then
        FirstControl = False
    Else
        Load txtPrompt(txtPrompt.count)
        Load txtData(txtData.count)
    End If
    CtlIndex = txtPrompt.count - 1

    ' Get the property type.
    strPropType = ExtractToken(strTag, "PropType")
    Select Case UCase$(strPropType)
        Case "HEADING", ""
            ' Set the fontbold for Txtprompt.
            With txtPrompt(CtlIndex)
                .FontBold = True
                .Caption = ""
            End With
            txtData(CtlIndex).Enabled = False

        Case "EDITABLE"
            ' Add 4 spaces for indentation purposes.
            With txtPrompt(CtlIndex)
                .Caption = IIf(gLangOffSet, Space(2), Space(4))
                .FontBold = False
                .Enabled = True
            End With
            txtData(CtlIndex).Enabled = True
        Case Else
            'MsgBox "Unknown Property type encountered " _
                    & "in Property file!", vbCritical
            MsgBox GetResourceString(603), vbCritical
            Exit Function
    End Select

    ' Set the PROPERTIES for controls.
    With txtPrompt(CtlIndex)
        strRet = PutToken(strTag, "Visible", "True")
        .Tag = strRet
        .Caption = .Caption & ExtractToken(.Tag, "Prompt")
        If CtlIndex = 0 Then
            .Top = 0
        Else
            .Top = txtPrompt(CtlIndex - 1).Top _
                + txtPrompt(CtlIndex - 1).Height + CTL_MARGIN
        End If
        .Left = 0
        .Visible = True
    End With
    With txtData(CtlIndex)
        .Top = txtPrompt(CtlIndex).Top
        .Left = txtPrompt(CtlIndex).Left + _
            txtPrompt(CtlIndex).Width + CTL_MARGIN
        .Visible = True
        ' Check the LockEdit property.
        strRet = ExtractToken(strTag, "LockEdit")
        If StrComp(strRet, "True", vbTextCompare) = 0 Then
            .Locked = True
        End If
    End With

    ' Get the display type. If its a List or Browse,
    ' then load a combo or a cmd button.
    Dim CmdLoaded As Boolean
    Dim ListLoaded As Boolean
    Dim ChkLoaded As Boolean
    strPropType = ExtractToken(strTag, "DisplayType")
    Select Case UCase$(strPropType)
        Case "LIST"
            'Load a combo.
            If Not ListLoaded Then
                ListLoaded = True
            Else
                Load cmb(cmb.count)
            End If
            ' Set the alignment.
            With cmb(cmb.count - 1)
                '.Index = i
                .Left = txtData(I).Left
                .Top = txtData(I).Top
                .Width = txtData(I).Width
                ' Set it's tab order.
                .TabIndex = txtData(I).TabIndex + 1
                ' Update the tag with the text index.
                .Tag = PutToken(.Tag, "TextIndex", CStr(I))
                ' Write back this button index to text tag.
                txtPrompt(I).Tag = PutToken(txtPrompt(I).Tag, _
                        "TextIndex", CStr(cmb.count - 1))
                'txtData(i).Visible = False
                ' If the list data is given, load it.
                Dim List() As String, j As Integer
                Dim strListData As String
                strListData = ExtractToken(strTag, "ListData")
                If strListData <> "" Then
                    ' Break up the data into array elements.
                    GetStringArray strListData, List(), ","
                    cmb(cmb.count - 1).Clear
                    For j = 0 To UBound(List)
                        cmb(cmb.count - 1).AddItem List(j)
                    Next
                End If
            End With

        Case "BROWSE"
            'Load a command button.
            If Not CmdLoaded Then
                CmdLoaded = True
            Else
                Load cmd(cmd.count)
            End If
            With cmd(cmd.count - 1)
                '.Index = i
                .Width = txtData(I).Height
                .Height = .Width
                .Left = txtData(I).Left + txtData(I).Width - .Width
                .Top = txtData(I).Top
                .TabIndex = txtData(I).TabIndex + 1
                .ZOrder 0
                '.Visible = True
                ' Update the tag with the text index.
                .Tag = PutToken(.Tag, "TextIndex", CStr(I))
                ' Write back this button index to text tag.
                txtPrompt(I).Tag = PutToken(txtPrompt(I).Tag, _
                        "TextIndex", CStr(cmd.count - 1))
                .Caption = "..."
                .Width = 350
            End With
        Case "BOOLEAN"
              ' Load a check box.
            If Not ChkLoaded Then
                ChkLoaded = True
            Else
                Load chk(chk.count)
            End If
            With chk(chk.count - 1)
                txtData(I).Text = "False"
                .Visible = True
                .Left = txtData(I).Left
                .Top = txtData(I).Top + CTL_MARGIN
                .Width = txtData(I).Width
                .Height = txtData(I).Height - 2 * CTL_MARGIN
                .Caption = String(txtData(I).Width / Me.TextWidth(" "), " ")
                .TabIndex = txtData(I).TabIndex + 1
                .ZOrder 0
                ' Update the tag with the text index.
                .Tag = PutToken(.Tag, "TextIndex", CStr(I))
                ' Write back this button index to text tag.
                txtPrompt(I).Tag = PutToken(txtPrompt(I).Tag, _
                        "TextIndex", CStr(chk.count - 1))
            End With
    End Select

    ' Increment the loop count.
    I = I + 1
Loop


'Now Set the captions of the Command Buttons
    'Set the Load
    I = Val(ExtractToken(txtPrompt(GetIndex("AccID")).Tag, "TextIndex"))
    With cmd(I)
        .Caption = GetResourceString(3) ' "Load"
        .Width = 1000
    End With
    'Set the Details
    I = Val(ExtractToken(txtPrompt(GetIndex("AccName")).Tag, "TextIndex"))
    With cmd(I)
        .Caption = GetResourceString(295) '"Details..."
        .Width = 1000
    End With
    
ArrangePropSheet

' Get a new account number and display it to accno textbox.
Dim txtIndex As Integer
txtIndex = GetIndex("AccID")
'txtData(txtIndex).Text = GetNewAccountNumber

' Show the current date wherever necessary.
txtIndex = GetIndex("CreateDate")
txtData(txtIndex).Text = gStrDate

' Set the default updation mode.
m_dbOperation = Insert

End Function

' Returns the index of the control bound to "strDatasrc".
Private Function GetIndex(strDataSrc As String) As Integer
GetIndex = -1
Dim strTmp As String
Dim I As Integer
For I = 0 To txtPrompt.count - 1
    ' Get the data source for this control.
    strTmp = ExtractToken(txtPrompt(I).Tag, "DataSource")
    If StrComp(strDataSrc, strTmp, vbTextCompare) = 0 Then
        GetIndex = I
        Exit For
    End If
Next
End Function
' Returns the text value from a control array
' bound the field "FieldName".
Private Function GetVal(FieldName As String) As String
Dim I As Integer
Dim strTxt As String
For I = 0 To txtData.count - 1
    strTxt = ExtractToken(txtPrompt(I).Tag, "DataSource")
    If StrComp(strTxt, FieldName, vbTextCompare) = 0 Then
        GetVal = txtData(I).Text
        Exit For
    End If
Next
End Function

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

TabStrip1.Tabs(1).Caption = GetResourceString(371, 295)
TabStrip1.Tabs(2).Caption = GetResourceString(383) '"Training "
TabStrip1.Tabs(3).Caption = GetResourceString(283)
'lblShgNum = GetResourceString(371,60)
'lblCustName = GetResourceString(35)
'lblDob.Caption = GetResourceString(37)

'fraMemNum.Caption = GetResourceString(49,60)
'lblTotmem = GetResourceString(52,49)
'lblScMem = "SC/ST " & GetResourceString(49)
'lblFemMem = GetResourceString(386,49)
'lblFemSc = "SC/ST " & GetResourceString(386)

'lblContact.Caption = GetResourceString(236)
'lblDay.Caption = GetResourceString(382,44)
'lblMeetPlace.Caption = GetResourceString(382,112)

'lblGender.Caption = GetResourceString(371,125)
'lblPlace.Caption = GetResourceString(371,112)
'lblCaste.Caption = GetResourceString(371,111)

'lblSb.Caption = GetResourceString(421,60)
'lblLoan.Caption = GetResourceString(80,60)
'lblRemark.Caption = GetResourceString(261)

'cmdLoad.Caption = GetResourceString(3)
cmdSave.Caption = GetResourceString(15)
cmdDelete.Caption = GetResourceString(14)


'Training
lblTrainFrom.Caption = GetResourceString(109)
lblTrainTo.Caption = GetResourceString(110)
lblTraining.Caption = GetResourceString(383, 295)
lblTrainingPlace.Caption = GetResourceString(383, 112)
cmdAdd.Caption = GetResourceString(10)
cmdTrainingDelete.Caption = GetResourceString(14)

'Reports
fraOrder.Caption = GetResourceString(287)
optAccId.Caption = GetResourceString(68)
optName.Caption = GetResourceString(69)
optReports(0).Caption = GetResourceString(370)
optReports(1).Caption = GetResourceString(384, 49)
optReports(2).Caption = GetResourceString(371, 64)
optReports(3).Caption = GetResourceString(370, 384)
optReports(4).Caption = GetResourceString(421, 61)
optReports(5).Caption = GetResourceString(80, 61)
optReports(6).Caption = GetResourceString(421, 463, 42)
optReports(7).Caption = GetResourceString(80, 463, 42)
optReports(8).Caption = GetResourceString(370, 463, 430)



lblDate1.Caption = GetResourceString(109)
lblDate2.Caption = GetResourceString(110)
cmdShow.Caption = GetResourceString(13)



cmdClose.Caption = GetResourceString(11)
cmdReset.Caption = GetResourceString(8)
cmdAdvance.Caption = GetResourceString(491)    'Options
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub chk_LostFocus(Index As Integer)
'
' Update the current text to the data text
'

Dim txtIndex As String
txtIndex = ExtractToken(chk(Index).Tag, "TextIndex")
If txtIndex <> "" Then
    txtData(Val(txtIndex)).Text = IIf(chk(Index).Value = vbChecked, True, False)
End If
End Sub

Private Sub cmb_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub cmb_LostFocus(Index As Integer)
'
' Update the current text to the data text
'

Dim txtIndex As String
txtIndex = ExtractToken(cmb(Index).Tag, "TextIndex")
If txtIndex <> "" Then
    txtData(Val(txtIndex)).Text = cmb(Index).Text
End If

End Sub

Private Sub cmd_Click(Index As Integer)
Screen.MousePointer = vbHourglass
Dim txtIndex As String
Dim count As Integer
Dim rst As ADODB.Recordset

' Check to which text index it is mapped.
txtIndex = ExtractToken(cmd(Index).Tag, "TextIndex")

' Extract the Bound field name.
Dim strField As String
strField = ExtractToken(txtPrompt(Val(txtIndex)).Tag, "DataSource")
Screen.MousePointer = vbDefault
Select Case UCase$(strField)
    Case "ACCID"
        RaiseEvent LoadClick
        'If m_accUpdatemode = wis_INSERT Then txtData(txtIndex).Text = GetNewAccountNumber

    Case "ACCNAME"
        RaiseEvent CustClick
'        If m_CustReg Is Nothing Then
'            Set m_CustReg = New clsCustReg
'            m_CustReg.NewCustomer
'            m_CustReg.ModuleID = wis_SBAcc
'        End If
'        m_CustReg.ShowDialog
'        txtData(txtIndex).Text = m_CustReg.FullName
    
    Case "CREATEDATE"
        With Calendar
            .Left = txtData(txtIndex).Left + Me.Left _
                    + Me.picViewport.Left + fra(1).Left + 50
            .Top = Me.Top + picViewport.Top + txtData(txtIndex).Top _
                + fra(1).Top + 300
            .Width = txtData(txtIndex).Width
            If .Top + .Height > Screen.Height Then .Top = .Top - .Height - txtData(txtIndex).Height
            .Height = .Width
            .selDate = txtData(txtIndex).Text
            .Show vbModal, Me
            If .selDate <> "" Then txtData(txtIndex).Text = .selDate
        End With
    
    
        
End Select
Screen.MousePointer = vbDefault

End Sub


Private Sub cmdAdd_Click()

RaiseEvent AddClick
End Sub


Private Sub cmdAdvance_Click()
    RaiseEvent RepOptionClick
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()
RaiseEvent SaveClick

End Sub


Private Sub cmdCust_Click()
    RaiseEvent CustClick
End Sub

Private Sub cmdDate1_Click()
With Calendar
    .Left = cmdDate1.Left + fra(3).Left + Left
    .Top = Top + fra(3).Top + cmdDate1.Top + fraOrder.Top - .Height / 2
    
    .Show 1
    txtDate1 = .selDate
End With

End Sub

Private Sub cmdDate2_Click()
With Calendar
    .Left = cmdDate2.Left + fra(3).Left + Left
    .Top = Top + fra(3).Top + cmdDate2.Top + fraOrder.Top - .Height / 2
    
    .Show 1
    txtDate2 = .selDate
End With

End Sub


Private Sub cmdDelete_Click()
    RaiseEvent DeleteClick
End Sub

Private Sub cmdLoad_Click()
    RaiseEvent LoadClick
End Sub

Private Sub cmdReset_Click()
    RaiseEvent ClearClick
End Sub


Private Sub cmdSave_Click()
    If Not Validate Then Exit Sub
    RaiseEvent SaveClick
End Sub

Private Function Validate() As Boolean
Validate = False

On Error GoTo Exit_Line

Dim strMsg As String

If m_CustomerID = 0 Then
   'No Customer Detials specified
    strMsg = GetResourceString(662)
    GoTo Exit_Line
End If

Dim rst As Recordset
Dim txtIndex As Byte

txtIndex = GetIndex("AccID")
'With m_frmShg
    'Check for the Shg Number
    txtIndex = GetIndex("AccID")
    If Len(Trim$(txtData(txtIndex))) = 0 Then
        strMsg = GetResourceString(500)
        ActivateTextBox txtData(txtIndex)
        GoTo Exit_Line
    End If
    strMsg = ""
    gDbTrans.SqlStmt = "Select * From ShgMaster " & _
                " Where AccNum = " & AddQuotes(Trim$(txtData(txtIndex)))
    If m_dbOperation = Update Then _
        gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And AccID <> " & m_AccID
        
    If gDbTrans.Fetch(rst, adOpenDynamic) > 0 Then
        strMsg = GetResourceString(545)
        ActivateTextBox txtData(txtIndex)
        'GoTo Exit_line
    End If

    'Check for the Date Of Inception
    txtIndex = GetIndex("CreateDate")
    With txtData(txtIndex)
        If Len(.Text) And (Not DateValidate(.Text, "/", True)) Then
            ActivateTextBox txtData(txtIndex)
            strMsg = strMsg & vbCrLf & GetResourceString(499)
            'GoTo Exit_line
        End If
    End With
    'Checkhe No of members
    txtIndex = GetIndex("TotalMem")
    If Not CurrencyValidate(txtData(txtIndex), True) Then
        strMsg = strMsg & vbCrLf & GetResourceString(760)
        ActivateTextBox txtData(txtIndex)
        'GoTo Exit_line
    End If
    
    'Checkthe No of female members
    txtIndex = GetIndex("FemaleMem")
    If Not CurrencyValidate(txtData(txtIndex), True) Then
        If Len(strMsg) Then strMsg = GetResourceString(760)
        ActivateTextBox txtData(txtIndex)
        'GoTo Exit_line
    End If
    
    'Checkthe No of Sc/St members
    txtIndex = GetIndex("SCSTMem")
    If Not CurrencyValidate(txtData(txtIndex), True) Then
        If Len(strMsg) Then strMsg = GetResourceString(760)
        ActivateTextBox txtData(txtIndex)
        'GoTo Exit_line
    End If
    
    'Check the Femeal Sc St members
    txtIndex = GetIndex("FemaleSCSTMem")
    If Not CurrencyValidate(txtData(txtIndex), True) Then
        If Len(strMsg) Then strMsg = GetResourceString(760)
        ActivateTextBox txtData(txtIndex)
        GoTo Exit_Line
    End If
    
    txtIndex = GetIndex("Gender")
    If Len(Trim$(txtData(txtIndex))) = 0 Then
        strMsg = strMsg & vbCrLf & GetResourceString(823)
        ActivateTextBox txtData(txtIndex)
        GoTo Exit_Line
    End If
    
'    If .cmbGender.ListIndex < 0 Then .cmbGender.ListIndex = 0
'    If .cmbGender.ListIndex < 0 Then
'        ActivateTextBox .cmbGender
'        GoTo Exit_line
'    End If
    
    txtIndex = GetIndex("Place")
    If Len(Trim$(txtData(txtIndex))) = 0 Then
        strMsg = strMsg & vbCrLf & GetResourceString(824)
        ActivateTextBox txtData(txtIndex)
        'GoTo Exit_line
    End If
    'If .cmbPlace.ListIndex < 0 Then .cmbPlace.ListIndex = 0
    
    txtIndex = GetIndex("Caste")
    If Len(Trim$(txtData(txtIndex))) = 0 Then
        'strMsg = "Specify the Caste"
        strMsg = strMsg & vbCrLf & GetResourceString(825)
        ActivateTextBox txtData(txtIndex)
        'GoTo Exit_line
    End If
    'If .cmbCaste.ListIndex < 0 Then .cmbCaste.ListIndex = 0
    
    txtIndex = GetIndex("MeetDay")
    'If .cmbDay.ListIndex < 0 Then
    If Len(txtData(txtIndex)) = 0 Then
        strMsg = strMsg & vbCrLf & "Specify the meeting day"
        ActivateTextBox txtData(txtIndex) '.cmbDay
        'GoTo Exit_line
    End If
    
    'txtIndex = GetIndex("MeetPlace")
    txtIndex = GetIndex("Contact")
    If Len(Trim$(txtData(txtIndex))) = 0 Then
        strMsg = strMsg & vbCrLf & "Contact person name not entered"
        ActivateTextBox txtData(txtIndex)
        'GoTo Exit_line
    End If
    
    If Len(strMsg) Then GoTo Exit_Line
    txtIndex = GetIndex("MeetPlace")
    If Len(Trim$(txtData(txtIndex))) = 0 Then
        strMsg = "Place of the weekly Meeting Not specified"
        ActivateTextBox txtData(txtIndex)
        GoTo Exit_Line
    End If

'End With

Validate = True
Exit Function

Exit_Line:

MsgBox strMsg, vbInformation, wis_MESSAGE_TITLE

End Function
Private Sub cmdShow_Click()
If txtDate1.Enabled Then
    If Not DateValidate(txtDate1.Text, "/", True) Then
        MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtDate1
        Exit Sub
    End If
End If

If txtDate2.Enabled Then
    If Not DateValidate(txtDate2.Text, "/", True) Then
        MsgBox GetResourceString(501), vbInformation, wis_MESSAGE_TITLE
        ActivateTextBox txtDate2
        Exit Sub
    End If
End If

RaiseEvent ShowClick

End Sub

Private Sub cmdTrainFromDate_Click()
With Calendar
    .Left = cmdTrainFromDate.Left + fra(2).Left + Left
    .Top = Top + fra(1).Top + cmdTrainFromDate.Top - .Height / 2
    
    
    .Show 1
    txtTrainFrom = .selDate
End With
End Sub

Private Sub cmdTrainingDelete_Click()
RaiseEvent DeleteTraining
End Sub

Private Sub cmdTrainToDate_Click()
With Calendar
    .Left = cmdTrainToDate.Left + fra(1).Left + Left
    .Top = Top + fra(1).Top + cmdTrainToDate.Top - .Height / 2
    
    .Show 1
    txtTrainTo = .selDate
End With

End Sub


Private Sub SetDescription(Ctl As Control)

' Extract the description title.
lblHeading.Caption = ExtractToken(Ctl.Tag, "DescTitle")
lblDesc.Caption = ExtractToken(Ctl.Tag, "Description")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' If the current tab is not Add/Modify, then exit.
'If TabStrip.SelectedItem.Key <> "AddModify" Then Exit Sub

Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0

If Not CtrlDown Then Exit Sub
If KeyCode <> vbKeyTab Then Exit Sub
Dim I As Byte

With TabStrip1
    I = .SelectedItem.Index
    If Shift = 2 Then
        I = I + 1
        If I > .Tabs.count Then I = 1
    Else
        I = I - 1
        If I = 0 Then I = .Tabs.count
    End If
    .Tabs(I).Selected = True
End With

End Sub

Private Sub Form_Load()

Call CenterMe(Me)
Call SetKannadaCaption

Call LoadPropSheet
Dim cmbIndex As Byte

cmbIndex = Val(ExtractToken(txtPrompt(GetIndex("Gender")).Tag, "TextIndex"))
Call LoadGender(cmb(cmbIndex))
cmbIndex = Val(ExtractToken(txtPrompt(GetIndex("Caste")).Tag, "TextIndex"))
Call LoadCastes(cmb(cmbIndex))
cmbIndex = Val(ExtractToken(txtPrompt(GetIndex("Place")).Tag, "TextIndex"))
Call LoadPlaces(cmb(cmbIndex))


'Now Load All the Week days to select the
'Meeting day
cmbIndex = Val(ExtractToken(txtPrompt(GetIndex("MeetDay")).Tag, "TextIndex"))
With cmb(cmbIndex)
    .Clear
    .AddItem GetWeekDayString(1)
    .AddItem GetWeekDayString(2)
    .AddItem GetWeekDayString(3)
    .AddItem GetWeekDayString(4)
    .AddItem GetWeekDayString(5)
    .AddItem GetWeekDayString(6)
    .AddItem GetWeekDayString(7)
End With

TabStrip1.Tabs(1).Selected = True
txtDate2 = DayBeginDate
optReports(0).Value = True

End Sub

Private Sub Form_Unload(cancel As Integer)
RaiseEvent CloseClick
End Sub




Private Sub grd_DblClick()
RaiseEvent GridClick
End Sub


Private Sub optReports_Click(Index As Integer)
Dim Dt As Boolean
Dim Caste As Boolean

Select Case Index
    Case 0
        Caste = True
    Case 1
        Caste = True
    Case 2
        Dt = True: Caste = True
    Case 3
        Dt = True
    Case 4, 5
        Caste = True
    Case 6, 7
        Dt = True: Caste = True
    Case 8
        
    Case 9
        Dt = True
    
    
End Select

lblDate1.Enabled = Dt
cmdDate1.Enabled = Dt

With txtDate1
    .Enabled = Dt
    .BackColor = IIf(Dt, wisWhite, wisGray)
End With

End Sub

Private Sub TabStrip1_Click()
Dim intIndex As Byte
intIndex = TabStrip1.SelectedItem.Index
fra(intIndex).ZOrder 0

If intIndex = 1 Then cmdSave.Default = True
If intIndex = 2 Then cmdAdd.Default = True
If intIndex = 3 Then cmdShow.Default = True
End Sub

Private Sub txtData_DblClick(Index As Integer)
    txtData_KeyPress Index, vbKeyReturn
End Sub

Private Sub ScrollWindow(Ctl As Control)

If picSlider.Top + Ctl.Top + Ctl.Height > picViewport.ScaleHeight Then
    ' The control is below the viewport.
    Do While picSlider.Top + Ctl.Top + Ctl.Height > picViewport.ScaleHeight
        ' scroll down by one row.
        With VScroll1
            If .Value + .SmallChange <= .Max Then
                .Value = .Value + .SmallChange
            Else
                .Value = .Max
            End If
        End With
    Loop

ElseIf picSlider.Top + Ctl.Top < 0 Then
    ' The control is above the viewport.
    ' Keep scrolling until it is in viewport.
    Do While picSlider.Top + Ctl.Top < 0
        With VScroll1
            If .Value - .SmallChange >= .Min Then
                .Value = .Value - .SmallChange
            Else
                .Value = .Min
            End If
        End With
    Loop
End If

End Sub

Private Sub txtData_GotFocus(Index As Integer)
txtPrompt(Index).ForeColor = vbBlue
SetDescription txtPrompt(Index)
'TabStrip.Tabs(2).Tag = Index
' Scroll the window, so that the
' control in focus is visible.
ScrollWindow txtData(Index)

' Select the text, if any.
With txtData(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With

' If the display type is Browse, then
' show the command button for this text.
Dim strDispType As String
Dim TextIndex As String
cmdSave.Default = True

strDispType = ExtractToken(txtPrompt(Index).Tag, "DisplayType")
If StrComp(strDispType, "Browse", vbTextCompare) = 0 Then
    ' Get the cmdbutton index.
    TextIndex = ExtractToken(txtPrompt(Index).Tag, "textindex")
    If TextIndex <> "" Then cmd(Val(TextIndex)).Visible = True
ElseIf StrComp(strDispType, "List", vbTextCompare) = 0 Then
    ' Get the cmdbutton index.
    cmdSave.Default = False
    TextIndex = ExtractToken(txtPrompt(Index).Tag, "textindex")
    ' Get the cmdbutton index.
    On Error Resume Next
    If TextIndex <> "" Then
        If cmb(Val(TextIndex)).Visible Then Exit Sub
        cmb(Val(TextIndex)).Visible = True
        cmb(Val(TextIndex)).SetFocus
    End If
End If


' Hide all other command buttons...
Dim I As Integer
For I = 0 To cmd.count - 1
    If I <> Val(TextIndex) Or TextIndex = "" Then
        cmd(I).Visible = False
    End If
Next

' Hide all other combo boxes.
For I = 0 To cmb.count - 1
    If I <> Val(TextIndex) Or TextIndex = "" Then
        cmb(I).Visible = False
    End If
Next

End Sub


Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strDisp As String
Dim strIndex As String
On Error Resume Next

If KeyAscii = vbKeyReturn Then
    ' Check if the display type is "LIST".
    strDisp = ExtractToken(txtPrompt(Index).Tag, "DisplayType")
    If StrComp(strDisp, "List", vbTextCompare) = 0 Then
        ' Get the index of the combo to display.
        strIndex = ExtractToken(txtPrompt(Index).Tag, "TextIndex")
        If Trim$(strIndex) <> "" Then
            cmb(Val(strIndex)).Visible = True
            cmb(Val(strIndex)).SetFocus
            cmb(Val(strIndex)).ZOrder 0
        End If
    Else
        SendKeys "{TAB}"
    End If
End If

End Sub

Private Sub txtData_LostFocus(Index As Integer)
txtPrompt(Index).ForeColor = vbBlack
Dim strDatSrc As String
Dim Lret As Long
Dim txtIndex As Integer
Dim rst As ADODB.Recordset

' If the item is IntroducerID, validate the
' ID and name.
strDatSrc = ExtractToken(txtPrompt(Index).Tag, "DataSource")
Select Case UCase(strDatSrc)
    Case "INTRODUCERID"
        ' Check if any data is found in this text.
        If Trim$(txtData(Index).Text) <> "" Then
            If Val(txtData(Index).Text) <= 0 Then
                'MsgBox "Invalid account number specified !", vbExclamation, gAppName & " - Error"
                MsgBox GetResourceString(500), vbExclamation, gAppName & " - Error"
                ActivateTextBox txtData(Index)
                Exit Sub
            End If
            gDbTrans.SqlStmt = "SELECT Title + FirstName + space(1) + " _
                & "MiddleName + space(1) + Lastname AS Name FROM " _
                & "NameTab WHERE CustomerID = " & m_CustomerID
                
            Lret = gDbTrans.Fetch(rst, adOpenForwardOnly)
            If Lret > 0 Then
                txtIndex = GetIndex("IntroducerName")
                txtData(txtIndex).Text = FormatField(rst("Name"))
            End If
        Else
            txtIndex = GetIndex("IntroducerName")
            txtData(txtIndex).Text = ""
        End If
    Case "ACCID"
        RaiseEvent LoadClick
End Select

End Sub

Private Sub VScroll1_Change()
' Move the picSlider.
picSlider.Top = -VScroll1.Value
End Sub


