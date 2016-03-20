VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B44203C3-6FD7-4C5A-B02A-E52525F0ECEA}#1.0#0"; "WISPrint.ocx"
Begin VB.Form wisMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "INDEX-2000  Bank Transaction Manager"
   ClientHeight    =   5670
   ClientLeft      =   3585
   ClientTop       =   2415
   ClientWidth     =   7275
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00808080&
   Icon            =   "wisMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   7275
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   4
      Top             =   5235
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   767
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4939
            MinWidth        =   4939
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2117
            MinWidth        =   2117
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   4
            Enabled         =   0   'False
            TextSave        =   "SCRL"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picViewport 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   1725
      ScaleHeight     =   5235
      ScaleWidth      =   5295
      TabIndex        =   6
      Top             =   30
      Width           =   5355
      Begin WIS_GRID_PrintNew.WISPrint grdPrint 
         Left            =   1800
         Top             =   3720
         _ExtentX        =   714
         _ExtentY        =   714
      End
      Begin MSComDlg.CommonDialog cdb 
         Left            =   990
         Top             =   3330
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin Threed.SSPanel sizeBar 
      Height          =   5625
      Left            =   1575
      TabIndex        =   2
      Top             =   45
      Width           =   30
      _Version        =   65536
      _ExtentX        =   53
      _ExtentY        =   9922
      _StockProps     =   15
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      MousePointer    =   9
   End
   Begin VB.PictureBox picToolbar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   30
      ScaleHeight     =   5595
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   30
      Width           =   1455
      Begin VB.CommandButton cmdDown 
         Height          =   330
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4875
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton cmdUP 
         Height          =   330
         Left            =   1065
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   285
         Visible         =   0   'False
         Width           =   330
      End
      Begin Threed.SSPanel pnlSlider 
         Height          =   300
         Index           =   0
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   2040
         _Version        =   65536
         _ExtentX        =   3598
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Accounts"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox picCanvas 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4410
         Left            =   15
         ScaleHeight     =   4350
         ScaleWidth      =   1980
         TabIndex        =   1
         Top             =   300
         Width           =   2040
         Begin VB.Image img 
            Height          =   480
            Index           =   0
            Left            =   420
            Top             =   45
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option title"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   15
            TabIndex        =   5
            Top             =   585
            Visible         =   0   'False
            Width           =   1320
         End
      End
   End
   Begin Threed.SSPanel resizeGuide 
      Height          =   5625
      Left            =   1665
      TabIndex        =   3
      Top             =   45
      Visible         =   0   'False
      Width           =   30
      _Version        =   65536
      _ExtentX        =   53
      _ExtentY        =   9922
      _StockProps     =   15
      ForeColor       =   -2147483630
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFin 
         Caption         =   "Change Fin Year"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDate 
         Caption         =   "Change Trans Date"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "wisMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' Constants used in this module.
Private Const CTL_MARGIN = 200
Private Const wis_TOP = 0
Private Const wis_BOTTOM = 1

Private WorkingWindowHandles(4) As Long
Private Const MaxHandles = 5

' Previously selected button.
Private m_PrevButton As Integer

' Mouse co-ordinates.
Private m_MouseX As Single
Private m_MouseY As Single

' Recurssion suppresent.
Private resizeNow As Boolean

' Control Index of the current canvas in focus.
Private m_CanvasIndex As Integer

Private m_CtrlDown As Boolean 'Idenifies whether CONTROL key is pressed or not
Private m_ShiftDown As Boolean 'Idenifies whether SHIFT key is pressed or not

' Declare an application class object.
' This class is the interface between the
' main form and the methods...
Private WithEvents wisAppObj As wisApp
Attribute wisAppObj.VB_VarHelpID = -1

Private Sub AlignButtons()


Dim I As Integer
Dim Alignment As Integer
Dim TopCount As Integer
Dim BottomCount As Integer

' Align all the sliding panels.
For I = 0 To pnlSlider.count - 1
    ' Get the alignment property.
    With pnlSlider(I)
        Alignment = Val(ExtractToken(.Tag, "Alignment"))
        If Alignment = wis_TOP Then
            .Top = (.Index - 1) * .Height
            If I <> 0 Then TopCount = TopCount + 1
        Else
            .Top = picToolbar.ScaleHeight - _
                    (pnlSlider.count - .Index) * .Height
            BottomCount = BottomCount + 1
        End If
        .Width = picToolbar.ScaleWidth
        .Left = 0
    End With
Next

End Sub
' Returns the no. of buttons on the specified canvas,
' identified by canvas index.
Private Function BtnCount(TabPanelIndex As Integer) As Integer
On Error Resume Next
Dim I As Integer
Dim ctIndex As Integer

' Loop throug the collection of image buttons.
For I = 0 To img.count - 1
    ' Get the Container index for the button.
    ctIndex = Val(ExtractToken(img(I).Tag, "Container"))
    ' If the container index is the same as
    ' the index of the tabpanel, increment the count.
    If ctIndex = TabPanelIndex Then
        BtnCount = BtnCount + 1
    End If
Next
'
End Function
Private Sub DrawBorder(btnindex As Integer, Bevel As Integer)
Const BORDER_MARGIN = 30

If btnindex < 0 Then Exit Sub
With img(btnindex)
    ' Set the bevel.
    If Bevel = 0 Then
        picCanvas.ForeColor = picCanvas.BackColor
        lbl(btnindex).ForeColor = &H80000012
        img(btnindex).BorderStyle = 0
    ElseIf Bevel = 1 Then
        picCanvas.ForeColor = vbWhite
        lbl(btnindex).ForeColor = &H80000012
    ElseIf Bevel = 2 Then
        picCanvas.ForeColor = vbBlack
    End If
    
    ' Draw the top line.
    picCanvas.Line (.Left - BORDER_MARGIN, .Top - BORDER_MARGIN) _
            -(.Left + .Width + BORDER_MARGIN, .Top - BORDER_MARGIN)
    ' Draw the left line.
    picCanvas.Line (.Left - BORDER_MARGIN, .Top - BORDER_MARGIN) _
            -(.Left - BORDER_MARGIN, .Top + .Height + BORDER_MARGIN)

    If Bevel = 0 Then
        picCanvas.ForeColor = picCanvas.BackColor
    ElseIf Bevel = 1 Then
        picCanvas.ForeColor = vbBlack
    ElseIf Bevel = 2 Then
        picCanvas.ForeColor = vbWhite
    End If
    ' Draw the bottom line.
    picCanvas.Line (.Left - BORDER_MARGIN, .Top + .Height + BORDER_MARGIN) _
            -(.Left + .Width + BORDER_MARGIN, .Top + .Height + BORDER_MARGIN)
    ' Draw the right side line.
    picCanvas.Line (.Left + .Width + BORDER_MARGIN, .Top - BORDER_MARGIN) _
            -(.Left + .Width + BORDER_MARGIN, .Top + .Height + BORDER_MARGIN)
End With

End Sub
Public Sub DrawLogo()
Dim ViewportWidth As Single
Dim ViewportHeight As Single
Dim ViewportLeft As Single
Dim ViewportTop As Single
Dim strBanner As String

strBanner = "INDEX-2000"

' Set the dimensions of the viewport.
ViewportLeft = picToolbar.Width + sizeBar.Width
ViewportTop = 0
ViewportWidth = Me.ScaleWidth - ViewportLeft
ViewportHeight = Me.ScaleHeight - StatusBar1.Height

    Cls
    ' Print the logo.
    CurrentX = ViewportLeft + (ViewportWidth - TextWidth(strBanner)) / 2
    CurrentY = ViewportTop + (ViewportHeight - TextHeight(strBanner)) / 2
    ForeColor = vbBlue
    Print strBanner
    
    ' Print the shadow.
    CurrentX = ViewportLeft + (ViewportWidth - TextWidth(strBanner)) / 2 - 25
    CurrentY = ViewportTop + (ViewportHeight - TextHeight(strBanner)) / 2 - 25
    ForeColor = &HC0C0C0
    Print strBanner

End Sub


Private Function PanelCount(AlignmentPos As Integer) As Integer
Dim I As Integer
Dim AlignmentVal As Integer

For I = 1 To pnlSlider.count - 1
    AlignmentVal = Val(ExtractToken(pnlSlider(I).Tag, "Alignment"))
    If AlignmentPos = AlignmentVal Then PanelCount = PanelCount + 1
Next

End Function
Private Sub ResetScrollButtons()
' Decide whether or not to show the scroll buttons.
If picCanvas.Top + picCanvas.Height > picToolbar.ScaleHeight _
            - PanelCount(wis_BOTTOM) * pnlSlider(0).Height Then
    cmdDown.Visible = True
Else
    cmdDown.Visible = False
End If

If picCanvas.Top < PanelCount(wis_TOP) * pnlSlider(0).Height Then
    cmdUP.Visible = True
Else
    cmdUP.Visible = False
End If
If picCanvas.Height < Me.Height - picCanvas.Top Then picCanvas.Height = Me.Height - picCanvas.Top

End Sub
Private Function Serialize(obj As Object, txt As String) As Boolean

Dim new_value As String
Dim token_name As String
Dim token_value As String
Dim ctl_Index As Integer

On Error Resume Next

new_value = txt
' Examine each token in turn.
 Do
     ' Get the token name and value.
     GetToken new_value, token_name, token_value
     If token_name = "" Then Exit Do

     ' Examine each token and initialize.
     Select Case UCase$(token_name)
        Case "TOOLTAB"
            ' Load a copy of panel tab.
            Load pnlSlider(pnlSlider.count)
            pnlSlider(pnlSlider.UBound).Visible = True
            ' Load a canvas for this tab.
            If Not Serialize(pnlSlider(pnlSlider.UBound), _
                    token_value) Then GoTo end_line

        Case "BUTTON"
            ' Load an img control
            Load img(img.count)
            If Not Serialize(img(img.UBound), _
                    token_value) Then GoTo end_line
            ' Update the tag of this button with the
            ' index of the tooltab.
            With img(img.UBound)
                .Tag = PutToken(.Tag, "Container", pnlSlider.UBound)
            End With
        Case "LABEL"
            ' Load a label.
            Load lbl(lbl.count)
            If Not Serialize(lbl(lbl.UBound), _
                    token_value) Then GoTo end_line
            With lbl(lbl.count - 1)
                If .Width <= TextWidth(.Caption) Then _
                    .Height = .Height + .Height * 1.25
            End With
        Case "ICON"
            img(img.UBound).Picture = LoadResPicture(Val(token_value), vbResIcon)
        Case "BITMAP"
            img(img.UBound).Picture = LoadResPicture(Val(token_value), vbResBitmap)
        Case "CAPTION"
            obj.Font.name = gFontName
            obj.Font.Size = gFontSize
            obj.Caption = token_value
        Case "ALIGNMENT"
            If StrComp(token_value, "Top", vbTextCompare) = 0 Then
                obj.Tag = "Alignment=" & wis_TOP
            Else
                obj.Tag = "Alignment=" & wis_BOTTOM
            End If
        Case "KEY"
            obj.Tag = PutToken(obj.Tag, "Key", token_value)
    End Select
Loop

Set obj = Nothing
Serialize = True

end_line:
    Exit Function

End Function
Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

End Sub


Private Sub ShowIcons(PanelIndex As Integer)

Dim I As Integer
Dim iconCount As Integer
Dim ctIndex As Integer
Dim PositionSet As Boolean
Dim PrevTop As Single
Dim PanelHeight As Single
'PrevTop = PanelCount(wis_TOP) * pnlSlider(0).Height
PrevTop = CTL_MARGIN

' Hide or display icons depending upon the
' paneltab index, to which they are bound.
For I = 0 To img.count - 1
    With img(I)
        ctIndex = Val(ExtractToken(.Tag, "Container"))
        If ctIndex = PanelIndex Then
            img(I).Visible = True
            lbl(I).Visible = True
            If Not PositionSet Then
                img(I).Top = PrevTop
                PositionSet = True
            Else
                img(I).Top = PrevTop + img(I - 1).Height + lbl(I - 1).Height + CTL_MARGIN
            End If
            lbl(I).Top = img(I).Top + img(I).Height + 10 '+ CTL_MARGIN
            iconCount = iconCount + 1
            PrevTop = img(I).Top
            PanelHeight = PrevTop + lbl(1).Height + img(I).Height + 10
        Else
            img(I).Visible = False
            lbl(I).Visible = False
        End If
    End With
Next

' Set the height of the canvas.
'resizeNow = False
'if PanelHeight <
picCanvas.Height = PanelHeight + CTL_MARGIN '(iconCount) * (img(0).Height + lbl(0).Height + CTL_MARGIN) + CTL_MARGIN
picCanvas.Top = PanelCount(wis_TOP) * pnlSlider(0).Height

End Sub

'THis sub Make the visible ivisible of the
'CmdUp & cmdDown buttons
Private Sub SetImagePosition(ImgNO As Integer)

Dim TopCount As Integer
Dim BottomCount As Integer
Dim I As Integer
Dim ImageTop As Integer
Dim ImageBottom As Integer

TopCount = PanelCount(wis_TOP)
BottomCount = PanelCount(wis_BOTTOM)

ImageTop = img(ImgNO).Top
ImageBottom = lbl(ImgNO).Top + lbl(ImgNO).Height

'Now che the Imgae Position is Viisble or not

With picCanvas
    'if image top is not visible then move the piccanvas down
    While ImageTop + .Top - TopCount * pnlSlider(0).Height < 0
        If Not cmdUP.Visible Then Exit Sub
        Call cmdUP_Click
    Wend
    
    'if image bottom is not visible then move the piccanvas up
    While ImageBottom + .Top > picToolbar.ScaleHeight - BottomCount * pnlSlider(0).Height
        If Not cmdDown.Visible Then Exit Sub
        Call cmdDown_Click
    Wend
    
End With


End Sub

Private Sub cmdDown_Click()

Dim MoveDistance As Single
Dim ScrollDiff As Single
Dim TopCount As Integer
Dim BottomCount As Integer
Dim I As Integer
Dim Alignment As Integer

' Get the count of panels that are top and bottom aligned.
TopCount = PanelCount(wis_TOP)
BottomCount = PanelCount(wis_BOTTOM)

' See if the canvas can be moved up.
With picCanvas
    ScrollDiff = .Top + .Height - (picToolbar.ScaleHeight _
                - BottomCount * pnlSlider(0).Height)
    ' If no scope for further movement upwards, exit.
    'If ScrollDiff <= 0 Then Exit Sub

    ' Set the move distance.
    MoveDistance = img(0).Height + lbl(0).Height + CTL_MARGIN
    'If ScrollDiff < MoveDistance Then
    If ScrollDiff < 0 Then
        .Top = .Top - ScrollDiff
        cmdDown.Visible = False
    Else
        .Top = .Top - MoveDistance
        cmdDown.Visible = ScrollDiff > MoveDistance
    End If
    
    ' If the top is less than 0, show the down scroll button.
    If .Top < TopCount * pnlSlider(0).Height Then
        cmdUP.Visible = True
    End If
End With

End Sub


Private Sub cmdUP_Click()

Dim MoveDistance As Single
Dim ScrollDiff As Single
picCanvas.SetFocus
Dim TopCount As Integer
Dim BottomCount As Integer
Dim I As Integer
Dim Alignment As Integer

' Get the count of panels which are TOP and BOTTOM aligned.
TopCount = PanelCount(wis_TOP)
BottomCount = PanelCount(wis_BOTTOM)

' See if the canvas can be moved down.
With picCanvas
    ' If no scope for further downward movement, exit.
    If .Top >= TopCount * pnlSlider(0).Height Then
        cmdUP.Visible = False
        Exit Sub
    End If
    ' Set the move distance.
    MoveDistance = img(0).Height + lbl(0).Height + CTL_MARGIN
    If MoveDistance > (TopCount * pnlSlider(0).Height) - .Top Then
        .Top = (TopCount * pnlSlider(0).Height)
        cmdUP.Visible = False
    Else
        .Top = .Top + MoveDistance
    End If

    ' If the canvas scrolls past the bottom of viewport,
    ' display the scroll button near the bottom.
    If .Top + .Height > picToolbar.ScaleHeight - BottomCount * pnlSlider(0).Height Then
        cmdDown.Visible = True
    Else
        cmdDown.Visible = False
    End If
End With

End Sub

Private Sub Form_Activate()
'Debug.Print Me.ActiveControl.Name
'Call SetActiveWindow(gWindowHandle)
Call wisAppObj.MakeWindowsActive
Exit Sub
Static buttonsAligned As Boolean
If Not buttonsAligned Then
    ShowIcons 1
    buttonsAligned = True
End If

End Sub

Private Sub Form_Click()
    Call SetActiveWindow(gWindowHandle)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If Any Form is loaded above this form
'then trnafer the key code event to that form
If gWindowHandle Then Exit Sub

If (KeyCode < Asc("1") Or KeyCode > Asc("9")) And KeyCode <> vbKeyTab And _
    (KeyCode <> vbKeyUp And KeyCode <> vbKeyDown) And KeyCode <> 13 Then Exit Sub

Dim CtrlDown  'Idenifies whether CONTROL key is pressed or not
Dim ShiftDown 'Idenifies whether SHIFT key is pressed or not

CtrlDown = (Shift And vbCtrlMask) > 0
ShiftDown = (Shift And vbShiftMask) > 0
m_CtrlDown = CtrlDown
m_ShiftDown = ShiftDown

Static ImgNO As Integer 'identifies the the first image no for this panel

'if user presses enter key then he had clicked the image button
'so execute the respective code and ext the sub
If Not CtrlDown And m_PrevButton > 0 And KeyCode = 13 Then _
            Call img_Click(m_PrevButton): Exit Sub

Dim lblNo As Integer
'if he presses only the keyup
If KeyCode = vbKeyUp And Not CtrlDown Then
    'If Val(lbl(0).Tag) <= 0 Then Exit Sub
    If m_PrevButton <= 0 Then Exit Sub
    lblNo = m_PrevButton 'Val(lbl(0).Tag)
    If lblNo < ImgNO Or Not lbl(lblNo).Visible Then lblNo = ImgNO
    lblNo = lblNo - 1
    If Not lbl(lblNo).Visible Then
        Do
            lblNo = lblNo + 1
            If lblNo = lbl.count Then Exit Do
            If Not lbl(lblNo).Visible Then Exit Do
        Loop
        lblNo = lblNo - 1
    End If
    If m_PrevButton > 0 Then DrawBorder m_PrevButton, 0
    SetImagePosition (lblNo)
    DrawBorder lblNo, 1
    m_PrevButton = lblNo
    Exit Sub
End If

If KeyCode = vbKeyDown And Not CtrlDown Then
    lblNo = m_PrevButton 'Val(lbl(0).Tag)
    If lblNo < 0 Then lblNo = 0
    If lblNo < ImgNO - 1 Then lblNo = ImgNO
    lblNo = lblNo + 1
    If lblNo = lbl.count Then lblNo = ImgNO
    If Not lbl(lblNo).Visible Then lblNo = ImgNO
    
    If m_PrevButton > 0 Then DrawBorder m_PrevButton, 0
    SetImagePosition (lblNo)
    DrawBorder lblNo, 1
    m_PrevButton = lblNo
    Exit Sub
End If

If Not CtrlDown Then Exit Sub

If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
  If KeyCode = vbKeyUp And cmdUP.Visible Then Call cmdUP_Click
  If KeyCode = vbKeyDown And cmdDown.Visible Then Call cmdDown_Click
  Exit Sub
End If


If ShiftDown And KeyCode <> vbKeyTab Then
    
'    If ImgNO = 0 Then ImgNO = 1
'    ImgNO = ImgNO + CInt(Chr(KeyCode)) - 1
'    If ImgNO >= img.COunt Then Exit Sub
'    If Not img(ImgNO).Visible Then Exit Sub
'    Call img_Click(ImgNO)
    lblNo = ImgNO
    If lblNo = 0 Then lblNo = 1
    lblNo = lblNo + CInt(Chr(KeyCode)) - 1
    If lblNo >= img.count Then Exit Sub
    If Not img(lblNo).Visible Then Exit Sub
    Call img_Click(lblNo)
    
    Exit Sub
End If

If KeyCode = vbKeyTab Then
    'find the which panel is loaded
    Dim PanelIndex As Integer
    For PanelIndex = 0 To pnlSlider.count - 1
        lblNo = Val(ExtractToken(pnlSlider(PanelIndex).Tag, "Alignment"))
        If lblNo = wis_BOTTOM Then Exit For
    Next
    PanelIndex = PanelIndex - 1
    PanelIndex = PanelIndex + IIf(ShiftDown, -1, 1)
    If PanelIndex < 1 Then PanelIndex = pnlSlider.count - 1
    If PanelIndex = pnlSlider.count Then PanelIndex = 1
    
Else
    If (KeyCode < Asc("1") Or KeyCode >= Asc(pnlSlider.count)) Then Exit Sub
    PanelIndex = Val((Chr(KeyCode)))
End If

Call pnlSlider_Click(PanelIndex)

For lblNo = (PanelIndex - 1) * 5 To img.count - 1
    If img(lblNo).Visible Then Exit For
Next
ImgNO = lblNo

End Sub

Private Sub Form_Load()
Dim I As Integer
'MsgBox "Befor picCanvas.BorderStyle "
picCanvas.BorderStyle = 0
'MsgBox "After picCanvas.BorderStyle "

'MsgBox "Before Center the window."
' Center the window.
Left = Screen.Width / 2 - Me.Width / 2
Top = Screen.Height / 2 - Me.Height / 2
'Set icon for the form caption
Icon = LoadResPicture(161, vbResIcon)

' Read the toolbar layout information
' from toolbar.lyt.
Dim strLayoutFile As String
Dim nFile As Integer
Dim txt As String

' Get a file handle.
 nFile = FreeFile
' Open the layout file
If gLangOffSet Then
    strLayoutFile = App.Path & "\tbarkan.lyt"
    Me.FontName = gFontName
Else
    strLayoutFile = App.Path & "\toolbar.lyt"
End If

On Error Resume Next
Open strLayoutFile For Input As nFile
If Err.Number = 53 Then
    MsgBox "Lay out file not found", , wis_MESSAGE_TITLE
    gDbTrans.CloseDB
    End
End If

' Read the contents at once.
txt = Input(LOF(nFile), #nFile)
Close #nFile

' Create the Picture-toolbar from serialization.
cmdUP.Picture = LoadResPicture(136, vbResBitmap)
cmdDown.Picture = LoadResPicture(137, vbResBitmap)
cmdUP.ZOrder
cmdDown.ZOrder

Serialize picToolbar, txt
pnlSlider_Click 1


' Create an instance of application object.
If wisAppObj Is Nothing Then Set wisAppObj = New wisApp

With Me.StatusBar1
    .Font.name = gFontName
    .Font.Size = 14
    .Panels(1).Text = gCompanyName
    .Panels(4).Style = sbrScrl
End With


End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = 0 Then Call ExitApplication(True, Cancel)

End Sub
Private Sub Form_Resize()
On Error Resume Next
With picToolbar
    .Left = 0
    .Top = 0
    .Height = Me.ScaleHeight - StatusBar1.Height
    sizeBar.Left = .Left + .Width
End With
sizeBar.Top = 0
sizeBar.Height = Me.ScaleHeight - StatusBar1.Height
resizeGuide.Height = sizeBar.Height

With picViewport
    .Left = sizeBar.Left + sizeBar.Width
    .Width = Me.ScaleWidth - picToolbar.Width - sizeBar.Width
    .Height = Me.ScaleHeight - StatusBar1.Height
    .Top = 0
End With


End Sub

Private Sub img_Click(Index As Integer)
'If any modules are runnnig then do not Show other module till closing them
'If gWindowHandle <> 0 And gWindowHandle <> Me.hwnd Then Exit Sub

If Val(img(0).Tag) Then
    img(img(0).Tag).BorderStyle = 0
    Call DrawBorder(Val(img(0).Tag), 0)
    lbl(img(0).Tag).ForeColor = &H80000012
End If

img(Index).BorderStyle = 1
img(0).Tag = Index
lbl(Index).ForeColor = vbBlue

'MsgBox Me.ButtonKey(Index)
Select Case UCase$(Me.ButtonKey(Index))
    Case "SBACC"
        wisAppObj.ShowSBDialog
    Case "CAACC"
        wisAppObj.ShowCADialog
    Case "FDACC"
        wisAppObj.ShowFDDialog
    Case "RDACC"
        wisAppObj.ShowRDDialog
    Case "PDACC"
        wisAppObj.ShowPDDialog
    Case "SHG"
        wisAppObj.ShowSHGDialog
        'wisAppObj.ShowSuspence
    Case "MEMBERS"
        wisAppObj.ShowMemberDialog
    Case "DEPLOANS"
        wisAppObj.ShowDepositLoan
    Case "KCC"
        wisAppObj.ShowKCCDialog
    Case "LNACCOUNT"
        wisAppObj.ShowLoanCreateDialog
    Case "LNTRANS"
        wisAppObj.ShowLoanTransDialog
    Case "LNSCHEME"
        wisAppObj.ShowLoanSchemeCreateDialog
    Case "LNREPORTS"
        wisAppObj.ShowLoanReportDialog
    Case "LNCREATE"
        wisAppObj.ShowLoanSchemeCreateDialog
    Case "CUSTINFO"
        wisAppObj.ShowCustInfo
    Case "TRACC"
        wisAppObj.ShowReportDialog (wisTradingAccount)
    Case "PLACC"
        wisAppObj.ShowReportDialog (wisProfitLossStatement)
    Case "TRIALBALANCE"
        wisAppObj.ShowReportDialog (wisTrialBalance)
    Case "BALANCESHEET"
        wisAppObj.ShowReportDialog (wisBalanceSheet)
    Case "CASHBOOK"
        wisAppObj.ShowReportDialog (wisDetailCashBook)
    Case "DAILYBOOK"
        wisAppObj.ShowReportDialog (wisDailyCashBook)
    Case "DEBITCREDIT"
        wisAppObj.ShowReportDialog (wisDebitCreditStatement)
    Case "GENLEDGER"
        wisAppObj.ShowReportDialog (wisBankBalance)
    Case "CRTRANS"
        wisAppObj.ShowReportDialog (wisRepCreditTrans)
    Case "DRTRANS"
        wisAppObj.ShowReportDialog (wisRepDebitTrans)
    Case "BALANCE"
        wisAppObj.ShowReportDialog (wisBalancing)
    Case "DDC"
        Call wisAppObj.ShowReportDialog
    Case "CONTRA"
        wisAppObj.ShowContra
    
    Case "SUSPACC"
        wisAppObj.ShowSuspence
        
    Case "UTILS"
        wisAppObj.ShowUtils
    Case "USERS"
        If gCurrUser Is Nothing Then Set gCurrUser = New clsUsers
        gCurrUser.ShowUserDialog
    Case "SEARCH"
        wisAppObj.ShowCustomerSearch
    Case "TRADECUST"
        wisAppObj.ShowCompanyCreation
    Case "PRODUCT"
        wisAppObj.ShowMaterialDialog
    Case "PURCHASE"
        wisAppObj.ShowMaterialPurchase
    Case "SALES"
        wisAppObj.ShowMaterialSales
    Case "INVOICE"
        wisAppObj.ShowMaterialInvoiceDetails
    Case "TRANSFER"
        wisAppObj.ShowMaterialTransfer
    Case "STOCKREP"
        wisAppObj.ShowMaterialReport
    Case "BANKACC"
        wisAppObj.ShowBankDialog
    Case "CLEARING"
        wisAppObj.ShowClearingDialog
    Case "BANKNAME"
        wisAppObj.ShowCompanyDetails
    Case "ABOUT"
        'wisAppObj.ShowSuspence
        If (gCurrUser.UserPermissions And perOnlyWaves) Then
            wisAppObj.ShowDataEntry
        End If
        
    Case "EXIT"
        Debug.Print m_CtrlDown
        Call ExitApplication(True, 0)
        Set wisMain = Nothing
    Case "HELP"
        '''Nothing Doing here
        wisAppObj.ShowPassing
    Case Else
        MsgBox "The trail period Of this Program is over " & vbCrLf & "Contact waves Information Systems., GADAG", vbInformation, wis_MESSAGE_TITLE
    
End Select

End Sub

Private Sub img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index <= 0 Then Exit Sub
DrawBorder Index, 2

End Sub
Private Sub img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If m_PrevButton = Index Then Exit Sub
If Index <= 0 Then Exit Sub
' Remove the border for previous button, if present.
If m_PrevButton > 0 Then DrawBorder m_PrevButton, 0
DrawBorder Index, 1
m_PrevButton = Index

End Sub

Private Sub img_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index <= 0 Then Exit Sub
DrawBorder Index, 1
End Sub





Private Sub mnuDate_Click()
If gCurrUser Is Nothing Then Exit Sub

gCurrUser.ShowDateChange

End Sub

Private Sub mnuExit_Click()

Call ExitApplication(True, 0)
End Sub

Private Sub mnuFin_Click()
If gCurrUser Is Nothing Then Exit Sub

gCurrUser.ShowFinChange

End Sub


Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If m_PrevButton <= 0 Then Exit Sub

DrawBorder m_PrevButton, 0
m_PrevButton = -1

End Sub

Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If m_PrevButton = -1 Then Exit Sub
DrawBorder m_PrevButton, 1
m_PrevButton = -1
End Sub
Private Sub picCanvas_Resize()

'If Not resizeNow Then Exit Sub
'resizeNow = False
Dim I As Integer
Dim TmpStr As String
' Set the position of image buttons.
For I = 1 To img.count - 1
    With img(I)
        .Left = (picCanvas.ScaleWidth - .Width) / 2
    End With
    With lbl(I)
        .Width = picCanvas.ScaleWidth - 20
        'adjust the height
        TmpStr = .Caption
        '.Height = TextHeight(tmpStr)
        '.AutoSize = True
        .Left = (picCanvas.ScaleWidth - .Width) / 2
    End With
Next

End Sub
Private Sub picToolbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picCanvas_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picToolbar_Resize()
Dim I As Integer
Dim TopCount As Integer
Dim BottomCount As Integer
Dim Alignment As Integer
On Error Resume Next
' If the form is minimized, exit.
If Me.WindowState = vbMinimized Then Exit Sub

' Align all the sliding panels.
For I = 0 To pnlSlider.count - 1
    ' Get the alignment property.
    With pnlSlider(I)
        Alignment = Val(ExtractToken(.Tag, "Alignment"))
        If Alignment = wis_TOP Then
            .Top = (.Index - 1) * .Height
            If I <> 0 Then TopCount = TopCount + 1
        Else
            .Top = picToolbar.ScaleHeight - _
                    (pnlSlider.count - .Index) * .Height
            BottomCount = BottomCount + 1
        End If
        .Width = picToolbar.ScaleWidth
        .Left = 0
    End With
Next

' Set the width of canvas.
With picCanvas
    resizeNow = True
    .Width = picToolbar.ScaleWidth
End With

' Position of the scroll buttons...
Const SCROLLMARGIN = 100
cmdUP.Left = picToolbar.ScaleWidth - cmdUP.Width - SCROLLMARGIN
cmdUP.Top = SCROLLMARGIN + TopCount * pnlSlider(0).Height
With cmdDown
    .Left = cmdUP.Left
    .Top = picToolbar.ScaleHeight - .Height - SCROLLMARGIN _
            - BottomCount * pnlSlider(0).Height
End With

' Hide or display the scroll buttons.
ResetScrollButtons

End Sub

Private Sub pnlSlider_Click(Index As Integer)
Dim Alignment As Integer
Dim strTmp As String
Dim I As Integer

' Get the alignment property.
Alignment = Val(ExtractToken(pnlSlider(Index).Tag, "Alignment"))
If Alignment = wis_TOP Then
    ' Change the alignment property for all the panels
    ' that have index greater than the current index.
    For I = Index + 1 To pnlSlider.count - 1
        With pnlSlider(I)
            If I > 1 Then .Tag = PutToken(.Tag, "Alignment", wis_BOTTOM)
        End With
    Next
    picToolbar_Resize
    ShowIcons (Index)

Else
    ' Change the alignment property for all the panels
    ' that have index less than the current index.
    For I = 1 To Index
        With pnlSlider(I)
            If I > 1 Then .Tag = PutToken(.Tag, "Alignment", wis_TOP)
        End With
    Next
    picToolbar_Resize
    ShowIcons (Index)
End If

' Hide or display the scroll buttons.
ResetScrollButtons
End Sub

Private Sub sizeBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
m_MouseX = X
m_MouseY = Y
With resizeGuide
    .Left = sizeBar.Left
    .Visible = True
End With
End Sub


Private Sub sizeBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    resizeGuide.Left = resizeGuide.Left + X - m_MouseX
    m_MouseX = X
End Sub
Private Sub sizeBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next
resizeGuide.Visible = False

' Check for min width of toolbar.
If resizeGuide.Left <= 2 * img(0).Width Then resizeGuide.Left = img(0).Left * 2

sizeBar.Left = resizeGuide.Left
picToolbar.Width = sizeBar.Left '- 25
With picViewport
    .Width = Me.ScaleWidth - _
        picToolbar.Width - sizeBar.Width
    .Left = sizeBar.Left + sizeBar.Width
End With

End Sub

Public Property Get ButtonKey(Indx As Integer) As String
    ButtonKey = ExtractToken(img(Indx).Tag, "Key")
End Property
Public Property Let ButtonKey(Indx As Integer, ByVal vNewValue As String)
    img(Indx).Tag = PutToken(img(Indx).Tag, "Key", vNewValue)
End Property

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)
If Panel.Key = "TransDate" Then Call mnuDate_Click
'If Panel.Key = "TransDate" Then Call mnuDate_Click

 
End Sub


