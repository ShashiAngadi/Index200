VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPrint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Print Wizard"
   ClientHeight    =   6600
   ClientLeft      =   1680
   ClientTop       =   1500
   ClientWidth     =   7680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   7680
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   6180
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "First"
            Object.Tag             =   ""
            ImageKey        =   "First"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Previous"
            Object.Tag             =   ""
            ImageKey        =   "Previous"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   1200
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Next"
            Object.Tag             =   ""
            ImageKey        =   "Next"
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Last"
            Object.Tag             =   ""
            ImageKey        =   "Last"
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.Tag             =   ""
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Close"
            Object.Tag             =   ""
            ImageKey        =   "Close"
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Saperator"
            Object.Tag             =   ""
            Object.Width           =   600
            Value           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   2300
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refershes with new print setting"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      Begin VB.ComboBox cmbFontName 
         Height          =   315
         Left            =   4650
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   30
         Width           =   1365
      End
      Begin VB.ComboBox cmbFontsize 
         Height          =   315
         Left            =   3750
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   30
         Width           =   765
      End
      Begin VB.TextBox txtPageCount 
         Height          =   330
         Left            =   720
         TabIndex        =   5
         Top             =   60
         Width           =   1170
      End
   End
   Begin VB.PictureBox picViewport 
      Height          =   6075
      Left            =   45
      ScaleHeight     =   6015
      ScaleWidth      =   7245
      TabIndex        =   1
      Top             =   15
      Width           =   7305
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         Left            =   0
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   5790
         Visible         =   0   'False
         Width           =   4545
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3795
         Left            =   6960
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   -15
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picPrint 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5700
         Left            =   15
         ScaleHeight     =   5670
         ScaleWidth      =   6165
         TabIndex        =   2
         Top             =   0
         Width           =   6195
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2445
      Top             =   2175
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrint.frx":0000
            Key             =   "First"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrint.frx":031A
            Key             =   "Previous"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrint.frx":0634
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrint.frx":094E
            Key             =   "Last"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrint.frx":0C68
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrint.frx":0D7A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmPrint.frx":0E8C
            Key             =   "Close"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As Integer
Public m_PrintOptDialog As frmPrintOpt

Public Event ProcessEvent(eventNo As Integer)

Public Sub Init()
' Set the margins.

With picPrint
    .Width = Printer.ScaleWidth
    .Height = Printer.ScaleHeight
    
    ' Center the printing canvas
    If .Width < Me.ScaleWidth Then
        .Left = Me.ScaleWidth / 2 - .Width / 2
    End If
    If .Height < Me.ScaleHeight Then
        .Top = Me.ScaleHeight / 2 - .Height / 2
    End If
End With

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

With VScroll1
    If KeyCode = vbKeyDown Then
        If .value + .SmallChange <= .Max Then
            .value = .value + .SmallChange
        Else
            .value = .Max
        End If
    ElseIf KeyCode = vbKeyUp Then
        If .value - .SmallChange >= .Min Then
            .value = .value - .SmallChange
        Else
            .value = .Min
        End If
    ElseIf KeyCode = vbKeyPageUp Then
        If .value - .LargeChange >= .Min Then
            .value = .value - .LargeChange
        Else
            .value = .Min
        End If
    ElseIf KeyCode = vbKeyPageDown Then
        If .value + .LargeChange <= .Max Then
            .value = .value + .LargeChange
        Else
            .value = .Max
        End If
    End If
End With

With HScroll1
    If KeyCode = vbKeyLeft Then
        If .value - .SmallChange >= .Min Then
            .value = .value - .SmallChange
        Else
            .value = .Min
        End If
    ElseIf KeyCode = vbKeyRight Then
        If .value + .SmallChange <= .Max Then
            .value = .value + .SmallChange
        Else
            .value = .Max
        End If
    End If
End With


End Sub

Private Sub Form_Load()

' Remove the border for picPrint.
picPrint.BorderStyle = 0
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
'Toolbar1.Buttons(10).Image = LoadResPicture(155, vbResIcon)
' Set the background color for viewport

picViewport.BackColor = Me.BackColor
picViewport.BorderStyle = 1


End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim vertScroll As Boolean
Dim horzScroll As Boolean

With picViewport
    .Visible = False
    .Left = 0
    .Top = 0
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight - Toolbar1.Height
    If picPrint.Width > picViewport.ScaleWidth Then
        horzScroll = True
    End If
    If picPrint.Height > picViewport.ScaleHeight Then
        vertScroll = True
    End If
    .Visible = True
End With

' Position the scrollbars...
With VScroll1
    .Left = picViewport.ScaleWidth - .Width
    .Top = 0
    If horzScroll Then
        .Height = picViewport.ScaleHeight - HScroll1.Height
        HScroll1.Visible = True
    Else
        .Height = picViewport.ScaleHeight
        HScroll1.Visible = False
    End If
    .Min = 0
    .Max = picPrint.Height - picViewport.ScaleHeight
    If horzScroll Then
        .Max = .Max + HScroll1.Height
    End If
    .SmallChange = picViewport.ScaleHeight / 10
    .LargeChange = picViewport.ScaleHeight / 2
End With
With HScroll1
    .Left = 0
    .Top = picViewport.ScaleHeight - .Height
    If vertScroll Then
        .Width = picViewport.ScaleWidth - VScroll1.Width
        VScroll1.Visible = True
    Else
        .Width = picViewport.ScaleWidth
        VScroll1.Visible = False
    End If
    .Min = 0
    .Max = picPrint.Width - picViewport.ScaleWidth
    If vertScroll Then
        .Max = .Max + VScroll1.Width
    End If
    .SmallChange = picViewport.ScaleWidth / 10
    .LargeChange = picViewport.ScaleWidth / 2
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
'""(Me.hwnd, False)
If Not m_PrintOptDialog Is Nothing Then
    Unload m_PrintOptDialog
End If
Set m_PrintOptDialog = Nothing
End Sub

Private Sub HScroll1_Change()
picPrint.Left = -HScroll1.value
End Sub

Private Sub lblFontSize_Click()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case LCase(Button.Key)
    Case "first"
        RaiseEvent ProcessEvent(wis_SHOW_FIRST)
    Case "previous"
        RaiseEvent ProcessEvent(wis_SHOW_PREVIOUS)
    Case "next"
        RaiseEvent ProcessEvent(wis_SHOW_NEXT)
    Case "last"
        RaiseEvent ProcessEvent(wis_SHOW_LAST)
    Case "print"

        ' Show a message box asking whether the user wants to
         ' print the current page or all pages.
         If m_PrintOptDialog Is Nothing Then
            Set m_PrintOptDialog = New frmPrintOpt
         End If
         If Not ExcelExists Then
            m_PrintOptDialog.optExcel.Enabled = False
         End If
         m_PrintOptDialog.Show vbModal
    
         If m_PrintOptDialog.Status = wis_OK Then
            ' Raise an appropriate depending upon the
            ' user selection in the print options dialog.
            With m_PrintOptDialog
                
                If .optPrintAllBegin And .chkPause.value = vbChecked Then
                    RaiseEvent ProcessEvent(wis_PRINT_ALL_PAUSE)
                
                ElseIf .optPrintAllBegin And .chkPause.value = vbUnchecked Then
                    RaiseEvent ProcessEvent(wis_PRINT_ALL)

                ElseIf .optPrintAllCur And .chkPause.value = vbChecked Then
                    RaiseEvent ProcessEvent(wis_PRINT_CURRENT_PAUSE)
                ElseIf .optPrintCur Then
                    RaiseEvent ProcessEvent(wis_PRINT_CURRENT)
                ElseIf .optExcel Then
                    RaiseEvent ProcessEvent(wis_Print_Excel)
                End If
            End With
         End If
    Case "close"
        Me.Status = wis_CANCEL
        Unload Me
End Select

' Hide this guy.
If Not m_PrintOptDialog Is Nothing Then
    Unload m_PrintOptDialog
End If
Set m_PrintOptDialog = Nothing

End Sub

Private Sub VScroll1_Change()
picPrint.Top = -VScroll1.value
End Sub


