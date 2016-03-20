VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   ClientHeight    =   4095
   ClientLeft      =   2520
   ClientTop       =   1515
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   7560
   Begin MSComctlLib.ListView lstResults 
      Height          =   1665
      Left            =   90
      TabIndex        =   13
      Top             =   2070
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   2937
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5940
      TabIndex        =   12
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   5940
      TabIndex        =   11
      Top             =   420
      Width           =   1215
   End
   Begin VB.Frame fraName 
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   270
      TabIndex        =   0
      Top             =   450
      Width           =   4485
      Begin VB.ComboBox cmbLocation 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   3285
      End
      Begin VB.TextBox txtName 
         Height          =   345
         Left            =   990
         TabIndex        =   2
         Top             =   240
         Width           =   3285
      End
      Begin VB.Label Label2 
         Caption         =   "Location"
         Height          =   285
         Left            =   150
         TabIndex        =   3
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   705
      End
   End
   Begin MSComctlLib.TabStrip tabStrip 
      Height          =   1875
      Left            =   90
      TabIndex        =   5
      Top             =   60
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3307
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Name && Locatioin"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Date Created"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraDate 
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   270
      TabIndex        =   6
      Top             =   450
      Width           =   4485
      Begin VB.TextBox txtTo 
         Height          =   345
         Left            =   1080
         TabIndex        =   10
         Top             =   720
         Width           =   1545
      End
      Begin VB.TextBox txtFrom 
         Height          =   345
         Left            =   1080
         TabIndex        =   7
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label4 
         Caption         =   "Between:"
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   330
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "and:"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   750
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private m_MinWidth As Single
'Private m_Minheight As Single

Private Const m_MinWidth = 4305
Private Const m_Minheight = 2430
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()

Dim FirstCond As Boolean
FirstCond = True

gDbTrans.SQLStmt = "Select * from NameTab "

If txtName.Text <> "" Then
    If FirstCond Then
        gDbTrans.SQLStmt = gDbTrans.SQLStmt & " where "
        FirstCond = False
    Else
        gDbTrans.SQLStmt = gDbTrans.SQLStmt & " and "
    End If
    
    gDbTrans.SQLStmt = gDbTrans.SQLStmt & "Firstname = '" & txtName.Text & "' or " & _
                "MiddleName = '" & txtName.Text & "' or " & _
                "LastName = '" & txtName.Text & "'"
End If

If cmbLocation.ListIndex >= 0 Then
    If FirstCond Then
        gDbTrans.SQLStmt = gDbTrans.SQLStmt & " where "
        FirstCond = False
    Else
        gDbTrans.SQLStmt = gDbTrans.SQLStmt & " and "
    End If
    gDbTrans.SQLStmt = gDbTrans.SQLStmt + "Place = '" & cmbLocation.List(cmbLocation.ListIndex) & "'"

End If

If Trim$(txtFrom.Text) <> "" Then
    If FirstCond Then
        gDbTrans.SQLStmt = gDbTrans.SQLStmt & " where "
        FirstCond = False
    Else
        gDbTrans.SQLStmt = gDbTrans.SQLStmt & " and "
    End If
    gDbTrans.SQLStmt = gDbTrans.SQLStmt + "Date >= #" & txtFrom.Text & "#"
End If

If Trim$(txtTo.Text) <> "" Then
    If FirstCond Then
        gDbTrans.SQLStmt = gDbTrans.SQLStmt & " where "
        FirstCond = False
    Else
        gDbTrans.SQLStmt = gDbTrans.SQLStmt & " and "
    End If
    
    gDbTrans.SQLStmt = gDbTrans.SQLStmt + "Date <= #" & txtTo.Text & "#"
End If
    
End Sub

Private Sub Form_Load()

'tabStrip.ZOrder 0
fraName.ZOrder 0
fraDate.ZOrder 1
'm_MinWidth = Me.Width
'm_Minheight = Me.Height
Dim I As Integer

lstResults.ColumnHeaders.Add , "Col1", "Name"
lstResults.ColumnHeaders.Add , "Col2", "Age"
lstResults.ColumnHeaders.Add , "Col3", "Gender"
lstResults.ColumnHeaders.Add , "Col4", "Profession"
lstResults.ColumnHeaders.Add , "Col5", "Home Address"
lstResults.ColumnHeaders.Add , "Col6", "Office Address"

End Sub

Private Sub Form_Resize()

Dim Margin As Single
Dim CmdPos As Single

Margin = tabStrip.Left
CmdPos = cmdSearch.Left

If frmSearch.WindowState = vbMinimized Then Exit Sub

If frmSearch.Width < m_MinWidth Then
    frmSearch.Width = m_MinWidth
    Exit Sub
Else
    cmdSearch.Left = frmSearch.Width - ((Margin * 2) + cmdSearch.Width)
    cmdClose.Left = frmSearch.Width - ((Margin * 2) + cmdSearch.Width)
    tabStrip.Width = tabStrip.Width + cmdSearch.Left - CmdPos
    fraName.Left = tabStrip.Left + (Margin * 2)
    fraDate.Left = tabStrip.Left + (Margin * 2)
    lstResults.Width = lstResults.Width + cmdSearch.Left - CmdPos
    lstResults.Left = tabStrip.Left
    If tabStrip.Width > (Margin * 4) Then
        fraName.Width = tabStrip.Width - ((Margin * 2) + (Margin * 2))
        fraDate.Width = tabStrip.Width - ((Margin * 2) + (Margin * 2))
    End If
End If

If frmSearch.Height < m_Minheight Then
    frmSearch.Height = m_Minheight
    Exit Sub
Else
    On Error Resume Next
    lstResults.Top = Margin + tabStrip.Height + (Margin)
    lstResults.Height = frmSearch.Height - lstResults.Top - (Margin * 5)
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmSearch = Nothing
End Sub

Private Sub tabStrip_Click()

If tabStrip.SelectedItem.Index = 1 Then
    tabStrip.ZOrder 1
    fraName.Visible = True
    fraDate.Visible = False
ElseIf tabStrip.SelectedItem.Index = 2 Then
    tabStrip.ZOrder 1
    fraName.Visible = False
    fraDate.Visible = True
End If

End Sub


