VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGrid 
   Caption         =   "Grid Form"
   ClientHeight    =   5805
   ClientLeft      =   2250
   ClientTop       =   1935
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   6105
   Begin VB.Frame fra 
      Height          =   795
      Left            =   60
      TabIndex        =   1
      Top             =   4980
      Width           =   5895
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&Web View"
         Height          =   400
         Left            =   180
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   400
         Left            =   1665
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   4620
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   3135
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4875
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8599
      _Version        =   393216
   End
End
Attribute VB_Name = "frmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event OkClicked()
Public Event CancelClicked()
Private m_Selected() As Boolean
Private m_selCol As Integer


Public Property Let Selected(Index As Integer, NewValue As Boolean)

If m_selCol <= 0 Then Exit Property
If Index >= grd.Rows Then Exit Property

With grd
    If UBound(m_Selected) < .Rows Then ReDim Preserve m_Selected(.Rows - 1)
    .Col = m_selCol
    .Row = Index
    If NewValue Then
        Set .CellPicture = LoadResPicture(143, vbResIcon)
    Else
        Set .CellPicture = LoadResPicture(144, vbResIcon)
    End If
    m_Selected(.Row) = NewValue
    
End With

End Property

Public Property Get Selected(Index As Integer) As Boolean
    
    If m_selCol <= 0 Then Exit Property
    If Index > UBound(m_Selected) Then Exit Property
    Selected = m_Selected(Index)

End Property

Public Property Let SelectionColoumn(NewValue As Integer)
    grd.Row = grd.Row
    m_selCol = NewValue
End Property




Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

cmdOk.Caption = GetResourceString(1)
cmdCancel.Caption = GetResourceString(2)
cmdPrint.Caption = GetResourceString(23)

Err.Clear

End Sub

Private Sub cmdCancel_Click()
RaiseEvent CancelClicked
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim MaxI As Integer
Dim I As Integer
With grd
    MaxI = UBound(m_Selected)
    For I = 1 To MaxI
        If Not m_Selected(I) Then .RowData(I) = 0
    Next
End With
    
RaiseEvent OkClicked

Me.Hide

End Sub

Private Sub Form_Load()
Call CenterMe(Me)

Call SetKannadaCaption

ReDim m_Selected(grd.Rows - 1)

grd.AllowUserResizing = flexResizeBoth

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Debug.Print " UnloadMode = " & UnloadMode
RaiseEvent CancelClicked
End Sub


Private Sub Form_Resize()

Screen.MousePointer = vbDefault
On Error Resume Next

With grd
    .Left = 0
    .Top = 150 'fra.Height
    .Width = Me.Width - 150
    .Height = Height - fra.Height * 2
End With
With fra
    .Top = Me.ScaleHeight - fra.Height
    .Left = 0
    .Width = Me.Width - (2 * fra.Left)
End With

cmdCancel.Left = fra.Width - cmdCancel.Width - (cmdCancel.Width / 4)
cmdOk.Left = cmdCancel.Left - cmdOk.Width - (cmdOk.Width / 4)
cmdPrint.Left = cmdOk.Left - cmdPrint.Width - (cmdPrint.Width / 4)
cmdWeb.Left = cmdPrint.Left - cmdWeb.Width - (cmdWeb.Width / 4)

Dim I As Integer
Dim ColWid As Single
For I = 0 To grd.Cols - 1
    ColWid = GetSetting(App.EXEName, "DAYEND", _
        "ColWidth" & I, 1 / grd.Cols) * grd.Width
    If ColWid < 10 Or ColWid > grd.Width * 0.9 Then ColWid = grd.Width / grd.Cols
    grd.ColWidth(I) = ColWid
Next I


End Sub

Private Sub grd_Click()

If m_selCol <= 0 Then Exit Sub

With grd
    If .Row < .FixedRows Then Exit Sub
    If .Col <> m_selCol Then Exit Sub
    
    If UBound(m_Selected) < .Rows Then ReDim Preserve m_Selected(.Rows - 1)
    
    If m_Selected(.Row) Then
        Set .CellPicture = LoadResPicture(143, vbResIcon)
    Else
        Set .CellPicture = LoadResPicture(144, vbResIcon)
    End If
    m_Selected(.Row) = Not m_Selected(.Row)
End With

End Sub


Private Sub grd_EnterCell()

If grd.Col = grd.Cols - 1 Then Call grd_Click

End Sub


Private Sub grd_LostFocus()

Dim ColCount As Integer

For ColCount = 0 To grd.Cols - 1
    Call SaveSetting(App.EXEName, "DayEnd", _
            "ColWidth" & ColCount, grd.ColWidth(ColCount) / grd.Width)
Next ColCount

End Sub


