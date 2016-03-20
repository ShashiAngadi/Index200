VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCLReport 
   Caption         =   "Clearing Report"
   ClientHeight    =   5325
   ClientLeft      =   2295
   ClientTop       =   2055
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   765
      Left            =   30
      TabIndex        =   2
      Top             =   4530
      Width           =   6495
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&Web view"
         Height          =   400
         Left            =   2220
         TabIndex        =   5
         Top             =   270
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   400
         Left            =   4680
         TabIndex        =   4
         Top             =   270
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   400
         Left            =   3420
         TabIndex        =   3
         Top             =   270
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   3795
      Left            =   90
      TabIndex        =   0
      Top             =   630
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   6694
      _Version        =   393216
   End
   Begin VB.Label lblReportTitle 
      Caption         =   "Report Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1590
      TabIndex        =   1
      Top             =   150
      Width           =   3975
   End
End
Attribute VB_Name = "frmCLReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_FromDate As String
Dim m_ToDate As String
Private WithEvents m_grdPrint As WISPrint
Attribute m_grdPrint.VB_VarHelpID = -1
Private m_TotalCount As Long
Private m_frmCancel As frmCancel
Private m_ReportType As wis_ClearingReports
Public Event WindowClosed()

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

cmdClose.Caption = GetResourceString(11)   '"Œâ³ú"  caption of ok converted to close
cmdPrint.Caption = GetResourceString(23)

End Sub

Private Sub InitGrid(Optional Resize As Boolean)
Dim ColCount As Integer
    For ColCount = 0 To grd.Cols - 1
        grd.ColWidth(ColCount) = GetSetting(App.EXEName, lblReportTitle.Caption, "ColWidth" & ColCount, 1 / grd.Cols) * grd.Width
    Next ColCount
   If grd.Cols = 1 Then grd.ColWidth(0) = grd.Width * 0.9
If Resize Then Exit Sub
    
    Dim ColWid As Single
    Dim count As Long
If Not Resize Then
    grd.Clear
    grd.Rows = 50
    grd.Cols = 2
    grd.FixedCols = 0
    grd.AllowUserResizing = flexResizeBoth
Else
    With grd
        .AllowUserResizing = flexResizeBoth
        .Row = 0
        ColWid = .Width / .Cols
        .Col = 0
        If .Text = GetResourceString(33) Then
            .ColWidth(0) = .Width * 0.1
            ColWid = (.Width - .ColWidth(0)) / (.Cols - 1)
        End If
    End With
End If


    If m_ReportType = repChqSubmitted Then     'Cheques Submitted
        With grd
            If Resize Then
                If .Cols = 1 Then
                    .ColWidth(0) = Me.Width
                    GoTo LastLine
                End If
                .ColWidth(0) = ColWid * 0.45
                .ColWidth(1) = ColWid * 0.75
                .ColWidth(2) = ColWid * 1.45
                .ColWidth(3) = ColWid * 1
                GoTo LastLine
            End If
            .Cols = 4
            .FixedCols = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(33) 'SLNO
            .Col = 1: .Text = GetResourceString(36) 'Account No
            .Col = 2: .Text = GetResourceString(275, 140) '" Cheque No " '
            .Col = 3: .Text = GetResourceString(42) '"Balance"
        End With
    ElseIf m_ReportType = repChqCleared Then    'Cheques cleared
        With grd
            If Resize Then
                If .Cols = 1 Then
                    .ColWidth(0) = Me.Width
                    GoTo LastLine
                End If
                .ColWidth(0) = ColWid * 0.45
                .ColWidth(1) = ColWid * 0.75
                .ColWidth(2) = ColWid * 1.45
                .ColWidth(3) = ColWid * 1
                
                GoTo LastLine
            End If
            .Cols = 4
            .FixedCols = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(33)
            .Col = 1: .Text = GetResourceString(36) '"Acc No"
            .Col = 2: .Text = GetResourceString(275, 140) '" Cheque No " '
            .Col = 3: .Text = GetResourceString(280) '"Balance"
        End With
    ElseIf m_ReportType = repChqNotCleared Then    'Cheques Bounced
        With grd
            If Resize Then
                If .Cols = 1 Then
                    .ColWidth(0) = Me.Width
                    GoTo LastLine
                End If
                .ColWidth(0) = ColWid * 0.45
                .ColWidth(1) = ColWid * 0.75
                .ColWidth(2) = ColWid * 1.45
                .ColWidth(3) = ColWid * 1
                
                GoTo LastLine
            End If
            .Cols = 4
            .FixedCols = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(33)
            .Col = 1: .Text = GetResourceString(36) '"Acc No"
            .Col = 2: .Text = GetResourceString(275, 140) '" Cheque No " '
            .Col = 3: .Text = GetResourceString(280) '"Balance"
        End With
    
    ElseIf m_ReportType = repChqPending Then       'Cheques Pending
        With grd
            If Resize Then
                If .Cols = 1 Then
                    .ColWidth(0) = Me.Width
                    GoTo LastLine
                End If
                .ColWidth(0) = ColWid * 0.45
                .ColWidth(1) = ColWid * 0.75
                .ColWidth(2) = ColWid * 1.45
                .ColWidth(3) = ColWid * 1
                
                GoTo LastLine
            End If
            .Cols = 4
            .FixedCols = 1
            .Row = 0
            .Col = 0: .Text = GetResourceString(33)
            .Col = 1: .Text = GetResourceString(36) '"Acc No"
            .Col = 2: .Text = GetResourceString(275, 140) '" Cheque No " '
            .Col = 3: .Text = GetResourceString(280) '"Balance"
        End With
   End If

    grd.Row = 0:
    For count = 0 To grd.Cols - 1
      grd.Col = count
      grd.CellAlignment = 4:  grd.CellFontBold = True
    Next
    Exit Sub

LastLine:
    ColWid = 0
    For count = 0 To grd.Cols - 2
        ColWid = ColWid + grd.ColWidth(count)
    Next count
    grd.ColWidth(grd.Cols - 1) = grd.Width - ColWid - grd.Width * 0.04 'Me.ScaleWidth * 0.03
    
End Sub

Private Sub ShowSubmittedCheques()
'Trap error handler
On Error GoTo ErrLine
'Declare variables
Dim SqlStmt  As String
Dim rst As Recordset


SqlStmt = "SELECT * from ClearingTab "
SqlStmt = "SELECT A.*, E.Status,E.TransDate from ClearingTab A Inner Join " & _
" (" & _
" Select B.ChequeID, Status,TransDate from ClearingTrans B Inner Join " & _
" (SELECT max(TransID) as MaxTransID,ChequeID from ClearingTrans group by ChequeID)" & _
" C On B.ChequeID = C.ChequeID and C.MaxTransId = B.TransID" & _
" )   E On A.ChequeID = E.ChequeID "

If m_FromDate <> "" And m_ToDate <> "" Then
        SqlStmt = SqlStmt & " where TransDate between #" & _
                    GetSysFormatDate(m_FromDate) & "# and #" & _
                    GetSysFormatDate(m_ToDate) & "#"
ElseIf m_FromDate = "" And m_ToDate <> "" Then
   SqlStmt = SqlStmt & " where TransDate  <= #" & GetSysFormatDate(m_ToDate) & "#"
ElseIf m_FromDate <> "" And m_ToDate = "" Then
   SqlStmt = SqlStmt & " where TransDate  >= #" & GetSysFormatDate(m_FromDate) & "#"
ElseIf m_FromDate = "" And m_ToDate = "" Then
   SqlStmt = SqlStmt & " where TransDate   <= #" & gStrDate & "#"
End If

gDbTrans.SqlStmt = SqlStmt & " order by AccID Asc "
If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub
    
'Initialize the grid

Call InitGrid
grd.ColAlignment(0) = 1
grd.ColAlignment(1) = 1
grd.ColAlignment(2) = 0
grd.ColAlignment(3) = 2

Dim SlNo As Long
Dim rowno As Integer
rowno = grd.Row
'Fill the grid
    While Not rst.EOF
        With grd
            'Set next row
            If .Rows = rowno + 1 Then .Rows = .Rows + 1
            rowno = rowno + 1: SlNo = SlNo + 1
            
            .TextMatrix(rowno, 0) = " " & Format(SlNo, "00")
            .TextMatrix(rowno, 1) = " " & FormatField(rst("AccID"))
            .TextMatrix(rowno, 2) = " " & FormatField(rst("ChequeNo"))
            .TextMatrix(rowno, 3) = " " & FormatField(rst("Amount"))
        End With
nextRecord:
        rst.MoveNext
    
    Wend

ExitLine:
   Exit Sub

ErrLine:
   
   If Err Then
      MsgBox "Cheques submitted" & vbCrLf & Err.Description, vbCritical, "ERROR"
      GoTo ExitLine
   End If
End Sub
Private Sub ShowClearedCheques()
'Trap error handler
On Error GoTo ErrLine
'Declare variables
Dim SqlStmt  As String
Dim rst As Recordset
Dim Status As wis_ChequeStatus

Status = wisCleared

SqlStmt = "SELECT A.*, E.Status,E.TransDate from ClearingTab A Inner Join " & _
" (" & _
" Select B.ChequeID, Status,TransDate from ClearingTrans B Inner Join " & _
" (SELECT max(TransID) as MaxTransID,ChequeID from ClearingTrans group by ChequeID)" & _
" C On B.ChequeID = C.ChequeID and C.MaxTransId = B.TransID" & _
" )   E On A.ChequeID = E.ChequeID "


'SQLStmt = "SELECT * from ClearingTab  "

If m_FromDate <> "" And m_ToDate <> "" Then
        SqlStmt = SqlStmt & " where TransDate between #" & _
                    GetSysFormatDate(m_FromDate) & "# and #" & _
                    GetSysFormatDate(m_ToDate) & "#"
ElseIf m_FromDate = "" And m_ToDate <> "" Then
   SqlStmt = SqlStmt & " where TransDate   <= #" & GetSysFormatDate(m_ToDate) & "# AND Status = " & Status
ElseIf m_FromDate <> "" And m_ToDate = "" Then
   SqlStmt = SqlStmt & " where TransDate  >= #" & GetSysFormatDate(m_FromDate) & "#  AND Status = " & Status
ElseIf m_FromDate = "" And m_ToDate = "" Then
   SqlStmt = SqlStmt & " where TransDate   <= #" & gStrDate & "# AND Status = " & Status
End If
gDbTrans.SqlStmt = SqlStmt '& "  order by AccID Asc "

If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub
    
'Initialize the grid
Call InitGrid
grd.ColAlignment(0) = 1
grd.ColAlignment(1) = 1
grd.ColAlignment(2) = 0
grd.ColAlignment(3) = 2

Dim SlNo As Long
Dim rowno As Integer
rowno = grd.Row
'Fill the grid
While Not rst.EOF
    With grd
        'Set next row
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1: SlNo = SlNo + 1
        
        .TextMatrix(rowno, 0) = " " & Format(SlNo, "00")
        .TextMatrix(rowno, 1) = " " & FormatField(rst("AccID"))
        .TextMatrix(rowno, 2) = " " & FormatField(rst("ChequeNo"))
        .TextMatrix(rowno, 3) = " " & FormatField(rst("Amount"))
    End With

nextRecord:
     rst.MoveNext
Wend

ExitLine:
   Exit Sub

ErrLine:
   If Err Then
      MsgBox "Cheques Cleared" & vbCrLf & Err.Description, vbCritical, "ERROR"
      GoTo ExitLine
   End If

End Sub
Private Sub ShowBouncedCheques()
'Trap error handler
On Error GoTo ErrLine
'Declare variables
Dim SqlStmt  As String
Dim rst As Recordset
Dim Status As wis_ChequeStatus

Status = wisBounced
SqlStmt = " SELECT * from ClearingTab "
SqlStmt = "SELECT A.*, E.Status,E.TransDate from ClearingTab A Inner Join " & _
" (" & _
" Select B.ChequeID, Status,TransDate from ClearingTrans B Inner Join " & _
" (SELECT max(TransID) as MaxTransID,ChequeID from ClearingTrans group by ChequeID)" & _
" C On B.ChequeID = C.ChequeID and C.MaxTransId = B.TransID" & _
" )   E On A.ChequeID = E.ChequeID "


If m_FromDate <> "" And m_ToDate <> "" Then
        SqlStmt = SqlStmt & " where TransDate between #" & _
                    GetSysFormatDate(m_FromDate) & "# and #" & _
                    GetSysFormatDate(m_ToDate) & "# AND Status = " & Status
ElseIf m_FromDate = "" And m_ToDate <> "" Then
   SqlStmt = SqlStmt & " where TransDate   <= #" & GetSysFormatDate(m_ToDate) & "# AND Status = " & Status
ElseIf m_FromDate <> "" And m_ToDate = "" Then
   SqlStmt = SqlStmt & " where TransDate  >= #" & GetSysFormatDate(m_FromDate) & "#  AND Status = " & Status
ElseIf m_FromDate = "" And m_ToDate = "" Then
   SqlStmt = SqlStmt & " where TransDate   <= # " & gStrDate & " # AND Status = " & Status
End If
   gDbTrans.SqlStmt = SqlStmt & " order by AccID Asc "
If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub
    
'Initialize the grid
Call InitGrid
grd.ColAlignment(0) = 1
grd.ColAlignment(1) = 1
grd.ColAlignment(2) = 0
grd.ColAlignment(3) = 2

Dim SlNo As Integer
Dim rowno As Integer
rowno = grd.Row
'Fill the grid
While Not rst.EOF
    With grd
        'Set next row
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1: SlNo = SlNo + 1
        
        .TextMatrix(rowno, 0) = " " & Format(SlNo, "00")
        .TextMatrix(rowno, 1) = " " & FormatField(rst("AccID"))
        .TextMatrix(rowno, 2) = " " & FormatField(rst("ChequeNo"))
        .TextMatrix(rowno, 3) = " " & FormatField(rst("Amount"))
    End With
    
nextRecord:
     rst.MoveNext
Wend

ExitLine:
   Exit Sub

ErrLine:
   If Err Then
      MsgBox "Cheques Bounced" & vbCrLf & Err.Description, vbCritical, "ERROR"
      GoTo ExitLine
   End If

End Sub

Private Sub ShowPendingCheques()
'Trap error handler
On Error GoTo ErrLine
'Declare variables
Dim SqlStmt  As String
Dim rst As Recordset
Dim Status As wis_ChequeStatus

Status = wisPending
SqlStmt = " SELECT * from ClearingTab "
SqlStmt = "SELECT A.*, E.Status,E.TransDate from ClearingTab A Inner Join " & _
" (" & _
" Select B.ChequeID, Status,TransDate from ClearingTrans B Inner Join " & _
" (SELECT max(TransID) as MaxTransID,ChequeID from ClearingTrans group by ChequeID)" & _
" C On B.ChequeID = C.ChequeID and C.MaxTransId = B.TransID" & _
" )   E On A.ChequeID = E.ChequeID "

If m_FromDate <> "" And m_ToDate <> "" Then
        SqlStmt = SqlStmt & " where TransDate between #" & _
                    GetSysFormatDate(m_FromDate) & "# and #" & _
                    GetSysFormatDate(m_ToDate) & "# AND Status = " & Status
ElseIf m_FromDate = "" And m_ToDate <> "" Then
   SqlStmt = SqlStmt & " where TransDate   <= #" & GetSysFormatDate(m_ToDate) & "# AND Status = " & Status
ElseIf m_FromDate <> "" And m_ToDate = "" Then
   SqlStmt = SqlStmt & " where TransDate  >= #" & GetSysFormatDate(m_FromDate) & "#  AND Status = " & Status
ElseIf m_FromDate = "" And m_ToDate = "" Then
   SqlStmt = SqlStmt & " where TransDate   <= # " & gStrDate & " # AND Status = " & Status
End If
   gDbTrans.SqlStmt = SqlStmt & "  order by AccID Asc "
If gDbTrans.Fetch(rst, adOpenDynamic) <= 0 Then Exit Sub
    
'Initialize the grid
Call InitGrid
grd.ColAlignment(0) = 1
grd.ColAlignment(1) = 1
grd.ColAlignment(2) = 0
grd.ColAlignment(3) = 2

Dim SlNo As Long
Dim rowno As Integer
rowno = grd.Row
'Fill the grid
While Not rst.EOF
    With grd
        'Set next row
        If .Rows = rowno + 1 Then .Rows = .Rows + 1
        rowno = rowno + 1: SlNo = SlNo + 1
        
        .TextMatrix(rowno, 0) = " " & Format(SlNo, "00")
        .TextMatrix(rowno, 1) = " " & FormatField(rst("AccID"))
        .TextMatrix(rowno, 2) = " " & FormatField(rst("ChequeNo"))
        .TextMatrix(rowno, 3) = " " & FormatField(rst("Amount"))
    End With
nextRecord:
     rst.MoveNext
Wend

ExitLine:
   Exit Sub

ErrLine:
   If Err Then
      MsgBox "Cheques Cleared" & vbCrLf & Err.Description, vbCritical, "ERROR"
      GoTo ExitLine
   End If

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
' Call the print class services...
Set m_grdPrint = wisMain.grdPrint
With m_grdPrint
    .Font.name = gFontName
    .Font.Size = gFontSize
    .CompanyName = gCompanyName
    .GridObject = grd
    .ReportTitle = lblReportTitle.Caption
    .PrintGrid
End With

End Sub


Private Sub cmdWeb_Click()
Dim clswebGrid As New clsgrdWeb
With clswebGrid
    Set .GridObject = grd
    .CompanyAddress = ""
    .CompanyName = gCompanyName
    .ReportTitle = lblReportTitle
    Call clswebGrid.ShowWebView '(grd)

End With

End Sub

Private Sub Form_Click()
Call grd_LostFocus
End Sub

Private Sub Form_Load()
 
'Center the form
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
'set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)
Call SetKannadaCaption
'Init the grid
grd.Rows = 50
grd.Cols = 1
grd.FixedCols = 0
grd.Row = 1

grd.Text = GetResourceString(278)   '"No Records Available"
grd.CellAlignment = 4: grd.CellFontBold = True
Me.lblReportTitle.FONTSIZE = 16
Screen.MousePointer = vbHourglass

      If m_ReportType = repChqSubmitted Then
          Me.lblReportTitle.Caption = GetResourceString(471)
           Call ShowSubmittedCheques
      ElseIf m_ReportType = repChqCleared Then
         Me.lblReportTitle.Caption = GetResourceString(469)
         Call ShowClearedCheques
      ElseIf m_ReportType = repChqNotCleared Then
         Me.lblReportTitle.Caption = GetResourceString(472) '"Cheques bounced"
         Call ShowBouncedCheques
      ElseIf m_ReportType = repChqPending Then
         Me.lblReportTitle.Caption = GetResourceString(473) '"Cheques  Pending "
         Call ShowPendingCheques
      End If

Screen.MousePointer = vbNormal
End Sub


Private Sub Form_Resize()

    Screen.MousePointer = vbDefault
    On Error Resume Next
    lblReportTitle.Top = 0
    lblReportTitle.Left = (Me.Width - lblReportTitle.Width) / 2
    grd.Left = 0
    grd.Top = lblReportTitle.Top + lblReportTitle.Height
    grd.Width = Me.Width - 150
    fra.Top = Me.ScaleHeight - fra.Height
    fra.Left = Me.Width - fra.Width
    grd.Height = Me.ScaleHeight - fra.Height - lblReportTitle.Height
    cmdWeb.Top = cmdPrint.Top
    cmdWeb.Left = cmdPrint.Left - cmdPrint.Width - (cmdPrint.Width / 4)

    Call InitGrid(True)

End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmCAReport = Nothing
   
End Sub


Private Sub grd_LostFocus()
Dim ColCount As Integer
    For ColCount = 0 To grd.Cols - 1
        Call SaveSetting(App.EXEName, Me.lblReportTitle.Caption, _
                "ColWidth" & ColCount, grd.ColWidth(ColCount) / grd.Width)
    Next ColCount

End Sub


Private Sub m_grdPrint_MaxProcessCount(MaxCount As Long)
m_TotalCount = MaxCount
Set m_frmCancel = New frmCancel
m_frmCancel.PicStatus.Visible = True
m_frmCancel.PicStatus.ZOrder 0

End Sub

Private Sub m_grdPrint_Message(strMessage As String)
m_frmCancel.lblMessage = strMessage
End Sub


Private Sub m_grdPrint_ProcessCount(count As Long)
On Error Resume Next

If (count / m_TotalCount) > 0.95 Then
    Unload m_frmCancel
    Exit Sub
End If
UpdateStatus m_frmCancel.PicStatus, count / m_TotalCount
Err.Clear

End Sub
Public Property Let fromDate(NewValue As String)
    m_FromDate = NewValue
End Property
Public Property Let toDate(NewValue As String)
    m_ToDate = NewValue
End Property


Public Property Let ReportType(NewValue As wis_ClearingReports)
    m_ReportType = NewValue
End Property
