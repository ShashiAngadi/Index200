VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReportview 
   Caption         =   "Reccuring Deposits Report.."
   ClientHeight    =   6810
   ClientLeft      =   1350
   ClientTop       =   1335
   ClientWidth     =   7470
   LinkTopic       =   "Form2"
   ScaleHeight     =   6810
   ScaleWidth      =   7470
   Begin VB.Frame FraDate 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   7200
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&Web View"
         Height          =   450
         Left            =   3240
         TabIndex        =   6
         Top             =   330
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   450
         Left            =   4860
         TabIndex        =   4
         Top             =   270
         Width           =   1215
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   450
         Left            =   6090
         TabIndex        =   2
         Top             =   255
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   735
      Left            =   -105
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   1296
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   5430
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9578
      _Version        =   393216
   End
   Begin VB.Label lblReportTitle 
      AutoSize        =   -1  'True
      Caption         =   " Report Title "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3060
      TabIndex        =   5
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmReportview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Processing(StrMessge As String, Ratio As Single)
Public Event Initialize(Min As Long, Max As Long)

Private m_FromIndianDate As String
Private m_ToIndianDate As String

Private m_FromDate As Date
Private m_ToDate As Date
Private m_Caste As String
Private m_Place As String
Private m_Gender As wis_Gender
Private m_ReportOrder As wis_ReportOrder

Private WithEvents m_grdPrint As WISPrint
Attribute m_grdPrint.VB_VarHelpID = -1
Private m_TotalCount As Long
Private m_frmCancel As frmCancel

Public ReportType As wisReports

'Declare All Classe of this project
    Private MatClass As clsMaterial
    Private MemClass As clsMMAcc
    Private SBClass As clsSBAcc
    Private CAClass As ClsCAAcc
    Private FdClass As clsFDAcc
    Private RDClass As clsRDAcc
    Private PDClass As clsPDAcc
    'Private DLClass As clsDLAcc

'Declare Varibles of class which are under Development
Private loanClass As clsLoan
'Private UtilClass As clsUtils
Private bankClass As clsBankAcc
Private SetupClass As clsSetup

Private M_SlNo As Long
'Private m_ShowCompact As Boolean
'Private m_Final As Boolean
Private m_ProcessCount As Integer
Private m_TotalProcess As Integer


Public Property Let Gender(NewValue As wis_Gender)
m_Gender = NewValue
End Property


Public Property Let Place(NewValue As String)
m_Place = NewValue
End Property


Public Property Let Caste(NewValue As String)
m_Caste = NewValue
End Property

Public Property Let ReportOrder(NewValue As wis_ReportOrder)
m_ReportOrder = NewValue
End Property

Private Function ShowCreditTrans()

Dim rstHead As Recordset
Dim rstTrans As Recordset
Dim headName As String
Dim headID As Long

Dim loopCount As Integer
Dim Total As Currency
    
RaiseEvent Processing("Initailiasing Trans List", m_ProcessCount / m_TotalProcess)

If gCancel Then Exit Function
Screen.MousePointer = vbHourglass
RaiseEvent Processing("Initailising Credit Trans List", 2.5 / 100)

'Now Set ths Date
FromIndianDate = "1/" & Month(m_ToDate) & "/" & Year(m_ToDate)
ToIndianDate = GetIndianDate(DateAdd("d", -1, DateAdd("M", 1, m_FromDate)))

''Get the Details OF Heads
gDbTrans.SqlStmt = "SELECT HeadID,HeadName From Heads " & _
            " Order By HeadID"
If gDbTrans.Fetch(rstHead, adOpenDynamic) < 1 Then Exit Function

''Get the Details OF Trans
gDbTrans.SqlStmt = "SELECT SUM(Credit) as CreditAmount,TransDate,HeadID " & _
                " From AccTrans Where TransDate >= #" & m_FromDate & "#" & _
                " AND TransDate <= #" & m_ToDate & "#" & _
                " Group BY HeadID,TransDate " & _
                " Order By HeadID,TransDate"
 'And Headid <> " & wis_CashHeadID
 
If gDbTrans.Fetch(rstTrans, adOpenDynamic) <= 0 Then Exit Function


With grd
    .Clear
    .Rows = DateDiff("d", m_FromDate, m_ToDate) + 4 '+ 2
            'No of days + (1title+1OB+1gap+1datediff) + (1gap+1total)
    .Cols = 2
    .FixedCols = 1
    .FixedRows = 1
    .CellFontBold = True
    .Row = 0
    .Col = 0: .Text = GetResourceString(37)
    .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .Text = GetResourceString(285)  'Closing Blance
    .CellAlignment = 4: .CellFontBold = True
    'Here we have to show the Closing Balanced
    .Row = 1: .Text = GetOpeningBalance(GetIndianDate(DateAdd("d", 1, m_ToDate)))
    .CellFontBold = True
    '.MergeCells = flexMergeFree
End With

m_ProcessCount = m_TotalProcess * 0.05
'Frmcancel
DoEvents
If gCancel Then Exit Function

Dim FstRow As Integer
Dim rowno  As Integer
Dim TransDate As Date

FstRow = grd.FixedRows
'First Count the Transaction of Member Share
DoEvents
If gCancel Then Exit Function

m_ProcessCount = 1
m_TotalProcess = rstTrans.RecordCount + 3

While Not rstTrans.EOF
    If headID <> rstTrans("HeadID") Then
        If headID Then
            If headID = wis_CashHeadID Then Total = 0
            With grd
                'Put the Grand TOtal At the end
                .Col = .Cols - 1
                .Row = .Rows - 1
                .CellFontBold = True
                .Text = FormatCurrency(Total)
                If Total = 0 Then .Cols = .Cols - 1
            End With
        End If
        
        headName = "": Total = 0
        headID = rstTrans("HeadID")
        rstHead.Find "HeadID = " & headID
        If Not rstHead.EOF Then headName = FormatField(rstHead("HeadName"))
        With grd
            .Cols = .Cols + 1
            .Col = .Cols - 1
            .Row = 0
            .Text = headName
        End With
    End If
    
    TransDate = rstTrans("TransDate")
    rowno = FstRow + DateDiff("D", m_FromDate, TransDate)
    With grd
        .Row = rowno
        .Text = FormatField(rstTrans("CreditAmount"))
        Total = Total + Val(.Text)
    End With
    
    If gCancel Then rstTrans.MoveLast
    
    rstTrans.MoveNext
    m_ProcessCount = m_ProcessCount + 1
    RaiseEvent Processing("Calculating Credits of " & headName, m_ProcessCount / m_TotalProcess)
    DoEvents

Wend

'Now Put the Total Of last head
With grd
    'Put the Grand TOtal At the end
    .Col = .Cols - 1
    .Row = .Rows - 1
    .CellFontBold = True
    .Text = FormatCurrency(Total)
    If Total = 0 Then .Cols = .Cols - 1
End With


'Now Put the Date and If no transaction on any Date
'then remove that row or else put the Horizontal row
Dim Retval As Integer
RaiseEvent Processing("Calculating Horizontal Total ", m_ProcessCount / m_TotalProcess)

With grd
    .Cols = .Cols + 1
    .CellFontBold = True
    .Row = 0
    .Col = .Cols - 1
    .Text = GetResourceString(52, 42)
    .CellFontBold = True: .CellAlignment = 4

    loopCount = DateDiff("d", m_FromDate, m_ToDate) + 1
    TransDate = DateAdd("d", -1, m_FromDate)
    rowno = .FixedRows
    Do
        .Row = rowno
        .Col = 0
        TransDate = DateAdd("d", 1, TransDate)
        .Text = GetIndianDate(TransDate)
        Total = 0
        For Retval = .FixedCols To .Cols - 2
            .Col = Retval
            Total = Total + Val(.Text)
        Next
        If Total Then
            .Col = .Cols - 1
            .Text = FormatCurrency(Total)
            .CellFontBold = True
        Else
            .RemoveItem rowno
            loopCount = loopCount - 1
            rowno = rowno - 1
        End If
        rowno = rowno + 1
        If TransDate >= m_ToDate Then Exit Do
    Loop
End With

'Now Put the Total
RaiseEvent Processing("Calculating Verticale Total ", m_ProcessCount / m_TotalProcess)
With grd
    .Row = .Rows - 1
    .Col = 1
    .Text = GetResourceString(52, 42)
    .RowHeight(0) = .RowHeight(0) * 3.5
    .WordWrap = True
End With

lblReportTitle = GetMonthString(Month(m_FromDate)) & " " & _
                GetResourceString(272, 283)

ExitLine:
Screen.MousePointer = vbDefault

End Function


Private Function ShowDebitTrans()
Dim rstHead As Recordset
Dim rstTrans As Recordset
Dim headName As String
Dim headID As Long
Dim loopCount As Integer
Dim Total As Currency
    

If gCancel Then Exit Function
Screen.MousePointer = vbHourglass

''Get the Details OF Head Name
gDbTrans.SqlStmt = "SELECT HeadID,HeadName From Heads " & _
            " Order By HeadID"
If gDbTrans.Fetch(rstHead, adOpenDynamic) < 1 Then Exit Function

''Get the Details OF Trans
'Now Set ths Date
FromIndianDate = "1/" & Month(m_ToDate) & "/" & Year(m_ToDate)
ToIndianDate = GetIndianDate(DateAdd("d", -1, DateAdd("M", 1, m_FromDate)))

gDbTrans.SqlStmt = "SELECT SUM(Debit) as DebitAmount,TransDate,HeadID " & _
            " From AccTrans Where TransDate >= #" & m_FromDate & "#" & _
            " AND TransDate <= #" & m_ToDate & "#" & _
            " Group BY HeadID,TransDate " & _
            " Order By HeadID,TransDate"
 'And Headid <> " & wis_CashHeadID
 
If gDbTrans.Fetch(rstTrans, adOpenDynamic) <= 0 Then Exit Function

RaiseEvent Processing("Initilising Trans List", 1 / 10)

With grd
    .Clear
    .Rows = DateDiff("d", m_FromDate, m_ToDate) + 4 '+ 2
            'No of days + (1title+1OB+1gap+1datediff) + (1gap+1total)
    .Cols = 2
    .FixedCols = 1
    .FixedRows = 1
    .CellFontBold = True
    .Row = 0
    .Col = 0: .Text = GetResourceString(37)
    .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .Text = GetResourceString(285)  'Closing Blance
    .CellAlignment = 4: .CellFontBold = True
    'Here we have to show the Closing Balanced
    .Row = 1: .Text = GetOpeningBalance(GetIndianDate(DateAdd("d", 1, m_ToDate)))
    .CellFontBold = True
    '.MergeCells = flexMergeFree
End With

m_ProcessCount = m_TotalProcess * 0.05
'Frmcancel
DoEvents
If gCancel Then Exit Function

Dim FstRow As Integer
Dim rowno  As Integer
Dim TransDate As Date

FstRow = grd.FixedRows
'First Count the Transaction of Member Share
DoEvents
If gCancel Then Exit Function

m_TotalProcess = rstTrans.RecordCount + 3
m_ProcessCount = 1

While Not rstTrans.EOF
    If headID <> rstTrans("HeadID") Then
        If headID Then
            If headID = wis_CashHeadID Then Total = 0
            With grd
                'Put the Grand TOtal At the end
                .Col = .Cols - 1
                .Row = .Rows - 1
                .CellFontBold = True
                .Text = FormatCurrency(Total)
                'If Total = 0 Then .Cols = .Cols - 1
                
            End With
        End If
        
        
        headName = "": Total = 0
        headID = rstTrans("HeadID")
        rstHead.Find "HeadID = " & headID
        If Not rstHead.EOF Then headName = FormatField(rstHead("HeadName"))
        With grd
            .Cols = .Cols + 1
            .Col = .Cols - 1
            .Row = 0
            .Text = headName
        End With
        
        m_ProcessCount = m_ProcessCount + 1
        RaiseEvent Processing("Calculating Debits of " & headName, m_ProcessCount / m_TotalProcess)
        DoEvents
        If gCancel Then Exit Function
    End If
    
    TransDate = rstTrans("TransDate")
    rowno = FstRow + DateDiff("D", m_FromDate, TransDate)
    With grd
        .Row = rowno
        .Text = FormatField(rstTrans("DebitAmount"))
        Total = Total + Val(.Text)
    End With
    rstTrans.MoveNext
    m_ProcessCount = m_ProcessCount + 1
    RaiseEvent Processing("Calculating Debits of " & headName, m_ProcessCount / m_TotalProcess)
    
Wend

'Now Put the Total Of last head
With grd
    'Put the Grand TOtal At the end
    .Col = .Cols - 1
    .Row = .Rows - 1
    .CellFontBold = True
    .Text = FormatCurrency(Total)
    If Total = 0 Then .Cols = .Cols - 1
End With


'Now Put the Date and If no transaction on any Date
'then remove that row or else put the Horizontal row
'Dim Count As Integer
Dim Retval As Integer
RaiseEvent Processing("Calculating Horizontal Total ", m_ProcessCount / m_TotalProcess)

With grd
    .Cols = .Cols + 1
    .CellFontBold = True
    .Row = 0
    .Col = .Cols - 1
    .Text = GetResourceString(52, 42)
    .CellFontBold = True: .CellAlignment = 4

    loopCount = DateDiff("d", m_FromDate, m_ToDate) + 1
    TransDate = DateAdd("d", -1, m_FromDate)
    rowno = .FixedRows
    Do
        .Row = rowno
        .Col = 0
        TransDate = DateAdd("d", 1, TransDate)
        .Text = GetIndianDate(TransDate)
        Total = 0
        For Retval = .FixedCols To .Cols - 2
            .Col = Retval
            Total = Total + Val(.Text)
        Next
        If Total Then
            .Col = .Cols - 1
            .Text = FormatCurrency(Total)
            .CellFontBold = True
        Else
            .RemoveItem rowno
            loopCount = loopCount - 1
            rowno = rowno - 1
        End If
        rowno = rowno + 1
        If TransDate >= m_ToDate Then Exit Do
    Loop
End With

'Now Put the Total
RaiseEvent Processing("Calculating Verticale Total ", m_ProcessCount / m_TotalProcess)
With grd
    .Row = .Rows - 1
    .Col = 1
    .Text = GetResourceString(52, 42)
    .RowHeight(0) = .RowHeight(0) * 3.5
    .WordWrap = True
End With

    lblReportTitle = GetMonthString(Month(m_FromDate)) & " " & _
                GetResourceString(271, 283)

ExitLine:
Screen.MousePointer = vbDefault

End Function

Public Property Let ToIndianDate(NewStrdate As String)

    If Not DateValidate(NewStrdate, "/", True) Then
        Err.Raise 1000000, "Invalid date format"
        Exit Property
    End If
    
    m_ToIndianDate = NewStrdate
    m_ToDate = GetSysFormatDate(NewStrdate)
    'm_ToIndianDate = GetAppFormatDate(m_ToDate)

End Property

Public Property Let FromIndianDate(NewStrdate As String)

    If Not DateValidate(NewStrdate, "/", True) Then
        Err.Raise 1000000, "Invalid date format"
        Exit Property
    End If
    
    m_FromIndianDate = NewStrdate
    m_FromDate = GetSysFormatDate(NewStrdate)
    'm_FromIndianDate = GetAppFormatDate(m_FromDate)

End Property

Private Function GetOpeningBalance(AsonIndianDate As String) As Currency
GetOpeningBalance = 0
End Function

Private Sub chkDetail_Click()
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim grdPrint As WISPrint
Set grdPrint = wisMain.grdPrint
With grdPrint
    .GridObject = grd
    .ReportTitle = Me.lblReportTitle
    .CompanyName = gCompanyName
    .Font.name = grd.Font.name
    .PrintGrid
End With
'
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
'Set icon for the form caption
Screen.MousePointer = vbHourglass
Me.Icon = LoadResPicture(161, vbResIcon)
RaiseEvent Processing("Initialising ...", 0)
Call SetKannadaCaption

Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
grd.AllowUserResizing = flexResizeBoth
grdTemp.Cols = 4: grdTemp.Rows = 10

m_TotalProcess = 0
m_ProcessCount = 0

'Find the Valu to initialise the frmcancel form
Dim BankHeadsCount As Integer
Dim DepositCount As Integer
Dim LoanCount As Integer
Dim MatCount As Integer

Dim Names() As String
Dim Ids() As Long
Dim clsObj As Object
gCancel = 0

Set clsObj = New clsBankAcc
BankHeadsCount = BankHeadsCount + clsObj.MainHeads(Names, Ids)
DepositCount = 11 'Members, SB, CA, fd/loan, rd/loan, Dl/loan,Pd/Loan
Set clsObj = New clsLoan
LoanCount = clsObj.LoanList(Names, Ids, wisAgriculural)
LoanCount = LoanCount + clsObj.LoanList(Names, Ids, wisNonAgriculural)


'Before Filling  in to the grid Find the required Values
M_SlNo = 1
If ReportType = wisRepCreditTrans Then
    grd.Cols = 4
    ' Now Count the total no of processes to display on Progressbar
    m_TotalProcess = MatCount + (DepositCount + LoanCount + BankHeadsCount)
    RaiseEvent Initialize(0, CLng(m_TotalProcess * 1.25))
    
    If gCancel Then Exit Sub
    FromIndianDate = m_ToIndianDate
    
    Call ShowCreditTrans
   If gCancel = 2 Then Exit Sub

ElseIf ReportType = wisRepDebitTrans Then
    grd.Cols = 4
    ' Now Count the total no of processes to display on Progressbar
    m_TotalProcess = MatCount + (DepositCount + LoanCount + BankHeadsCount)
    RaiseEvent Initialize(0, CLng(m_TotalProcess * 1.25))
    
    If gCancel Then Exit Sub
    FromIndianDate = m_ToIndianDate
    
    Call ShowDebitTrans
   If gCancel = 2 Then Exit Sub

End If
Screen.MousePointer = vbDefault
End Sub

Private Sub InitGrid(Optional Resize As Boolean)
Dim Wid As Single
Dim colno As Integer


If Resize Then
    For colno = 0 To grd.Cols - 1
        Wid = 1 / grd.Cols
        'wid = GetSetting(App.EXEName, "GLReportType" & ReportType, "COLWIDTH" & ColNo, wid) * grd.Width
        Wid = GetSetting(App.EXEName, "GLReportType" & ReportType, "COLWIDTH" & colno, 1 / grd.Cols) * grd.Width
        If Wid < 10 Or Wid > grd.Width * 0.95 Then Wid = (grd.Width / grd.Cols)
        grd.ColWidth(colno) = Wid
    Next
    Exit Sub
End If
RaiseEvent Processing("INITIALISING GRID", 0)
On Error GoTo ErrLine
If Resize = True Then GoTo LastLine
With grd
    .Clear
    .Cols = 4
    .Rows = 5
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) '"Sl No"
    .CellAlignment = 4: .CellFontBold = True
    'For Trading Account Grid
    If ReportType = wisTradingAccount Then
        .Col = 1: .Text = GetResourceString(400) '" Materails "
        .CellAlignment = 4: .CellFontBold = True
        .Col = 2: .Text = GetResourceString(401) '" Purchase"
        .CellAlignment = 4: .CellFontBold = True
        .Col = 3: .Text = GetResourceString(226) '" Sales "
        .CellAlignment = 4: .CellFontBold = True
        lblReportTitle.Caption = GetResourceString(162, 36) & _
                GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        .ColAlignment(3) = 1
    End If
    
    'For the Debit&Credit statement
    If ReportType = wisDebitCreditStatement Then
        .Row = 0
        .Col = 1: .Text = GetResourceString(39) '" Particulars "
        .CellAlignment = 4: .CellFontBold = True
        .Col = 2: .Text = GetResourceString(276) '" Debit "
        .CellAlignment = 4: .CellFontBold = True
        .Col = 3: .Text = GetResourceString(277) '" Credit"
        .CellAlignment = 4: .CellFontBold = True
        lblReportTitle.Caption = GetResourceString(276) & " - " & GetResourceString(277) & " " & _
            GetResourceString(430) & _
            GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        .ColAlignment(3) = 1
    End If
    
    If ReportType = wisDailyDebitCredit Then
        .Cols = 6: .Row = 0
        .Col = 1: .Text = GetResourceString(39) '" Particulars "
        .CellAlignment = 4: .CellFontBold = True
        .Col = 2: .Text = GetResourceString(36)
        .CellAlignment = 4: .CellFontBold = True
        .Col = 3: .Text = GetResourceString(381) & _
            "/" & GetResourceString(270)
        .CellAlignment = 4: .CellFontBold = True
        .Col = 4: .Text = GetResourceString(276) '" Debit "
        .CellAlignment = 4: .CellFontBold = True
        .Col = 5: .Text = GetResourceString(277) '" Credit"
        .CellAlignment = 4: .CellFontBold = True
        lblReportTitle.Caption = GetResourceString(416) & GetFromDateString(m_ToIndianDate)
        .ColAlignment(3) = 1
    End If
    'For Profit & loss Statment
    If ReportType = wisProfitLossStatement Then
        .Col = 1: .Text = GetResourceString(39) '" Particulars "
        .CellAlignment = 4: .CellFontBold = True
        .Col = 2: .Text = GetResourceString(403) '" Profit "
        .CellAlignment = 4: .CellFontBold = True
        .Col = 3: .Text = GetResourceString(404) '" Loss "
        .CellAlignment = 4: .CellFontBold = True
        Me.lblReportTitle.Caption = GetResourceString(443) & _
                      " " & GetFromDateString(m_FromIndianDate, m_ToIndianDate)
        .ColAlignment(3) = 1
    End If
    
    'For Balance Sheet
    If ReportType = wisBalanceSheet Then
        .Cols = .Cols + 3
        .Col = 1: .Text = GetResourceString(405) '" Laibilities "
        .CellAlignment = 4: .CellFontBold = True
        .Col = 3: .Text = GetResourceString(40) '" Amount"
        .CellAlignment = 4: .CellFontBold = True
        .Col = 4: .Text = GetResourceString(406) '" Assets "
        .CellAlignment = 4: .CellFontBold = True
        .Col = 6: .Text = GetResourceString(40) '" amount"
        .CellAlignment = 4: .CellFontBold = True
        Me.lblReportTitle.Caption = GetResourceString(408) & " " & _
                    GetFromDateString(m_ToIndianDate)
        .ColAlignment(2) = 1
        .ColAlignment(4) = 1
    End If
    
    'For General Ledger
    If ReportType = 1 Then
        .Cols = 6: .Row = 0
        .Col = 1: .Text = GetResourceString(39) '" Particulars "
        .CellAlignment = 4: .CellFontBold = True
        .Col = 2: .Text = GetResourceString(284) '" Opening balance"
        .CellAlignment = 4: .CellFontBold = True
        .Col = 3: .Text = GetResourceString(276) '" Debit "
        .CellAlignment = 4: .CellFontBold = True
        .Col = 4: .Text = GetResourceString(277) '" Credit"
        .CellAlignment = 4: .CellFontBold = True
        .Col = 5: .Text = GetResourceString(285) '" Closing balance"
        .CellAlignment = 4: .CellFontBold = True
        lblReportTitle.Caption = GetResourceString(86)
        .ColAlignment(3) = 1
    End If
End With

Me.lblReportTitle.FONTSIZE = 14
LastLine:

With grd
    .ColWidth(0) = 1000
    Wid = (.Width - .ColWidth(0) * 1.1) / (.Cols - 1)
    If ReportType <> wisBalanceSheet Then
        .ColWidth(1) = Wid * 1.65
        .ColAlignment(1) = 0
        .ColWidth(2) = Wid * 0.66
        .ColAlignment(2) = 1
        .ColWidth(3) = Wid * 0.66
        .ColAlignment(3) = 1
    Else
        .ColWidth(0) = Wid * 0.66
        .ColWidth(1) = Wid * 2.3
        .ColAlignment(1) = 0
        .ColWidth(2) = Wid * 0.99
        .ColWidth(3) = Wid * 1.22
        .ColAlignment(3) = 1
        .ColWidth(4) = Wid * 2.3
        .ColAlignment(4) = 0
        .ColWidth(5) = Wid * 0.99
        .ColWidth(6) = Wid * 1.22
        .ColAlignment(6) = 1
        .ColAlignment(5) = 1
    End If
End With


Wid = 0
Exit Sub

ErrLine:

If Err Then MsgBox Err.Number & vbCrLf & "   Description : " & Err.Description

End Sub

Private Sub Form_Resize()

Screen.MousePointer = vbDefault
On Error Resume Next
lblReportTitle.Top = 0
lblReportTitle.Left = (Me.Width - lblReportTitle.Width) / 2

With grd
    .Left = 0
    .Top = lblReportTitle.Top + lblReportTitle.Height
    .Width = Me.Width - 150
    FraDate.Top = Me.ScaleHeight - FraDate.Height
    FraDate.Left = 100 'Me.Width - FraDate.Width
    FraDate.Width = Me.ScaleWidth - 100
    .Height = Me.ScaleHeight - FraDate.Height - lblReportTitle.Height
End With

'Adjust the Text boxes & Labels
cmdOk.Left = FraDate.Width - cmdOk.Width - 100
cmdPrint.Top = cmdOk.Top
cmdPrint.Left = cmdOk.Left - cmdPrint.Width - 100
cmdWeb.Top = cmdPrint.Top
cmdWeb.Left = cmdPrint.Left - (cmdWeb.Width + (cmdPrint.Width / 4))

Call InitGrid(True)

End Sub

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

cmdOk.Caption = GetResourceString(11)
cmdPrint.Caption = GetResourceString(23)

End Sub

Private Sub grd_LostFocus()
Dim ColCount As Integer
    For ColCount = 0 To grd.Cols - 1
        Call SaveSetting(App.EXEName, "GLReportType" & ReportType, "COLWIDTH" & ColCount, grd.ColWidth(ColCount) / grd.Width)
    Next
    
End Sub

Private Sub m_grdPrint_MaxProcessCount(MaxCount As Long)
m_TotalCount = MaxCount

Set m_frmCancel = New frmCancel
m_frmCancel.PicStatus.Visible = True
m_frmCancel.PicStatus.ZOrder 0

End Sub

Private Sub m_grdPrint_Message(strMessage As String)

On Error Resume Next
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


