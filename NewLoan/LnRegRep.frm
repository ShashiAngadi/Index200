VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRegReport 
   Caption         =   "Grids"
   ClientHeight    =   6495
   ClientLeft      =   660
   ClientTop       =   1740
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   9585
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   3060
      TabIndex        =   4
      Top             =   5850
      Width           =   5145
      Begin VB.CommandButton cmdWeb 
         Caption         =   "&Web view"
         Height          =   450
         Left            =   690
         TabIndex        =   7
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Close"
         Height          =   450
         Left            =   3300
         TabIndex        =   6
         Top             =   90
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   450
         Left            =   1680
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   4755
      Left            =   330
      TabIndex        =   0
      Top             =   1020
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   8387
      _Version        =   393216
   End
   Begin VB.Label lblReportTitle 
      AutoSize        =   -1  'True
      Caption         =   "TheBijapur District Co-Operative Central Bank Limited, Bijapur."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   540
      TabIndex        =   2
      Top             =   480
      Width           =   7530
   End
   Begin VB.Label lblTypeLoan 
      Caption         =   "Type of Loan :"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   1140
      Width           =   1245
   End
   Begin VB.Label lblBankName 
      AutoSize        =   -1  'True
      Caption         =   "From1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   510
      TabIndex        =   1
      Top             =   150
      Width           =   1380
   End
End
Attribute VB_Name = "frmRegReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Initialise(Min As Long, Max As Long)
Public Event Processing(strMessage As String, Ratio As Single)

Private WithEvents m_grdPrint As WISPrint
Attribute m_grdPrint.VB_VarHelpID = -1
Private m_TotalCount As Long
Private m_frmCancel As frmCancel

Private m_Count As Integer
Private m_MaxCount As Integer
Private m_SchemeId As Integer

Private m_repType As wis_LoanReports
Private m_Place As String
Private m_Caste As String
Private m_FromAmount As Currency
Private m_ToAmount As Currency
Private m_FromIndianDate As String
Private m_ToIndianDate As String
Private m_ReportOrder As wis_ReportOrder

Private m_FromDate As Date
Private m_ToDate As Date

Private xlWorkBook As Object
Private xlWorkSheet As Object

Public Property Let Caste(NewCaste As String)
    m_Caste = NewCaste
End Property

Public Property Let ToAmount(curTo As Currency)
    m_ToAmount = curTo
End Property

Public Property Let FromAmount(curFrom As Currency)
    m_FromAmount = curFrom
End Property

Public Property Let ToIndianDate(NewDate As String)
    If DateValidate(NewDate, "/", True) Then
        m_ToIndianDate = NewDate
        m_ToDate = GetSysFormatDate(NewDate)
    Else
        m_ToIndianDate = ""
        m_ToDate = vbNull
    End If
End Property

Public Property Let FromIndianDate(NewDate As String)
    If DateValidate(NewDate, "/", True) Then
        m_FromDate = GetSysFormatDate(NewDate)
        m_FromIndianDate = NewDate
    Else
        m_FromIndianDate = ""
        m_FromDate = vbNull
    End If
End Property

Public Property Let Place(NewPlace As String)
    m_Place = NewPlace
End Property

Public Property Let ReportOrder(RepOrder As wis_ReportOrder)
    m_ReportOrder = RepOrder
End Property

Public Property Let ReportType(RepType As wis_LoanReports)
    m_repType = RepType
End Property

Public Property Let LoanSchemeType(LoanType As Integer)
    m_SchemeId = LoanType
End Property

Private Sub GridCols(HeadArray() As String, Optional LExcel As Boolean, Optional lSlNo As Boolean)
          
Dim ColNum As Integer
Dim RowNum As Integer
Dim Cols As Integer
Dim Items As Integer

With grd
    RowNum = .Row
    .Col = 0
                
    ' put the main header column wise
    For Items = LBound(HeadArray) To UBound(HeadArray)
        .Col = Items: .Text = HeadArray(Items): .CellFontBold = True: .CellAlignment = 7
        If LExcel Then
            With xlWorkSheet
                .cells(RowNum + 1, ColNum + 1) = HeadArray(ColNum)
                .cells(RowNum + 1, ColNum + 1).Font.Bold = True
            End With
        End If
    Next
    
    ' if lslno is true
    If lSlNo Then
        RowNum = RowNum + 1
        .Row = RowNum
        .Col = 0
        For ColNum = LBound(HeadArray) To UBound(HeadArray)
            .Col = ColNum: .Text = ColNum + 1: .CellFontBold = True: .CellAlignment = 4
            If LExcel Then
                With xlWorkSheet
                    .cells(RowNum + 1, ColNum + 1) = ColNum + 1
                    .cells(RowNum + 1, ColNum + 1).Font.Bold = True
                End With
            End If
        Next ColNum
    End If
    
End With

End Sub

Private Sub GridColsKeys(HeadArray() As String, Optional LExcel As Boolean, Optional lSlNo As Boolean)
          
Dim ColNum As Integer
Dim RowNum As Integer
Dim Cols As Integer
Dim Items As Integer

With grd
    RowNum = .Row
    .Col = 0
                
    ' put the main header column wise
    For Items = LBound(HeadArray) To UBound(HeadArray)
        .Col = Items: .Text = HeadArray(Items): .CellFontBold = True: .CellAlignment = 4
        If LExcel Then
            With xlWorkSheet
                .cells(RowNum + 1, ColNum + 1) = HeadArray(ColNum)
                .cells(RowNum + 1, ColNum + 1).Font.Bold = True
            End With
        End If
    Next
    
    ' if lslno is true
    If lSlNo Then
        RowNum = RowNum + 1
        .Row = RowNum
        .Col = 0
        For ColNum = LBound(HeadArray) To UBound(HeadArray)
            .Col = ColNum: .Text = ColNum + 1: .CellFontBold = True: .CellAlignment = 4
            If LExcel Then
                With xlWorkSheet
                    .cells(RowNum + 1, ColNum + 1) = ColNum + 1
                    .cells(RowNum + 1, ColNum + 1).Font.Bold = True
                End With
            End If
        Next ColNum
    End If
    
End With
          
End Sub
Private Sub MoreRows(RowNum As Integer)

With grd
    If .Rows < .Row + RowNum Then
        .Rows = .Rows + RowNum
    End If
End With

End Sub

Private Function Shed4CRowCol() As Boolean

Dim COlHeader() As String
Dim RowNum As Integer
Dim ColNum As Integer

Shed4CRowCol = False

RowNum = 0
With grd
    .Cols = 10
    .Rows = 20
    .AllowUserResizing = flexResizeBoth
    .WordWrap = True
    .Clear
    .FixedCols = 1
    .FixedRows = 4
    .Row = 0
    RowNum = 0
End With

ReDim COlHeader(0 To 9)

COlHeader(0) = ""
COlHeader(1) = "Castewise Distribution of Loans"
COlHeader(2) = "Castewise Distribution of Loans"
COlHeader(3) = "Castewise Distribution of Loans"
COlHeader(4) = "Castewise Distribution of Loans"
COlHeader(5) = "Castewise Distribution of Loans"
COlHeader(6) = "Castewise Distribution of Loans"
COlHeader(7) = "Castewise Distribution of Loans"
COlHeader(8) = "Castewise Distribution of Loans"
COlHeader(9) = "Castewise Distribution of Loans"

Call GridCols(COlHeader())

COlHeader(0) = "Sl No"
COlHeader(1) = "Bank Name"
COlHeader(2) = "Crop Name"
COlHeader(3) = "Caste Name"
COlHeader(4) = "Male"
COlHeader(5) = "Male"
COlHeader(6) = "Female"
COlHeader(7) = "Female"
COlHeader(8) = "Other Female"
COlHeader(9) = "Other Female"

RowNum = RowNum + 1
grd.Row = RowNum
Call GridCols(COlHeader())

COlHeader(0) = "Sl No"
COlHeader(1) = "Bank Name"
COlHeader(2) = "Crop Name"
COlHeader(3) = "Caste Name"
COlHeader(4) = "Number"
COlHeader(5) = "Amount"
COlHeader(6) = "Number"
COlHeader(7) = "Amount"
COlHeader(8) = "Number"
COlHeader(9) = "Amount"

RowNum = RowNum + 1
grd.Row = RowNum
Call GridCols(COlHeader())

' other settings
With grd
    .ColWidth(1) = 2500
    .ColWidth(2) = 2500
        
    .MergeCells = flexMergeRestrictColumns
    .MergeRow(0) = True
    .MergeRow(1) = True
    .MergeRow(2) = True
    .MergeCells = flexMergeRestrictAll
End With

Shed4CRowCol = True
End Function

Private Function Shed4BRowCol() As Boolean

Dim COlHeader() As String
Dim RowNum As Integer
Dim ColNum As Integer

Shed4BRowCol = False

RowNum = 0
With grd
    .Cols = 25
    .Rows = 20
    .AllowUserResizing = flexResizeBoth
    .WordWrap = True
    .Clear
    .FixedCols = 1
    .FixedRows = 4
    .Row = 0
    RowNum = 0
End With

ReDim COlHeader(0 To 24)

COlHeader(0) = ""
COlHeader(1) = ""
COlHeader(2) = ""
COlHeader(3) = "Loan Given to the New Members"
COlHeader(4) = "Loan Given to the New Members"
COlHeader(5) = "Loan Given to the New Members"
COlHeader(6) = "Loan Given to the New Members"
COlHeader(7) = "Loan Given to the New Members"
COlHeader(8) = "Loan Given to the New Members"
COlHeader(9) = "Loan Given to the New Members"
COlHeader(10) = "Loan Given to the New Members"
COlHeader(11) = "Castewise Distribution of Total Loan"
COlHeader(12) = "Castewise Distribution of Total Loan"
COlHeader(13) = "Castewise Distribution of Total Loan"
COlHeader(14) = "Castewise Distribution of Total Loan"
COlHeader(15) = "Castewise Distribution of Total Loan"
COlHeader(16) = "Castewise Distribution of Total Loan"
COlHeader(17) = "Castewise Distribution of Total Loan"
COlHeader(18) = "Castewise Distribution of Total Loan"
COlHeader(19) = "Castewise Distribution of Total Loan"
COlHeader(20) = "Castewise Distribution of Total Loan"
COlHeader(21) = "Castewise Distribution of Total Loan"
COlHeader(22) = "Castewise Distribution of Total Loan"
COlHeader(23) = "Castewise Distribution of Total Loan"
COlHeader(24) = "Castewise Distribution of Total Loan"

' first row and first header
Call GridCols(COlHeader())


COlHeader(0) = ""
COlHeader(1) = ""
COlHeader(2) = ""
COlHeader(3) = "Big Farmers"
COlHeader(4) = "Big Farmers"
COlHeader(5) = "Small Farmers"
COlHeader(6) = "Small Farmers"
COlHeader(7) = "SC/ST Farmers"
COlHeader(8) = "SC/ST Farmers"
COlHeader(9) = "Total Farmers"
COlHeader(10) = "Total Farmers"
COlHeader(11) = "Muslim Members"
COlHeader(12) = "Muslim Members"
COlHeader(13) = "Muslim Members"
COlHeader(14) = "Muslim Members"
COlHeader(15) = "Christian Members"
COlHeader(16) = "Christian Members"
COlHeader(17) = "Christian Members"
COlHeader(18) = "Christian Members"
COlHeader(19) = "Jain Members"
COlHeader(20) = "Jain Members"
COlHeader(21) = "Jain Members"
COlHeader(22) = "Jain Members"
COlHeader(23) = "Other Caste Female"
COlHeader(24) = "Other Caste Female"

RowNum = RowNum + 1
grd.Row = RowNum
Call GridCols(COlHeader())


COlHeader(0) = ""
COlHeader(1) = ""
COlHeader(2) = ""
COlHeader(3) = "Big Farmers"
COlHeader(4) = "Big Farmers"
COlHeader(5) = "Small Farmers"
COlHeader(6) = "Small Farmers"
COlHeader(7) = "SC/ST Farmers"
COlHeader(8) = "SC/ST Farmers"
COlHeader(9) = "Total Farmers"
COlHeader(10) = "Total Farmers"
COlHeader(11) = "Male"
COlHeader(12) = "Male"
COlHeader(13) = "Female"
COlHeader(14) = "Female"
COlHeader(15) = "Male"
COlHeader(16) = "Male"
COlHeader(17) = "Female"
COlHeader(18) = "Female"
COlHeader(19) = "Male"
COlHeader(20) = "Male"
COlHeader(21) = "Female"
COlHeader(22) = "Female"
COlHeader(23) = "Other Caste Female"
COlHeader(24) = "Other Caste Female"


RowNum = RowNum + 1
grd.Row = RowNum
Call GridCols(COlHeader())

COlHeader(0) = "Sl No"
COlHeader(1) = "Branch Name"
COlHeader(2) = "Crop Name"
COlHeader(3) = "Number"
COlHeader(4) = "Amount"
COlHeader(5) = "Number"
COlHeader(6) = "Amount"
COlHeader(7) = "Number"
COlHeader(8) = "Amount"
COlHeader(9) = "Number"
COlHeader(10) = "Amount"
COlHeader(11) = "Number"
COlHeader(12) = "Amount"
COlHeader(13) = "Number"
COlHeader(14) = "Amount"
COlHeader(15) = "Number"
COlHeader(16) = "Amount"
COlHeader(17) = "Number"
COlHeader(18) = "Amount"
COlHeader(19) = "Number"
COlHeader(20) = "Amount"
COlHeader(21) = "Number"
COlHeader(22) = "Amount"
COlHeader(23) = "Number"
COlHeader(24) = "Amount"

RowNum = RowNum + 1
grd.Row = RowNum
Call GridCols(COlHeader())

' other settings
With grd
    .ColWidth(1) = 2500
    .ColWidth(2) = 2500
    
    .MergeCells = flexMergeRestrictColumns
    .MergeRow(0) = True
    .MergeRow(1) = True
    .MergeRow(2) = True
    .MergeRow(3) = True
    
    .Row = 1
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .MergeCol(6) = True
    .MergeCol(7) = True
    .MergeCol(8) = True
    .MergeCol(9) = True
    .MergeCol(10) = True
    
    .MergeCol(23) = True
    .MergeCol(24) = True
    
    .Row = 2
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .MergeCol(6) = True
    .MergeCol(7) = True
    .MergeCol(8) = True
    .MergeCol(9) = True
    .MergeCol(10) = True
    
    .MergeCol(23) = True
    .MergeCol(24) = True
    
    .MergeCells = flexMergeRestrictAll
End With


Shed4BRowCol = True

End Function

Public Function ShowConsoleBalance() As Boolean

m_repType = repConsBalance

Dim SqlStr As String
Dim rst As Recordset
Dim Total As Currency

Err.Clear
On Error GoTo ExitLine


RaiseEvent Processing("Fetching the record", 0)

gDbTrans.SqlStmt = "Select MAX(TransID),BankID,LoanID From LoanTrans Where " & _
    " TransDate <= #" & m_ToDate & "# "
gDbTrans.CreateView ("MaxLoanTransID")

SqlStr = "select Sum(Balance) as Toal_Balance,D.BankName,B.SchemeID " & _
    " From LoanTrans A Inner Join (MaxLoanTransID C Inner join " & _
    " (LoanMaster B Inner Join BankDet D ON B.BankID= D.BankID)" & _
    " ON C.LoanId = B.LoanId AND C.BankId = B.BankId) ON " & _
    " C.LoanID =A.LoanId AND C.BankId = A.BankId  "
'SqlStr = "SELECT SUM(Balance) as Total_Balance,D.BankName,B.SchemeID " & _
    " From LoanTrans A,LoanMaster B,BankDet D  WHERE TransID = " & _
    " (Select MAX(TransID) From LoanTrans C Where C.LoanID =A.LoanId" & _
    " AND C.BankId=A.BankId  AND C.BankID = D.BankID and transdate <= #" & m_ToDate & "# ) " & _
    " And A.LoanId = B.LoanId AND A.BankId = B.BankId AND A.BankID= D.BankID "
    
If m_SchemeId Then SqlStr = SqlStr & " HAVING B.SchemeID = " & m_SchemeId

SqlStr = SqlStr & " GROUP BY B.SchemeID,D.BankName ORDER BY B.SchemeID,D.Bankname"

gDbTrans.SqlStmt = SqlStr
  If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Function
  
RaiseEvent Initialise(0, rst.RecordCount)

With grd
    .Clear
    .Cols = 1
    .Rows = 25
    .Row = 0
    .FormatString = ">SlNo |> BankName |> SchemeName |< Balance "
    .FixedCols = 1
    .AllowUserResizing = flexResizeBoth
End With

Call InitGrid
m_repType = repConsBalance

Dim SINO As Long
Dim Balance As Currency
Dim SubBalance As Currency
Dim TotalBalance As Currency
Dim l_SchemeID As Integer
Dim SchemName As String
Dim rstTemp As Recordset

l_SchemeID = FormatField(rst("Schemeid"))
gDbTrans.SqlStmt = "SELECT SchemeName From LoanScheme WHERE SchemeID = " & l_SchemeID
Call gDbTrans.Fetch(rstTemp, adOpenStatic)
SchemName = FormatField(rstTemp("SchemeName"))

lblReportTitle = ""
If m_SchemeId Then lblReportTitle = SchemName & " "
lblReportTitle = lblReportTitle & GetResourceString(67) & _
    " " & GetFromDateString(m_ToIndianDate)

'Write in to the grid
 SINO = 1
 Dim rowno As Long
 rowno = grd.Row
 
 While Not rst.EOF
    If l_SchemeID <> FormatField(rst("Schemeid")) Then
        l_SchemeID = FormatField(rst("Schemeid"))
        With grd
           rowno = rowno + 1
           If .Rows <= rowno Then .Rows = rowno + 1
           
           .Row = rowno
           .Col = 1: .Text = "Sub Total "
           .CellAlignment = 1: .CellFontBold = True
           .Col = 2: .Text = SchemName: .CellAlignment = 1
           .CellAlignment = 1: .CellFontBold = True
           .Col = 3: .Text = FormatCurrency(SubBalance):
           .CellAlignment = 7: .CellFontBold = True
            TotalBalance = TotalBalance + SubBalance: SubBalance = 0
           
           rowno = rowno + 1
           
        End With
        SINO = 1
        gDbTrans.SqlStmt = "SELECT SchemeName From LoanScheme WHERE SchemeID = " & l_SchemeID
        Call gDbTrans.Fetch(rstTemp, adOpenStatic)
        SchemName = FormatField(rstTemp("SchemeName"))
    End If
    
    With grd
        rowno = rowno + 1
       If .Rows <= rowno Then .Rows = rowno + 1
       
       .TextMatrix(rowno, 0) = Format(SINO, "00"): .CellAlignment = 1
       .TextMatrix(rowno, 1) = FormatField(rst("BankName")): .CellAlignment = 1
       .TextMatrix(rowno, 2) = SchemName: .CellAlignment = 1
       .TextMatrix(rowno, 3) = FormatField(rst("Total_Balance")): .CellAlignment = 7
        SubBalance = SubBalance + Val(.TextMatrix(rowno, 3))
    End With
    SINO = SINO + 1
    
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the data ", SINO / rst.RecordCount)
    rst.MoveNext

Wend

With grd
   rowno = rowno + 1
   If .Rows <= rowno Then .Rows = rowno + 1
   .Row = rowno
   .Col = 1: .Text = GetResourceString(42)
   .CellAlignment = 4: .CellFontBold = True
   .Col = 2: .Text = SchemName: .CellAlignment = 1
   .CellAlignment = 4: .CellFontBold = True
   .Col = 3: .Text = FormatCurrency(SubBalance):
   .CellAlignment = 7: .CellFontBold = True
    TotalBalance = TotalBalance + SubBalance: SubBalance = 0
   
    rowno = rowno + 2
    If .Rows <= rowno Then .Rows = rowno + 1
    .Row = rowno
    .Col = 3: .Text = GetResourceString(42, 52)
    .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .Text = FormatCurrency(TotalBalance)
    .CellAlignment = 4: .CellFontBold = True

End With
 grd.Visible = True
Call grd_LostFocus

ShowConsoleBalance = True

ExitLine:
    If Err Then
        MsgBox "ERROR ReportConsoleBalance" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
    End If
End Function

Private Function ShowConsoleInstOverDue() As Boolean
Dim SqlStr As String
Dim rst As Recordset
RaiseEvent Processing("Reading the data", 0)

m_repType = repConsInstOD

SqlStr = "SELECT SUM(InstBalance) as Total_Balance,B.SchemeID " & _
    " From LoanInst A Inner Join LoanMaster B ON And A.LoanId = B.LoanId " & _
    " WHERE InstDate <= #" & m_ToDate & "#"

If m_SchemeId Then SqlStr = SqlStr & " AND B.SchemeID = " & m_SchemeId

SqlStr = SqlStr & " GROUP BY B.SchemeID order by B.SchemeId"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Function

RaiseEvent Initialise(0, rst.RecordCount)
With grd
    .Clear
    .Rows = 25
    .FixedCols = 0
    .FixedRows = 1
    .Row = 0
    .FormatString = ">Sl No |>BankName |<Over Due Balance"
    .FixedCols = 1
    .AllowUserResizing = flexResizeBoth
End With
Call InitGrid

Dim l_SchemeID As Integer
Dim SchemeName As String
Dim SlNo  As Integer
Dim SubBalance As Double
Dim TotalBalance As Double
Dim rstTemp As Recordset

l_SchemeID = rst("SchemeiD")
gDbTrans.SqlStmt = "SELECT SchemeName From LoanScheme WHERE SchemeId = " & l_SchemeID

Call gDbTrans.Fetch(rstTemp, adOpenStatic)
SchemeName = FormatField(rstTemp("SchemeName"))

grd.Row = 1
grd.Col = 1: grd.Text = SchemeName: grd.CellFontBold = True

lblReportTitle = GetResourceString(113) & " " & GetFromDateString(m_ToIndianDate)
Dim rowno As Long
rowno = grd.Row

While Not rst.EOF
    If l_SchemeID <> rst("SchemeiD") Then
        With grd
            rowno = rowno + 1
            If .Row <= rowno Then .Rows = rowno + 1
            .Row = rowno
            .Col = 1: .Text = GetResourceString(304)
            .CellFontBold = True: .CellAlignment = 1
            .Col = 2: .Text = FormatCurrency(SubBalance)
            .CellFontBold = True: .CellAlignment = 7
            TotalBalance = TotalBalance + SubBalance
            SubBalance = 0
            rowno = rowno + 2
            If .Rows <= rowno Then .Rows = rowno + 1
            .Row = rowno
        
            l_SchemeID = rst("SchemeiD")
            gDbTrans.SqlStmt = "SELECT SchemeName From LoanScheme " & _
                    "WHERE SchemeId = " & l_SchemeID
            Call gDbTrans.Fetch(rstTemp, adOpenStatic)
            SchemeName = FormatField(rstTemp("SchemeName"))
            .Col = 1: .Text = SchemeName
            .CellFontBold = True: .CellAlignment = 4
        End With
    End If
    SlNo = SlNo + 1
    With grd
        rowno = rowno + 1
        If .Rows <= rowno Then .Rows = rowno + 1
        .TextMatrix(rowno, 0) = SlNo
        .TextMatrix(rowno, 2) = FormatField(rst("Total_Balance")): .CellAlignment = 7
        SubBalance = SubBalance + Val(.Text)
    End With
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("writing the record", rst.AbsolutePosition / rst.RecordCount)
    rst.MoveNext
Wend

With grd
    rowno = rowno + 2
    If .Rows <= rowno Then .Rows = rowno + 1
    .Row = rowno
    .Col = 1: .Text = GetResourceString(304)
    .CellFontBold = True: .CellAlignment = 1
    .Col = 2: .Text = FormatCurrency(SubBalance)
    .CellFontBold = True: .CellAlignment = 7
    TotalBalance = TotalBalance + SubBalance
    SubBalance = 0
    
    rowno = rowno + 2
    If .Rows <= rowno - 1 Then .Rows = rowno + 1
    .Row = rowno
    .Col = 1: .Text = GetResourceString(286)
    .CellFontBold = True: .CellAlignment = 4
    .Col = 2: .Text = FormatCurrency(TotalBalance)
    .CellFontBold = True: .CellAlignment = 7
    SubBalance = 0
End With

ShowConsoleInstOverDue = True
Call grd_LostFocus

End Function

Private Function ShowConsoleODBalance() As Boolean

Dim SqlStr As String
Dim rst As Recordset

RaiseEvent Processing("Fetching the record", 0)
m_repType = repConsOD

gDbTrans.SqlStmt = "Select max(TransID) as MaxTransID,LoanID From LoanTrans C Where " & _
                " Transdate <= #" & m_ToDate & "# GROUP BY LoanID"
gDbTrans.CreateView ("MaxODLoanTransID")

If m_SchemeId Then
    SqlStr = "SELECT SUM(Balance) as Total_Balance,B.SchemeID " & _
        " From LoanMaster B Inner Join (LoanTrans A INNER JOIN " & _
            " MaxODLoanTransID C ON C.LoanID =A.LoanId AND C.MaxTransID = A.TransID)" & _
        " ON A.LoanId = B.LoanId WHERE A.LoanId In (SELECT LoanID From LoanMaster D " & _
            " WHERE LoanDueDate <= #" & m_ToDate & "# AND D.SchemeId = " & m_SchemeId & ")" & _
        " GROUP BY B.SchemeID order by B.SchemeId"

    'SqlStr = "SELECT SUM(Balance) as Total_Balance,B.SchemeID " & _
        " From LoanTrans A,LoanMaster B  WHERE TransID = " & _
            " (Select max(TransID) From LoanTrans C Where C.LoanID =A.LoanId" & _
            " AND Transdate <= #" & m_ToDate & "# ) " & _
        " AND A.LoanId In (SELECT LoanID From LoanMaster D WHERE " & _
            " D.LoanID = A.LoanID AND LoanDueDate <= #" & m_ToDate & "#" & _
            " AND D.SchemeId = " & m_SchemeId & ")" & _
        " And A.LoanId = B.LoanId " & _
        " GROUP BY B.SchemeID order by B.SchemeId"
Else
    SqlStr = "SELECT SUM(Balance) as Total_Balance,B.SchemeID " & _
        " From LoanMaster B  INner JOIN (LoanTrans A INNER JOIN MaxODLoanTransID " & _
            " ON C.LoanID =A.LoanId AND C.MaxTransID = A.TransID)" & _
        " ON A.LoanId = B.LoanId WHERE A.LoanId In (SELECT distinct LoanID From LoanMaster D " & _
            " WHERE LoanDueDate <= #" & m_ToDate & "#)" & _
        " GROUP BY B.SchemeID order by B.SchemeId"
        
    'SqlStr = "SELECT SUM(Balance) as Total_Balance,B.SchemeID " & _
        " From LoanTrans A,LoanMaster B  WHERE TransID = " & _
            " (Select max(TransID) From LoanTrans C Where C.LoanID =A.LoanId" & _
            " AND Transdate <= #" & m_ToDate & "# ) " & _
        " AND A.LoanId In (SELECT LoanID From LoanMaster D WHERE " & _
            " D.LoanID = A.LoanID AND LoanDueDate <= #" & m_ToDate & "#)" & _
        " And A.LoanId = B.LoanId " & _
        " GROUP BY B.SchemeID order by B.SchemeId"
End If
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rst, adOpenStatic) < 1 Then Exit Function

RaiseEvent Initialise(0, rst.RecordCount)
With grd
    .Clear
    .Rows = 25
    .FixedCols = 0
    .FixedRows = 1
    .Row = 0
    .FormatString = ">Sl No |<Loan Name |<BankName |>Over Due Balance"
    .FixedCols = 1
    .AllowUserResizing = flexResizeBoth
End With
Call InitGrid

Dim l_SchemeID As Integer
Dim BankID As Long
Dim SchemeName As String
Dim BankName As String
Dim SlNo  As Integer

Dim SubBalance As Double
Dim TotalBalance As Double
Dim rstTemp As Recordset

l_SchemeID = rst("SchemeiD")
gDbTrans.SqlStmt = "SELECT SchemeName From LoanScheme WHERE SchemeId = " & l_SchemeID
Call gDbTrans.Fetch(rstTemp, adOpenStatic)
SchemeName = FormatField(rstTemp("SchemeName"))

lblReportTitle = "Over due balance As on " & m_FromIndianDate
If m_SchemeId Then lblReportTitle = "Over due balance of " & SchemeName & " As on " & m_FromIndianDate

Dim rowno As Long
rowno = grd.Row
While Not rst.EOF
    If l_SchemeID <> rst("SchemeiD") Then
        With grd
            rowno = rowno + 1
            If .Rows <= rowno Then .Rows = rowno + 1
            .Row = rowno
            .Col = 1: .Text = SchemeName: .CellFontBold = True
            .Col = 2: .Text = "Total": .CellFontBold = True
            .Col = 3: .Text = FormatCurrency(SubBalance): .CellFontBold = True
            TotalBalance = TotalBalance + SubBalance
            SubBalance = 0
            rowno = rowno + 1
        End With
        SlNo = 0
        l_SchemeID = rst("SchemeiD")
        gDbTrans.SqlStmt = "SELECT SchemeName From LoanScheme WHERE SchemeId = " & l_SchemeID
        Call gDbTrans.Fetch(rstTemp, adOpenStatic)
        SchemeName = FormatField(rstTemp("SchemeName"))
    End If
    
    If BankID <> rst("BankID") Then
        BankID = rst("BankID")
        gDbTrans.SqlStmt = "SELECT BankName From BankDet WHERE BankID = " & BankID
        Call gDbTrans.Fetch(rstTemp, adOpenStatic)
        BankName = FormatField(rstTemp("BankName"))
    End If
    SlNo = SlNo + 1
    With grd
        rowno = rowno + 1
        If .Rows <= rowno Then .Rows = rowno + 1
        .TextMatrix(rowno, 0) = SlNo
        .TextMatrix(rowno, 1) = SchemeName
        .TextMatrix(rowno, 2) = BankName
        .TextMatrix(rowno, 3) = FormatField(rst("Total_Balance"))
        SubBalance = SubBalance + Val(.TextMatrix(rowno, 3))
    End With
    DoEvents
    If gCancel Then rst.MoveLast
    RaiseEvent Processing("Writing the records", SlNo / rst.RecordCount)
    rst.MoveNext
Wend

With grd
    rowno = rowno + 1
    If .Rows <= rowno Then .Rows = rowno + 1
    .Row = rowno
    .Col = 1: .Text = SchemeName: .CellFontBold = True
    .Col = 2: .Text = "Total": .CellFontBold = True
    .Col = 3: .Text = FormatCurrency(SubBalance): .CellFontBold = True
    TotalBalance = TotalBalance + SubBalance
    SubBalance = 0
    
    rowno = rowno + 2
    If .Rows <= rowno Then .Rows = rowno + 1
    .Row = rowno
    .Col = 1: .Text = "Grand Total": .CellFontBold = True
    .Col = 3: .Text = FormatCurrency(TotalBalance): .CellFontBold = True
    SubBalance = 0
    .ColAlignment(0) = 1
    .ColAlignment(1) = 1
    .ColAlignment(2) = 1
    .ColAlignment(3) = 8
End With

ShowConsoleODBalance = True
Call grd_LostFocus

End Function

Private Function ShowShed4C() As Boolean

'contact pradeep bellubbi for this function
Dim SlNo As Integer
Dim RowNum As Integer
Dim ColNum As Integer
Dim RefID As Long
Dim CasteName As String
Dim CropName As String
Dim BankName As String

Dim Male_No As Integer
Dim Male_Amount As Currency
Dim Female_No As Integer
Dim Female_Amount As Currency
Dim SCF_No As Integer
Dim SCF_Amount As Currency

Dim totMale_No As Long
Dim totMale_Amount As Currency
Dim totFemale_No As Long
Dim totFemale_Amount As Currency
Dim totSCF_No As Long
Dim totSCF_Amount As Currency

Dim grdMale_No As Long
Dim grdMale_Amount As Currency
Dim grdFemale_No As Long
Dim grdFemale_Amount As Currency
Dim grdSCF_No As Long
Dim grdSCF_Amount As Currency

Dim rstCasteWise As Recordset
Dim rstLoanDetail As Recordset

Dim SqlStr As String

ShowShed4C = False
Call Shed4CRowCol

SqlStr = " SELECT a.BankId,BankName,a.CropId,CropName,SCF_No,SCF_Amount,a.RefId" & _
         " FROM Shed4 a,BranchDet b,Crops c" & _
         " WHERE a.BankId = b.BankId" & _
         " AND a.CropId=c.CropId" & _
         " AND a.LoanDate >= " & "#" & m_FromDate & "#" & _
         " AND a.Loandate <= " & "#" & m_ToDate & "#" & _
         " ORDER BY a.RefId"
         
         
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstLoanDetail, adOpenStatic) < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
End If

' caste wise rst
SqlStr = " SELECT a.RefID,a.CasteID,c.CasteName,a.Male_No," & _
         " a.Male_Amount,a.Female_No,a.Female_Amount" & _
         " FROM LoanCasteWise a,NewLoans b,Caste c" & _
         " WHERE a.RefId=b.RefId AND a.CasteId=c.CasteId " & _
         " AND b.LoanDate >= " & "#" & m_FromDate & "#" & _
         " AND b.Loandate <= " & "#" & m_ToDate & "#"
         
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstCasteWise, adOpenStatic) < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
End If

SlNo = 1
RowNum = grd.FixedRows
ColNum = 0
' start the main loop
While Not rstLoanDetail.EOF

    RefID = FormatField(rstLoanDetail("RefId"))
    BankName = FormatField(rstLoanDetail("BankName"))
    ' find if loan has got data
    rstCasteWise.MoveFirst
    rstCasteWise.Find "RefId=" & RefID
    
    If Not rstCasteWise.EOF Then   ' if found and not eof
        ' start the inner loop
        While Not rstCasteWise.EOF And RefID = rstCasteWise("RefId")
        
            CasteName = FormatField(rstCasteWise("CasteName"))
            Male_No = FormatField(rstCasteWise("Male_No"))
            Male_Amount = FormatField(rstCasteWise("Male_Amount"))
            Female_No = FormatField(rstCasteWise("Female_No"))
            Female_Amount = FormatField(rstCasteWise("Female_Amount"))
                               
            With grd
                ' check the total rows
                MoreRows (2)
                grd.Row = RowNum
                ColNum = 0
                .Col = ColNum: .Text = SlNo: .CellAlignment = 4: ColNum = ColNum + 1
                .Col = ColNum: .Text = BankName: ColNum = ColNum + 1
                .Col = ColNum: .Text = CropName: ColNum = ColNum + 1
                .Col = ColNum: .Text = CasteName: ColNum = ColNum + 1
                .Col = ColNum: .Text = Male_No: ColNum = ColNum + 1
                .Col = ColNum: .Text = Male_Amount: ColNum = ColNum + 1
                .Col = ColNum: .Text = Female_No: ColNum = ColNum + 1
                .Col = ColNum: .Text = Female_Amount: ColNum = ColNum + 1
                        
                ' get the totals
                totMale_No = totMale_No + Male_No
                totMale_Amount = totMale_Amount + Male_Amount
                totFemale_No = totFemale_No + Female_No
                totFemale_Amount = totFemale_Amount + Female_Amount
                
            End With
            
            rstCasteWise.MoveNext
            RowNum = RowNum + 1
            SlNo = SlNo + 1
        Wend
        
        ' print the other women
        With grd
            SCF_No = FormatField(rstLoanDetail("SCF_No"))
            SCF_Amount = FormatField(rstLoanDetail("SCF_Amount"))
            
            .Col = ColNum: .Text = SCF_No: ColNum = ColNum + 1
            .Col = ColNum: .Text = SCF_Amount: ColNum = ColNum + 1
            
            totSCF_No = totSCF_No + SCF_No
            totSCF_Amount = totSCF_Amount + SCF_Amount
        End With
        
        RowNum = RowNum + 1
        
        With grd
            ' check the total rows
            MoreRows (2)
            grd.Row = RowNum
            ColNum = 0
            .Col = ColNum: .Text = "": ColNum = ColNum + 1
            .Col = ColNum: .Text = "Loan Total": ColNum = ColNum + 1
            .Col = ColNum: .Text = "": ColNum = ColNum + 1
            .Col = ColNum: .Text = "": ColNum = ColNum + 1
            .Col = ColNum: .Text = totMale_No: ColNum = ColNum + 1
            .Col = ColNum: .Text = totMale_Amount: ColNum = ColNum + 1
            .Col = ColNum: .Text = totFemale_No: ColNum = ColNum + 1
            .Col = ColNum: .Text = totFemale_Amount: ColNum = ColNum + 1
            .Col = ColNum: .Text = totSCF_No: ColNum = ColNum + 1
            .Col = ColNum: .Text = totSCF_Amount: ColNum = ColNum + 1
                    
            ' get the grand totals
            grdMale_No = grdMale_No + totMale_No
            grdMale_Amount = grdMale_Amount + totMale_Amount
            grdFemale_No = grdFemale_No + totFemale_No
            grdFemale_Amount = grdFemale_Amount + totFemale_Amount
            grdSCF_No = grdSCF_No + totSCF_No
            grdSCF_Amount = grdSCF_Amount + totSCF_Amount
            
        End With
        
    End If
    ' move to next loan
    rstLoanDetail.MoveNext
Wend

' now print the grand total
With grd
    ' check the total rows
    
    RowNum = RowNum + 1
    MoreRows (2)
    MoreRows (4) ' double check
    
    grd.Row = RowNum
    ColNum = 0
    .Col = ColNum: .Text = "": ColNum = ColNum + 1
    .Col = ColNum: .Text = "Grand Total": .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = "": .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = "": ColNum = ColNum + 1
    .Col = ColNum: .Text = totMale_No: .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = totMale_Amount: .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = totFemale_No: .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = totFemale_Amount: .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = totSCF_No: .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = totSCF_Amount: .CellFontBold = True: ColNum = ColNum + 1
            
End With

grd.Visible = True

ShowShed4C = True

End Function

Private Function ShowShed4B() As Boolean

Dim RowNum As Integer
Dim ColNum As Integer
Dim SlNo As Integer

Dim rstNewMems As Recordset
Dim rstCasteWise As Recordset
Dim rstLoanDetail As Recordset

Dim RefID As Long
Dim CropId As Byte
Dim CasteId As Byte

Dim SqlStr As String
Dim CropName As String
Dim BankName As String

Dim BF_No As Integer
Dim BF_Amount As Currency
Dim SF_No As Integer
Dim SF_Amount As Currency
Dim SC_No As Integer
Dim SC_Amount As Currency
Dim totNumber As Long
Dim TotAmount As Currency

Dim totBF_No As Integer
Dim totBF_Amount As Currency
Dim totSF_No As Integer
Dim totSF_Amount As Currency
Dim totSC_No As Integer
Dim totSC_Amount As Currency
Dim grdNumber As Long
Dim grdAmount As Currency


ShowShed4B = False

' set the main headers
Call Shed4BRowCol

' get the all loan details for the given period and bank
SqlStr = " SELECT a.BankId,BankName,a.CropId,CropName,SCF_No,SCF_Amount,a.RefId" & _
         " FROM Shed4 a,BranchDet b,Crops c" & _
         " WHERE a.BankId = b.BankId AND a.CropId=c.CropId" & _
         " AND a.LoanDate >= " & "#" & m_FromDate & "#" & _
         " AND a.Loandate <= " & "#" & m_ToDate & "#" & _
         " ORDER BY a.RefId"
         
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstLoanDetail, adOpenStatic) < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
End If

' this main rst
SqlStr = " SELECT a.RefID,a.BF_No,a.BF_Amount,a.SF_No,a.SF_Amount,a.SCST_No,a.SCST_Amount" & _
         " FROM NewLoanMembers a,NewLoans b " & _
         " WHERE a.RefID=b.RefID" & _
         " AND b.LoanDate >= " & "#" & m_FromDate & "#" & _
         " AND b.Loandate <= " & "#" & m_ToDate & "#" & _
         " ORDER BY a.RefID"
         
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstNewMems, adOpenStatic) < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
End If

' caste wise rst
SqlStr = " SELECT a.RefID,a.CasteID,c.CasteName,a.Male_No,a.Male_Amount,a.Female_No,a.Female_Amount" & _
         " FROM LoanCasteWise a,NewLoans b,Caste c" & _
         " WHERE a.RefId=b.RefId" & _
         " AND b.LoanDate >= " & "#" & m_FromDate & "#" & _
         " AND b.Loandate <= " & "#" & m_ToDate & "#" & _
         " AND a.CasteId=c.CasteId "
         
         
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstCasteWise, adOpenStatic) < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
End If
         
RowNum = grd.FixedRows
SlNo = 1
ColNum = 0

' start the main loop
While Not rstNewMems.EOF
    RefID = FormatField(rstNewMems("RefID"))
    BF_No = FormatField(rstNewMems("BF_No"))
    BF_Amount = FormatField(rstNewMems("BF_Amount"))
    SF_No = FormatField(rstNewMems("SF_No"))
    SF_Amount = FormatField(rstNewMems("SF_Amount"))
    SC_No = FormatField(rstNewMems("SCST_No"))
    SC_Amount = FormatField(rstNewMems("SCST_Amount"))
    
    ' total for the record
    totNumber = BF_No + SF_No + SC_No
    TotAmount = BF_Amount + SF_Amount + SC_Amount
    
    ' grand totals
    totBF_Amount = totBF_Amount + BF_Amount
    totBF_No = totBF_No + BF_No
    totSF_Amount = totSF_Amount + SF_Amount
    totSF_No = totSF_No + SF_No
    totSC_Amount = totSC_Amount + SC_Amount
    totSC_No = totSC_No + SC_No
    
    ' get the bank name and crop name
    rstLoanDetail.MoveFirst
    rstLoanDetail.Find "RefId=" & RefID
    If Not rstLoanDetail.EOF Then
        CropName = FormatField(rstLoanDetail("CropName"))
        BankName = FormatField(rstLoanDetail("BankName"))
    End If
       
    With grd
        ' check for the rows
         MoreRows (2)
         ColNum = 0
         .Row = RowNum
         
        .Col = ColNum: .Text = Str(SlNo): .CellAlignment = 4: ColNum = ColNum + 1
        .Col = ColNum: .Text = BankName: .CellAlignment = 1: ColNum = ColNum + 1
        .Col = ColNum: .Text = CropName: .CellAlignment = 1: ColNum = ColNum + 1
        .Col = ColNum: .Text = BF_No: ColNum = ColNum + 1
        .Col = ColNum: .Text = FormatCurrency(BF_Amount): ColNum = ColNum + 1
        .Col = ColNum: .Text = SF_No: ColNum = ColNum + 1
        .Col = ColNum: .Text = FormatCurrency(SF_Amount): ColNum = ColNum + 1
        .Col = ColNum: .Text = SC_No: ColNum = ColNum + 1
        .Col = ColNum: .Text = FormatCurrency(SC_Amount): ColNum = ColNum + 1
        .Col = ColNum: .Text = totNumber: ColNum = ColNum + 1
        .Col = ColNum: .Text = FormatCurrency(TotAmount): ColNum = ColNum + 1
        
        ' left other data half because of confusion
        
    End With
    
    rstNewMems.MoveNext
    SlNo = SlNo + 1
    RowNum = RowNum + 1
Wend


' for the grand total
With grd
    MoreRows (4)
    MoreRows (2) ' just a double check
    
    ' total fieidls
    grdNumber = totBF_No + totSF_No + totSC_No
    grdAmount = totBF_Amount + totSF_Amount + totSC_Amount
    
    RowNum = RowNum + 2
    .Row = RowNum
    ColNum = 0
    .Col = ColNum: .Text = "": .CellAlignment = 4: ColNum = ColNum + 1
    .Col = ColNum: .Text = "Grand Total": .CellFontBold = True: .CellAlignment = 1: ColNum = ColNum + 1
    .Col = ColNum: .Text = "": .CellAlignment = 1: ColNum = ColNum + 1
    .Col = ColNum: .Text = totBF_No: .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = FormatCurrency(totBF_Amount): .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = totSF_No: .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = FormatCurrency(totSF_Amount): .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = totSC_No: .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = FormatCurrency(totSC_Amount): .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = grdNumber: .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = FormatCurrency(grdAmount): .CellFontBold = True: ColNum = ColNum + 1
End With

grd.Visible = True

ShowShed4B = True

End Function


Private Function ShowShed4A() As Boolean

Dim rstLoanDetail As Recordset
Dim SqlStr As String
Dim RowNum As Integer
Dim ColNum As Integer
Dim SlNo As String

Dim totNumber As Long
Dim TotAmount As Currency

Dim BF_Amount As Currency
Dim BF_No As Integer
Dim SF_Amount As Currency
Dim SF_No As Integer
Dim SC_Amount As Currency
Dim SC_No As Integer

Dim totBF_Amount As Currency
Dim totBF_No As Long
Dim totSF_Amount As Currency
Dim totSF_No As Long
Dim totSC_Amount As Currency
Dim totSC_No As Long

Dim grdNumber As Long
Dim grdAmount As Currency


ShowShed4A = False

' set the headers
Call Shed4ARowCol

SqlStr = " SELECT a.BankId,BankName,a.CropId,CropName,LoanDate,LoanDueDate," & _
         " BF_No,BF_Amount,SF_No,SF_Amount,SCST_No,SCST_Amount" & _
         " FROM Shed4 a,BranchDet b,Crops c" & _
         " WHERE a.BankId = b.BankId" & _
         " AND a.CropId=c.CropId" & _
         " AND a.LoanDate >= " & "#" & m_FromDate & "#" & _
         " AND a.Loandate <= " & "#" & m_ToDate & "#"
         
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstLoanDetail, adOpenStatic) < 1 Then
    MsgBox "Unable to fetch the Data"
    Exit Function
End If


' start the main loop
SlNo = 1
RowNum = grd.FixedRows
ColNum = 0
While Not rstLoanDetail.EOF
    ColNum = 0
    With grd
        ' all the data into variables
        BF_Amount = FormatField(rstLoanDetail("BF_Amount"))
        BF_No = FormatField(rstLoanDetail("BF_No"))
        SF_Amount = FormatField(rstLoanDetail("SF_Amount"))
        SF_No = FormatField(rstLoanDetail("SF_No"))
        SC_Amount = FormatField(rstLoanDetail("SCST_Amount"))
        SC_No = FormatField(rstLoanDetail("SCST_No"))
        
        ' total fieidls
        totNumber = BF_No + SF_No + SC_No
        TotAmount = BF_Amount + SF_Amount + SC_Amount
              
        ' get the data into grand totals
        totBF_Amount = totBF_Amount + BF_Amount
        totBF_No = totBF_No + BF_No
        totSF_Amount = totSF_Amount + SF_Amount
        totSF_No = totSF_No + SF_No
        totSC_Amount = totSC_Amount + SC_Amount
        totSC_No = totSC_No + SC_No
                      
        
        ' check the total rows
        MoreRows (2)
        .Row = RowNum
        .Col = ColNum: .Text = Str(SlNo): .CellAlignment = 4: ColNum = ColNum + 1
        .Col = ColNum: .Text = FormatField(rstLoanDetail("BankName")): ColNum = ColNum + 1
        .Col = ColNum: .Text = FormatField(rstLoanDetail("CropName")): ColNum = ColNum + 1
        .Col = ColNum: .Text = FormatField(rstLoanDetail("LoanDate")): ColNum = ColNum + 1
        .Col = ColNum: .Text = FormatField(rstLoanDetail("LoanDueDate")): ColNum = ColNum + 1
        .Col = ColNum: .Text = BF_No: ColNum = ColNum + 1
        .Col = ColNum: .Text = FormatCurrency(BF_Amount): ColNum = ColNum + 1
        .Col = ColNum: .Text = SF_No: ColNum = ColNum + 1
        .Col = ColNum: .Text = FormatCurrency(SF_Amount): ColNum = ColNum + 1
        .Col = ColNum: .Text = SC_No: ColNum = ColNum + 1
        .Col = ColNum: .Text = FormatCurrency(SC_Amount): ColNum = ColNum + 1
        .Col = ColNum: .Text = totNumber: ColNum = ColNum + 1
        .Col = ColNum: .Text = FormatCurrency(TotAmount): ColNum = ColNum + 1
        
    End With
    
    rstLoanDetail.MoveNext
    RowNum = RowNum + 1
    SlNo = SlNo + 1
Wend

' for the grand total
With grd
    MoreRows (4)
    MoreRows (2) ' just a double check
    
       
    ' total fieidls
    grdNumber = totBF_No + totSF_No + totSC_No
    grdAmount = totBF_Amount + totSF_Amount + totSC_Amount
    
    RowNum = RowNum + 2
    .Row = RowNum
    ColNum = 0
    .Col = ColNum: .Text = "": .CellAlignment = 4: ColNum = ColNum + 1
    .Col = ColNum: .Text = "Grand Total": .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = "": .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = "": .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = "": .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = totBF_No: .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = FormatCurrency(totBF_Amount): .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = totSF_No: .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = FormatCurrency(totSF_Amount): .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = totSC_No: .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = FormatCurrency(totSC_Amount): .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = grdNumber: .CellFontBold = True: ColNum = ColNum + 1
    .Col = ColNum: .Text = FormatCurrency(grdAmount): .CellFontBold = True: ColNum = ColNum + 1
End With

ShowShed4A = True

End Function

Private Sub Shed1RowCOl()

With grd
    .Clear
    .Rows = 1: .Cols = 1
    .Cols = 17: .Rows = 10
    .FixedCols = 2: .FixedRows = 3
    .WordWrap = True: .AllowUserResizing = flexResizeBoth
    .Row = 0
    .Col = 0: .Text = GetResourceString(33)
    .Col = 1: .Text = GetResourceString(80, 60)
    .Col = 2: .Text = GetResourceString(35) '"Name of the customer"
    .Col = 3: .Text = GetResourceString(397) '"Demand"
    .Col = 4: .Text = GetResourceString(397)
    .Col = 5: .Text = GetResourceString(397)
    .Col = 6: .Text = GetResourceString(398) '"Recovery against"
    .Col = 7: .Text = GetResourceString(398) '"Recovery against"
    .Col = 8: .Text = GetResourceString(398) '"Recovery against"
    .Col = 9: .Text = GetResourceString(398) '"Recovery against"
    .Col = 10: .Text = GetResourceString(84) '"Overdue"
    .Col = 11: .Text = GetResourceString(84) '
    .Col = 12: .Text = GetResourceString(84)
    .Col = 13: .Text = GetResourceString(84)
    .Col = 14: .Text = GetResourceString(84)
    .Col = 15: .Text = GetResourceString(84)
    .Col = 16: .Text = GetResourceString(84)
    
    .Row = 1
    .Col = 0: .Text = GetResourceString(33)
    .Col = 1: .Text = GetResourceString(80, 60) '"Loan No"
    .Col = 2: .Text = GetResourceString(35) '"Name of the customer"
    .Col = 3: .Text = GetResourceString(399) '"Arrears"
    .Col = 4: .Text = GetResourceString(374) '"Current"
    .Col = 5: .Text = GetResourceString(52) & vbCrLf & "(3+4)" '"Total" & vbCrLf & "(3+4)"
    .Col = 6: .Text = GetResourceString(399, 397) ' "Arrears Demand"
    .Col = 7: .Text = GetResourceString(374, 397) '"Current Demand"
    .Col = 8: .Text = "Advance Recovery If any"
    .Col = 9: .Text = GetResourceString(52) & vbCrLf & "(6+7+8)"
    .Col = 10: .Text = GetResourceString(84, 67) & vbCrLf & "(5-(6+7))" _
            ' "Balance of Overdue" & vbCrLf & "(5-(6+7))"
    .Col = 11: .Text = "0 " & GetResourceString(107) & " 1 " & GetResourceString(108) '"Less than One year"
    .Col = 12: .Text = "1 " & GetResourceString(107) & " 2 " & GetResourceString(108) '"1 to 2 years"
    .Col = 13: .Text = "2 " & GetResourceString(107) & " 3 " & GetResourceString(108) '"2 to 3 years"
    .Col = 14: .Text = "3 " & GetResourceString(107) & " 4 " & GetResourceString(108) '"3 to 4 years"
    .Col = 15: .Text = "4 " & GetResourceString(107) & " 5 " & GetResourceString(108) '"4 to 5 years"
    .Col = 16: .Text = GetChangeString(GetResourceString(193), "5 " & GetResourceString(208)) '"Above 5 Years"
    .RowHeight(1) = 800
    
    Dim I As Integer
    Dim j As Integer
    .Row = 2
    For j = 3 To .Cols - 1
        .Col = j: .Text = Format(j, "00")
    Next
    .Col = 0: .Text = "01"
    .Col = 1: .Text = "02"
    .Col = 2: .Text = "2a"
    
    .MergeCells = flexMergeRestrictRows
    For I = 0 To .FixedRows - 1
        .Row = I
        For j = 0 To .Cols - 1
            .Col = j: .MergeCol(j) = True
            .CellFontBold = True
            .CellAlignment = 4
        Next
        .MergeRow(I) = True
    Next
    
    If m_SchemeId = 0 Then
        .ColWidth(1) = 0
        Call SaveSetting(App.EXEName, "LoanReport" & m_repType, _
                    "ColWidth" & 1, "0")
        .Row = 0: .Col = 2: .Text = GetResourceString(60) & _
                            " " & GetResourceString(35)
        .Row = 1: .Col = 2: .Text = GetResourceString(80) & _
                            " " & GetResourceString(35)
        .Row = 2: .Col = 2: .Text = GetResourceString(80) & _
                            " " & GetResourceString(35)
    Else
        If GetSetting(App.EXEName, "LoanReport" & m_repType, "ColWidth" & 1, "0") = "0" Then
            Call SaveSetting(App.EXEName, "LoanReport" & m_repType, _
                "ColWidth" & 1, 1 / grd.Cols)
        End If
    End If

End With

End Sub

Private Sub Shed2RowCol()

With grd
    .Clear
    .Rows = 1: .Cols = 1
    .Rows = 13: .Cols = 14
    .FixedCols = 2: .FixedRows = 3
    .AllowUserResizing = flexResizeBoth
    .Row = 0
    .Col = 0: .Text = GetResourceString(33)
    .Col = 1: .Text = GetResourceString(80, 60)
    .Col = 2: .Text = "Name of the Customer"
    .Col = 3: .Text = "Loan Outstanding as on 1 st July"
    .Col = 4: .Text = "Loan Advanced "
    .Col = 5: .Text = "Loan Advanced "
    .Col = 6: .Text = "Loan Advanced "
    .Col = 7: .Text = "Outstanding at the end of the month" & vbCrLf & "3 + 6"
    .Col = 8: .Text = GetResourceString(398)  '"Recovery " 'Upto Previous month"
    .Col = 9: .Text = GetResourceString(398) '"Recovery " 'during the month"
    .Col = 10: .Text = GetResourceString(398) '"Recovery " 'upto the end of the the month" & vbCrLf & "8 + 9 "
    .Col = 11: .Text = GetResourceString(52) '"Balance at the end of the month" & vbCrLf & "7-10"
    .Col = 12: .Text = "Out of Which overdue as at the end of the month"
    '.RowHeight(0) = 200
    
    .Row = 1
    .Col = 0: .Text = GetResourceString(33)
    .Col = 1: .Text = GetResourceString(80, 60)
    .Col = 2: .Text = GetResourceString(35) '"Name of the Customer"
    .Col = 3: .Text = "Loan Outstanding as on 1 st July"
    .Col = 4: .Text = "Up to the previous month"
    .Col = 5: .Text = "During the month"
    .Col = 6: .Text = "Total upto the end of the month" & vbCrLf & "(4+5)"
    .Col = 7: .Text = "Outstanding at the end of the month" & vbCrLf & "(3+6)"
    .Col = 8: .Text = "Upto Previous month"
    .Col = 9: .Text = "during the month"
    .Col = 10: .Text = "Total upto the end of the the month" & vbCrLf & "8 + 9 "
    .Col = 11: .Text = "Balance at the end of the month" & vbCrLf & "7 - 10"
    .Col = 12: .Text = "Out of Which overdue as at the end of the month"
    .RowHeight(1) = 1000
    Dim I As Integer, j As Integer
    .WordWrap = True
    .Row = 2
    For j = 0 To .Cols - 1
         .Col = j: .Text = j - 1
    Next
    .Col = 0: .Text = "1"
    .Col = 1: .Text = "2"
    .Col = 2: .Text = "2a"
'    .Col = 3: .Text = "2b"
    For I = 0 To .FixedRows - 1
        .Row = I
        For j = 0 To .Cols - 1
            .Col = j
            .CellAlignment = 4
            .CellFontBold = True
        Next
    Next
    
End With

End Sub

Private Sub Shed2RowColKan()

With grd
    .Clear
    .Rows = 1: .Cols = 1
    .Rows = 13: .Cols = 14
    .FixedCols = 2: .FixedRows = 2
    .AllowUserResizing = flexResizeBoth
    .Row = 0
    .Col = 0: .Text = GetResourceString(33)
    .Col = 1: .Text = GetResourceString(80, 60)
    .Col = 2: .Text = GetResourceString(35) '"Name of the Customer"
    .Col = 3: .Text = GetResourceString(58) & " " & _
        GetResourceString(67) & " " & _
        GetFromDateString(GetMonthString(7) & " 1 ")  '"Loan Outstanding as on 1 st July"
    .Col = 4: .Text = GetResourceString(290, 250) & _
            GetResourceString(192) '"Loan Advanced Up to the previous month"
    .Col = 5: .Text = GetResourceString(290) & _
            GetResourceString(374, 192) '"Loan Advanced during the month"
    .Col = 6: .Text = GetResourceString(52, 290) _
            & vbCrLf & "4 + 5)" '"Total Loan Advanced upto the end of the month" & vbCrLf & "(4 + 5)"
    .Col = 7: .Text = GetResourceString(58) & _
        GetResourceString(192) & vbCrLf & "(3 + 6)" ' " " "Outstanding at the end of the month" & vbCrLf & "3 + 6"
    .Col = 8: .Text = GetResourceString(20) & " " & _
        GetFromDateString(GetResourceString(250, 192)) '"Recovery Upto Previous month"
    .Col = 9: .Text = GetResourceString(20, 374) & _
        GetResourceString(192) & " " '"Recovery during the month"
    .Col = 10: .Text = GetResourceString(52, 20) _
            & vbCrLf & "(8 + 9)" '"Total Recovery upto the end of the the month" & vbCrLf & "8 + 9 "
    .Col = 11: .Text = "Balance at the end of the month" & vbCrLf & "7 - 10"
    .Col = 12: .Text = "Out of Which overdue as at the end of the month"
    .RowHeight(0) = 1200
    Dim I As Integer, j As Integer
    .WordWrap = True
    .Row = 1
    For j = 0 To .Cols - 1
         .Col = j: .Text = j - 1
    Next
    .Col = 0: .Text = "1"
    .Col = 1: .Text = "2"
    .Col = 2: .Text = "2a"
'    .Col = 3: .Text = "2b"
    For I = 0 To .FixedRows - 1
        .Row = I
        For j = 0 To .Cols - 1
            .Col = j
            .CellAlignment = 4
            .CellFontBold = True
        Next
    Next
    
End With

End Sub


Private Sub shGridCols(FCol As Integer, FRow As Integer, HeadArray() As String, FirstHead() As String, Optional LExcel As Boolean, Optional lSlNo As Boolean)
          
Dim ColNum As Integer
Dim RowNum As Integer

With grd
    .Clear
    .Cols = UBound(HeadArray) + 1
    .Rows = 20
    .FixedCols = FCol
    .FixedRows = FRow
    .AllowUserResizing = flexResizeBoth
    .WordWrap = True
    .Row = 0
    RowNum = .Row
    .Col = 0
    
    
    If UBound(FirstHead()) >= 0 And Len(FirstHead(0)) > 0 Then ' to avoid blank first header
        For ColNum = LBound(FirstHead) To UBound(FirstHead)
            .Col = ColNum: .Text = FirstHead(ColNum): .CellFontBold = True: .CellAlignment = 4
            If LExcel Then
                With xlWorkSheet
                    .cells(RowNum + 1, ColNum + 1) = FirstHead(ColNum)
                    .cells(RowNum + 1, ColNum + 1).Font.Bold = True
                End With
            End If
        Next
        RowNum = RowNum + 1
        .Row = RowNum
    End If
                    
                
    For ColNum = LBound(HeadArray) To UBound(HeadArray)
        .Col = ColNum: .Text = HeadArray(ColNum, 0): .CellFontBold = True: .CellAlignment = 4
        .ColWidth(ColNum) = Val(HeadArray(ColNum, 1))
        If LExcel Then
            With xlWorkSheet
                .cells(RowNum + 1, ColNum + 1) = HeadArray(ColNum, 0)
                .cells(RowNum + 1, ColNum + 1).Font.Bold = True
            End With
        End If
    Next
    
    Dim SlNo As Integer
    
    If lSlNo Then
        RowNum = RowNum + 1
        .Row = RowNum
        .Col = 0
        SlNo = 0
        For ColNum = LBound(HeadArray) To UBound(HeadArray)
            If ColNum <> 2 Then SlNo = SlNo + 1
            .Col = ColNum: .Text = SlNo: .CellFontBold = True: .CellAlignment = 4
            If LExcel Then
                With xlWorkSheet
                    .cells(RowNum + 1, ColNum + 1) = ColNum + 1
                    .cells(RowNum + 1, ColNum + 1).Font.Bold = True
                End With
            End If
        Next
        .Col = 2: .Text = "2a"
    End If
End With
          

End Sub

Private Sub GridRows(HeadArray() As String, Optional LExcel As Boolean)
          
Dim ColNum As Integer
Dim RowNum As Integer
Dim Items As Integer
With grd
    .AllowUserResizing = flexResizeBoth
    .WordWrap = True
    .Row = .FixedRows
    .Visible = True
    
    RowNum = .Row
    .Col = 0
    For Items = LBound(HeadArray) To UBound(HeadArray)
         If .Rows < RowNum + 2 Then .Rows = RowNum + 4
        .Row = RowNum: .Text = HeadArray(Items, 0): .CellFontBold = True: .CellAlignment = 0
        .RowHeight(RowNum) = Val(HeadArray(Items, 1))
        If LExcel Then
            With xlWorkSheet
                .cells(RowNum + 1, ColNum + 1) = HeadArray(RowNum)
                .cells(RowNum + 1, ColNum + 1).Font.Bold = True
            End With
        End If
        RowNum = RowNum + 1
    Next Items
    
End With

End Sub

Private Sub shGridRows(HeadArray() As String, Optional LExcel As Boolean)
          
Dim ColNum As Integer
Dim RowNum As Integer

With grd
    .AllowUserResizing = flexResizeBoth
    .WordWrap = True
    .Row = 0
    .Rows = UBound(HeadArray) + 1
    '.RowHeight(.Row) = 800
    '.ColWidth(1) = 3000
    RowNum = .Row
    .Col = 0
    For RowNum = LBound(HeadArray) To UBound(HeadArray)
        .Row = RowNum: .Text = HeadArray(RowNum, 0): .CellFontBold = True: .CellAlignment = 0
        .RowHeight(RowNum) = Val(HeadArray(RowNum, 1))
        
        If LExcel Then
            With xlWorkSheet
                .cells(RowNum + 1, ColNum + 1) = HeadArray(RowNum)
                .cells(RowNum + 1, ColNum + 1).Font.Bold = True
            End With
        End If
    Next RowNum
    
End With

End Sub

Private Sub Shed6RowCol()

Dim StrFirst As String
Dim strLast As String

'Get The First day Of this month & Lasrt Day OF this month
StrFirst = GetAppFirstDate(m_ToIndianDate)
strLast = GetAppLastDate(m_ToIndianDate)

With grd
    .Clear
    .Rows = 1: .Cols = 1
    .Rows = 10: .Cols = 16
    .FixedCols = 2: .FixedRows = 3
    .AllowUserResizing = flexResizeBoth
    
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) '"slno"
    .Col = 1: .Text = GetResourceString(80, 60) '"Loan No"
    .Col = 2: .Text = GetResourceString(35) '"Name of Customer"
    .Col = 3: .Text = GetResourceString(340) '"Date of Sanction"
    .Col = 4: .Text = GetResourceString(209) '"Due Date"
    .Col = 5: .Text = GetResourceString(67) & " " & GetFromDateString(StrFirst)  '  "Balance at the Begining of the Month"
    .Col = 6: .Text = GetMonthString(Month(m_ToDate))  '"During the Month"
    .Col = 7: .Text = GetMonthString(Month(m_ToDate))  '"During the Month"
    .Col = 8: .Text = GetResourceString(67) & " " & GetFromDateString(strLast) '"Outstanding at the End of the Month"
    .Col = 9: .Text = GetResourceString(84, 58)  '"of which Over Dues"
    .Col = 10: .Text = GetResourceString(84, 58) ' "Classification of Overdue"
    .Col = 11: .Text = GetResourceString(84, 58)
    .Col = 12: .Text = GetResourceString(84, 58)
    .Col = 13: .Text = GetResourceString(84, 58)
    .Col = 14: .Text = GetResourceString(84, 58)
    .Col = 15: .Text = GetResourceString(84, 58)
    
    .Row = 1
    .Col = 0: .Text = GetResourceString(33) '"slno"
    .Col = 1: .Text = GetResourceString(80, 60) '"Loan No"
    .Col = 2: .Text = GetResourceString(35) '"Name of Customer"
    .Col = 3: .Text = GetResourceString(340) '"Date of Sanction"
    .Col = 4: .Text = GetResourceString(209) '"Due Date"
    .Col = 5: .Text = GetResourceString(67) & " " & GetFromDateString(StrFirst) '"Balance at the Begining of the Month"
    .Col = 6: .Text = GetResourceString(289) 'Advance
    .Col = 7: .Text = GetResourceString(20)  'Repayments
    .Col = 8: .Text = GetResourceString(67) & " " & GetFromDateString(strLast) '"Outstanding at the End of the Month"
    .Col = 9: .Text = GetResourceString(52) '"of which Over Dues"
    .Col = 10: .Text = GetFromDateString("0 ", "1 ") & GetResourceString(208)   '"Under 1 Year"
    .Col = 11: .Text = GetFromDateString("1 ", "2 ") & GetResourceString(208) '"1 to 2 Years"
    .Col = 12: .Text = GetFromDateString("2 ", "3 ") & GetResourceString(208) '"2 to 3 Years"
    .Col = 13: .Text = GetFromDateString("3 ", "4 ") & GetResourceString(208) '"3 to 4 Years"
    .Col = 14: .Text = GetFromDateString("4 ", "5 ") & GetResourceString(208) '"4 to 5 Years"
    .Col = 15: .Text = GetChangeString(GetResourceString(193), "5 " & GetResourceString(208))  '"Above 5 Years"

    .RowHeight(1) = 700
    
    Dim I As Integer, j As Integer
    .Row = 2
    For I = 3 To .Cols - 1
        .Col = I: .Text = (I)
    Next
    .Col = 0: .Text = "1"
    .Col = 1: .Text = "2"
    .Col = 2: .Text = "2a"
    
    .MergeCells = flexMergeFree
    For I = 0 To .FixedRows - 1
        .Row = I
        For j = 0 To .Cols - 1
             .MergeCol(j) = True
            .Col = j: .CellAlignment = 4: .CellFontBold = True
        Next
        .MergeRow(I) = True
    Next
    .MergeCol(6) = False: .MergeCol(7) = False
    .MergeCol(11) = False: .MergeCol(12) = False
    .MergeCol(13) = False: .MergeCol(14) = False
    .MergeCol(15) = False ':.MergeCol(10) = False
    .WordWrap = True
    .AllowUserResizing = flexResizeBoth
End With


End Sub

Private Function Shed5RowCol() As Boolean

With grd
    .Clear
    .Rows = .Cols = 1
    .Rows = 10: .Cols = 19
    .FixedCols = 2: .FixedRows = 3
    .Row = 0
    .Col = 0: .Text = "Slno"
    .Col = 1: .Text = "Loan No"
    .Col = 2: .Text = "Name of Customer"
    .Col = 3: .Text = "Limit Sanctioned"
    .Col = 4: .Text = "Date of Sanction"
    .Col = 5: .Text = "Due Date"
    .Col = 6: .Text = "Purpose"
    .Col = 7: .Text = "Balance at the Begining of the Month"
    .Col = 8: .Text = "During the Month"
    .Col = 9: .Text = "During the Month"
    .Col = 10: .Text = "Outstanding at the End of the Month"
    .Col = 11: .Text = "Maximum Outstanding During the Month"
    .Col = 12: .Text = "of which Over Dues"
    .Col = 13: .Text = "Calssification Over Dues"
    .Col = 14: .Text = "Calssification Over Dues"
    .Col = 15: .Text = "Calssification Over Dues"
    .Col = 16: .Text = "Calssification Over Dues"
    .Col = 17: .Text = "Calssification Over Dues"
    .Col = 18: .Text = "Calssification Over Dues"
    
    .Row = 1
    .Col = 0: .Text = "Slno"
    .Col = 1: .Text = "Loan No"
    .Col = 2: .Text = "Name of Customer"
    .Col = 3: .Text = "Limit Sanctioned"
    .Col = 4: .Text = "Date of Sanction"
    .Col = 5: .Text = "Due Date"
    .Col = 6: .Text = "Purpose"
    .Col = 7: .Text = "Balance at the Begining of the Month"
    .Col = 8: .Text = "Advances"
    .Col = 9: .Text = "Recovered "
    .Col = 10: .Text = "Outstanding at the End of the Month"
    .Col = 11: .Text = "Maximum Outstanding During the Month"
    .Col = 12: .Text = "of which Over Dues"
    '.Col = 12: .Text = "Max Outstanding during Month"
    '.Col = 13: .Text = "Of Which Overdue"
    .Col = 13: .Text = "Under 1 Year"
    .Col = 14: .Text = "1 to 2 Years"
    .Col = 15: .Text = "2 to 3 Years"
    .Col = 16: .Text = "3 to 4 Years"
    .Col = 17: .Text = "4 to 5 Years"
    .Col = 18: .Text = "Above 5 Years"
    .RowHeight(1) = 700
    
    Dim I As Integer, j As Integer
    .Row = 2
    For I = 4 To .Cols - 1
        .Col = I: .Text = (I - 1)
    Next
    .Col = 0: .Text = "1"
    .Col = 1: .Text = "2"
    .Col = 2: .Text = "2a"
    
    .MergeCells = flexMergeFree
    For I = 0 To .FixedRows - 1
        .Row = I
        For j = 0 To .Cols - 1
             .MergeCol(j) = True
             .Col = j
            .CellAlignment = 4: .CellFontBold = True
        Next
        .MergeRow(I) = True
    Next
    .MergeCol(14) = False: .MergeCol(15) = False
    .MergeCol(16) = False: .MergeCol(17) = False
    .MergeCol(18) = False ':.MergeCol(11) = False
    .WordWrap = True
    .AllowUserResizing = flexResizeBoth
End With

End Function

Private Function Shed4ARowCol() As Boolean

Dim RowNum As Integer
Dim ColNum As Integer
Dim colHeads(0 To 12) As String

'important : left half because the database is not specified.
' will be completed in the later stage.

Shed4ARowCol = False

RowNum = 0
With grd
    .Cols = 13
    .Rows = 20
    .AllowUserResizing = flexResizeBoth
    .WordWrap = True
    .Clear
    .FixedCols = 1
    .FixedRows = 3
    .Row = 0
    RowNum = 0
End With


colHeads(0) = "Sl No"
colHeads(1) = "Name of the Society"
colHeads(2) = "Crop Name"
colHeads(3) = "Sanction Date"
colHeads(4) = "Due Date"
colHeads(5) = "Big Farmers"
colHeads(6) = "Big Farmers"
colHeads(7) = "Small Farmers"
colHeads(8) = "Small Farmers"
colHeads(9) = "SC/ST Farmers"
colHeads(10) = "SC/ST Farmers"
colHeads(11) = "Total Farmers"
colHeads(12) = "Total Farmers"

Call GridCols(colHeads())


colHeads(0) = "Sl No"
colHeads(1) = "Name of the Society"
colHeads(2) = "Crop Name"
colHeads(3) = "Sanction Date"
colHeads(4) = "Due Date"
colHeads(5) = "Number"
colHeads(6) = "Amount"
colHeads(7) = "Number"
colHeads(8) = "Amount"
colHeads(9) = "Number"
colHeads(10) = "Amount"
colHeads(11) = "Number"
colHeads(12) = "Amount"


RowNum = RowNum + 1
grd.Row = RowNum
Call GridCols(colHeads(), , True)
RowNum = RowNum + 1 ' for the sl no is true


' other settings
With grd
    .MergeCells = flexMergeRestrictColumns
    .MergeRow(0) = True
    .MergeRow(1) = True
    
    .Row = 0
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    
    .Row = 1
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    
    .ColWidth(1) = 2500
    .MergeCells = flexMergeRestrictAll
End With


Shed4ARowCol = True

End Function

Private Sub GridResize(choice As String)

Dim ColWidth As Double
Dim ColCount As Integer
Dim Ratio As Double

Select Case choice
    Case "Shedule2"
        Ratio = grd.Width / grd.Cols

        grd.ColWidth(0) = 500
        grd.ColWidth(1) = 3000
        grd.ColWidth(2) = 1200
        grd.ColWidth(3) = 1200
        grd.ColWidth(4) = 1200
        grd.ColWidth(5) = 1200
        grd.ColWidth(6) = 1200
        grd.ColWidth(7) = 1200
        grd.ColWidth(8) = 1200
        grd.ColWidth(9) = 1200
        grd.ColWidth(10) = 1300
        grd.ColWidth(11) = 1300
    Case "shedule1"
        grd.ColWidth(0) = 500
        grd.ColWidth(1) = 3000
        grd.ColWidth(2) = 1200
        grd.ColWidth(3) = 1200
        grd.ColWidth(4) = 1200
        grd.ColWidth(5) = 1200
        grd.ColWidth(6) = 1200
        grd.ColWidth(7) = 1050
        grd.ColWidth(8) = 1095
        grd.ColWidth(9) = 1200
        grd.ColWidth(10) = 1125
        grd.ColWidth(11) = 1125
        grd.ColWidth(12) = 1125
        grd.ColWidth(13) = 1125
        grd.ColWidth(14) = 1125
        grd.ColWidth(15) = 1125
        
End Select

End Sub



Private Sub InitGrid()

Dim ColCount As Long
Dim Wid As Single
For ColCount = 0 To grd.Cols - 1
    Wid = GetSetting(App.EXEName, "LoanReport" & m_repType, "ColWidth" & ColCount, grd.Width / grd.Cols) * grd.Width
    If Wid >= grd.Width * 0.9 Then Wid = grd.Width / grd.Cols
    If Wid < 20 And Wid <> 0 Then Wid = grd.Width / grd.Cols * 2
    grd.ColWidth(ColCount) = Wid
Next ColCount

End Sub


Private Function ShowShed6() As Boolean

ShowShed6 = False
Err.Clear
On Error GoTo ExitLine:
RaiseEvent Initialise(0, 10)
RaiseEvent Processing("Fetching record", 0)

Dim SqlStr As String
Dim transType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes

Dim rstOpBalance As Recordset
Dim rstClBalance As Recordset
Dim rstAdvance As Recordset
Dim rstRecovery As Recordset

Dim FirstDate As Date

Dim ColAmount() As Currency
Dim GrandTotal() As Currency


'Get the First day of the Month
FirstDate = GetSysFirstDate(m_ToDate)

Dim LoanType As wis_LoanType
Dim LoanTerm As wisLoanTerm
Dim LoanCategary As wisLoanCategories

LoanTerm = wisLongTerm
LoanType = wisIndividualLoan
LoanCategary = wisNonAgriculural


RaiseEvent Processing("Fetching the record", 0.1)
'Get The LoanDetails And THier balance as on Date
SqlStr = "SELECT A.LoanID,AccNum,IssueDate,LoanDueDate," & _
    " LoanAmount,Balance,LoanPurpose,Name " & _
    " FROM LoanMaster A,LoanTrans B, QRyName C " & _
    " WHERE A.SchemeId IN (SELECT SchemeID FROM LoanScheme WHERE " & _
        " Category = " & LoanCategary & " AND (LoanType = " & LoanType & _
        " OR TermType = " & LoanTerm & "))" & _
    " AND B.LoanID = A.LoanID" & _
    " AND C.CustomerId =A.CustomerID AND TransID = (SELECT MAX(TransId) FROM " & _
        " LoanTrans D WHERE D.TransDate <= #" & m_ToDate & "# " & _
        " AND D.LoanId = A.LoanId ) " & _
    " AND (LoanClosed is NULL OR LOanClosed = 0 ) "

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstClBalance, adOpenStatic) < 1 Then Exit Function

DoEvents

RaiseEvent Processing("Fetching the record", 0.25)
If gCancel Then Exit Function

'Get The LoanDetails And THier balance as on first day of the given month
SqlStr = "SELECT LoanID,TransDate,Balance FROM LoanTrans A WHERE " & _
    " TransID = (SELECT MAX(TransId) FROM LoanTrans B WHERE B.TransDate <" & _
        " #" & FirstDate & "# AND B.LoanId = A.LoanId)"
    
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstOpBalance, adOpenStatic) < 1 Then Exit Function
DoEvents
RaiseEvent Processing("Fetching the record", 0.5)
If gCancel Then Exit Function

'Get The Advances During the MOnth
transType = wWithdraw
ContraTrans = wContraWithdraw
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans WHERE TransDate >= #" & FirstDate & "#" & _
    " AND TransDate <= #" & m_ToDate & "# AND (TransType = " & transType & _
    " OR TransType = " & ContraTrans & ") " & _
    " GROUP BY LoanID "
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstAdvance, adOpenStatic) < 1 Then Set rstAdvance = Nothing

DoEvents
RaiseEvent Processing("Fetching the record", 0.65)
If gCancel Then Exit Function


'Get The Recovery During the MOnth
transType = wDeposit
ContraTrans = wContraDeposit
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans WHERE TransDate >= #" & FirstDate & "#" & _
    " AND TransDate <= #" & m_ToDate & "#" & _
    " AND (TransType = " & transType & " OR TransType = " & ContraTrans & ") " & _
    " GROUP BY LoanID "
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstRecovery, adOpenStatic) < 1 Then Set rstRecovery = Nothing

DoEvents
RaiseEvent Processing("Fetching the record", 0.85)
If gCancel Then Exit Function

'Now Align the grid
Call Shed6RowCol

ReDim ColAmount(5 To grd.Cols - 1)
ReDim GrandTotal(5 To grd.Cols - 1)
'Now Start to writing to the grid
Dim LoanID As Long
Dim AddRow As Boolean
Dim L_clsLoan As New clsLoan
Dim PrevOD As Currency
Dim ODAmount As Currency
Dim count As Long
Dim SlNo As Long

RaiseEvent Initialise(0, rstClBalance.RecordCount)

SlNo = 0
While Not rstClBalance.EOF
    
    LoanID = FormatField(rstClBalance("LoanId"))
    rstOpBalance.MoveFirst
    rstOpBalance.Find "LoanID = " & LoanID
    ColAmount(5) = 0
    If Not rstOpBalance.EOF Then _
        ColAmount(5) = FormatField(rstOpBalance("Balance")) 'Balance as on 31/3/yyyy
    
    ColAmount(6) = 0
    If Not rstAdvance Is Nothing Then
        rstAdvance.MoveFirst
        rstAdvance.Find "LoanId = " & LoanID
        If Not rstAdvance.EOF Then _
            ColAmount(6) = FormatField(rstAdvance(0))  'Advances During the MOnth
        
    End If
    ColAmount(7) = 0
    If Not rstRecovery Is Nothing Then
        rstRecovery.MoveFirst
        rstRecovery.Find "LoanId = " & LoanID
        If Not rstRecovery.EOF Then _
            ColAmount(7) = FormatField(rstRecovery(0)) 'Recovery During the mOnth
        
    End If
    ColAmount(8) = FormatField(rstClBalance("Balance")) 'Balance at the end of month
    ODAmount = L_clsLoan.OverDueAmount(LoanID, , m_ToDate)   'Over due
    ColAmount(9) = ODAmount 'Over due amount of the loan as on given date
    
    'Over due amount calssifiacation
    PrevOD = ODAmount
    ColAmount(15) = L_clsLoan.OverDueSince(5, LoanID, , m_ToDate)
                'Over due since & above 5 Years
    ODAmount = ODAmount - ColAmount(15)
    
    ColAmount(14) = L_clsLoan.OverDueSince(4, LoanID, , m_ToDate) - ColAmount(15)
    If ColAmount(14) < 0 Then ColAmount(14) = ODAmount 'Over due since 4 Years
    ODAmount = ODAmount - ColAmount(14)
    
    ColAmount(13) = L_clsLoan.OverDueSince(3, LoanID, , m_ToDate) - ColAmount(14)
    If ColAmount(13) < 0 Then ColAmount(13) = ODAmount 'Over due since 3 Years
    ODAmount = ODAmount - ColAmount(13)
    
    ColAmount(12) = L_clsLoan.OverDueSince(2, LoanID, , m_ToDate) - ColAmount(13)
    If ColAmount(12) < 0 Then ColAmount(12) = ODAmount 'Over due since 2 Years
    ODAmount = ODAmount - ColAmount(12)
    
    ColAmount(11) = L_clsLoan.OverDueSince(1, LoanID, , m_ToDate) - ColAmount(12)
    If ColAmount(11) < 0 Then ColAmount(11) = ODAmount 'Over due since a Year
    ODAmount = ODAmount - ColAmount(11)
    
    ColAmount(10) = ODAmount 'Over due under one Year
    
    'Check whther this row has to be write or not
    AddRow = False
    For count = 5 To grd.Cols - 1
        If ColAmount(count) Then
            AddRow = True
            SlNo = SlNo + 1
            Exit For
        End If
    Next
    If AddRow Then
        With grd
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            .MergeRow(.Row) = False
            .Col = 0: .Text = SlNo
            .Col = 1: .Text = FormatField(rstClBalance("AccNum"))
            .Col = 2: .Text = FormatField(rstClBalance("Name"))
            .Col = 3: .Text = FormatField(rstClBalance("IssueDate"))
            .Col = 4: .Text = FormatField(rstClBalance("LoanDueDate"))
            For count = 5 To grd.Cols - 1
                .Col = count:
                If ColAmount(count) < 0 Then ColAmount(count) = 0
                .Text = FormatCurrency(ColAmount(count))
                GrandTotal(count) = GrandTotal(count) + ColAmount(count)
            Next
        End With
    End If
    
    DoEvents
    If gCancel Then rstClBalance.MoveLast
    RaiseEvent Processing("Writing the records", rstClBalance.AbsolutePosition / rstClBalance.RecordCount)
    rstClBalance.MoveNext
    
Wend
Set L_clsLoan = Nothing
With grd
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 2: .CellFontBold = True
    .Text = "Grand Total"
    For count = 5 To .Cols - 1
        .Col = count: .CellFontBold = True
        .Text = FormatCurrency(GrandTotal(count))
    Next
End With

ShowShed6 = True
lblReportTitle.Caption = "Statement showing the long term and other loans for the month of " & _
    GetMonthString(Month(m_ToDate)) & " " & GetFromDateString(m_ToIndianDate)

ExitLine:
If Err Then
    MsgBox "ERROR SHED 5" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
    Err.Clear
    'Resume
End If

End Function
Private Function ShowMeetingRegistar() As Boolean

If m_SchemeId = 0 Then ShowMeetingRegistarAll

Dim SqlPrin As String
Dim SqlInt As String
Dim SqlStr As String
Dim PrinRepay As Currency
Dim IntRepay As Currency

Dim rst As Recordset
Dim rstLoanScheme As Recordset
Dim SchemeName  As String
Dim Date31_3  As Date
Dim DateLastMonth As Date

'INDIAN dATE FORMAT OF ABOVE VARIABLES
Dim IndDate31_3  As String
Dim IndDateLastMonth As String

Dim transType As wisTransactionTypes
Dim LoanType As wis_LoanType
Dim SchemeStr  As String

Err.Clear
On Error GoTo ErrLine
'Get all date in the format of system format
IndDate31_3 = "31/3/" & Val(Year(m_ToDate) - IIf(Month(m_ToDate) > 3, 0, 1))
Date31_3 = GetSysFormatDate(IndDate31_3)

DateLastMonth = GetSysLastDate(DateAdd("m", -1, m_ToDate))
IndDateLastMonth = GetIndianDate(DateLastMonth)

RaiseEvent Processing("Fetching the records", 0)
m_repType = repMonthlyRegister
SqlStr = "SELECT * FROM LoanScheme Where SchemeID = " & m_SchemeId
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstLoanScheme, adOpenStatic) < 1 Then Set rstLoanScheme = Nothing

SchemeName = FormatField(rstLoanScheme("SchemeName"))
SchemeStr = " SchemeID = " & m_SchemeId & " "
LoanType = FormatField(rstLoanScheme("LoanType"))

Dim rstMaster As Recordset

Dim rstPrin31_3 As Recordset
Dim rstInt31_3 As Recordset

Dim rstPrinLastMonth As Recordset
Dim rstIntLastMonth As Recordset

Dim rstPrinAsOn As Recordset
Dim rstIntAsOn As Recordset

Dim rstPrinTransLast As Recordset
Dim rstIntTransLast As Recordset

Dim rstCurPrinTrans As Recordset
Dim rstCurIntTrans As Recordset

Dim rstPrinTransAsOn As Recordset
Dim rstIntTransAsOn As Recordset

Screen.MousePointer = vbHourglass

DoEvents
RaiseEvent Initialise(0, 10)
RaiseEvent Processing("Fetching the record", 0.1)
If gCancel Then Exit Function

'Get The details of loan
SqlStr = "SELECT A.LoanID,AccNum,IntRate,Guarantor1,Guarantor2, " & _
    " CustomerID, IssueDate,LoanAmount,OtherDets,B.AbnDate,b.EpDate " & _
    " From LoanMaster A left Join LoanAbnEp B On A.lOanID = B.LoanID WHERE " & _
    " A.LoanId IN (SELECT Distinct LoanID From LoanTrans)" & _
    " AND (B.BKCC = 0 or BKCC is NULL)" & _
    " AND SchemeID = " & m_SchemeId

SqlStr = SqlStr & " ORDER BY SchemeID,val(AccNum)"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstMaster, adOpenStatic) <= 0 Then GoTo ErrLine
    
    
'Get the Loan balance on 31/3/yyyy
SqlPrin = "SELECT A.LoanID,Balance FROM LoanTrans A WHERE " & _
        " TransDate <= #" & Date31_3 & "#" & _
        " ORDER BY LoanId, TransID Desc"

gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstPrin31_3, adOpenStatic) < 1 Then Set rstPrin31_3 = Nothing

DoEvents
RaiseEvent Processing("Fetching the record", 0.15)
If gCancel Then Exit Function

'Get the Interest Balance 31/3/yyyy
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B " & _
    " WHERE B.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanIntTrans C WHERE TransDate <= #" & Date31_3 & "# " & _
        " AND C.LoanID = B.LoanID And C.BankID = B.BankID )"
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B WHERE " & _
    " TransDate <= #" & Date31_3 & "# ORDER BY LOanID,TransID Desc"
gDbTrans.SqlStmt = SqlInt

Call gDbTrans.Fetch(rstInt31_3, adOpenStatic) '< 1 Then Set rstInt31_3 = Nothing

'Get the Loan balance as on lastMonth
SqlPrin = "SELECT A.LoanID,TransDate,Balance FROM LoanTrans A WHERE " & _
     " A.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanTrans C WHERE TransDate <= #" & DateLastMonth & "# " & _
        " AND C.LoanID = A.LoanID )"
SqlPrin = "SELECT B.LoanID,Balance,TransDate FROM LoanTrans B WHERE " & _
        " TransDate <= #" & DateLastMonth & "# ORDER BY LOanID,TransID Desc"

'Get the Interest Balance ON  LAST MONTH
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B " & _
    " WHERE B.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanIntTrans C WHERE TransDate <= #" & DateLastMonth & "# " & _
        " AND C.LoanID = B.LoanID )"
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B WHERE " & _
        " TransDate <= #" & DateLastMonth & "# ORDER BY LOanID,TransID Desc"

gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstPrinLastMonth, adOpenStatic) < 1 Then Set rstPrinLastMonth = Nothing

DoEvents
RaiseEvent Processing("Fetching record", 0.25)
If gCancel Then Exit Function
    
gDbTrans.SqlStmt = SqlInt
If gDbTrans.Fetch(rstIntLastMonth, adOpenStatic) < 1 Then Set rstIntLastMonth = Nothing

DoEvents
RaiseEvent Processing("Fetching record", 0.35)
If gCancel Then Exit Function

'Get the Loan balance as on date
SqlPrin = "SELECT A.LoanID,TransDate,Balance FROM LoanTrans A WHERE " & _
     " A.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanTrans C WHERE TransDate <= #" & m_ToDate & "# " & _
        " AND C.LoanID = A.LoanID )"
SqlPrin = "SELECT B.LoanID,Balance,TransDate FROM LoanTrans B WHERE " & _
        " TransDate <= #" & m_ToDate & "# ORDER BY LOanID,TransID Desc"
'Get the Interest Balance ON Date
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B " & _
    " WHERE B.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanIntTrans C WHERE TransDate <= #" & m_ToDate & "# " & _
        " AND C.LoanID = B.LoanID )"
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B WHERE " & _
    " TransDate <= #" & m_ToDate & "# ORDER BY LOanID,TransID Desc"
gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstPrinAsOn, adOpenStatic) < 1 Then Set rstPrinAsOn = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.45)
If gCancel Then Exit Function

gDbTrans.SqlStmt = SqlInt
If gDbTrans.Fetch(rstIntAsOn, adOpenStatic) < 1 Then Set rstIntAsOn = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.55)
If gCancel Then Exit Function

'GEt the Transacted amount After 31/3/yyyy till last month
SqlPrin = "SELECT SUM(AMOUNT) as SumAmount,LoanID,TransType FROM LoanTrans WHERE " & _
    " TransDate > #" & Date31_3 & "# AND TransDate <= #" & DateLastMonth & "# " & _
    " GROUP BY LoanId,TransType"
SqlInt = "SELECT SUM(IntAmount) as SumIntAmount,SUM(PenalIntAmount) as SumPenalIntAmount," & _
    " LoanID,TransType FROM LoanIntTrans WHERE " & _
    " TransDate > #" & Date31_3 & "# AND TransDate <= #" & DateLastMonth & "# " & _
    " GROUP BY LoanId,TransType"

gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstPrinTransLast, adOpenStatic) < 1 Then GoTo ErrLine

DoEvents
RaiseEvent Processing("Writing the record", 0.65)
If gCancel Then Exit Function

gDbTrans.SqlStmt = SqlInt
If gDbTrans.Fetch(rstIntTransLast, adOpenStatic) < 1 Then Set rstIntTransLast = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.75)
If gCancel Then Exit Function

'GEt the Transacted amount From last month to till Today
SqlPrin = "SELECT SUM(AMOUNT)as SumAmount,LoanID,TransType FROM LoanTrans WHERE " & _
    " TransDate > #" & DateLastMonth & "# AND TransDate <= #" & m_ToDate & "# " & _
    " GROUP BY LoanId,TransType"
SqlInt = "SELECT SUM(IntAmount) as SumIntAmount,SUM(PenalIntAmount) as SumPenalIntAmount," & _
    " LoanID,TransType FROM LoanIntTrans WHERE " & _
    " TransDate > #" & DateLastMonth & "# AND TransDate <= #" & m_ToDate & "# " & _
    " GROUP BY LoanId,TransType"

gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstPrinTransAsOn, adOpenStatic) < 1 Then Set rstPrinTransAsOn = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.85)
If gCancel Then Exit Function

gDbTrans.SqlStmt = SqlInt
If gDbTrans.Fetch(rstIntTransAsOn, adOpenStatic) < 1 Then Set rstIntTransAsOn = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.95)
If gCancel Then Exit Function

'Now Initialise the grid
grd.Clear
grd.Cols = 23
grd.Rows = 20

Dim SlNo As Integer
Dim LoanID As Long
Dim L_clsCust As New clsCustReg
Dim L_clsLoan As New clsLoan
Dim retstr As String
Dim strArr() As String
Dim TransDate As Date
Dim IntRate As Single
Dim Amount As Currency
Dim IntAmount As Currency
Dim Balance As Currency
Dim PrevDate As Date
Dim Balance31_3 As Currency
Dim BalanceLastMonth As Currency
Dim BalanceNow As Currency
Dim IntBal31_3 As Currency
Dim IntBalLastMonth As Currency
Dim IntBalNow As Currency

Dim ODAmount As Currency
Dim ODInt As Currency

Call SetGrid(m_SchemeId)

Call InitGrid
RaiseEvent Initialise(0, rstMaster.RecordCount)
Dim rowno As Long

grd.Row = grd.FixedRows
rowno = grd.Row
lblReportTitle = "Meeting register As on " & m_FromIndianDate
If LoanType <> wisVehicleloan Then
    grd.ColWidth(7) = 0: grd.ColWidth(8) = 0:
Else
    grd.ColWidth(7) = grd.Width / grd.Cols: grd.ColWidth(8) = grd.Width / grd.Cols
End If

Call grd_LostFocus
    
lblReportTitle = "Meeting register of " & SchemeName & " As on " & m_FromIndianDate

Do
    If rstMaster.EOF Then Exit Do
    
    Balance31_3 = 0: BalanceLastMonth = 0: BalanceNow = 0
    IntBal31_3 = 0: IntBalLastMonth = 0: IntBalNow = 0
    LoanID = FormatField(rstMaster("LoanID"))
    IntRate = FormatField(rstMaster("IntRate"))
    SlNo = SlNo + 1
  With grd
    rowno = rowno + 1
    If .Rows <= rowno Then .Rows = rowno + 2
    .Row = .Row + 1
    .MergeRow(rowno) = False
    .TextMatrix(rowno, 0) = Format(SlNo, "00")
    .TextMatrix(rowno, 1) = FormatField(rstMaster("AccNum"))
    .TextMatrix(rowno, 2) = L_clsCust.CustomerName(FormatField(rstMaster("CustomerID")))
    retstr = FormatField(rstMaster("Guarantor1"))
    On Error Resume Next
    If Val(retstr) > 0 Then
        .TextMatrix(rowno, 3) = L_clsCust.CustomerName(Val(retstr))
    End If
    retstr = FormatField(rstMaster("Guarantor2"))
    If Val(retstr) > 0 Then
        .TextMatrix(rowno, 4) = L_clsCust.CustomerName(Val(retstr))
    End If
    On Error GoTo ErrLine
    'Loan Advance Deatails
    .TextMatrix(rowno, 5) = FormatField(rstMaster("IssueDate"))
    .TextMatrix(rowno, 6) = FormatField(rstMaster("LoanAmount"))
    
    'Incase of vehicle loan Reg No & Insur details
    retstr = FormatField(rstMaster("otherDets"))
    Call GetStringArray(retstr, strArr, gDelim)
    If LoanType = wisVehicleloan Then
        ReDim Preserve strArr(1)
        On Error Resume Next
        .TextMatrix(rowno, 7) = strArr(0)
        .TextMatrix(rowno, 8) = strArr(1)
        On Error GoTo ErrLine
    End If
    
    'Out standing Loan Balance as on 31/3/yyyy
    PrevDate = Date31_3
    Balance31_3 = 0: IntBal31_3 = 0
    If Not rstPrin31_3 Is Nothing Then
        rstPrin31_3.MoveFirst
        rstPrin31_3.Find " LoanID = " & LoanID
        If Not rstPrin31_3.EOF Then
            If rstPrin31_3("LoanID") = LoanID Then
                rstInt31_3.MoveFirst
                rstInt31_3.Find " LoanID = " & LoanID
                Balance31_3 = rstPrin31_3("Balance")
                On Error Resume Next
                If rstInt31_3.EOF Then Exit Function
                TransDate = rstInt31_3("TransDate")
                IntBal31_3 = FormatField(rstInt31_3("IntBalance"))
                PrevDate = TransDate
            End If
        End If
    End If
    IntBal31_3 = IntBal31_3 + L_clsLoan.RegularInterest(LoanID, , Date31_3)
    .TextMatrix(rowno, 9) = Balance31_3
    
    'Over due as on 31/3/yyyy
    ODAmount = L_clsLoan.OverDueAmount(LoanID, , Date31_3)
    ODInt = L_clsLoan.OverDueInterest(LoanID, Date31_3)
    .TextMatrix(rowno, 10) = L_clsLoan.DueInstallments(LoanID, Date31_3)
    .TextMatrix(rowno, 11) = FormatCurrency(ODAmount)
    .TextMatrix(rowno, 12) = FormatCurrency(ODInt)
    .TextMatrix(rowno, 13) = FormatCurrency(ODAmount + ODInt)
    
    'Loan Repayment From 1/4/yyyy to last month
    transType = wDeposit
    Amount = 0: IntAmount = 0
    If Not rstPrinTransLast Is Nothing Then
      If rstPrinTransLast.EOF Then GoTo ErrLine
        rstPrinTransLast.MoveNext
        Call gDbTrans.FindRecord(rstPrinTransLast, " LoanID = " & LoanID & " ,TransType = " & transType)
        If Not rstPrinLastMonth.EOF Then
            If rstPrinLastMonth("loanID") = LoanID Then
                Amount = FormatField(rstPrinTransLast("SumAmount"))
                If Not rstIntTransLast Is Nothing Then
                    rstIntTransLast.Find " LoanID = " & LoanID & " AND TransType = " & transType
                    IntAmount = FormatField(rstIntTransLast("SumIntAmount"))
                    IntAmount = IntAmount + FormatField(rstIntTransLast("SumPenalIntAmount"))
                End If
            End If
        End If
    End If
    '.Col= 14:
    .TextMatrix(rowno, 15) = FormatCurrency(Amount)
    .TextMatrix(rowno, 16) = FormatCurrency(IntAmount)
    .TextMatrix(rowno, 17) = FormatCurrency(Amount + IntAmount)  'Out standing Loan Balance as on last month
    
    'LOan Balance as on last month
    Amount = 0: IntAmount = 0
    If Not rstPrinLastMonth Is Nothing Then
        rstPrinLastMonth.Find " LoanID = " & LoanID
        If Not rstPrinLastMonth.EOF Then
            If rstPrinLastMonth("loanID") = LoanID Then
                Amount = FormatField(rstPrinLastMonth("Balance"))
                IntAmount = 0
                If Not rstIntLastMonth Is Nothing Then
                    rstIntLastMonth.Find " LoanID = " & LoanID
                    IntAmount = FormatField(rstIntLastMonth("IntBalance"))
                End If
                '.Textmatrix(rowNo,18)= IntBalLastMonth
            End If
        End If
    End If
    IntAmount = IntAmount + L_clsLoan.RegularInterest(LoanID, , DateLastMonth)
    .TextMatrix(rowno, 19) = FormatCurrency(Amount)
    .TextMatrix(rowno, 20) = FormatCurrency(IntAmount)
    .TextMatrix(rowno, 21) = FormatCurrency(Amount + IntAmount)
    
    'Over due as on last month
    ODAmount = L_clsLoan.OverDueAmount(LoanID, , DateLastMonth)
    ODInt = L_clsLoan.PenalInterest(LoanID, , DateLastMonth)
    .TextMatrix(rowno, 22) = L_clsLoan.DueInstallments(LoanID, DateLastMonth)
    .TextMatrix(rowno, 23) = FormatCurrency(ODAmount)
    .TextMatrix(rowno, 24) = FormatCurrency(ODInt)
    .TextMatrix(rowno, 25) = FormatCurrency(ODAmount + ODInt)
    
    'If Case has filed Then Date Of Filing the case
    .TextMatrix(rowno, 26) = FormatField(rstMaster("ABNDate"))
    .TextMatrix(rowno, 27) = FormatField(rstMaster("EpDate"))
  End With
    DoEvents
    RaiseEvent Processing("Writing the record", rstMaster.AbsolutePosition / (rstMaster.RecordCount * 2))
    If gCancel Then rstMaster.MoveLast
    rstMaster.MoveNext
Loop
Set L_clsLoan = Nothing

ShowMeetingRegistar = True
Screen.MousePointer = vbDefault

Exit Function

ErrLine:
    Screen.MousePointer = vbDefault
    If Err Then
        MsgBox Err.Number & vbCrLf & Err.Description, , wis_MESSAGE_TITLE
       Resume
        Exit Function
    End If

End Function


Private Function ShowMeetingRegistarAll() As Boolean

Dim SqlPrin As String
Dim SqlInt As String
Dim SqlStr As String
Dim PrinRepay As Currency
Dim IntRepay As Currency

Dim rst As Recordset
Dim rstLoanScheme As Recordset
Dim SchemeName  As String
Dim Date31_3  As Date
Dim DateLastMonth As Date

'INDIAN dATE FORMAT OF ABOVE VARIABLES
Dim IndDate31_3  As String
Dim IndDateLastMonth As String

Dim transType As wisTransactionTypes
Dim LoanType As wis_LoanType
Dim SchemeStr  As String

Err.Clear
On Error GoTo ErrLine
'Get all date in the format of system format
IndDate31_3 = "31/3/" & Val(Year(m_ToDate) - IIf(Month(m_ToDate) > 3, 0, 1))
Date31_3 = GetSysFormatDate(IndDate31_3)

DateLastMonth = GetSysLastDate(DateAdd("m", -1, m_ToDate))
IndDateLastMonth = GetIndianDate(DateLastMonth)

RaiseEvent Processing("Fetching the records", 0)
m_repType = repMonthlyRegisterAll
SchemeStr = " SchemeID  <> " & m_SchemeId & " "

Dim rstMaster As Recordset

Dim rstPrin31_3 As Recordset
Dim rstInt31_3 As Recordset

Dim rstPrinLastMonth As Recordset
Dim rstIntLastMonth As Recordset

Dim rstPrinAsOn As Recordset
Dim rstIntAsOn As Recordset

Dim rstPrinTransLast As Recordset
Dim rstIntTransLast As Recordset

Dim rstCurPrinTrans As Recordset
Dim rstCurIntTrans As Recordset

Dim rstPrinTransAsOn As Recordset
Dim rstIntTransAsOn As Recordset

Screen.MousePointer = vbHourglass

DoEvents
RaiseEvent Initialise(0, 10)
RaiseEvent Processing("Fetching the record", 0.1)
If gCancel Then Exit Function

'Get The details of loan
SqlStr = "SELECT A.LoanID,AccNum,IntRate,Guarantor1,Guarantor2, " & _
    " CustomerID, IssueDate,LoanAmount,OtherDets,B.AbnDate,b.EpDate " & _
    " From LoanMaster A left Join LoanAbnEp B On A.lOanID = B.LoanID WHERE " & _
    " A.LoanId IN (SELECT Distinct LoanID From LoanTrans)" & _
    " AND (B.BKCC = 0 or BKCC is NuLL)"

SqlStr = SqlStr & " ORDER BY SchemeID,val(AccNum)"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstMaster, adOpenStatic) <= 0 Then GoTo ErrLine
       
'Get the Loan balance on 31/3/yyyy
SqlPrin = "SELECT A.LoanID,Balance FROM LoanTrans A WHERE " & _
        " TransDate <= #" & Date31_3 & "#" & _
        " ORDER BY LoanId, TransID Desc"

gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstPrin31_3, adOpenStatic) < 1 Then Set rstPrin31_3 = Nothing

DoEvents
RaiseEvent Processing("Fetching the record", 0.15)
If gCancel Then Exit Function

'Get the Interest Balance 31/3/yyyy
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B " & _
    " WHERE B.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanIntTrans C WHERE TransDate <= #" & Date31_3 & "# " & _
        " AND C.LoanID = B.LoanID And C.BankID = B.BankID )"
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B WHERE " & _
    " TransDate <= #" & Date31_3 & "# ORDER BY LOanID,TransID Desc"
gDbTrans.SqlStmt = SqlInt

Call gDbTrans.Fetch(rstInt31_3, adOpenStatic) '< 1 Then Set rstInt31_3 = Nothing

'Get the Loan balance as on lastMonth
SqlPrin = "SELECT A.LoanID,TransDate,Balance FROM LoanTrans A WHERE " & _
     " A.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanTrans C WHERE TransDate <= #" & DateLastMonth & "# " & _
        " AND C.LoanID = A.LoanID )"
SqlPrin = "SELECT B.LoanID,Balance,TransDate FROM LoanTrans B WHERE " & _
        " TransDate <= #" & DateLastMonth & "# ORDER BY LOanID,TransID Desc"

'Get the Interest Balance ON  LAST MONTH
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B " & _
    " WHERE B.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanIntTrans C WHERE TransDate <= #" & DateLastMonth & "# " & _
        " AND C.LoanID = B.LoanID )"
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B WHERE " & _
        " TransDate <= #" & DateLastMonth & "# ORDER BY LOanID,TransID Desc"

gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstPrinLastMonth, adOpenStatic) < 1 Then Set rstPrinLastMonth = Nothing

DoEvents
RaiseEvent Processing("Fetching record", 0.25)
If gCancel Then Exit Function
    
gDbTrans.SqlStmt = SqlInt
If gDbTrans.Fetch(rstIntLastMonth, adOpenStatic) < 1 Then Set rstIntLastMonth = Nothing

DoEvents
RaiseEvent Processing("Fetching record", 0.35)
If gCancel Then Exit Function

'Get the Loan balance as on date
SqlPrin = "SELECT A.LoanID,TransDate,Balance FROM LoanTrans A WHERE " & _
     " A.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanTrans C WHERE TransDate <= #" & m_ToDate & "# " & _
        " AND C.LoanID = A.LoanID )"
SqlPrin = "SELECT B.LoanID,Balance,TransDate FROM LoanTrans B WHERE " & _
        " TransDate <= #" & m_ToDate & "# ORDER BY LOanID,TransID Desc"
'Get the Interest Balance ON Date
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B " & _
    " WHERE B.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanIntTrans C WHERE TransDate <= #" & m_ToDate & "# " & _
        " AND C.LoanID = B.LoanID )"
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B WHERE " & _
    " TransDate <= #" & m_ToDate & "# ORDER BY LOanID,TransID Desc"
gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstPrinAsOn, adOpenStatic) < 1 Then Set rstPrinAsOn = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.45)
If gCancel Then Exit Function

gDbTrans.SqlStmt = SqlInt
If gDbTrans.Fetch(rstIntAsOn, adOpenStatic) < 1 Then Set rstIntAsOn = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.55)
If gCancel Then Exit Function

'GEt the Transacted amount After 31/3/yyyy till last month
SqlPrin = "SELECT SUM(AMOUNT) as SumAmount,LoanID,TransType FROM LoanTrans WHERE " & _
    " TransDate > #" & Date31_3 & "# AND TransDate <= #" & DateLastMonth & "# " & _
    " GROUP BY LoanId,TransType"
SqlInt = "SELECT SUM(IntAmount) as SumIntAmount,SUM(PenalIntAmount) as SumPenalIntAmount," & _
    " LoanID,TransType FROM LoanIntTrans WHERE " & _
    " TransDate > #" & Date31_3 & "# AND TransDate <= #" & DateLastMonth & "# " & _
    " GROUP BY LoanId,TransType"

gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstPrinTransLast, adOpenStatic) < 1 Then GoTo ErrLine

DoEvents
RaiseEvent Processing("Writing the record", 0.65)
If gCancel Then Exit Function

gDbTrans.SqlStmt = SqlInt
If gDbTrans.Fetch(rstIntTransLast, adOpenStatic) < 1 Then Set rstIntTransLast = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.75)
If gCancel Then Exit Function

'Get the Transacted amount From last month to till Today
SqlPrin = "SELECT SUM(AMOUNT)as SumAmount,LoanID,TransType FROM LoanTrans WHERE " & _
    " TransDate > #" & DateLastMonth & "# AND TransDate <= #" & m_ToDate & "# " & _
    " GROUP BY LoanId,TransType"
SqlInt = "SELECT SUM(IntAmount) as SumIntAmount,SUM(PenalIntAmount) as SumPenalIntAmount," & _
    " LoanID,TransType FROM LoanIntTrans WHERE " & _
    " TransDate > #" & DateLastMonth & "# AND TransDate <= #" & m_ToDate & "# " & _
    " GROUP BY LoanId,TransType"

gDbTrans.SqlStmt = SqlPrin
If gDbTrans.Fetch(rstPrinTransAsOn, adOpenStatic) < 1 Then Set rstPrinTransAsOn = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.85)
If gCancel Then Exit Function

gDbTrans.SqlStmt = SqlInt
If gDbTrans.Fetch(rstIntTransAsOn, adOpenStatic) < 1 Then Set rstIntTransAsOn = Nothing

DoEvents
RaiseEvent Processing("Writing the record", 0.95)
If gCancel Then Exit Function

'Now Initialise the grid
grd.Clear
grd.Cols = 23
grd.Rows = 20

Dim SlNo As Integer
Dim LoanID As Long
Dim L_clsCust As New clsCustReg
Dim L_clsLoan As New clsLoan
Dim retstr As String
Dim strArr() As String
Dim TransDate As Date
Dim IntRate As Single
Dim Amount As Currency
Dim IntAmount As Currency
Dim Balance As Currency
Dim PrevDate As Date
Dim Balance31_3 As Currency
Dim BalanceLastMonth As Currency
Dim BalanceNow As Currency
Dim IntBal31_3 As Currency
Dim IntBalLastMonth As Currency
Dim IntBalNow As Currency

Dim ODAmount As Currency
Dim ODInt As Currency

Call SetGrid(m_SchemeId)

Call InitGrid
RaiseEvent Initialise(0, rstMaster.RecordCount)
Dim rowno As Long

grd.Row = grd.FixedRows
rowno = grd.Row
lblReportTitle = "Meeting register As on " & m_FromIndianDate
'*********

RaiseEvent Initialise(0, rstMaster.RecordCount)
Do
    If rstMaster.EOF Then Exit Do
    Balance31_3 = 0: BalanceLastMonth = 0: BalanceNow = 0
    IntBal31_3 = 0: IntBalLastMonth = 0: IntBalNow = 0
    LoanID = FormatField(rstMaster("LoanID"))
    IntRate = FormatField(rstMaster("IntRate"))
    SlNo = SlNo + 1
  With grd
    If .Rows < .Row + 3 Then .Rows = .Rows + 3
    .Row = .Row + 1
    .MergeRow(.Row) = False
    .Col = 0: .Text = Format(SlNo, "00")
    .Col = 1: .Text = FormatField(rstMaster("AccNum"))
    .Col = 2: .Text = L_clsCust.CustomerName(FormatField(rstMaster("CustomerID")))
    retstr = FormatField(rstMaster("Guarantor1"))
    On Error Resume Next
    If Val(retstr) Then
        .Col = 3: .Text = L_clsCust.CustomerName(Val(retstr))
        .Row = .Row + 1: .Text = "": .Row = .Row - 1
    End If
    retstr = FormatField(rstMaster("Guarantor2"))
    If Val(retstr) Then
        .Col = 4: .Text = L_clsCust.CustomerName(Val(retstr))
        .Row = .Row + 1: .Text = "": .Row = .Row - 1
    End If
    On Error GoTo ErrLine
    'Loan Advance Date
    .Col = 5: .Text = FormatField(rstMaster("IssueDate"))
    .Col = 6: .Text = FormatField(rstMaster("LoanAmount"))
    
    'Loan Balance as on 31/3
    PrevDate = Date31_3
    Balance31_3 = 0: IntBal31_3 = 0
    
    If Not rstPrin31_3 Is Nothing Then
        rstPrin31_3.Find " LoanID = " & LoanID
        If Not rstPrin31_3.EOF Then
            If rstPrin31_3("LoanID") = LoanID Then
                rstInt31_3.Find " LoanID = " & LoanID
                Balance31_3 = FormatField(rstPrin31_3("Balance"))
               ' Transdate = FormatField(rstInt31_3("TransDate"))
                IntBal31_3 = FormatField(rstInt31_3("IntBalance"))
                PrevDate = TransDate
            End If
        End If
    End If
    IntBal31_3 = IntBal31_3 + L_clsLoan.RegularInterest(LoanID, , Date31_3)
    IntBal31_3 = IntBal31_3 + L_clsLoan.PenalInterest(LoanID, , Date31_3)
    .Col = 7: .Text = Balance31_3
    .Col = 8: .Text = IntBal31_3
    .Col = 9: .Text = Val(Balance31_3 + IntBal31_3)
            
    'Recovery from 1/4/yyyy to LastMonth
    PrinRepay = 0: IntRepay = 0: transType = wDeposit
    If Not rstPrinTransLast Is Nothing Then
       On Error Resume Next
        rstPrinTransLast.Find "LoanID = " & LoanID '& " AND Transtype = " & TransType
        If (Not rstPrinTransLast.EOF) Then
         If rstPrinTransLast("LoanID") = LoanID Then
            rstIntTransLast.Find "LoanID = " & LoanID ' & " AND Transtype = " & TransType
            PrinRepay = FormatField(rstPrinTransLast("SumAmount"))
            IntRepay = FormatField(rstIntTransLast("SumIntAmount"))
         End If
        End If
    End If
    .Col = 10: .Text = FormatCurrency(PrinRepay)
    .Col = 11: .Text = FormatCurrency(IntRepay)
    .Col = 12: .Text = FormatCurrency(PrinRepay + IntRepay)
            
    'Loan Balance as on end of last month
    BalanceLastMonth = Balance31_3: IntBalLastMonth = 0
    If Not rstPrinLastMonth Is Nothing Then
        rstPrinLastMonth.Find " LoanID = " & LoanID
        If Not rstPrinLastMonth.EOF Then
            If rstPrinLastMonth("LoanID") = LoanID Then
                rstIntLastMonth.Find " LoanID = " & LoanID
                BalanceLastMonth = rstPrinLastMonth("Balance")
                TransDate = rstPrinLastMonth("TransDate")
                IntBalLastMonth = FormatField(rstIntLastMonth("IntBalance"))
                PrevDate = TransDate
                PrevDate = TransDate
            End If
        End If
    End If
    IntBalLastMonth = IntBalLastMonth + L_clsLoan.RegularInterest(LoanID, , DateLastMonth)
    .Col = 13: .Text = BalanceLastMonth
    .Col = 14: .Text = IntBalLastMonth
    .Col = 15: .Text = Val(BalanceLastMonth + IntBalLastMonth)
    
    .Col = 10: .Text = Balance31_3 - BalanceLastMonth
    If Val(.Text) < 0 Then .Text = "0.00"
    
    'Recovery during this month
    PrinRepay = 0: IntRepay = 0: transType = wDeposit
    If Not rstPrinTransAsOn Is Nothing Then
        rstPrinTransAsOn.Find "LoanID = " & LoanID '& " AND Transtype = " & TransType
        If Not rstPrinTransAsOn.EOF Then
        If rstPrinTransAsOn("LoanID") = LoanID Then
            rstIntTransAsOn.Find "LoanID = " & LoanID & " AND Transtype = " & transType
            PrinRepay = FormatField(rstPrinTransAsOn("SumAmount"))
            IntRepay = FormatField(rstIntTransAsOn("SumIntAmount"))
        End If
        End If
    End If
    .Col = 16: .Text = FormatCurrency(PrinRepay)
    .Col = 17: .Text = FormatCurrency(IntRepay)
    .Col = 18: .Text = FormatCurrency(PrinRepay + IntRepay)
    
    'Balance as of now
    BalanceNow = BalanceLastMonth: IntBalNow = 0
    If Not rstPrinAsOn Is Nothing Then
        rstPrinAsOn.Find " LoanID = " & LoanID
        If Not rstPrinAsOn.EOF Then
            If rstPrinAsOn("LoanID") = LoanID Then
                rstIntAsOn.Find " LoanID = " & LoanID
                BalanceNow = rstPrinAsOn("Balance")
                TransDate = rstPrinAsOn("TransDate")
                IntBalNow = FormatField(rstIntAsOn("IntBalance"))
                PrevDate = TransDate
            End If
        End If
    End If
    IntBalNow = IntBalNow + L_clsLoan.RegularInterest(LoanID, , m_ToDate)
    .Col = 19: .Text = BalanceNow
    .Col = 20: .Text = IntBalNow
    .Col = 21: .Text = Val(BalanceNow + IntBalNow)
    
    'Recovery during this Month
    .Col = 16
    Debug.Assert Val(.Text) = BalanceLastMonth - BalanceNow
    '.Col = 17: .Text = IntBalLastMonth - IntBalNow
    '.Col = 18: .Text = Val((BalanceLastMonth - BalanceNow) + (IntBalLastMonth - IntBalNow))

  End With
    DoEvents
    RaiseEvent Processing("Writing the record", rstMaster.AbsolutePosition / rstMaster.RecordCount)
    If gCancel Then rstMaster.MoveLast
    rstMaster.MoveNext
Loop
Set L_clsLoan = Nothing
ShowMeetingRegistarAll = True
Screen.MousePointer = vbDefault

ErrLine:
    Screen.MousePointer = vbDefault
    If Err Then
        MsgBox Err.Number & vbCrLf & Err.Description, , wis_MESSAGE_TITLE
       Resume
        Exit Function
    End If

End Function

Private Sub SetGrid(SchemeID As Integer)

Dim count As Integer
Dim strText As String
Dim rst As Recordset
Dim strYear As String

Dim str31_3 As String
Dim str1_4 As String
Dim strLastMonth As String

strYear = Year(FinUSFromDate) 'CStr(Year("3/31/" & Val(Year(m_ToDate) - IIf(Month(m_ToDate) > 3, 0, 1))))

str31_3 = GetIndianDate(DateAdd("d", -1, FinUSFromDate))
'"31/3/" & Val(Year(m_ToDate) - IIf(Month(m_ToDate) > 3, 0, 1))
str1_4 = FinIndianFromDate
'"1/4/" & Val(Year(m_ToDate) - IIf(Month(m_ToDate) > 3, 0, 1))
'Last day Of the Previos month
strLastMonth = GetIndianDate(GetSysLastDate(DateAdd("m", -1, m_ToDate)))

With grd
    .Clear
    .AllowUserResizing = flexResizeBoth
    .WordWrap = True
    .FixedCols = 0
    .FixedRows = 0
    .MergeCells = flexMergeNever
    .Cols = 1: .Row = 1
End With


'Get the Details of the Loan scheme
If SchemeID = 0 Then GoTo CommonSetting

Dim LoanType As wis_LoanType

    gDbTrans.SqlStmt = "SELECT * FROM LoanScheme Where SchemeID = " & m_SchemeId
    Call gDbTrans.Fetch(rst, adOpenStatic)
    'SchemeName = FormatField(gDbTrans.Rst("SchemeName"))
    'SchemeStr = " SchemeID = " & m_SchemeId & " "
    LoanType = FormatField(rst("LoanType"))


With grd
    .Cols = 28
    .Rows = 10
    .FixedCols = 2
    .FixedRows = 4
    .Row = 0
    .Col = 0: .Text = GetResourceString(33) ' "Sl No"
    .Col = 1: .Text = GetResourceString(80, 60) '"Loan No"
    .Col = 2: .Text = GetResourceString(35) ' "Name of Customer"
    .Col = 3: .Text = GetResourceString(389) '"Name OF Sureties with full address"
    .Col = 4: .Text = GetResourceString(389) '"Name OF Sureties with full address"
    .Col = 5: .Text = GetResourceString(80, 295) '"Loan Disbursment particulars"
    .Col = 6: .Text = GetResourceString(80, 295) '"Loan Disbursment particulars"
    .Col = 7: .Text = "Vehicle " & GetResourceString(295) 'Detail"
    .Col = 8: .Text = "Vehicle " & GetResourceString(295) 'Detail"
    .Col = 9: .Text = GetResourceString(67) & " " & GetFromDateString(str31_3) '"Loans out standing as on 31/3/" & strYear
    .Col = 10: .Text = GetResourceString(67) & " " & GetFromDateString(str31_3) '"Loans out standing as on 31/3/" & strYear
    .Col = 11: .Text = GetResourceString(67) & " " & GetFromDateString(str31_3) '"Loans out standing as on 31/3/" & strYear
    .Col = 12: .Text = GetResourceString(67) & " " & GetFromDateString(str31_3) '"Loans out standing as on 31/3/" & strYear
    .Col = 13: .Text = GetResourceString(67) & " " & GetFromDateString(str31_3) '"Loans out standing as on 31/3/" & strYear
    .Col = 14: .Text = "'"
    .Col = 15: .Text = "'"
    .Col = 16: .Text = "'"
    .Col = 17: .Text = "'"
    .Col = 18: .Text = GetResourceString(67) & " " & GetFromDateString(strLastMonth)  '"Loans out standing as on end of last month"
    .Col = 19: .Text = GetResourceString(67) & " " & GetFromDateString(strLastMonth)  '"Loans out standing as on end of last month"
    .Col = 20: .Text = GetResourceString(67) & " " & GetFromDateString(strLastMonth)  '"Loans out standing as on end of last month"
    .Col = 21: .Text = GetResourceString(67) & " " & GetFromDateString(strLastMonth)  '"Loans out standing as on end of last month"
    .Col = 22: .Text = GetResourceString(67) & " " & GetFromDateString(strLastMonth)  '"Loans out standing as on end of last month"
    .Col = 23: .Text = GetResourceString(67) & " " & GetFromDateString(strLastMonth)  '"Loans out standing as on end of last month"
    .Col = 24: .Text = GetResourceString(67) & " " & GetFromDateString(strLastMonth)  '"Loans out standing as on end of last month"
    .Col = 25: .Text = GetResourceString(67) & " " & GetFromDateString(strLastMonth)  '"Loans out standing as on end of last month"
    .Col = 26: .Text = "."
    .Col = 27: .Text = "."
    
    ''2nd Row
    For count = 0 To .Cols - 1
        .Col = count
        .Row = 0: strText = .Text
        .Row = 1: .Text = strText
    Next
    .Row = 1
    .Col = 3: .Text = "1" '"Suruty No. 1"
    .Col = 4: .Text = "2" '"Suruty No. 2"
    .Col = 5: .Text = GetResourceString(340) '"Date of Loan advance"
    .Col = 6: .Text = GetResourceString(80, 40) '"Amount disubursed"
    .Col = 7: .Text = "Registration No (RTO)"
    .Col = 8: .Text = "Insurence Renewed up to"
    .Col = 9: .Text = GetResourceString(52, 67) '"Total loan out standing"
    .Col = 10: .Text = GetResourceString(84) '"Of which over dues"
    .Col = 11: .Text = GetResourceString(84) '"Of which over dues"
    .Col = 12: .Text = GetResourceString(84) '"Of which over dues"
    .Col = 13: .Text = GetResourceString(84) '"Of which over dues"
    
    .Col = 14: .Text = "Repayments From 1/4/" & strYear & " to upto end of last month"
    .Col = 14: .Text = GetResourceString(20) & " " & GetFromDateString(str1_4, strLastMonth)
    .Col = 15: .Text = GetResourceString(20) & " " & GetFromDateString(str1_4, strLastMonth)
    .Col = 16: .Text = GetResourceString(20) & " " & GetFromDateString(str1_4, strLastMonth)
    .Col = 17: .Text = GetResourceString(20) & " " & GetFromDateString(str1_4, strLastMonth)
    .Col = 18: .Text = GetResourceString(84) '"Of which over due"
    .Col = 19: .Text = GetResourceString(84) '"Of which over due"
    .Col = 20: .Text = GetResourceString(84) '"Of which over due"
    .Col = 21: .Text = GetResourceString(84) '"Of which over due"
    .Col = 22: .Text = "Dates of filing"
    .Col = 23: .Text = "Dates of filing"

    
    ''2nd  Row
    For count = 0 To .Cols - 1
        .Col = count
        .Row = 1: strText = .Text
        .Row = 2: .Text = strText
    Next
    .Row = 2
    .Col = 10: .Text = GetResourceString(55) '"No of Inst"
    .Col = 11: .Text = GetResourceString(310) '"Principal"
    .Col = 12: .Text = GetResourceString(47) '"Interest"
    .Col = 13: .Text = GetResourceString(52) '"Total"
    .Col = 22: .Text = GetResourceString(55) '"No of Inst"
    .Col = 23: .Text = GetResourceString(310) '"Principal"
    .Col = 24: .Text = GetResourceString(47) '"Interest"
    .Col = 25: .Text = GetResourceString(52) '"Total"
    
    .Col = 14: .Text = GetResourceString(55) '"No of Inst"
    .Col = 15: .Text = GetResourceString(310) '"Principal"
    .Col = 16: .Text = GetResourceString(47) '"Interest"
    .Col = 17: .Text = GetResourceString(52) '"Total"
    .Col = 18: .Text = GetResourceString(55) '"No of Inst"
    .Col = 19: .Text = GetResourceString(310) '"Principal"
    .Col = 20: .Text = GetResourceString(47) '"Interest"
    .Col = 21: .Text = GetResourceString(52) '"Total"
    .Col = 26: .Text = GetResourceString(372) '"ABN"
    .Col = 27: .Text = GetResourceString(373) '"Ep"

    ''3rd row
    .Row = 3
    Dim K As Integer
    K = 1
    For count = 1 To 27
        .Col = count: .Text = K
        If .Col = 8 And LoanType <> wisVehicleloan Then K = K - 2
        K = K + 1
    Next
    .Col = 0: .Text = "1"
    .Col = 1: .Text = "2"
    .Col = 2: .Text = "2a"
    
    .MergeCells = flexMergeRestrictRows
    
    Dim RowCount As Integer
    For RowCount = 0 To .FixedRows - 1
        .MergeRow(RowCount) = True
        .Row = RowCount
        For count = 0 To grd.Cols - 1
            .MergeCol(count) = True
            .Col = count
            .CellAlignment = 4: .CellFontBold = True
        Next
    Next
    .Row = 1
    For count = 19 To 22
        .Col = count
        .Text = ".."
    Next
    
End With
Exit Sub

CommonSetting:
With grd
    .Cols = 22
    .Rows = 10
    .FixedCols = 2
    .FixedRows = 3
    .MergeRow(0) = True
    .MergeRow(1) = True
    .MergeRow(2) = True
    .Row = 0:
    .Col = 0: .Text = GetResourceString(33) '"Sl No"
    .Col = 1: .Text = GetResourceString(80, 60) '"Loan No"
    .Col = 2: .Text = GetResourceString(35) '"Name of Customer"
    .Col = 3: .Text = GetResourceString(389) '"Guarantor Name & Address"
    .Col = 4: .Text = GetResourceString(389) '"Guarantor Name & Address"
    .Col = 5: .Text = GetResourceString(340) '"Issue Date"
    .Col = 6: .Text = GetResourceString(80, 40) '"Loan Amount"

    .Col = 7: .Text = "Out standing as 31/3/" & strYear & ""
    .Col = 7: .Text = GetResourceString(67) & " " & GetFromDateString(str31_3)
    .Col = 8: .Text = GetResourceString(67) & " " & GetFromDateString(str31_3)
    .Col = 9: .Text = GetResourceString(67) & " " & GetFromDateString(str31_3)
    
    '.Col = 10: .Text = "Repayments from 1/4/" & strYear & " to up to the end of last month"
    .Col = 10: .Text = GetResourceString(20) & " " & GetFromDateString(str1_4, strLastMonth)
    .Col = 11: .Text = GetResourceString(20) & " " & GetFromDateString(str1_4, strLastMonth)
    .Col = 12: .Text = GetResourceString(20) & " " & GetFromDateString(str1_4, strLastMonth)
    
    '.Col = 13: .Text = "Balance OutStanding as on the end of last month"
    .Col = 13: .Text = GetResourceString(67) & " " & GetFromDateString(strLastMonth)
    .Col = 14: .Text = GetResourceString(67) & " " & GetFromDateString(strLastMonth)
    .Col = 15: .Text = GetResourceString(67) & " " & GetFromDateString(strLastMonth)
    
    '.Col = 16: .Text = "Repayments during this month"
    .Col = 16: .Text = GetResourceString(374, 192, 20)
    .Col = 17: .Text = GetResourceString(374, 192, 20)
    .Col = 18: .Text = GetResourceString(374, 192, 20)
             
    
    '.Col = 19: .Text = "Balance OutStanding end of this month"
    .Col = 19: .Text = GetResourceString(67) & " " & GetFromDateString(m_ToIndianDate)
    .Col = 20: .Text = GetResourceString(67) & " " & GetFromDateString(m_ToIndianDate)
    .Col = 21: .Text = GetResourceString(67) & " " & GetFromDateString(m_ToIndianDate)
    
    ''2nd row
    .Row = 1: .MergeRow(3) = True
    .Col = 0: .Text = GetResourceString(33) '"Sl No"
    .Col = 1: .Text = GetResourceString(80, 60) '"Loan No"
    .Col = 2: .Text = GetResourceString(35) '"Name of Customer"
    .Col = 3:  .Text = "1" '"Guarantor 1"
    .Col = 4:  .Text = "2" '"Guarantor 2"
    .Col = 5:  .Text = GetResourceString(340) '"Issue Date"
    .Col = 6: .Text = GetResourceString(80, 40) '"Loan Amount"

    .Col = 7: .Text = GetResourceString(310) '"Principal"
    .Col = 8: .Text = GetResourceString(47) '"Interest"
    .Col = 9: .Text = GetResourceString(52) '"Total"
    
    .Col = 10: .Text = GetResourceString(310) '"Principal"
    .Col = 11: .Text = GetResourceString(47) '"Interest"
    .Col = 12: .Text = GetResourceString(52) '"Total"
    .Col = 13: .Text = GetResourceString(310) '"Principal"
    .Col = 14: .Text = GetResourceString(47) '"Interest"
    .Col = 15: .Text = GetResourceString(52) '"Total"
    
    .Col = 16: .Text = GetResourceString(310) '"Principal"
    .Col = 17: .Text = GetResourceString(47) '"Interest"
    .Col = 18: .Text = GetResourceString(52) '"Total"
    .Col = 19: .Text = GetResourceString(310) '"Principal"
    .Col = 20: .Text = GetResourceString(47) '"Interest"
    .Col = 21: .Text = GetResourceString(52) '"Total"
    
    
    .Row = 2: .MergeRow(4) = True
    .MergeCells = flexMergeFree
    For count = 3 To .Cols - 1
        .Col = count: .Text = (count)
        .CellAlignment = 4
        .MergeCol(count) = True
    Next
    '.Col = 3: .Text = "2b"
    .Col = 2: .Text = "2a"
    .Col = 1: .Text = "2"
    .Col = 0: .Text = "1"
    
    Dim I As Integer, j As Integer
    For I = 0 To .FixedRows - 1
        .Row = I
        For j = 0 To .Cols - 1
            .Col = j
            .CellAlignment = 4: .CellFontBold = True
        Next
    Next
End With

End Sub

Private Function ShowShed5() As Boolean

RaiseEvent Processing("Fetching the record", 0)
ShowShed5 = False
Err.Clear
On Error GoTo ExitLine:


Dim SqlStr As String
Dim transType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes

Dim rstOpBalance As Recordset
Dim rstClBalance As Recordset
Dim rstAdvance As Recordset
Dim rstRecovery As Recordset

Dim FirstDate As Date

Dim ColAmount() As Currency
Dim GrandTotal() As Currency


'Get the First day of the Month
FirstDate = GetSysFirstDate(m_ToDate)
'Fetch Only Cash Credit Loans
Dim LoanType As wis_LoanType
LoanType = wisCashCreditLoan
RaiseEvent Initialise(0, 10)

'Get The LoanDetails And THier balance as on Date
SqlStr = "SELECT A.LoanID,AccNum,IssueDate,LoanDueDate,LoanAmount,Balance," & _
    " LoanPurpose, Title+' '+ FirstName+' '+ MiddleName+' '+LastName As Name" & _
    " FROM LoanMaster A,LoanTrans B, NameTab C WHERE A.SchemeId" & _
        " IN (SELECT SchemeId FROM LoanScheme Ls WHERE Ls.LoanType = " & LoanType & ")" & _
    "  AND B.LoanID = A.LoanID" & _
    " AND C.CustomerId =A.CustomerID AND TransID = (SELECT MAX(TransId) FROM " & _
        " LoanTrans D WHERE D.TransDate <= #" & m_ToDate & "# " & _
        " AND D.LoanId = A.LoanId ) " & _
    " AND (LoanClosed is NULL OR LOanClosed = 0 ) "

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstClBalance, adOpenStatic) < 1 Then gCancel = True: Exit Function

DoEvents
RaiseEvent Processing("Fetching the record", 0.25)
If gCancel Then Exit Function

'Get The LoanDetails And Thier balance as on first day of the given month
SqlStr = "SELECT LoanID,Balance,TransDate FROM LoanTrans WHERE " & _
    " TransDate < #" & FirstDate & "# ORDER BY LoanID,TransId Desc"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstOpBalance, adOpenStatic) < 1 Then Exit Function


DoEvents
RaiseEvent Processing("Fetching the record", 0.5)
If gCancel Then Exit Function

'Get The Advances During the MOnth
transType = wWithdraw
ContraTrans = wContraWithdraw
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans WHERE TransDate >= #" & FirstDate & "#" & _
    " AND TransDate <= #" & m_ToDate & "# AND (TransType = " & transType & _
    " OR TransType = " & ContraTrans & ")" & _
    " GROUP BY LoanID "
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstAdvance, adOpenStatic) < 1 Then Set rstAdvance = Nothing

DoEvents
RaiseEvent Processing("Fetching the record", 0.65)
If gCancel Then Exit Function


'Get The Recovery During the MOnth
transType = wDeposit
ContraTrans = wContraDeposit
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans WHERE TransDate >= #" & FirstDate & "#" & _
    " AND TransDate <= #" & m_ToDate & "# AND " & _
    " (TransType = " & transType & " OR TransType = " & ContraTrans & ")" & _
    " GROUP BY LoanID "
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstRecovery, adOpenStatic) < 1 Then Set rstRecovery = Nothing

DoEvents
RaiseEvent Processing("Fetching the record", 0.85)
If gCancel Then Exit Function

'Now Align the grid
Call Shed5RowCol
ReDim ColAmount(7 To grd.Cols - 1)
ReDim GrandTotal(7 To grd.Cols - 1)
Dim LoanID As Long
Dim AddRow As Boolean
Dim L_clsLoan As New clsLoan
Dim PrevOD As Currency
Dim ODAmount As Currency
Dim count As Long
Dim SlNo As Long
RaiseEvent Initialise(0, rstClBalance.RecordCount)

'Now Start to writing to the grid
SlNo = 0
While Not rstClBalance.EOF
    LoanID = FormatField(rstClBalance("LoanId"))
    rstOpBalance.MoveFirst
    rstOpBalance.Find "LoanID = " & LoanID
    ColAmount(7) = 0
    If Not rstOpBalance.EOF Then
        ColAmount(7) = FormatField(rstOpBalance("Balance")) 'Balance as on 31/3/yyyy
    End If
    ColAmount(8) = 0
    If Not rstAdvance Is Nothing Then
        rstAdvance.MoveFirst
        rstAdvance.Find "LoanId = " & LoanID
        If Not rstAdvance.EOF Then _
            ColAmount(8) = FormatField(rstAdvance(0))  'Advances During the MOnth
        
    End If
    ColAmount(9) = 0
    If Not rstRecovery Is Nothing Then
        rstRecovery.MoveFirst
        rstRecovery.Find "LoanId = " & LoanID
        If Not rstRecovery.EOF Then _
            ColAmount(9) = FormatField(rstRecovery(0)) 'Recovery During the mOnth
        
    End If
    ColAmount(10) = FormatField(rstClBalance("Balance")) 'Balance at the end of month
    
    ColAmount(11) = ColAmount(7) + ColAmount(8)  'Maximum O/s Balance
    
    
    ODAmount = L_clsLoan.OverDueAmount(LoanID, , m_ToDate)  'Over due
    ColAmount(12) = ODAmount 'Over due amount of the loan as on given date
    
    'Over due amount classification
    PrevOD = 0
'    ODAmount = L_clsLoan.OverDueSince(5, LoanID, , AsOnIndianDate)
'    ColAmount(17) = ODAmount - PrevOD 'Over due since & above 5 Years
'    PrevOD = ColAmount(17) + PrevOD
    ColAmount(18) = L_clsLoan.OverDueSince(5, LoanID, , m_ToDate)
    If ColAmount(18) < 0 Then ColAmount(18) = ODAmount  'Over due since & above 5 Years
    ODAmount = ODAmount - ColAmount(18)
    
    ColAmount(17) = L_clsLoan.OverDueSince(4, LoanID, , m_ToDate) - ColAmount(18)
    If ColAmount(17) < 0 Then ColAmount(17) = ODAmount  'Over due since 5 Years
    ODAmount = ODAmount - ColAmount(17)
    
    ColAmount(16) = L_clsLoan.OverDueSince(3, LoanID, , m_ToDate) - ColAmount(17)
    If ColAmount(16) < 0 Then ColAmount(16) = ODAmount  'Over due since 3 Years
    ODAmount = ODAmount - ColAmount(16)
    
    ColAmount(15) = L_clsLoan.OverDueSince(2, LoanID, , m_ToDate) - ColAmount(16)
    If ColAmount(15) < 0 Then ColAmount(15) = ODAmount  'Over due since 2 Years
    ODAmount = ODAmount - ColAmount(15)
    
    ColAmount(14) = L_clsLoan.OverDueSince(1, LoanID, , m_ToDate) - ColAmount(15)
    If ColAmount(14) < 0 Then ColAmount(14) = ODAmount  'Over due since 1 Years
    ODAmount = ODAmount - ColAmount(14)
    
    ColAmount(13) = ODAmount   'Over due Under 1 Years
    
    'Check whether this row has to be write or not
    AddRow = False
    For count = 7 To grd.Cols - 1
        If ColAmount(count) Then
            AddRow = True
            SlNo = SlNo + 1
            Exit For
        End If
    Next
    If AddRow Then
      With grd
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        '.MergeRow(.Row) = False
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = FormatField(rstClBalance("AccNum"))
        .Col = 2: .Text = FormatField(rstClBalance("Name"))
        .Col = 3: .Text = FormatField(rstClBalance("LoanAmount"))
        .Col = 4: .Text = FormatField(rstClBalance("IssueDate"))
        .Col = 5: .Text = FormatField(rstClBalance("LoanDueDate"))
        .Col = 6: .Text = FormatField(rstClBalance("LoanPurpose"))
        For count = 7 To .Cols - 1
            .Col = count
            If ColAmount(count) < 0 Then ColAmount(count) = 0
            .Text = FormatCurrency(ColAmount(count))
            GrandTotal(count) = GrandTotal(count) + ColAmount(count)
        Next
      End With
    End If
    DoEvents
    If gCancel Then rstClBalance.MoveLast
    RaiseEvent Processing("Writing the records", rstClBalance.AbsolutePosition / rstClBalance.RecordCount)
    rstClBalance.MoveNext
Wend

Set L_clsLoan = Nothing

With grd
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 2
    .Text = GetResourceString(286): .CellFontBold = True
    For count = 7 To .Cols - 1
        .Col = count: .CellFontBold = True
        .Text = FormatCurrency(GrandTotal(count))
    Next
End With
ShowShed5 = True


ExitLine:
lblReportTitle.Caption = "Statement showing the Cash Credit loans for the month of " & _
    GetMonthString(Month(m_ToDate)) & " " & GetFromDateString(m_ToIndianDate)

    If Err Then
        MsgBox "ERROR SHED 5" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
        'Resume
        Err.Clear
    End If

End Function

Private Function ShowShed1All() As Boolean

RaiseEvent Processing("Fetching the records", 0)


On Error GoTo ErrLine
'Declarations
Dim rstLoan As Recordset
Dim RstRepay  As Recordset
Dim rstAdvance As Recordset

Dim SqlStr As String
Dim count As Integer

'Decalration By SHashi
Dim ODAmount As Currency

Dim ColAmount() As Currency
Dim SubTotal() As Currency
Dim GrandTotal() As Currency
Dim fromDate As Date
Dim LastDate As Date
Dim ObDate As Date

Dim ProcCount As Integer
Dim totalCount As Integer


ShowShed1All = False

'ObDate = "4/1/" & IIf(Month(m_ToDate) > 3, Year(m_ToDate), Year(m_ToDate) - 1)
ObDate = FinUSFromDate

SqlStr = "SELECT LoanId From LoanMaster "
    
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstLoan, adOpenStatic) < 1 Then gCancel = True: Exit Function

RaiseEvent Initialise(0, rstLoan.RecordCount * 2)
totalCount = rstLoan.RecordCount + 10

'Now fix the headers for the shedule
Call Shed1RowCOl

'Now Get the Loan RepayMent Of the during this period
Dim transType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes
transType = wDeposit
ContraTrans = wContraDeposit
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans " & _
            " Where (TransType = " & transType & _
                " OR TransType = " & ContraTrans & ") " & _
            " AND TransDate >= #" & ObDate & "# " & _
            " AND TransDate <= #" & m_ToDate & "# " & _
            " GROUP BY LoanID"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(RstRepay, adOpenDynamic) < 1 Then Set RstRepay = Nothing

'Now Get the Loan Advance Of the during this period
transType = wWithdraw
ContraTrans = wContraWithdraw
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans Where TransType = " & transType & _
    " OR TransType = " & ContraTrans & " AND TransDate >= #" & ObDate & "# " & _
    " AND TransDate <= #" & m_ToDate & "#" & _
    " GROUP BY LoanID"
gDbTrans.SqlStmt = SqlStr
Call gDbTrans.Fetch(rstAdvance, adOpenStatic)

ReDim ColAmount(3 To grd.Cols - 1)
ReDim SubTotal(3 To grd.Cols - 1)
ReDim GrandTotal(3 To grd.Cols - 1)

Dim LoanID As Long
Dim L_clsLoan As New clsLoan
Dim OBIndiandate As String
Dim curRepay As Currency
Dim AddRow As Boolean

OBIndiandate = GetIndianDate(ObDate)

grd.Row = grd.FixedRows - 1
Dim rstScheme As Recordset

gDbTrans.SqlStmt = "Select * From LoanScheme"
If gDbTrans.Fetch(rstScheme, adOpenDynamic) < 1 Then gCancel = True: Exit Function

Do
    If rstScheme.EOF Then Exit Do
    
    
    SqlStr = "SELECT A.LoanId, Balance " & _
        " From LoanMaster A,LoanTrans B WHERE A.SchemeID = " & rstScheme("SChemeID") & _
        " AND A.LoanID = B.LoanID " & _
        " AND TransID = (SELECT MAX(TransID) " & _
            " From LoanTrans D Where D.LoanId = A.LoanId" & _
            " AND D.TransDate <= #" & m_ToDate & "#)"
    gDbTrans.SqlStmt = SqlStr
    If gDbTrans.Fetch(rstLoan, adOpenStatic) < 1 Then GoTo NextScheme
    
    If Not RstRepay Is Nothing Then RstRepay.MoveLast
    While Not rstLoan.EOF
        'Initialise the Varible
        LoanID = FormatField(rstLoan("LoanID"))
        ColAmount(3) = L_clsLoan.OverDueAmount(LoanID, , ObDate)  'Over due as on Opeing date
        ColAmount(4) = L_clsLoan.LoanDemand(LoanID, ObDate, m_ToDate)    'Amount which falls as over due between OB Date & today
        ColAmount(5) = L_clsLoan.OverDueAmount(LoanID, , m_ToDate) 'Over due as on today
        
        ColAmount(5) = ColAmount(3) + ColAmount(4)
        
        curRepay = 0: ' curAdvance = 0
        If Not RstRepay Is Nothing Then
            RstRepay.MoveFirst
            'RstRepay.Find "LoanID = " & Loanid
            If gDbTrans.FindRecord(RstRepay, _
                    "LoanID = " & LoanID) Then curRepay = FormatField(RstRepay(0))
        End If
        
        Debug.Assert ColAmount(5) - curRepay = L_clsLoan.OverDueAmount(LoanID, , m_ToDate)
        
        ColAmount(9) = ColAmount(9) + FormatCurrency(curRepay)
        ColAmount(10) = 0: ColAmount(6) = 0
        ColAmount(7) = 0: ColAmount(8) = 0
        'Now Calculate the Recovery demand
        If curRepay > 0 Then
            'Now calculate the recovery against arrears demand
            If ColAmount(3) >= curRepay Then
                ColAmount(6) = curRepay: curRepay = 0
            Else
                ColAmount(6) = ColAmount(3): curRepay = curRepay - ColAmount(6)
            End If
            'Now calculate the recovery against current demand
            If ColAmount(4) >= curRepay Then
                ColAmount(7) = curRepay: curRepay = 0
            Else
                ColAmount(7) = ColAmount(4): curRepay = curRepay - ColAmount(7)
            End If
            'Remianinog amount is the advance recovery
            ColAmount(8) = curRepay
        End If
        
        ColAmount(10) = ColAmount(5) - ColAmount(6) - ColAmount(7)  'OVer due amount as on date
        ODAmount = ColAmount(10)
        
        'Over due amount as on date
        ODAmount = L_clsLoan.OverDueAmount(LoanID, , m_ToDate)
        ColAmount(16) = L_clsLoan.OverDueSince(5, LoanID, , m_ToDate)
        If ColAmount(16) > ODAmount Then ColAmount(16) = ODAmount 'Over due since 5 & above 5 Years
        ODAmount = ODAmount - ColAmount(16)
        
        ColAmount(15) = L_clsLoan.OverDueSince(4, LoanID, , m_ToDate) - ColAmount(16)
        If ColAmount(15) > ODAmount Then ColAmount(15) = ODAmount 'Over due since 4 Years
        ODAmount = ODAmount - ColAmount(15)
        
        ColAmount(14) = L_clsLoan.OverDueSince(3, LoanID, , m_ToDate) - ColAmount(15)  'Over due since 3 Years
        If ColAmount(14) > ODAmount Then ColAmount(14) = ODAmount 'Over due since 3 Years
        ODAmount = ODAmount - ColAmount(14)
        
        ColAmount(13) = L_clsLoan.OverDueSince(2, LoanID, , m_ToDate) - ColAmount(14)  'Over due since 2 Years
        If ColAmount(13) > ODAmount Then ColAmount(13) = ODAmount 'Over due since 2 Years
        ODAmount = ODAmount - ColAmount(13)
        
        ColAmount(12) = L_clsLoan.OverDueSince(1, LoanID, , m_ToDate) - ColAmount(13)
        If ColAmount(12) > ODAmount Then ColAmount(12) = ODAmount 'Over due since 1 Year
        ODAmount = ODAmount - ColAmount(12)
        
        ColAmount(11) = ODAmount 'Over due under one year
        
        For count = 3 To grd.Cols - 1
            SubTotal(count) = SubTotal(count) + ColAmount(count)
        Next
        DoEvents
        ProcCount = ProcCount + 1
        RaiseEvent Processing("Writing the record", (ProcCount / totalCount))
        rstLoan.MoveNext
    Wend

    AddRow = False
    For count = 3 To grd.Cols - 1
        SubTotal(count) = SubTotal(count) + ColAmount(count)
        If SubTotal(count) Then AddRow = True
    Next
    If AddRow Then
        With grd
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            .Col = 0: .Text = Format(.Row, "00")
            .Col = 2: .Text = FormatField(rstScheme("SchemeName"))
            For count = 3 To .Cols - 1
                If SubTotal(count) < 0 Then SubTotal(count) = 0
                .Col = count
                .Text = FormatCurrency(SubTotal(count))
            Next
        End With
    End If
    For count = 3 To grd.Cols - 1
        GrandTotal(count) = GrandTotal(count) + SubTotal(count)
        SubTotal(count) = 0
    Next

NextScheme:
    rstScheme.MoveNext

Loop

    AddRow = False
    For count = 3 To grd.Cols - 1
        If GrandTotal(count) Then AddRow = True
    Next
    If grd.Row <= 1 Then AddRow = False
    
    If AddRow Then
        With grd
            
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            .Col = 0: .Text = Format(.Row, "00")
            .Col = 2: .Text = GetResourceString(286): .CellFontBold = True
            For count = 3 To .Cols - 1
                If GrandTotal(count) < 0 Then GrandTotal(count) = 0
                .Col = count: .CellFontBold = True
                .Text = FormatCurrency(GrandTotal(count))
            Next
        End With
    End If


Set L_clsLoan = Nothing
lblReportTitle = "Demand, collecion and blance register of fo the month of " & _
    GetMonthString(Month(m_ToDate)) & " " & GetFromDateString(m_ToIndianDate)
ShowShed1All = True
Exit Function

ErrLine:
    MsgBox "error Showshed1" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
  'Resume
End Function

Private Function ShowShed1() As Boolean

RaiseEvent Processing("Fetching the records", 0)

If m_SchemeId = 0 Then
    ShowShed1 = ShowShed1All
    Exit Function
End If

On Error GoTo ErrLine
'Declarations
Dim rstLoan As Recordset
Dim RstRepay  As Recordset
Dim rstAdvance As Recordset

Dim SqlStr As String
Dim count As Integer

Dim SlNo As Integer
Dim strSocName As String
Dim strLoanName As String
Dim strBranchName As String
Dim BankID As Long
Dim OBalance As Currency
Dim RefID As Long
Dim TransID As Long
Dim LoanSchemeID

Dim RowNum As Integer
Dim ColNum As Integer

'Decalration By SHashi
Dim ColAmount() As Currency
Dim SubTotal() As Currency
Dim GrandTotal() As Currency
Dim fromDate As Date
Dim LastDate As Date
Dim ObDate As Date

ShowShed1 = False


Dim LoanCategary As wisLoanCategories
'Get the Loan Details
gDbTrans.SqlStmt = "Select Category From LoanScheme Where SchemeID = " & m_SchemeId
Call gDbTrans.Fetch(rstLoan, adOpenDynamic)

'LoanCategary = wisAgriculural
LoanCategary = FormatField(rstLoan("Category"))

If LoanCategary = wisAgriculural Then
    'ObDate = "7/1/" & IIf(Month(m_ToDate) > 6, Year(m_ToDate), Year(m_ToDate) - 1)
    ObDate = GetSysFormatDate("1/7/" & IIf(Month(m_ToDate) > 6, Year(m_ToDate), Year(m_ToDate) - 1))
Else
    'ObDate = "4/1/" & IIf(Month(m_ToDate) > 3, Year(m_ToDate), Year(m_ToDate) - 1)
    ObDate = FinUSFromDate
End If

'Get The Details Of Loans, and Balance as ondate
SqlStr = "SELECT A.LoanId, AccNum, Balance,  " & _
    " Title + ' ' + FirstName +' ' + MiddleName + ' ' + LastName As Name " & _
    " From LoanMaster A,LoanTrans B,NameTab C WHERE A.SchemeID = " & m_SchemeId & _
    " AND A.LoanID = B.LoanID AND C.CustomerID = A.CustomerId " & _
    " AND TransID = (SELECT MAX(TransID) " & _
        " From LoanTrans D Where D.LoanId = A.LoanId" & _
        " AND D.TransDate <= #" & m_ToDate & "#)" & _
    " AND A.LoanId In (Select LoanId From LoanMaster" & _
            " Where SchemeID = " & m_SchemeId & ") "

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstLoan, adOpenStatic) < 1 Then gCancel = True: Exit Function

RaiseEvent Initialise(0, rstLoan.RecordCount * 2)

'Now fix the headers for the shedule
Call Shed1RowCOl

'Now Get the Loan RepayMent Of the during this period
Dim transType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes
transType = wDeposit
ContraTrans = wContraDeposit
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans " & _
    " Where (TransType = " & transType & _
        " OR TransType = " & ContraTrans & ") " & _
    " AND TransDate >= #" & ObDate & "#" & _
    " AND TransDate <= #" & m_ToDate & "#" & _
    " AND LoanId In (Select LoanId From LoanMaster" & _
            " Where SchemeID = " & m_SchemeId & ") " & _
    " GROUP BY LoanID"
    
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(RstRepay, adOpenDynamic) < 1 Then Set RstRepay = Nothing

'Now Get the Loan Advance Of the during this period
transType = wWithdraw
ContraTrans = wContraWithdraw
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans " & _
    " Where TransType = " & transType & " OR TransType = " & ContraTrans & _
    " AND TransDate >= #" & ObDate & "# AND TransDate <= #" & m_ToDate & "#" & _
    " AND LoanId In (Select LoanId From LoanMaster" & _
            " Where SchemeID = " & m_SchemeId & ") " & _
    " GROUP BY LoanID"
    
gDbTrans.SqlStmt = SqlStr
Call gDbTrans.Fetch(rstAdvance, adOpenStatic)

ReDim ColAmount(3 To grd.Cols - 1)
ReDim SubTotal(3 To grd.Cols - 1)
ReDim GrandTotal(3 To grd.Cols - 1)

Dim LoanID As Long
Dim L_clsLoan As New clsLoan
Dim OBIndiandate As String
Dim curRepay As Currency
Dim ODAmount As Currency
Dim PrevOD As Currency
Dim AddRow As Boolean

OBIndiandate = GetIndianDate(ObDate)

grd.Row = grd.FixedRows
SlNo = 0
If Not RstRepay Is Nothing Then RstRepay.MoveLast
While Not rstLoan.EOF
    'Initialise the Varible
    LoanID = FormatField(rstLoan("LoanID"))
    ColAmount(3) = L_clsLoan.OverDueAmount(LoanID, , ObDate)   'Over due as on Opeing date
    ColAmount(4) = L_clsLoan.LoanDemand(LoanID, ObDate, m_ToDate)     'Amount which falls as over due between OB Date & today
    ColAmount(5) = L_clsLoan.OverDueAmount(LoanID, , m_ToDate)   'Over due as on today
    
    ColAmount(5) = ColAmount(3) + ColAmount(4)
    
    'The difference between two date is the Loan demand of that period
    'Therefore
    'ColAmount(4) = ColAmount(5) - ColAmount(3)
    
    curRepay = 0: ' curAdvance = 0
    If Not RstRepay Is Nothing Then
        RstRepay.MoveFirst
        RstRepay.Find "LoanID = " & LoanID
        If Not RstRepay.EOF Then curRepay = FormatField(RstRepay(0))
        
    End If
    Debug.Assert ColAmount(5) - curRepay = L_clsLoan.OverDueAmount(LoanID, , m_ToDate)
'    If Not rstAdvance Is Nothing Then
'        rstRepay.FindFirst "LoanID = " & LoanID
'        If Not rstAdvance.NoMatch And rstAdvance("LoanID") = LoanID Then
'            curAdvance = FormatField(rstAdvance(0))
'        End If
'    End If
    ColAmount(9) = FormatCurrency(curRepay)
    ColAmount(10) = 0: ColAmount(6) = 0: ColAmount(7) = 0
    ColAmount(8) = 0
    'Now Calculate the Recovery demand
    If curRepay > 0 Then
        'Now calculate the recovery against arrears demand
        If ColAmount(3) >= curRepay Then
            ColAmount(6) = curRepay: curRepay = 0
        Else
            ColAmount(6) = ColAmount(3): curRepay = curRepay - ColAmount(6)
        End If
        'Now calculate the recovery against current demand
        If ColAmount(4) >= curRepay Then
            ColAmount(7) = curRepay: curRepay = 0
        Else
            ColAmount(7) = ColAmount(4): curRepay = curRepay - ColAmount(7)
        End If
        'Remianinog amount is the advance recovery
        ColAmount(8) = curRepay
    End If
    
    ColAmount(10) = ColAmount(5) - ColAmount(6) - ColAmount(7)  'OVer due amount as on date
    ODAmount = ColAmount(10)
    
    'Over due amount as on date
    ODAmount = L_clsLoan.OverDueAmount(LoanID, , m_ToDate)
    PrevOD = 0
    ColAmount(16) = L_clsLoan.OverDueSince(5, LoanID, , m_ToDate)
    If ColAmount(16) > ODAmount Then ColAmount(16) = ODAmount 'Over due since 5 & above 5 Years
    ODAmount = ODAmount - ColAmount(16)
    
    ColAmount(15) = L_clsLoan.OverDueSince(4, LoanID, , m_ToDate) - ColAmount(16)
    If ColAmount(15) > ODAmount Then ColAmount(15) = ODAmount 'Over due since 4 Years
    ODAmount = ODAmount - ColAmount(15)
    
    ColAmount(14) = L_clsLoan.OverDueSince(3, LoanID, , m_ToDate) - ColAmount(15)  'Over due since 3 Years
    If ColAmount(14) > ODAmount Then ColAmount(14) = ODAmount 'Over due since 3 Years
    ODAmount = ODAmount - ColAmount(14)
    
    ColAmount(13) = L_clsLoan.OverDueSince(2, LoanID, , m_ToDate) - ColAmount(14)  'Over due since 2 Years
    If ColAmount(13) > ODAmount Then ColAmount(13) = ODAmount 'Over due since 2 Years
    ODAmount = ODAmount - ColAmount(13)
    
    ColAmount(12) = L_clsLoan.OverDueSince(1, LoanID, , m_ToDate) - ColAmount(13)
    If ColAmount(12) > ODAmount Then ColAmount(12) = ODAmount 'Over due since 1 Year
    ODAmount = ODAmount - ColAmount(12)
    
    ColAmount(11) = ODAmount 'Over due under one year
    AddRow = False
    For count = 3 To grd.Cols - 1
        If ColAmount(count) Then
            AddRow = True
            Exit For
        End If
    Next
    If AddRow Then
        With grd
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            SlNo = SlNo + 1
            .Col = 0: .Text = SlNo
            .Col = 1: .Text = FormatField(rstLoan("AccNum"))
            .Col = 2: .Text = FormatField(rstLoan("Name"))
            For count = 3 To .Cols - 1
                If ColAmount(count) < 0 Then ColAmount(count) = 0
                .Col = count: .Text = FormatCurrency(ColAmount(count))
                GrandTotal(count) = GrandTotal(count) + ColAmount(count)
            Next
        End With
    End If
    DoEvents
    RaiseEvent Processing("Writing the record", (rstLoan.AbsolutePosition / rstLoan.RecordCount))
    rstLoan.MoveNext
Wend

AddRow = False
For count = 3 To grd.Cols - 1
    If GrandTotal(count) Then
        AddRow = True
        Exit For
    End If
Next
If AddRow Then
    With grd
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 2: .Text = "Grand Total": .CellFontBold = True
        For count = 3 To .Cols - 1
            If GrandTotal(count) < 0 Then ColAmount(count) = 0
            .Col = count: .CellFontBold = True
            .Text = FormatCurrency(GrandTotal(count))
        Next
    End With
End If

Set L_clsLoan = Nothing
lblReportTitle = "Demand, collecion and blance register of fo the month of " & _
    GetMonthString(Month(m_ToDate)) & " " & GetFromDateString(m_ToIndianDate)
ShowShed1 = True
Exit Function

ErrLine:
    MsgBox "error Showshed1" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
  'Resume
End Function



Private Function ShowShed1_All() As Boolean

RaiseEvent Processing("Fetching the records", 0)


On Error GoTo ErrLine
'Declarations
Dim rstScheme As Recordset
Dim rstLoan As Recordset
Dim RstRepay  As Recordset
Dim rstAdvance As Recordset

Dim SqlStr As String
Dim count As Integer

Dim SlNo As Integer
Dim strSocName As String
Dim strLoanName As String
Dim strBranchName As String
Dim BankID As Long
Dim OBalance As Currency
Dim RefID As Long
Dim TransID As Long
Dim LoanSchemeID

Dim RowNum As Integer
Dim ColNum As Integer

'Decalration By SHashi
Dim ColAmount() As Currency
Dim SubTotal() As Currency
Dim GrandTotal() As Currency
Dim fromDate As Date
Dim LastDate As Date
Dim ObDate As Date

ShowShed1_All = False


Dim LoanCategary As wisLoanCategories
'Get the Loan Details
'gDbTrans.SQLStmt = "Select Category From LoanScheme Where SchemeID = " & m_SchemeId
gDbTrans.SqlStmt = "Select * From LoanScheme Where Category = " & wisAgriculural

Call gDbTrans.Fetch(rstLoan, adOpenDynamic)

LoanCategary = wisAgriculural
'LoanCategary = FormatField(rstLoan("Category"))

If LoanCategary = wisAgriculural Then
    'ObDate = "7/1/" & IIf(Month(m_ToDate) > 6, Year(m_ToDate), Year(m_ToDate) - 1)
    ObDate = GetSysFormatDate("1/7/" & IIf(Month(m_ToDate) > 6, Year(m_ToDate), Year(m_ToDate) - 1))
Else
    'ObDate = "4/1/" & IIf(Month(m_ToDate) > 3, Year(m_ToDate), Year(m_ToDate) - 1)
    ObDate = FinUSFromDate
End If

'Get The Details Of Loans, and Balance as ondate
SqlStr = "SELECT A.LoanId, AccNum, Balance,  " & _
    " Title + ' ' + FirstName +' ' + MiddleName + ' ' + LastName As Name " & _
    " From LoanMaster A,LoanTrans B,NameTab C WHERE A.SchemeID = " & m_SchemeId & _
    " AND A.LoanID = B.LoanID AND C.CustomerID = A.CustomerId " & _
    " AND TransID = (SELECT MAX(TransID) " & _
        " From LoanTrans D Where D.LoanId = A.LoanId" & _
        " AND D.TransDate <= #" & m_ToDate & "#)" & _
    " AND A.LoanId In (Select LoanId From LoanMaster" & _
            " Where SchemeID = " & m_SchemeId & ") "

'gDbTrans.SQLStmt = SqlStr
'If gDbTrans.Fetch(rstLoan, adOpenStatic) < 1 Then gCancel = True: Exit Function

RaiseEvent Initialise(0, rstLoan.RecordCount * 2)

'Now fix the headers for the shedule
Call Shed1RowCOl

'Now Get the Loan RepayMent Of the during this period
Dim transType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes
transType = wDeposit
ContraTrans = wContraDeposit
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans " & _
    " Where (TransType = " & transType & _
        " OR TransType = " & ContraTrans & ") " & _
    " AND TransDate >= #" & ObDate & "#" & _
    " AND TransDate <= #" & m_ToDate & "#" & _
    " AND LoanId In (Select LoanId From LoanMaster" & _
            " Where SchemeID = " & m_SchemeId & ") " & _
    " GROUP BY LoanID"
    
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(RstRepay, adOpenDynamic) < 1 Then Set RstRepay = Nothing

'Now Get the Loan Advance Of the during this period
transType = wWithdraw
ContraTrans = wContraWithdraw
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans " & _
    " Where TransType = " & transType & " OR TransType = " & ContraTrans & _
    " AND TransDate >= #" & ObDate & "# AND TransDate <= #" & m_ToDate & "#" & _
    " AND LoanId In (Select LoanId From LoanMaster" & _
            " Where SchemeID = " & m_SchemeId & ") " & _
    " GROUP BY LoanID"
    
gDbTrans.SqlStmt = SqlStr
Call gDbTrans.Fetch(rstAdvance, adOpenStatic)

ReDim ColAmount(3 To grd.Cols - 1)
ReDim SubTotal(3 To grd.Cols - 1)
ReDim GrandTotal(3 To grd.Cols - 1)

Dim LoanID As Long
Dim SchemeID As Integer
Dim L_clsLoan As New clsLoan
Dim OBIndiandate As String
Dim curRepay As Currency
Dim ODAmount As Currency
Dim PrevOD As Currency
Dim AddRow As Boolean

OBIndiandate = GetIndianDate(ObDate)

grd.Row = grd.FixedRows
SlNo = 0
If Not RstRepay Is Nothing Then RstRepay.MoveLast
While Not rstLoan.EOF
    
    'Initialise the Varible
    'LoanID = FormatField(rstLoan("LoanID"))
    SchemeID = FormatField(rstLoan("SchemeID"))
    
    'Get the Loan Demand od this Scheme

    
    ColAmount(3) = L_clsLoan.OverDueAmount(, SchemeID, ObDate)     'Over due as on Opeing date
    ColAmount(4) = L_clsLoan.LoanDemand(LoanID, ObDate, m_ToDate)  'Amount which falls as over due between OB Date & today
    ColAmount(5) = L_clsLoan.OverDueAmount(, SchemeID, m_ToDate)    'Over due as on today
    
    ColAmount(5) = ColAmount(3) + ColAmount(4)
    
    'The difference between two date is the Loan demand of that period
    'Therefore
    'ColAmount(4) = ColAmount(5) - ColAmount(3)
    
    curRepay = 0: ' curAdvance = 0
    If Not RstRepay Is Nothing Then
        RstRepay.MoveFirst
        RstRepay.Find "LoanID = " & LoanID
        If Not RstRepay.EOF Then curRepay = FormatField(RstRepay(0))
        
    End If
    Debug.Assert ColAmount(5) - curRepay = L_clsLoan.OverDueAmount(LoanID, , m_ToDate)
'    If Not rstAdvance Is Nothing Then
'        rstRepay.FindFirst "LoanID = " & LoanID
'        If Not rstAdvance.NoMatch And rstAdvance("LoanID") = LoanID Then
'            curAdvance = FormatField(rstAdvance(0))
'        End If
'    End If
    ColAmount(9) = FormatCurrency(curRepay)
    ColAmount(10) = 0: ColAmount(6) = 0: ColAmount(7) = 0
    ColAmount(8) = 0
    'Now Calculate the Recovery demand
    If curRepay > 0 Then
        'Now calculate the recovery against arrears demand
        If ColAmount(3) >= curRepay Then
            ColAmount(6) = curRepay: curRepay = 0
        Else
            ColAmount(6) = ColAmount(3): curRepay = curRepay - ColAmount(6)
        End If
        'Now calculate the recovery against current demand
        If ColAmount(4) >= curRepay Then
            ColAmount(7) = curRepay: curRepay = 0
        Else
            ColAmount(7) = ColAmount(4): curRepay = curRepay - ColAmount(7)
        End If
        'Remianinog amount is the advance recovery
        ColAmount(8) = curRepay
    End If
    
    ColAmount(10) = ColAmount(5) - ColAmount(6) - ColAmount(7)  'OVer due amount as on date
    ODAmount = ColAmount(10)
    
    'Over due amount as on date
    ODAmount = L_clsLoan.OverDueAmount(LoanID, , m_ToDate)
    PrevOD = 0
    ColAmount(16) = L_clsLoan.OverDueSince(5, LoanID, , m_ToDate)
    If ColAmount(16) > ODAmount Then ColAmount(16) = ODAmount 'Over due since 5 & above 5 Years
    ODAmount = ODAmount - ColAmount(16)
    
    ColAmount(15) = L_clsLoan.OverDueSince(4, LoanID, , m_ToDate) - ColAmount(16)
    If ColAmount(15) > ODAmount Then ColAmount(15) = ODAmount 'Over due since 4 Years
    ODAmount = ODAmount - ColAmount(15)
    
    ColAmount(14) = L_clsLoan.OverDueSince(3, LoanID, , m_ToDate) - ColAmount(15)  'Over due since 3 Years
    If ColAmount(14) > ODAmount Then ColAmount(14) = ODAmount 'Over due since 3 Years
    ODAmount = ODAmount - ColAmount(14)
    
    ColAmount(13) = L_clsLoan.OverDueSince(2, LoanID, , m_ToDate) - ColAmount(14)  'Over due since 2 Years
    If ColAmount(13) > ODAmount Then ColAmount(13) = ODAmount 'Over due since 2 Years
    ODAmount = ODAmount - ColAmount(13)
    
    ColAmount(12) = L_clsLoan.OverDueSince(1, LoanID, , m_ToDate) - ColAmount(13)
    If ColAmount(12) > ODAmount Then ColAmount(12) = ODAmount 'Over due since 1 Year
    ODAmount = ODAmount - ColAmount(12)
    
    ColAmount(11) = ODAmount 'Over due under one year
    AddRow = False
    For count = 3 To grd.Cols - 1
        If ColAmount(count) Then
            AddRow = True
            Exit For
        End If
    Next
    If AddRow Then
        With grd
            If .Rows <= .Row + 2 Then .Rows = .Rows + 1
            .Row = .Row + 1
            SlNo = SlNo + 1
            .Col = 0: .Text = SlNo
            .Col = 1: .Text = FormatField(rstLoan("AccNum"))
            .Col = 2: .Text = FormatField(rstLoan("Name"))
            For count = 3 To .Cols - 1
                If ColAmount(count) < 0 Then ColAmount(count) = 0
                .Col = count: .Text = FormatCurrency(ColAmount(count))
                GrandTotal(count) = GrandTotal(count) + ColAmount(count)
            Next
        End With
    End If
    DoEvents
    RaiseEvent Processing("Writing the record", (rstLoan.AbsolutePosition / rstLoan.RecordCount))
    rstLoan.MoveNext
Wend

AddRow = False
For count = 3 To grd.Cols - 1
    If GrandTotal(count) Then
        AddRow = True
        Exit For
    End If
Next
If AddRow Then
    With grd
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        If .Rows <= .Row + 2 Then .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 2: .Text = "Grand Total": .CellFontBold = True
        For count = 3 To .Cols - 1
            If GrandTotal(count) < 0 Then ColAmount(count) = 0
            .Col = count: .CellFontBold = True
            .Text = FormatCurrency(GrandTotal(count))
        Next
    End With
End If

Set L_clsLoan = Nothing
lblReportTitle = "Demand, collecion and blance register of fo the month of " & _
    GetMonthString(Month(m_ToDate)) & " " & GetFromDateString(m_ToIndianDate)
ShowShed1_All = True
Exit Function

ErrLine:
    MsgBox "error Showshed1" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
  'Resume
End Function


Private Function ShowShed2() As Boolean

RaiseEvent Processing("Fetching the record", 0)
m_repType = repShedule_2

On Error GoTo ExitLine

'Declarations

Dim rstCurrBalance As Recordset
Dim rstOpBalance As Recordset
Dim rstRepayTill  As Recordset
Dim rstAdvanceTill As Recordset
Dim rstRepayLast  As Recordset
Dim rstAdvanceLast As Recordset

Dim SqlStr As String
Dim count As Integer

Dim SlNo As Integer
Dim strBranchName As String


'Decalration By SHashi
Dim ColAmount() As Currency
Dim GrandTotal() As Currency
Dim fromDate As Date
Dim LastDate As Date
Dim ObDate As Date
Dim LoanCategary As wisLoanCategories

Dim LExcel As Boolean ' to be removed later and get the data from outside. - pradeep

'LExcel = True
If LExcel Then
    'Set xlWorkBook = Workbooks.Add
    Set xlWorkSheet = xlWorkBook.Sheets(1)
End If

ShowShed2 = False


LoanCategary = wisAgriculural
LastDate = GetSysLastDate(m_ToDate)

'GetLast Friday date
'm_toDate = DateAdd("M", 1, m_toDate)
'm_toDate = Month(m_toDate) & "/1/" & Year(m_toDate)
''Get the End Of LastMonth
'LastDate = DateAdd("m", -2, m_toDate)
'Do
'    m_toDate = DateAdd("d", -1, m_toDate)
'    If Format(m_toDate, "dddd") = "Friday" Then Exit Do
'Loop

If LoanCategary = wisAgriculural Then
    'ObDate = "7/1/" & IIf(Month(m_ToDate) > 6, Year(m_ToDate), Year(m_ToDate) - 1)
    ObDate = GetSysFormatDate("1/7/" & IIf(Month(m_ToDate) > 6, Year(m_ToDate), Year(m_ToDate) - 1))
Else
    'ObDate = "4/1/" & IIf(Month(m_ToDate) > 3, Year(m_ToDate), Year(m_ToDate) - 1)
    ObDate = FinUSFromDate
End If


'Me.lblBranchName = strBranchName
RaiseEvent Initialise(0, 15)

'Get The Details Of Loans, and Balance as ondate
    RaiseEvent Processing("Fetching the record", 0.15)
SqlStr = "SELECT A.LoanId, AccNum, Balance, " & _
    " Title + ' ' + FirstName +' ' + MiddleName + ' ' + LastName As Name " & _
    " From LoanMaster A,LoanTrans B,NameTab C WHERE A.Schemeid IN " & _
        " (SELECT SchemeID From LoanScheme WHERE Category = " & LoanCategary & ")" & _
    " AND A.LoanID = B.LoanID" & _
    " AND C.CustomerID = A.CustomerId AND TransID = (SELECT MAX(TransID) " & _
        " From LoanTrans D Where D.LoanId = A.LoanId" & _
        " AND D.TransDate <= #" & m_ToDate & "#)" & _
    " ORDER BY A.LoanID"

gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstCurrBalance, adOpenStatic) < 1 Then Exit Function

'RaiseEvent Initialise(0, rstCurrBalance.RecordCount )

RaiseEvent Processing("Fetching the record", 0.55)
'Get The Details Of Loans, and Balance as on 31/M/yyyy
SqlStr = "SELECT A.LoanId, AccNum, Balance,TransDate" & _
    " From LoanMaster A,LoanTrans B WHERE A.LoanID = B.LoanID" & _
    " AND A.LoanID = B.LoanID AND TransID = " & _
            " (SELECT MAX(TransID) From LoanTrans D Where " & _
            " D.LoanId = A.LoanId AND D.TransDate < #" & ObDate & "#)"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstOpBalance, adOpenStatic) < 1 Then Exit Function

Dim transType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes
'Now Get the Loan Repayment From Ob Date to Last Month
transType = wDeposit
ContraTrans = wContraDeposit
RaiseEvent Processing("Fetching the record", 0.65)
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans Where " & _
    " TransDate >= #" & ObDate & "# AND TransDate <= #" & LastDate & "#" & _
    " AND (TransType = " & transType & " OR TransType = " & ContraTrans & ")" & _
    " GROUP BY LoanID"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstRepayLast, adOpenStatic) < 1 Then Set rstRepayLast = Nothing

'Now Get the Loan Repayment From Lastdate to till date
RaiseEvent Processing("Fetching the record", 0.75)
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans Where TransDate > #" & LastDate & "# " & _
    " AND TransDate <= #" & m_ToDate & "# AND (TransType = " & transType & _
    " OR TransType = " & ContraTrans & ") " & _
    " GROUP BY LoanID"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstRepayTill, adOpenStatic) < 1 Then Set rstRepayTill = Nothing

'Now Get the Loan Advance From Ob date to last month
transType = wWithdraw
ContraTrans = wContraWithdraw
RaiseEvent Processing("Fetching the record", 0.85)
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans Where " & _
    " TransDate >= #" & ObDate & "# AND TransDate <= #" & LastDate & "#" & _
    " AND (TransType = " & transType & " OR TransType = " & ContraTrans & ")" & _
    " GROUP BY LoanID"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstAdvanceLast, adOpenStatic) < 1 Then Set rstAdvanceLast = Nothing

'Now Get the Loan Advance during this month
RaiseEvent Processing("Fetching the record", 0.95)
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans Where " & _
    " TransDate > #" & LastDate & "# AND TransDate <= #" & m_ToDate & "#" & _
    " AND (TransType = " & transType & " OR TransType = " & ContraTrans & ")" & _
    " GROUP BY LoanID"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.Fetch(rstAdvanceTill, adOpenStatic) < 1 Then Set rstAdvanceTill = Nothing


SlNo = 0
grd.Visible = True

' now set the shed headers
Call Shed2RowColKan

grd.Row = grd.FixedRows

' beginning of the loantrans loop
ReDim ColAmount(3 To grd.Cols - 1)
ReDim SubTotal(3 To grd.Cols - 1)
ReDim GrandTotal(3 To grd.Cols - 1)

Dim LoanID As Long
Dim AddRow As Boolean
Dim L_clsLoan As New clsLoan

RaiseEvent Initialise(0, rstCurrBalance.RecordCount)

While Not rstCurrBalance.EOF
    LoanID = FormatField(rstCurrBalance("LoanID"))
    'Get the opening balance
    ColAmount(3) = 0
    If Not rstOpBalance Is Nothing Then
        rstOpBalance.MoveFirst
        rstOpBalance.Find "LoanID = " & LoanID
        If Not rstOpBalance.EOF Then
            ColAmount(3) = FormatField(rstOpBalance("Balance")) ' Opening Balance
        End If
    End If
    
    'loan advanced in the last month
    ColAmount(4) = 0
    If Not rstAdvanceLast Is Nothing Then
        rstAdvanceLast.MoveFirst
        rstAdvanceLast.Find "LoanId = " & LoanID
        If Not rstAdvanceLast.EOF Then
            ColAmount(4) = FormatField(rstAdvanceLast(0)) ' Advanced up to Previous month
        End If
    End If
    
    'loan advanced during this month
    ColAmount(5) = 0
    If Not rstAdvanceTill Is Nothing Then
        rstAdvanceTill.MoveFirst
        rstAdvanceTill.Find "LoanId = " & LoanID
        If Not rstAdvanceTill.EOF Then
            ColAmount(5) = FormatField(rstAdvanceTill(0)) ' Advanced during the month
        End If
    End If
    'Toal Loan Advance upto end of this month
    ColAmount(6) = ColAmount(4) + ColAmount(5)
    'Maxmum Loan Balance during the month
    ColAmount(7) = ColAmount(3) + ColAmount(6)
    
    'Loan recoverd up to last month
    ColAmount(8) = 0
    If Not rstRepayLast Is Nothing Then
        rstRepayLast.MoveFirst
        rstRepayLast.Find "LoanId = " & LoanID
        If Not rstRepayLast.EOF Then
            ColAmount(8) = FormatField(rstRepayLast(0)) ' Recovery up to Previous month
        End If
    End If
    
    'Loan recoverd during this month
    ColAmount(9) = 0
    If Not rstRepayTill Is Nothing Then
        rstRepayTill.MoveFirst
        rstRepayTill.Find "LoanId = " & LoanID
        If Not rstRepayTill.EOF Then
            ColAmount(9) = FormatField(rstRepayTill(0)) ' Recovery during this month
        End If
    End If
    
    
    'Total recovery at the end of month
    ColAmount(10) = ColAmount(8) + ColAmount(9)
    'Loan Balance at the end this month
    ColAmount(11) = ColAmount(7) - ColAmount(10)
    
    Debug.Assert ColAmount(11) = FormatField(rstCurrBalance("Balance"))
    'OVER DUE amount as on end of this month
    'i.e. OD of the abave balance
    ColAmount(12) = L_clsLoan.OverDueAmount(LoanID, , m_ToDate)
    
    AddRow = False
    For count = 3 To grd.Cols - 1
        If ColAmount(count) Then
            AddRow = True
            Exit For
        End If
    Next
    If AddRow Then
      With grd
        If .Rows <= .Row + 2 Then .Rows = .Rows + 2
        .Row = .Row + 1: SlNo = SlNo + 1
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = FormatField(rstCurrBalance("AccNum"))
        .Col = 2: .Text = FormatField(rstCurrBalance("Name"))
        '.Col = 3: .Text = FormatField(rstCurrBalance("firstName"))
        For count = 3 To grd.Cols - 1
            grd.Col = count: grd.Text = FormatCurrency(ColAmount(count))
            GrandTotal(count) = GrandTotal(count) + ColAmount(count)
        Next
      End With
    End If
    DoEvents
    If gCancel Then rstCurrBalance.MoveLast
    RaiseEvent Processing("Writing the record", rstCurrBalance.AbsolutePosition / rstCurrBalance.RecordCount)
    rstCurrBalance.MoveNext
Wend
AddRow = False
For count = 4 To grd.Cols - 1
    If GrandTotal(count) Then
        AddRow = True
        Exit For
    End If
Next
If AddRow Then
  With grd
    If .Rows <= .Row + 3 Then .Rows = .Rows + 2
    .Row = .Row + 2: SlNo = SlNo + 1
    .Col = 2: .Text = "Grand Total": .CellFontBold = True
    For count = 4 To grd.Cols - 1
        grd.Col = count: .CellFontBold = True
        grd.Text = FormatCurrency(GrandTotal(count))
    Next
  End With
End If

Set L_clsLoan = Nothing

lblReportTitle.Caption = "Statement showing the short term loans for the month of " & _
    GetMonthString(Month(m_ToDate)) & " " & GetFromDateString(m_ToIndianDate)

grd.Visible = True
         
If LExcel Then
    xlWorkBook.SaveAs App.Path & "|" & "shed2.xls"
    xlWorkBook.Close savechanges:=True
            
    Set xlWorkSheet = Nothing
    Set xlWorkBook = Nothing
End If
         
         
ShowShed2 = True

ExitLine:
Screen.MousePointer = vbDefault
grd.Visible = True

If Err Then
    MsgBox "ERROR ShowShed2" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
End If

End Function

Private Sub SetKannadaCaption()

Call SetFontToControls(Me)

cmdPrint.Caption = GetResourceString(23)  'Print
cmdOK.Caption = GetResourceString(1) 'OK


End Sub

Private Sub cmdOk_Click()
Unload Me

End Sub

Private Sub cmdPrint_Click()
Set m_grdPrint = wisMain.grdPrint
With m_grdPrint
    .GridObject = grd
    .ReportTitle = Me.lblReportTitle
    .CompanyName = gCompanyName
    .Font.name = grd.Font.name
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

Private Sub Form_Load()

Call CenterMe(Me)
Call SetKannadaCaption

lblBankName.Caption = gCompanyName

If m_repType = repConsBalance Then ShowConsoleBalance
If m_repType = repConsInstOD Then ShowConsoleInstOverDue
If m_repType = repConsOD Then ShowConsoleODBalance
If m_repType = repMonthlyRegister Then ShowMeetingRegistar
If m_repType = repShedule_1 Then ShowShed1
If m_repType = repShedule_2 Then ShowShed2
'If m_repType = repShedule_3 Then ShowShed3
If m_repType = repShedule_4A Then ShowShed4A
If m_repType = repShedule_4B Then ShowShed4B
If m_repType = repShedule_4C Then ShowShed4C
If m_repType = repShedule_5 Then ShowShed5
If m_repType = repShedule_6 Then ShowShed6

End Sub

Private Sub Form_Resize()
Const Margin = 50
Const CTL_MARGIN = 50
'Const BOTTOM_MARGIN = 600
    Screen.MousePointer = vbDefault
    On Error Resume Next
    lblBankName.Top = 0
    lblBankName.Left = (Me.Width - lblBankName.Width) / 2
    grd.Left = 10
    'lblBranch.Top = lblFrom.Top + lblFrom.Height + 50
    lblReportTitle.Top = lblBankName.Top + lblBankName.Height + 50
    lblReportTitle.Left = (Me.Width - lblReportTitle.Width) / 2
    'lblBranch.Left = lblBankname.Left
    grd.Top = lblReportTitle.Top + lblReportTitle.Height + 200
    grd.Width = Me.Width - 120
    grd.Height = Me.ScaleHeight - (lblBankName.Height + lblBankName.Height + lblReportTitle.Height + cmdPrint.Height + 370)
    
    fra.Top = Me.ScaleHeight - fra.Height
    fra.Left = Me.Width - fra.Width
    grd.Height = Me.ScaleHeight - fra.Height - grd.Top
    cmdOK.Left = fra.Width - cmdOK.Width - (cmdOK.Width / 4)
    cmdPrint.Left = cmdOK.Left - cmdPrint.Width - (cmdPrint.Width / 8)
    cmdPrint.Top = cmdOK.Top
    cmdWeb.Top = cmdPrint.Top
    cmdWeb.Left = cmdPrint.Left - cmdPrint.Width - (cmdPrint.Width / 4)
    ' removed the call for personal use - pradeep
    Call InitGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim count  As Integer
For count = 0 To grd.Cols - 1
    Call SaveSetting(App.EXEName, "LoanReport" & m_repType, "ColWidth", grd.ColWidth(count))
Next
End Sub

Private Sub grd_LostFocus()
Dim ColCount As Integer
For ColCount = 0 To grd.Cols - 1
Call SaveSetting(App.EXEName, "LoanReport" & m_repType, _
    "ColWidth" & ColCount, grd.ColWidth(ColCount) / grd.Width)
Next ColCount

End Sub

Private Sub PrintClass_MaxProcessCount(MaxCount As Long)
m_Count = 1
m_MaxCount = MaxCount
RaiseEvent Initialise(0, MaxCount)
End Sub

Private Sub PrintClass_ProcessCount(count As Long)
m_Count = count

End Sub


Private Sub PrintClass_ProcessingMessage(strMessage As String)
On Error Resume Next
RaiseEvent Processing(strMessage, m_Count / m_MaxCount)

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


