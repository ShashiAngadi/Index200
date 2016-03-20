VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBKCCRegReport 
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
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Close"
         Height          =   345
         Left            =   3300
         TabIndex        =   6
         Top             =   90
         Width           =   1155
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   345
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
Attribute VB_Name = "frmBKCCRegReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Initialise(Min As Long, Max As Long)
Public Event Processing(strMessage As String, Ratio As Single)

'Private WithEvents PrintClass As clsPrint

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

'Private xlWorkBook As Workbook
'Private xlWorkSheet As Worksheet
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
        m_ToDate = FormatDate(NewDate)
    Else
        m_ToIndianDate = ""
        m_ToDate = vbNull
    End If
End Property

Public Property Let FromIndianDate(NewDate As String)
    If DateValidate(NewDate, "/", True) Then
        m_FromIndianDate = NewDate
        m_FromDate = FormatDate(NewDate)
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
        '.ColWidth(Items) = Val(HeadArray(Items, 1))
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
        
        '.ColWidth(Items) = Val(HeadArray(Items, 1))
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
COlHeader(1) = "Catewise Distribution of Loans"
COlHeader(2) = "Catewise Distribution of Loans"
COlHeader(3) = "Catewise Distribution of Loans"
COlHeader(4) = "Catewise Distribution of Loans"
COlHeader(5) = "Catewise Distribution of Loans"
COlHeader(6) = "Catewise Distribution of Loans"
COlHeader(7) = "Catewise Distribution of Loans"
COlHeader(8) = "Catewise Distribution of Loans"
COlHeader(9) = "Catewise Distribution of Loans"

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

Private Function ShowShed4C() As Boolean

' contact pradeep for this function
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
         
         
gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(rstLoanDetail, adOpenDynamic) < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
Else
    'Set rstLoanDetail = gDBTrans.Rst.Clone
End If

' caste wise rst
SqlStr = " SELECT a.RefID,a.CasteID,c.CasteName,a.Male_No,a.Male_Amount,a.Female_No,a.Female_Amount" & _
         " FROM LoanCasteWise a,NewLoans b,Caste c" & _
         " WHERE a.RefId=b.RefId" & _
         " AND a.CasteId=c.CasteId " & _
         " AND b.LoanDate >= " & "#" & m_FromDate & "#" & _
         " AND b.Loandate <= " & "#" & m_ToDate & "#"
         
gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(rstCasteWise, adOpenDynamic) < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
Else
    'Set rstCasteWise = gDBTrans.Rst.Clone
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
    
    If (Not rstCasteWise.EOF) And (Not rstCasteWise.EOF) Then  ' if found and not eof
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
Dim totAmount As Currency

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
         " WHERE a.BankId = b.BankId" & _
         " AND a.CropId=c.CropId" & _
         " AND a.LoanDate >= " & "#" & m_FromDate & "#" & _
         " AND a.Loandate <= " & "#" & m_ToDate & "#" & _
         " ORDER BY a.RefId"
         
gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(rstLoanDetail, adOpenDynamic) < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
Else
    'Set rstLoanDetail = gDBTrans.Rst.Clone
End If

' this main rst
SqlStr = " SELECT a.RefID,a.BF_No,a.BF_Amount,a.SF_No,a.SF_Amount,a.SCST_No,a.SCST_Amount" & _
         " FROM NewLoanMembers a,NewLoans b " & _
         " WHERE a.RefID=b.RefID" & _
         " AND b.LoanDate >= " & "#" & m_FromDate & "#" & _
         " AND b.Loandate <= " & "#" & m_ToDate & "#" & _
         " ORDER BY a.RefID"
         
gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(rstNewMems, adOpenDynamic) < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
Else
    'Set rstNewMems = gDBTrans.Rst.Clone
End If

' caste wise rst
SqlStr = " SELECT a.RefID,a.CasteID,c.CasteName,a.Male_No,a.Male_Amount,a.Female_No,a.Female_Amount" & _
         " FROM LoanCasteWise a,NewLoans b,Caste c" & _
         " WHERE a.RefId=b.RefId" & _
         " AND b.LoanDate >= " & "#" & m_FromDate & "#" & _
         " AND b.Loandate <= " & "#" & m_ToDate & "#" & _
         " AND a.CasteId=c.CasteId "
         
         
gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(rstCasteWise, adOpenDynamic) < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
Else
    'Set rstCasteWise = gDBTrans.Rst.Clone
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
    totAmount = BF_Amount + SF_Amount + SC_Amount
    
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
        .Col = ColNum: .Text = FormatCurrency(totAmount): ColNum = ColNum + 1
        
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
Dim totAmount As Currency

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

SqlStr = " SELECT a.BankId,BankName,a.CropId,CropName,LoanDate,LoanDueDate,BF_No,BF_Amount," & _
         " SF_No,SF_Amount,SCST_No,SCST_Amount" & _
         " FROM Shed4 a,BranchDet b,Crops c" & _
         " WHERE a.BankId = b.BankId" & _
         " AND a.CropId=c.CropId" & _
         " AND a.LoanDate >= " & "#" & m_FromDate & "#" & _
         " AND a.Loandate <= " & "#" & m_ToDate & "#"
         
gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(rstLoanDetail, adOpenDynamic) < 1 Then
    MsgBox "Unable to fetch the Data"
    Exit Function
Else
    'Set rstLoanDetail = gDBTrans.Rst.Clone
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
        totAmount = BF_Amount + SF_Amount + SC_Amount
              
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
        .Col = ColNum: .Text = FormatCurrency(totAmount): ColNum = ColNum + 1
        
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
    .Col = 0: .Text = "Sl No"
    .Col = 1: .Text = "Loan No"
    .Col = 2: .Text = "Name of the customer"
    .Col = 3: .Text = "Demand"
    .Col = 4: .Text = "Demand"
    .Col = 5: .Text = "Demand"
    .Col = 6: .Text = "Recovery against"
    .Col = 7: .Text = "Recovery against"
    .Col = 8: .Text = "Recovery against"
    .Col = 9: .Text = "Recovery against"
    .Col = 10: .Text = "Overdue"
    .Col = 11: .Text = "Overdue"
    .Col = 12: .Text = "Overdue"
    .Col = 13: .Text = "Overdue"
    .Col = 14: .Text = "Overdue"
    .Col = 15: .Text = "Overdue"
    .Col = 16: .Text = "Overdue"
    
    .Row = 1
    .Col = 0: .Text = "Sl No"
    .Col = 1: .Text = "Loan No"
    .Col = 2: .Text = "Name of the customer"
    .Col = 3: .Text = "Arrears"
    .Col = 4: .Text = "Current"
    .Col = 5: .Text = "Total" & vbCrLf & "(3+4)"
    .Col = 6: .Text = "Arrears Demand"
    .Col = 7: .Text = "Current Demand"
    .Col = 8: .Text = "Advance Recovery If any"
    .Col = 9: .Text = "Total" & vbCrLf & "(6 + 7 + 8)"
    .Col = 10: .Text = "Balance of Overdue" & vbCrLf & "(5-(6+7))"
    .Col = 11: .Text = "Less than One year"
    .Col = 12: .Text = "1 to 2 years"
    .Col = 13: .Text = "2 to 3 years"
    .Col = 14: .Text = "3 to 4 years"
    .Col = 15: .Text = "4 to 5 years"
    .Col = 16: .Text = "Above 5 Years"
    .RowHeight(1) = 800
    
    
    Dim i As Integer
    Dim j As Integer
    .Row = 2
    For j = 3 To .Cols - 1
        .Col = j: .Text = Format(j, "00")
    Next
    .Col = 0: .Text = "01"
    .Col = 1: .Text = "02"
    .Col = 2: .Text = "2a"
    
    .MergeCells = flexMergeRestrictRows
    For i = 0 To .FixedRows - 1
        .Row = i
        For j = 0 To .Cols - 1
            .Col = j: .MergeCol(j) = True
            .CellFontBold = True
            .CellAlignment = 4
        Next
        .MergeRow(i) = True
    Next
End With

End Sub

Private Sub Shed2RowCol()

With grd
    .Clear
    .Rows = 1: .Cols = 1
    .Rows = 10: .Cols = 14
    .FixedCols = 2: .FixedRows = 2
    .AllowUserResizing = flexResizeBoth
    .Row = 0
    .Col = 0: .Text = "Sl No"
    .Col = 1: .Text = "Loan No"
    .Col = 2: .Text = "Name of the Customer"
    .Col = 3: .Text = "Name of the Society"
    .Col = 4: .Text = "Loan Outstanding as on 1 st July"
    .Col = 5: .Text = "Loan Advanced Up to the previous month"
    .Col = 6: .Text = "Loan Advanced during the month"
    .Col = 7: .Text = "Total Loan Advanced upto the end of the month" & vbCrLf & "(4 + 5)"
    .Col = 8: .Text = "Outstanding at the end of the month" & vbCrLf & "3 + 6"
    .Col = 9: .Text = "Recovery Upto Previous month"
    .Col = 10: .Text = "Recovery during the month"
    .Col = 11: .Text = "Total Recovery upto the end of the the month" & vbCrLf & "8 + 9 "
    .Col = 12: .Text = "Balance at the end of the month" & vbCrLf & "7 - 10"
    .Col = 13: .Text = "Out of Which overdue as at the end of the month"
    .RowHeight(0) = 1200
    Dim i As Integer, j As Integer
    .WordWrap = True
    .Row = 1
    For j = 0 To .Cols - 1
         .Col = j: .Text = j - 1
    Next
    .Col = 0: .Text = "1"
    .Col = 1: .Text = "2"
    .Col = 2: .Text = "2a"
    .Col = 3: .Text = "2b"
    For i = 0 To .FixedRows - 1
        .Row = i
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

With grd
    .Clear
    .Rows = 1: .Cols = 1
    .Rows = 10: .Cols = 16
    .FixedCols = 2: .FixedRows = 3
    .AllowUserResizing = flexResizeBoth
    
    .Row = 0
    .Col = 0: .Text = "slno"
    .Col = 1: .Text = "Loan No"
    .Col = 2: .Text = "Name of Customer"
    .Col = 3: .Text = "Date of Sanction"
    .Col = 4: .Text = "Due Date"
    .Col = 5: .Text = "Balance at the Begining of the Month"
    .Col = 6: .Text = "During the Month"
    .Col = 7: .Text = "During the Month"
    .Col = 8: .Text = "Outstanding at the End of the Month"
    .Col = 9: .Text = "of which Over Dues"
    .Col = 10: .Text = "Classification of Overdue"
    .Col = 11: .Text = "Classification of Overdue"
    .Col = 12: .Text = "Classification of Overdue"
    .Col = 13: .Text = "Classification of Overdue"
    .Col = 14: .Text = "Classification of Overdue"
    .Col = 15: .Text = "Classification of Overdue"
    
    .Row = 1
    .Col = 0: .Text = "slno"
    .Col = 1: .Text = "Loan No"
    .Col = 2: .Text = "Name of Customer"
    .Col = 3: .Text = "Date of Sanction"
    .Col = 4: .Text = "Due Date"
    .Col = 5: .Text = "Balance at the Begining of the Month"
    .Col = 6: .Text = "Advances "
    .Col = 7: .Text = "Recovered"
    .Col = 8: .Text = "Outstanding at the End of the Month"
    .Col = 9: .Text = "of which Over Dues"
    .Col = 10: .Text = "Under 1 Year"
    .Col = 11: .Text = "1 to 2 Years"
    .Col = 12: .Text = "2 to 3 Years"
    .Col = 13: .Text = "3 to 4 Years"
    .Col = 14: .Text = "4 to 5 Years"
    .Col = 15: .Text = "Above 5 Years"

    .RowHeight(1) = 700
    
    Dim i As Integer, j As Integer
    .Row = 2
    For i = 3 To .Cols - 1
        .Col = i: .Text = (i)
    Next
    .Col = 0: .Text = "1"
    .Col = 1: .Text = "2"
    .Col = 2: .Text = "2a"
    
    .MergeCells = flexMergeFree
    For i = 0 To .FixedRows - 1
        .Row = i
        For j = 0 To .Cols - 1
             .MergeCol(j) = True
            .Col = j: .CellAlignment = 4: .CellFontBold = True
        Next
        .MergeRow(i) = True
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
    .Rows = 10: .Cols = 20
    .FixedCols = 2: .FixedRows = 3
    .Row = 0
    .Col = 0: .Text = "Slno"
    .Col = 1: .Text = "Loan No"
    .Col = 2: .Text = "Name of Customer"
    .Col = 3: .Text = "Name of Society"
    .Col = 4: .Text = "Limit Sanctioned"
    .Col = 5: .Text = "Date of Sanction"
    .Col = 6: .Text = "Due Date"
    .Col = 7: .Text = "Purpose"
    .Col = 8: .Text = "Balance at the Begining of the Month"
    .Col = 9: .Text = "During the Month"
    .Col = 10: .Text = "During the Month"
    .Col = 11: .Text = "Outstanding at the End of the Month"
    .Col = 12: .Text = "Maximum Outstanding During the Month"
    .Col = 13: .Text = "of which Over Dues"
    .Col = 14: .Text = "Calssification Over Dues"
    .Col = 15: .Text = "Calssification Over Dues"
    .Col = 16: .Text = "Calssification Over Dues"
    .Col = 17: .Text = "Calssification Over Dues"
    .Col = 18: .Text = "Calssification Over Dues"
    .Col = 19: .Text = "Calssification Over Dues"
    
    .Row = 1
    .Col = 0: .Text = "Slno"
    .Col = 1: .Text = "Loan No"
    .Col = 2: .Text = "Name of Customer"
    .Col = 3: .Text = "Name of Society"
    .Col = 4: .Text = "Limit Sanctioned"
    .Col = 5: .Text = "Date of Sanction"
    .Col = 6: .Text = "Due Date"
    .Col = 7: .Text = "Purpose"
    .Col = 8: .Text = "Balance at the Begining of the Month"
    .Col = 9: .Text = "Advances"
    .Col = 10: .Text = "Recovered "
    .Col = 11: .Text = "Outstanding at the End of the Month"
    .Col = 12: .Text = "Maximum Outstanding During the Month"
    .Col = 13: .Text = "of which Over Dues"
    '.Col = 12: .Text = "Max Outstanding during Month"
    '.Col = 13: .Text = "Of Which Overdue"
    .Col = 14: .Text = "Under 1 Year"
    .Col = 15: .Text = "1 to 2 Years"
    .Col = 16: .Text = "2 to 3 Years"
    .Col = 17: .Text = "3 to 4 Years"
    .Col = 18: .Text = "4 to 5 Years"
    .Col = 19: .Text = "Above 5 Years"
    .RowHeight(1) = 700
    
    Dim i As Integer, j As Integer
    .Row = 2
    For i = 4 To .Cols - 1
        .Col = i: .Text = (i - 1)
    Next
    .Col = 0: .Text = "1"
    .Col = 1: .Text = "2"
    .Col = 2: .Text = "2a"
    .Col = 3: .Text = "2b"
    
    .MergeCells = flexMergeFree
    For i = 0 To .FixedRows - 1
        .Row = i
        For j = 0 To .Cols - 1
             .MergeCol(j) = True
             .Col = j
            .CellAlignment = 4: .CellFontBold = True
        Next
        .MergeRow(i) = True
    Next
    .MergeCol(15) = False: .MergeCol(16) = False
    .MergeCol(17) = False: .MergeCol(18) = False
    .MergeCol(19) = False ':.MergeCol(11) = False
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
colHeads(1) = "Name of the Society "
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
colHeads(1) = "Name of the Society "
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
Dim wid As Single
For ColCount = 0 To grd.Cols - 1
    wid = GetSetting(App.EXEName, "LoanReport" & m_repType, "ColWidth" & ColCount, grd.Width / grd.Cols) * grd.Width
    If wid >= grd.Width * 0.9 Then wid = grd.Width / grd.Cols
    If wid < 20 And wid <> 0 Then wid = grd.Width / grd.Cols * 2
    grd.ColWidth(ColCount) = wid
Next ColCount

End Sub


Private Function ShowShed6() As Boolean


ShowShed6 = False
Err.Clear
On Error GoTo Exitline:
RaiseEvent Initialise(0, 10)
RaiseEvent Processing("Fetching record", 0)

Dim SqlStr As String
Dim TransType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes

Dim rstOpBalance As Recordset
Dim rstClBalance As Recordset
Dim rstAdvance As Recordset
Dim rstRecovery As Recordset

Dim FirstDate As Date

Dim ColAmount() As Currency
Dim GrandTotal() As Currency

'Get the First day of the Month
FirstDate = Month(m_FromDate) & "/1/" & Year(m_FromDate)
Dim LoanType As wis_LoanType
Dim LoanTerm As wisLoanTerm
Dim LoanCategary As wisLoanCategories

LoanTerm = wisLongTerm
LoanType = wisIndividualLoan
LoanCategary = wisNonAgriculural


RaiseEvent Processing("Fetching the record", 0.1)
'Get The LoanDetails And THier balance as on Date
SqlStr = "SELECT A.LoanID,AccNum,IssueDate,RenewDate,Balance," & _
    " Title +' ' + FirstName +' ' + MiddleName + ' ' + LastName As Name" & _
    " FROM BKCCMaster A,BKCCTrans B, NameTab C " & _
    " WHERE B.LoanID = A.LoanID" & _
    " AND C.CustomerId =A.CustomerID AND TransID = (SELECT MAX(TransId) FROM " & _
        " BKCCTrans D WHERE D.TransDate <= #" & m_FromDate & "# " & _
        " AND D.LoanId = A.LoanId ) " & _
    " AND (LoanClosed IS NULL OR LOanClosed = False ) "

gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(rstClBalance, adOpenDynamic) < 1 Then Exit Function

DoEvents
RaiseEvent Processing("Fetching the record", 0.25)
If gCancel Then Exit Function

'Get The LoanDetails And THier balance as on first day of the given month
SqlStr = "SELECT LoanID,TransDate,Balance FROM BKCCTrans A WHERE " & _
    " TransID = (SELECT MAX(TransId) FROM BKCCTrans B WHERE B.TransDate <" & _
        " #" & FirstDate & "# AND B.LoanId = A.LoanId)"
    
gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(rstOpBalance, adOpenDynamic) < 1 Then Exit Function
'Set rstOpBalance = gDBTrans.Rst.Clone
DoEvents
RaiseEvent Processing("Fetching the record", 0.5)
If gCancel Then Exit Function

'Get The Advances During the MOnth
TransType = wWithdraw
ContraTrans = wContraWithdraw
SqlStr = "SELECT SUM(Amount),LoanID FROM BKCCTrans WHERE TransDate >= #" & FirstDate & "#" & _
    " AND TransDate <= #" & m_FromDate & "# AND (TransType = " & TransType & _
    " OR TransType = " & ContraTrans & ") " & _
    " GROUP BY LoanID "
gDbTrans.SQLStmt = SqlStr

Call gDbTrans.Fetch(rstAdvance, adOpenDynamic) ' > 0 Then

DoEvents
RaiseEvent Processing("Fetching the record", 0.65)
If gCancel Then Exit Function


'Get The Recovery During the MOnth
TransType = wDeposit
ContraTrans = wContraDeposit
SqlStr = "SELECT SUM(Amount),LoanID FROM BKCCTrans WHERE TransDate >= #" & FirstDate & "#" & _
    " AND TransDate <= #" & m_FromDate & "# AND (TransType = " & TransType & _
    " OR TransType = " & ContraTrans & ") " & _
    " GROUP BY LoanID "
gDbTrans.SQLStmt = SqlStr
Call gDbTrans.Fetch(rstRecovery, adOpenDynamic)

DoEvents
RaiseEvent Processing("Fetching the record", 0.85)
If gCancel Then Exit Function

'Now Align the grid
Call Shed6RowCol


ReDim ColAmount(5 To grd.Cols - 1)
ReDim GrandTotal(5 To grd.Cols - 1)
'Now Start to writing to the grid
Dim Loanid As Long
Dim AddRow As Boolean
Dim L_clsLoan As New clsBkcc
Dim PrevOD As Currency
Dim ODAmount As Currency
Dim Count As Long
Dim SlNo As Long

RaiseEvent Initialise(0, rstClBalance.RecordCount)

SlNo = 0
While Not rstClBalance.EOF
    
    Loanid = FormatField(rstClBalance("LoanId"))
    
    rstOpBalance.MoveFirst
    rstOpBalance.Find "LoanID = " & Loanid
    ColAmount(5) = 0
    If Not rstOpBalance.EOF Then _
        ColAmount(5) = FormatField(rstOpBalance("Balance")) 'Balance as on 31/3/yyyy
    
    ColAmount(6) = 0
    If Not rstAdvance Is Nothing Then
        rstAdvance.MoveFirst
        rstAdvance.Find "LoanId = " & Loanid
        If Not rstAdvance.EOF Then _
            ColAmount(6) = FormatField(rstAdvance(0))  'Advances During the MOnth
        
    End If
    ColAmount(7) = 0
    If Not rstRecovery Is Nothing Then
        rstRecovery.MoveFirst
        rstRecovery.Find "LoanId = " & Loanid
        If Not rstRecovery.EOF Then _
            ColAmount(7) = FormatField(rstRecovery(0)) 'Recovery During the mOnth
        
    End If
    ColAmount(8) = FormatField(rstClBalance("Balance")) 'Balance at the end of month
    ODAmount = L_clsLoan.OverDueAmount(Loanid, m_FromDate)    'Over due
    ColAmount(9) = ODAmount 'Over due amount of the loan as on given date
    
    'Over due amount calssifiacation
    PrevOD = ODAmount
    ColAmount(15) = L_clsLoan.OverDueSince(5, Loanid, m_FromDate)
                'Over due since & above 5 Years
    ODAmount = ODAmount - ColAmount(15)
    
    ColAmount(14) = L_clsLoan.OverDueSince(4, Loanid, m_FromDate) - ColAmount(15)
    If ColAmount(14) < 0 Then ColAmount(14) = ODAmount 'Over due since 4 Years
    ODAmount = ODAmount - ColAmount(14)
    
    ColAmount(13) = L_clsLoan.OverDueSince(3, Loanid, m_FromDate) - ColAmount(14)
    If ColAmount(13) < 0 Then ColAmount(13) = ODAmount 'Over due since 3 Years
    ODAmount = ODAmount - ColAmount(13)
    
    ColAmount(12) = L_clsLoan.OverDueSince(2, Loanid, m_FromDate) - ColAmount(13)
    If ColAmount(12) < 0 Then ColAmount(12) = ODAmount 'Over due since 2 Years
    ODAmount = ODAmount - ColAmount(12)
    
    ColAmount(11) = L_clsLoan.OverDueSince(1, Loanid, m_FromDate) - ColAmount(12)
    If ColAmount(11) < 0 Then ColAmount(11) = ODAmount 'Over due since a Year
    ODAmount = ODAmount - ColAmount(11)
    
    ColAmount(10) = ODAmount 'Over due under one Year
    
    'Check whther this row has to be write or not
    AddRow = False
    For Count = 5 To grd.Cols - 1
        If ColAmount(Count) Then
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
            .Col = 3: .Text = FormatField(rstClBalance("IssueDate")) ': .MergeCol(4) = False
            .Col = 4: .Text = FormatField(rstClBalance("RenewDate")) ': .MergeCol(5) = False
            For Count = 5 To grd.Cols - 1
                .Col = Count:
                If ColAmount(Count) < 0 Then ColAmount(Count) = 0
                .Text = FormatCurrency(ColAmount(Count))
                GrandTotal(Count) = GrandTotal(Count) + ColAmount(Count)
            Next
        End With
    End If
    DoEvents
    If gCancel Then Exit Function
    RaiseEvent Processing("Writing the records", _
            rstClBalance.AbsolutePosition / rstClBalance.RecordCount)
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
    For Count = 5 To .Cols - 1
        .Col = Count: .CellFontBold = True
        .Text = FormatCurrency(GrandTotal(Count))
    Next
End With

ShowShed6 = True
lblReportTitle.Caption = "Statement showing the long term and other loans for the month of " & _
    GetMonthString(Month(m_FromDate)) & " as on " & m_FromIndianDate

Exitline:
If Err Then
    MsgBox "ERROR SHED 5" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
    Err.Clear
    'Resume
End If

End Function
Private Function ShowMeetingRegistar() As Boolean

Dim SqlPrin As String
Dim SqlInt As String
Dim SqlStr As String
Dim PrinRepay As Currency
Dim IntRepay As Currency

Dim Rst As Recordset
Dim rstLoanScheme As Recordset
Dim SchemeName  As String
Dim Date31_3  As Date
Dim DateLastMonth As Date
'INDIAN dATE FORMAT OF ABOVE VARIABLES
Dim IndDate31_3  As String
Dim IndDateLastMonth As String

Dim TransType As wisTransactionTypes
Dim LoanType As wis_LoanType
Dim SchemeStr  As String

Err.Clear
On Error GoTo ErrLine
'Get all date in the format of system format
Date31_3 = "3/31/" & Val(Year(m_FromDate) - IIf(Month(m_FromDate) > 3, 0, 1))
IndDate31_3 = FormatDate(CStr(Date31_3))

DateLastMonth = FormatDate("1/" & Month(m_FromDate) & "/" & Year(m_FromDate))
DateLastMonth = DateAdd("d", -1, DateLastMonth)
IndDateLastMonth = FormatDate(CStr(DateLastMonth))

RaiseEvent Processing("Fetching the records", 0)
If m_SchemeId Then
    m_repType = repMonthlyRegister
    SqlStr = "SELECT * FROM LoanScheme Where SchemeID = " & m_SchemeId
    gDbTrans.SQLStmt = SqlStr
    Call gDbTrans.Fetch(rstLoanScheme, adOpenDynamic)
    'Set rstLoanScheme = gDBTrans.Rst.Clone
    SchemeName = FormatField(Rst("SchemeName"))
    SchemeStr = " SchemeID = " & m_SchemeId & " "
    LoanType = FormatField(Rst("LoanType"))
Else
    m_repType = repMonthlyRegisterAll
    SchemeStr = " SchemeID  <> " & m_SchemeId & " "
End If

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
'Get The details of loan
SqlStr = "SELECT A.LOanID,AccNUm,IntRate,Guarantor1,Guarantor2, " & _
    " CustomerID, IssueDate,SanctionAmount,B.AbnDate,b.EpDate " & _
    " From BKCCMaster A left Join LoanAbnEp B On A.LoanID = B.LoanID WHERE " & _
    " A.LoanId IN (SELECT Distinct LoanID From BKCCTrans)" & _
    " AND (B.BKCC = 1 or BKCC is NuLL)"
If m_SchemeId Then SqlStr = SqlStr & " AND SchemeID = " & m_SchemeId


SqlStr = SqlStr & " ORDER BY SchemeID,val(AccNum)"

DoEvents
RaiseEvent Initialise(0, 10)
RaiseEvent Processing("Fetching the record", 0.1)
If gCancel Then Exit Function

gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(rstMaster, adOpenDynamic) <= 0 Then GoTo ErrLine

'Get the Loan balance on 31/3/yyyy
SqlPrin = "SELECT A.LoanID,Balance FROM BKCCTrans A WHERE " & _
        " TransDate <= #" & Date31_3 & "#" & _
        " ORDER BY LoanId, TransID Desc"

gDbTrans.SQLStmt = SqlPrin

Call gDbTrans.Fetch(rstPrin31_3, adOpenDynamic)

DoEvents
RaiseEvent Processing("Fetching the record", 0.15)
If gCancel Then Exit Function

'Get the Interest Balance 31/3/yyyy
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM BkccIntTrans B " & _
    " WHERE B.TransId = (SELECT MAX(TransID) FROM " & _
        " BKCCIntTrans C WHERE TransDate <= #" & Date31_3 & "# " & _
        " AND C.LoanID = B.LoanID And C.BankID = B.BankID )"
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM BKCCIntTrans B WHERE " & _
    " TransDate <= #" & Date31_3 & "# ORDER BY LOanID,TransID Desc"
gDbTrans.SQLStmt = SqlInt

Call gDbTrans.Fetch(rstInt31_3, adOpenDynamic)

'Get the Loan balance as on lastMonth
SqlPrin = "SELECT A.LoanID,TransDate,Balance FROM BKCCTrans A WHERE " & _
     " A.TransId = (SELECT MAX(TransID) FROM " & _
        " BkccTrans C WHERE TransDate <= #" & DateLastMonth & "# " & _
        " AND C.LoanID = A.LoanID )"
SqlPrin = "SELECT B.LoanID,Balance,TransDate FROM BKCCTrans B WHERE " & _
        " TransDate <= #" & DateLastMonth & "# ORDER BY LOanID,TransID Desc"

'Get the Interest Balance ON  LAST MONTH
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM BKCCIntTrans B " & _
    " WHERE B.TransId = (SELECT MAX(TransID) FROM " & _
        " BKCCIntTrans C WHERE TransDate <= #" & DateLastMonth & "# " & _
        " AND C.LoanID = B.LoanID )"
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM BKCCIntTrans B WHERE " & _
        " TransDate <= #" & DateLastMonth & "# ORDER BY LOanID,TransID Desc"

gDbTrans.SQLStmt = SqlPrin
Call gDbTrans.Fetch(rstPrinLastMonth, adOpenDynamic)

DoEvents
RaiseEvent Processing("Fetching record", 0.25)
If gCancel Then Exit Function
    
gDbTrans.SQLStmt = SqlInt
Call gDbTrans.Fetch(rstIntLastMonth, adOpenDynamic)

DoEvents
RaiseEvent Processing("Fetching record", 0.35)
If gCancel Then Exit Function

'Get the Loan balance as on date
SqlPrin = "SELECT A.LoanID,TransDate,Balance FROM BKCCTrans A WHERE " & _
     " A.TransId = (SELECT MAX(TransID) FROM " & _
        " BKCCTrans C WHERE TransDate <= #" & m_FromDate & "# " & _
        " AND C.LoanID = A.LoanID )"
SqlPrin = "SELECT B.LoanID,Balance,TransDate FROM BKCCTrans B WHERE " & _
        " TransDate <= #" & m_FromDate & "# ORDER BY LOanID,TransID Desc"
'Get the Interest Balance ON  Date
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM BKCCIntTrans B " & _
    " WHERE B.TransId = (SELECT MAX(TransID) FROM " & _
        " BKCCIntTrans C WHERE TransDate <= #" & m_FromDate & "# " & _
        " AND C.LoanID = B.LoanID )"
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM BKCCIntTrans B WHERE " & _
    " TransDate <= #" & m_FromDate & "# ORDER BY LOanID,TransID Desc"
gDbTrans.SQLStmt = SqlPrin
Call gDbTrans.Fetch(rstPrinAsOn, adOpenDynamic)

DoEvents
RaiseEvent Processing("Writing the record", 0.45)
If gCancel Then Exit Function

gDbTrans.SQLStmt = SqlInt
Call gDbTrans.Fetch(rstIntAsOn, adOpenDynamic)

DoEvents
RaiseEvent Processing("Writing the record", 0.55)
If gCancel Then Exit Function

'GEt the Transacted amount After 31/3/yyyy till last month
SqlPrin = "SELECT SUM(AMOUNT) as SumAmount,LoanID,TransType FROM BKCCTrans WHERE " & _
    " TransDate > #" & Date31_3 & "# AND TransDate <= #" & DateLastMonth & "# " & _
    " GROUP BY LoanId,TransType"
SqlInt = "SELECT SUM(IntAmount) as SumIntAmount,SUM(PenalIntAmount) as SumPenalIntAmount," & _
    " LoanID,TransType FROM BKCCIntTrans WHERE " & _
    " TransDate > #" & Date31_3 & "# AND TransDate <= #" & DateLastMonth & "# " & _
    " GROUP BY LoanId,TransType"

gDbTrans.SQLStmt = SqlPrin
Call gDbTrans.Fetch(rstPrinTransLast, adOpenDynamic)

DoEvents
RaiseEvent Processing("Writing the record", 0.65)
If gCancel Then Exit Function

gDbTrans.SQLStmt = SqlInt
Call gDbTrans.Fetch(rstIntTransLast, adOpenDynamic)

DoEvents
RaiseEvent Processing("Writing the record", 0.75)
If gCancel Then Exit Function

'GEt the Transacted amount From last month to till Today
SqlPrin = "SELECT SUM(AMOUNT)as SumAmount,LoanID,TransType FROM BKCCTrans WHERE " & _
    " TransDate > #" & DateLastMonth & "# AND TransDate <= #" & m_FromDate & "# " & _
    " GROUP BY LoanId,TransType"
SqlInt = "SELECT SUM(IntAmount) as SumIntAmount,SUM(PenalIntAmount) as SumPenalIntAmount," & _
    " LoanID,TransType FROM BKCCIntTrans WHERE " & _
    " TransDate > #" & DateLastMonth & "# AND TransDate <= #" & m_FromDate & "# " & _
    " GROUP BY LoanId,TransType"

gDbTrans.SQLStmt = SqlPrin
Call gDbTrans.Fetch(rstPrinTransAsOn, adOpenDynamic)

DoEvents
RaiseEvent Processing("Writing the record", 0.85)
If gCancel Then Exit Function

gDbTrans.SQLStmt = SqlInt
Call gDbTrans.Fetch(rstIntTransAsOn, adOpenDynamic)

DoEvents
RaiseEvent Processing("Writing the record", 0.95)
If gCancel Then Exit Function

'Now Initialise the grid
grd.Clear
grd.Cols = 23
grd.Rows = 20

Dim SlNo As Integer
Dim Loanid As Long
Dim L_clsCust As New clsCustReg
Dim L_clsLoan As New clsBkcc
Dim RetStr As String
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

Call SetGrid(m_SchemeId, CStr(Year(Date31_3)))

Call InitGrid
'grd.MergeCells = flexMergeNever
RaiseEvent Initialise(0, rstMaster.RecordCount)

grd.Row = grd.FixedRows
lblReportTitle = "Meeting register As on " & m_FromIndianDate


'*********

If m_SchemeId = 0 Then GoTo WithoutSchemeID

''**********
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
    Loanid = FormatField(rstMaster("LoanID"))
    IntRate = FormatField(rstMaster("IntRate"))
    SlNo = SlNo + 1
  With grd
    If .Rows < .Row + 2 Then .Rows = .Rows + 2
    .Row = .Row + 1
    .MergeRow(.Row) = False
    .Col = 0: .Text = Format(SlNo, "00")
    .Col = 1: .Text = FormatField(rstMaster("AccNum"))
    .Col = 2: .Text = L_clsCust.CustomerName(FormatField(rstMaster("CustomerID")))
    RetStr = FormatField(rstMaster("Guarantor1"))
    On Error Resume Next
    If Val(RetStr) > 0 Then
        .Col = 3: .Text = L_clsCust.CustomerName(Val(RetStr))
'        .Row = .Row + 1: .Text = strArr(1): .Row = .Row - 1
    End If
    RetStr = FormatField(rstMaster("Guarantor2"))
    If Val(RetStr) > 0 Then
        .Col = 4: .Text = L_clsCust.CustomerName(Val(RetStr))
'        .Row = .Row + 1: .Text = StrArr(1): .Row = .Row - 1
    End If
    On Error GoTo ErrLine
    'Loan Advance Deatails
    .Col = 5: .Text = FormatField(rstMaster("IssueDate"))
    .Col = 6: .Text = FormatField(rstMaster("SanctionAmount"))
    
    'Out standing Loan Balance as on 31/3/yyyy
    PrevDate = Date31_3
    Balance31_3 = 0: IntBal31_3 = 0
    rstPrin31_3.MoveFirst
    rstPrin31_3.Find " LoanID = " & Loanid
    If Not rstPrin31_3.EOF Then
        If rstPrin31_3("LoanID") = Loanid Then
            rstInt31_3.MoveFirst
            rstInt31_3.Find " LoanID = " & Loanid
            Balance31_3 = rstPrin31_3("Balance")
            TransDate = rstInt31_3("TransDate")
            IntBal31_3 = FormatField(rstInt31_3("IntBalance"))
            'IntBal31_3 = IntBal31_3 + DateDiff("d", TransDate, Date31_3) / 365 * IntRate / 100 * Balance
            PrevDate = TransDate
        End If
    End If
    IntBal31_3 = IntBal31_3 + L_clsLoan.RegularInterest(Loanid, Date31_3)      'FormatField(rstInt31_3("IntBalance"))
    .Col = 9: .Text = Balance31_3
    
    'Over due as on 31/3/yyyy
    ODAmount = L_clsLoan.OverDueAmount(Loanid, Date31_3)
    ODInt = L_clsLoan.OverDueInterest(Loanid, Date31_3)
    '.Col = 10: .Text = L_clsLoan.DueInstallments(LoanId, FormatDate(CStr(Date31_3)))
    .Col = 11: .Text = FormatCurrency(ODAmount)
    .Col = 12: .Text = FormatCurrency(ODInt)
    .Col = 13: .Text = FormatCurrency(ODAmount + ODInt)
    
    'Loan Repayment From 1/4/yyyy to last month
    TransType = wDeposit
    Amount = 0: IntAmount = 0
    If Not rstPrinTransLast Is Nothing Then
        rstPrinTransLast.MoveFirst
        rstPrinTransLast.Find " LoanID = " & Loanid & " AND TransType = " & TransType
        If Not rstPrinLastMonth.EOF Then
            If rstPrinLastMonth("loanID") = Loanid Then
                rstIntTransLast.MoveFirst
                rstIntTransLast.Find " LoanID = " & Loanid & " AND TransType = " & TransType
                Amount = FormatField(rstPrinTransLast("SumAmount"))
                IntAmount = FormatField(rstIntTransLast("SumIntAmount"))
                IntAmount = IntAmount + FormatField(rstIntTransLast("SumPenalIntAmount"))
            End If
        End If
    End If
    .Col = 14: '.Text = IntBalLastMonth
    .Col = 15: .Text = FormatCurrency(Amount)
    .Col = 16: .Text = FormatCurrency(IntAmount)
    .Col = 17: .Text = FormatCurrency(Amount + IntAmount)    'Out standing Loan Balance as on last month
    
    'LOan Balance as on last month
    Amount = 0: IntAmount = 0
    If Not rstPrinLastMonth Is Nothing Then
        rstPrinLastMonth.MoveFirst
        rstPrinLastMonth.Find " LoanID = " & Loanid
        If Not rstPrinLastMonth.EOF Then
            If rstPrinLastMonth("loanID") = Loanid Then
                rstIntLastMonth.MoveFirst
                rstIntLastMonth.Find " LoanID = " & Loanid
                Amount = FormatField(rstPrinLastMonth("Balance"))
                IntAmount = FormatField(rstIntLastMonth("IntBalance"))
                .Col = 18: '.Text = IntBalLastMonth
            End If
        End If
    End If
    
    IntAmount = IntAmount + L_clsLoan.RegularInterest(Loanid, CDate(FormatDate(IndDateLastMonth)))
    .Col = 19: .Text = FormatCurrency(Amount)
    .Col = 20: .Text = FormatCurrency(IntAmount)
    .Col = 21: .Text = FormatCurrency(Amount + IntAmount)
    
    'Over due as on last month
    ODAmount = L_clsLoan.OverDueAmount(Loanid, CDate(FormatDate(IndDateLastMonth)))
    ODInt = L_clsLoan.PenalInterest(Loanid, CDate(FormatDate(IndDateLastMonth)))
    '.Col = 22: .Text = L_clsLoan.DueInstallments(LoanId, CDate(FormatDate(IndDateLastMonth)))
    .Col = 23: .Text = FormatCurrency(ODAmount)
    .Col = 24: .Text = FormatCurrency(ODInt)
    .Col = 25: .Text = FormatCurrency(ODAmount + ODInt)
    
    'If Case has filed Then Date Of Filing the case
    .Col = 26: .Text = FormatField(rstMaster("ABNDate"))
    .Col = 27: .Text = FormatField(rstMaster("EpDate"))
  End With
    DoEvents
    RaiseEvent Processing("Writing the record", _
        rstMaster.AbsolutePosition / (rstMaster.RecordCount * 2))
    rstMaster.MoveNext
    If gCancel Then Exit Function
Loop
Set L_clsLoan = Nothing

ShowMeetingRegistar = True
Screen.MousePointer = vbDefault
Debug.Print Now
Exit Function


'"THE MEETING REGISTER OF ALL LOANS

WithoutSchemeID:

RaiseEvent Initialise(0, rstMaster.RecordCount)

Do
    If rstMaster.EOF Then Exit Do
    Balance31_3 = 0: BalanceLastMonth = 0: BalanceNow = 0
    IntBal31_3 = 0: IntBalLastMonth = 0: IntBalNow = 0
    Loanid = FormatField(rstMaster("LoanID"))
    IntRate = FormatField(rstMaster("IntRate"))
    SlNo = SlNo + 1
  With grd
    If .Rows < .Row + 3 Then .Rows = .Rows + 3
    .Row = .Row + 1
    .MergeRow(.Row) = False
    .Col = 0: .Text = Format(SlNo, "00")
    .Col = 1: .Text = FormatField(rstMaster("AccNum"))
    .Col = 2: .Text = L_clsCust.CustomerName(FormatField(rstMaster("CustomerID")))
    RetStr = FormatField(rstMaster("Guarantor1"))
    On Error Resume Next
    If Val(RetStr) Then
        .Col = 3: .Text = L_clsCust.CustomerName(Val(RetStr))
        .Row = .Row + 1: .Text = "": .Row = .Row - 1
    End If
    RetStr = FormatField(rstMaster("Guarantor2"))
    If Val(RetStr) Then
        .Col = 4: .Text = L_clsCust.CustomerName(Val(RetStr))
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
        rstPrin31_3.MoveFirst
        rstPrin31_3.Find " LoanID = " & Loanid
        If Not rstPrin31_3.EOF Then
            If rstPrin31_3("LoanID") = Loanid Then
                'rstInt31_3.FindFirst " LoanID = " & LoanID
                Balance31_3 = FormatField(rstPrin31_3("Balance"))
                TransDate = FormatField(rstInt31_3("TransDate"))
                IntBal31_3 = FormatField(rstInt31_3("IntBalance"))
                PrevDate = TransDate
            End If
        End If
    End If
    IntBal31_3 = IntBal31_3 + L_clsLoan.RegularInterest(Loanid, Date31_3)
    IntBal31_3 = IntBal31_3 + L_clsLoan.PenalInterest(Loanid, Date31_3)
    .Col = 7: .Text = Balance31_3
    .Col = 8: .Text = IntBal31_3
    .Col = 9: .Text = Val(Balance31_3 + IntBal31_3)
            
    'Recovery from 1/4/yyyy to LastMonth
    PrinRepay = 0: IntRepay = 0: TransType = wDeposit
    If Not rstPrinTransLast Is Nothing Then
        rstPrinTransLast.MoveFirst
        rstPrinTransLast.Find "LoanID = " & Loanid & " AND Transtype = " & TransType
        If (Not rstPrinTransLast.EOF) And rstPrinTransLast("LoanID") = Loanid Then
            rstIntTransLast.MoveFirst
            rstIntTransLast.Find "LoanID = " & Loanid & " AND Transtype = " & TransType
            PrinRepay = FormatField(rstPrinTransLast("SumAmount"))
            IntRepay = FormatField(rstIntTransLast("SumIntAmount"))
        End If
    End If
    .Col = 10: .Text = FormatCurrency(PrinRepay)
    .Col = 11: .Text = FormatCurrency(IntRepay)
    .Col = 12: .Text = FormatCurrency(PrinRepay + IntRepay)
            
    'Loan Balance as on end of last month
    BalanceLastMonth = Balance31_3: IntBalLastMonth = 0
    If Not rstPrinLastMonth Is Nothing Then
        rstPrinLastMonth.MoveFirst
        rstPrinLastMonth.Find " LoanID = " & Loanid
        If Not rstPrinLastMonth.EOF Then
            If rstPrinLastMonth("LoanID") = Loanid Then
                rstIntLastMonth.MoveFirst
                rstIntLastMonth.Find " LoanID = " & Loanid
                BalanceLastMonth = rstPrinLastMonth("Balance")
                TransDate = rstPrinLastMonth("TransDate")
                IntBalLastMonth = FormatField(rstIntLastMonth("IntBalance"))
                PrevDate = TransDate
                PrevDate = TransDate
            End If
        End If
    End If
    IntBalLastMonth = IntBalLastMonth + L_clsLoan.RegularInterest(Loanid, CDate(FormatDate(IndDateLastMonth)))
    .Col = 13: .Text = BalanceLastMonth
    .Col = 14: .Text = IntBalLastMonth
    .Col = 15: .Text = Val(BalanceLastMonth + IntBalLastMonth)
    
    .Col = 10: .Text = Balance31_3 - BalanceLastMonth
    If Val(.Text) < 0 Then .Text = "0.00"
    
    'Recovery during this month
    PrinRepay = 0: IntRepay = 0: TransType = wDeposit
    If Not rstPrinTransAsOn Is Nothing Then
        rstPrinTransAsOn.MoveFirst
        rstPrinTransAsOn.Find "LoanID = " & Loanid & " AND Transtype = " & TransType
        If Not rstPrinTransAsOn.EOF And rstPrinTransAsOn("LoanID") = Loanid Then
            rstIntTransAsOn.MoveFirst
            rstIntTransAsOn.Find "LoanID = " & Loanid & " AND Transtype = " & TransType
            PrinRepay = FormatField(rstPrinTransAsOn("SumAmount"))
            IntRepay = FormatField(rstIntTransAsOn("SumIntAmount"))
        End If
    End If
    .Col = 16: .Text = FormatCurrency(PrinRepay)
    .Col = 17: .Text = FormatCurrency(IntRepay)
    .Col = 18: .Text = FormatCurrency(PrinRepay + IntRepay)
    
    'Balance as of now
    BalanceNow = BalanceLastMonth: IntBalNow = 0
    If Not rstPrinAsOn Is Nothing Then
        rstPrinAsOn.MoveFirst
        rstPrinAsOn.Find " LoanID = " & Loanid
        If Not rstPrinAsOn.EOF Then
            If rstPrinAsOn("LoanID") = Loanid Then
                rstIntAsOn.MoveFirst
                rstIntAsOn.Find " LoanID = " & Loanid
                BalanceNow = rstPrinAsOn("Balance")
                TransDate = rstPrinAsOn("TransDate")
                IntBalNow = FormatField(rstIntAsOn("IntBalance"))
                PrevDate = TransDate
            End If
        End If
    End If
    IntBalNow = IntBalNow + L_clsLoan.RegularInterest(Loanid, m_FromDate)
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
    rstMaster.MoveNext
    If gCancel Then Exit Function
Loop
Set L_clsLoan = Nothing
ShowMeetingRegistar = True
Screen.MousePointer = vbDefault
ErrLine:
    Screen.MousePointer = vbDefault
    If Err Then
        MsgBox Err.Number & vbCrLf & Err.Description, , wis_MESSAGE_TITLE
       Resume
        Exit Function
    End If


End Function

Private Sub SetGrid(SchemeId As Integer, Optional strYear As String = "YYYY")
Dim Count As Integer
Dim StrText As String
Dim Rst As Recordset

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
If SchemeId = 0 Then GoTo CommonSetting

Dim LoanType As wis_LoanType

    gDbTrans.SQLStmt = "SELECT * FROM LoanScheme Where SchemeID = " & m_SchemeId
    Call gDbTrans.Fetch(Rst, adOpenDynamic)
    'SchemeName = FormatField(gDbTrans.Rst("SchemeName"))
    'SchemeStr = " SchemeID = " & m_SchemeId & " "
    LoanType = FormatField(Rst("LoanType"))


With grd
    .Cols = 28
    .Rows = 10
    .FixedCols = 2
    .FixedRows = 4
    .Row = 0
    .Col = 0: .Text = "Sl No"
    .Col = 1: .Text = "Loan No"
    .Col = 2: .Text = "Name of Customer"
    .Col = 3: .Text = "Name OF Sureties with full address"
    .Col = 4: .Text = "Name OF Sureties with full address"
    .Col = 5: .Text = "Loan Disbursment particulars"
    .Col = 6: .Text = "Loan Disbursment particulars"
    .Col = 7: .Text = "Vehicle Detail"
    .Col = 8: .Text = "Vehicle Detail"
    .Col = 9: .Text = "Loans out standing as on 31/3/" & strYear
    .Col = 10: .Text = "Loans out standing as on 31/3/" & strYear
    .Col = 11: .Text = "Loans out standing as on 31/3/" & strYear
    .Col = 12: .Text = "Loans out standing as on 31/3/" & strYear
    .Col = 13: .Text = "Loans out standing as on 31/3/" & strYear
    .Col = 14: .Text = "'"
    .Col = 15: .Text = "'"
    .Col = 16: .Text = "'"
    .Col = 17: .Text = "'"
    .Col = 18: .Text = "Loans out standing as on end of last month"
    .Col = 19: .Text = "Loans out standing as on end of last month"
    .Col = 20: .Text = "Loans out standing as on end of last month"
    .Col = 21: .Text = "Loans out standing as on end of last month"
    .Col = 22: .Text = "Loans out standing as on end of last month"
    .Col = 23: .Text = "Loans out standing as on end of last month"
    .Col = 24: .Text = "Loans out standing as on end of last month"
    .Col = 25: .Text = "Loans out standing as on end of last month"
    .Col = 26: .Text = "."
    .Col = 27: .Text = "."
    
    ''2nd Row
    For Count = 0 To .Cols - 1
        .Col = Count
        .Row = 0: StrText = .Text
        .Row = 1: .Text = StrText
    Next
    .Row = 1
    .Col = 3: .Text = "Suruty No. 1"
    .Col = 4: .Text = "Suruty No. 2"
    .Col = 5: .Text = "Date of Loan advance"
    .Col = 6: .Text = "Amount disubursed"
    .Col = 7: .Text = "Registration No (RTO)"
    .Col = 8: .Text = "Insurence Renewed up to"
    .Col = 9: .Text = "Total loan out standing"
    .Col = 10: .Text = "Of which over dues"
    .Col = 11: .Text = "Of which over dues"
    .Col = 12: .Text = "Of which over dues"
    .Col = 13: .Text = "Of which over dues"
    .Col = 14: .Text = "Repayments From 1/4/" & strYear & " to upto end of last month"
    .Col = 15: .Text = "Repayments From 1/4/" & strYear & " to upto end of last month"
    .Col = 16: .Text = "Repayments From 1/4/" & strYear & " to upto end of last month"
    .Col = 17: .Text = "Repayments From 1/4/" & strYear & " to upto end of last month"
    .Col = 18: .Text = "Of which over due"
    .Col = 19: .Text = "Of which over due"
    .Col = 20: .Text = "Of which over due"
    .Col = 21: .Text = "Of which over due"
    .Col = 22: .Text = "Dates of filing"
    .Col = 23: .Text = "Dates of filing"

    
    ''2nd  Row
    For Count = 0 To .Cols - 1
        .Col = Count
        .Row = 1: StrText = .Text
        .Row = 2: .Text = StrText
    Next
    .Row = 2
    .Col = 10: .Text = "No of Inst"
    .Col = 11: .Text = "Principal"
    .Col = 12: .Text = "Interest"
    .Col = 13: .Text = "Total"
    .Col = 22: .Text = "No of Inst"
    .Col = 23: .Text = "Principal"
    .Col = 24: .Text = "Interest"
    .Col = 25: .Text = "Total"
    
    .Col = 14: .Text = "No of Inst"
    .Col = 15: .Text = "Principal"
    .Col = 16: .Text = "Interest"
    .Col = 17: .Text = "Total"
    .Col = 18: .Text = "No of Inst"
    .Col = 19: .Text = "Principal"
    .Col = 20: .Text = "Interest"
    .Col = 21: .Text = "Total"
    .Col = 26: .Text = "ABN"
    .Col = 27: .Text = "Ep"

    ''3rd row
    .Row = 3
    Dim K As Integer
    K = 1
    For Count = 1 To 27
        .Col = Count: .Text = K
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
        For Count = 0 To grd.Cols - 1
            .MergeCol(Count) = True
            .Col = Count
            .CellAlignment = 4: .CellFontBold = True
        Next
    Next
    .Row = 1
    For Count = 19 To 22
        .Col = Count
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
    .Col = 0: .Text = "Sl No"
    .Col = 1: .Text = "Loan No"
    .Col = 2: .Text = "Name of Customer"
    .Col = 3: .Text = "Guarantor Name & Address"
    .Col = 4: .Text = "Guarantor Name & Address"
    .Col = 5: .Text = "Issue Date"
    .Col = 6: .Text = "Loan Amount"

    .Col = 7: .Text = "Out standing as 31/3/" & strYear & ""
    .Col = 8: .Text = "Out standing as 31/3/" & strYear & ""
    .Col = 9: .Text = "Out standing as 31/3/" & strYear & ""
    
    '.Col = 6: .Text = "Disbursuments from 1/4/" & strYear & " to upto end of last month"
    .Col = 10: .Text = "Repayments from 1/4/" & strYear & " to up to the end of last month"
    .Col = 11: .Text = "Repayments from 1/4/" & strYear & " to up to the end of last month"
    .Col = 12: .Text = "Repayments from 1/4/" & strYear & " to up to the end of last month"
    .Col = 13: .Text = "Balance OutStanding as on the end of last month"
    .Col = 14: .Text = "Balance OutStanding as on the end of last month"
    .Col = 15: .Text = "Balance OutStanding as on the end of last month"
    
    .Col = 16: .Text = "Repayments during this month"
    .Col = 17: .Text = "Repayments during this month"
    .Col = 18: .Text = "Repayments during this month"
    .Col = 19: .Text = "Balance OutStanding end of this month"
    .Col = 20: .Text = "Balance OutStanding end of this month"
    .Col = 21: .Text = "Balance OutStanding end of this month"
    
    ''2nd row
    .Row = 1: .MergeRow(3) = True
    .Col = 0: .Text = "Sl No"
    .Col = 1: .Text = "Loan No"
    .Col = 2: .Text = "Name of Customer"
    .Col = 3:  .Text = "Guarantor 1"
    .Col = 4:  .Text = "Guarantor 2"
    .Col = 5:  .Text = "Issue Date"
    .Col = 6: .Text = "Loan Amount"

    '.Col = 6: .Text = "Disbursuments from 1/4/" & strYear & " to upto end of last month"
    .Col = 7: .Text = "Principal"
    .Col = 8: .Text = "Interest"
    .Col = 9: .Text = "Total"
    
    .Col = 10: .Text = "Principal"
    .Col = 11: .Text = "Interest"
    .Col = 12: .Text = "Total"
    .Col = 13: .Text = "Principal"
    .Col = 14: .Text = "Interest"
    .Col = 15: .Text = "Total"
    
    .Col = 16: .Text = "Principal"
    .Col = 17: .Text = "Interest"
    .Col = 18: .Text = "Total"
    .Col = 19: .Text = "Principal"
    .Col = 20: .Text = "Interest"
    .Col = 21: .Text = "Total"
    
    
    .Row = 2: .MergeRow(4) = True
    .MergeCells = flexMergeFree
    For Count = 3 To .Cols - 1
        .Col = Count: .Text = (Count)
        .CellAlignment = 4
        .MergeCol(Count) = True
    Next
    '.Col = 3: .Text = "2b"
    .Col = 2: .Text = "2a"
    .Col = 1: .Text = "2"
    .Col = 0: .Text = "1"
    
    Dim i As Integer, j As Integer
    For i = 0 To .FixedRows - 1
        .Row = i
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
On Error GoTo Exitline:


Dim SqlStr As String
Dim TransType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes

Dim rstOpBalance As Recordset
Dim rstClBalance As Recordset
Dim rstAdvance As Recordset
Dim rstRecovery As Recordset

Dim FirstDate As Date

Dim ColAmount() As Currency
Dim GrandTotal() As Currency


'If AsOnIndianDate = "" Then
'    m_fromDate = gStrDate
'Else
'    m_fromDate = FormatDate(AsOnIndianDate)
'End If
'Get the First day of the Month
FirstDate = Month(m_FromDate) & "/1/" & Year(m_FromDate)
'Fetch Only Cash Credit Loans
Dim LoanType As wis_LoanType
LoanType = wisCashCreditLoan
RaiseEvent Initialise(0, 10)

'Get The LoanDetails And THier balance as on Date
SqlStr = "SELECT A.LoanID,AccNum,IssueDate,RenewDate,SanctionAmount,Balance," & _
    " Title+' '+ FirstName+' '+ MiddleName+' '+LastName As Name" & _
    " FROM BKCCCMaster A,BKCCTrans B, NameTab C " & _
    " WHERE AND B.LoanID = A.LoanID" & _
    " AND C.CustomerId =A.CustomerID AND TransID = (SELECT MAX(TransId) FROM " & _
        " BKCCTrans D WHERE D.TransDate <= #" & m_FromDate & "# " & _
        " AND D.LoanId = A.LoanId ) " & _
    " AND (ClosedDate is NULL OR LoanClosed = False ) "

gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(rstClBalance, adOpenDynamic) < 1 Then gCancel = True: Exit Function
'Set rstClBalance = gDBTrans.Rst.Clone
DoEvents
RaiseEvent Processing("Fetching the record", 0.25)
If gCancel Then Exit Function

'Get The LoanDetails And Thier balance as on first day of the given month
SqlStr = "SELECT LoanID,Balance,TransDate FROM BKCCTrans WHERE " & _
    " TransDate < #" & FirstDate & "# ORDER BY LoanID,TransId Desc"
gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(rstOpBalance, adOpenDynamic) < 1 Then Exit Function
'Set rstOpBalance = gDBTrans.Rst.Clone

DoEvents
RaiseEvent Processing("Fetching the record", 0.5)
If gCancel Then Exit Function

'Get The Advances During the MOnth
TransType = wWithdraw
ContraTrans = wContraWithdraw
SqlStr = "SELECT SUM(Amount),LoanID FROM BKCCTrans WHERE TransDate >= #" & FirstDate & "#" & _
    " AND TransDate <= #" & m_FromDate & "# AND (TransType = " & TransType & _
    " OR TransType = " & ContraTrans & ")" & _
    " GROUP BY LoanID "
gDbTrans.SQLStmt = SqlStr

Call gDbTrans.Fetch(rstAdvance, adOpenDynamic)

DoEvents
RaiseEvent Processing("Fetching the record", 0.65)
If gCancel Then Exit Function


'Get The Recovery During the MOnth
TransType = wDeposit
ContraTrans = wContraDeposit
SqlStr = "SELECT SUM(Amount),LoanID FROM BKCCTrans WHERE TransDate >= #" & FirstDate & "#" & _
    " AND TransDate <= #" & m_FromDate & "# AND (TransType = " & TransType & _
    " OR TransType = " & ContraTrans & ")" & _
    " GROUP BY LoanID "
gDbTrans.SQLStmt = SqlStr

Call gDbTrans.Fetch(rstRecovery, adOpenDynamic)

DoEvents
RaiseEvent Processing("Fetching the record", 0.85)
If gCancel Then Exit Function

'Now Align the grid
Call Shed5RowCol
ReDim ColAmount(8 To grd.Cols - 1)
ReDim GrandTotal(8 To grd.Cols - 1)
Dim Loanid As Long
Dim AddRow As Boolean
Dim L_clsLoan As New clsBkcc
Dim PrevOD As Currency
Dim ODAmount As Currency
Dim Count As Long
Dim SlNo As Long
RaiseEvent Initialise(0, rstClBalance.RecordCount)

'Now Start to writing to the grid
SlNo = 0
While Not rstClBalance.EOF
   Loanid = FormatField(rstClBalance("LoanId"))
    rstOpBalance.MoveFirst
    rstOpBalance.Find "LoanID = " & Loanid
    ColAmount(8) = 0
    If Not rstOpBalance.EOF Then
        ColAmount(8) = FormatField(rstOpBalance("Balance")) 'Balance as on 31/3/yyyy
    End If
    ColAmount(9) = 0
    If Not rstAdvance Is Nothing Then
        rstAdvance.MoveFirst
        rstAdvance.Find "LoanId = " & Loanid
        If Not rstAdvance.EOF Then
            ColAmount(9) = FormatField(rstAdvance(0))  'Advances During the MOnth
        End If
    End If
    ColAmount(10) = 0
    If Not rstRecovery Is Nothing Then
        rstRecovery.MoveFirst
        rstRecovery.Find "LoanId = " & Loanid
        If Not rstRecovery.EOF Then
            ColAmount(10) = FormatField(rstRecovery(0)) 'Recovery During the mOnth
        End If
    End If
    ColAmount(11) = FormatField(rstClBalance("Balance")) 'Balance at the end of month
    
    ColAmount(12) = ColAmount(8) + ColAmount(9)  'Maximum O/s Balance
    
    
    ODAmount = L_clsLoan.OverDueAmount(Loanid, m_FromDate)  'Over due
    ColAmount(13) = ODAmount 'Over due amount of the loan as on given date
    
    'Over due amount classification
    PrevOD = 0
'    ODAmount = L_clsLoan.OverDueSince(5, LoanID, , AsOnIndianDate)
'    ColAmount(18) = ODAmount - PrevOD 'Over due since & above 5 Years
'    PrevOD = ColAmount(18) + PrevOD
    ColAmount(19) = L_clsLoan.OverDueSince(5, Loanid, m_FromDate)
    If ColAmount(19) < 0 Then ColAmount(19) = ODAmount  'Over due since & above 5 Years
    ODAmount = ODAmount - ColAmount(19)
    
    ColAmount(18) = L_clsLoan.OverDueSince(4, Loanid, m_FromDate) - ColAmount(19)
    If ColAmount(18) < 0 Then ColAmount(18) = ODAmount  'Over due since 5 Years
    ODAmount = ODAmount - ColAmount(18)
    
    ColAmount(17) = L_clsLoan.OverDueSince(3, Loanid, m_FromDate) - ColAmount(18)
    If ColAmount(17) < 0 Then ColAmount(17) = ODAmount  'Over due since 3 Years
    ODAmount = ODAmount - ColAmount(17)
    
    ColAmount(16) = L_clsLoan.OverDueSince(2, Loanid, m_FromDate) - ColAmount(17)
    If ColAmount(16) < 0 Then ColAmount(16) = ODAmount  'Over due since 2 Years
    ODAmount = ODAmount - ColAmount(16)
    
    ColAmount(15) = L_clsLoan.OverDueSince(1, Loanid, m_FromDate) - ColAmount(16)
    If ColAmount(15) < 0 Then ColAmount(15) = ODAmount  'Over due since 1 Years
    ODAmount = ODAmount - ColAmount(15)
    
    ColAmount(14) = ODAmount   'Over due Under 1 Years
    
    'Check whether this row has to be write or not
    AddRow = False
    For Count = 8 To grd.Cols - 1
        If ColAmount(Count) Then
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
        '.Col = 3: .Text = FormatField(rstClBalance("CustBankName"))
        .Col = 4: .Text = FormatField(rstClBalance("LoanAmount"))
        .Col = 5: .Text = FormatField(rstClBalance("IssueDate"))
        .Col = 6: .Text = FormatField(rstClBalance("LoanDueDate"))
        .Col = 7: .Text = FormatField(rstClBalance("LoanPurpose"))
        For Count = 8 To grd.Cols - 1
            .Col = Count
            If ColAmount(Count) < 0 Then ColAmount(Count) = 0
            .Text = FormatCurrency(ColAmount(Count))
            GrandTotal(Count) = GrandTotal(Count) + ColAmount(Count)
        Next
      End With
    End If
    DoEvents
    If gCancel Then Exit Function
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
    .Text = "Grand Total": .CellFontBold = True
    For Count = 8 To .Cols - 1
        .Col = Count: .CellFontBold = True
        .Text = FormatCurrency(GrandTotal(Count))
    Next
End With
ShowShed5 = True

lblReportTitle.Caption = "Statement showing the Cash Credit loans for the month of " & _
    GetMonthString(Month(m_FromDate)) & " as on " & m_FromIndianDate

Exitline:
    If Err Then
        MsgBox "ERROR SHED 5" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
        'Resume
        Err.Clear
    End If

End Function

Private Function ShowShed1() As Boolean

RaiseEvent Processing("Fetching the records", 0)


On Error GoTo ErrLine
'Declarations
Dim rstLoan As Recordset
Dim RstRepay  As Recordset
Dim rstAdvance As Recordset

Dim SqlStr As String
Dim Count As Integer

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
Dim FromDate As Date
Dim LastDate As Date
Dim OBDate As Date

Dim LExcel As Boolean ' to be removed later and get the data from outside. - pradeep

'LExcel = True
If LExcel Then
    'Set xlWorkBook = Workbooks.Add
    Set xlWorkSheet = xlWorkBook.Sheets(1)
End If

ShowShed1 = False

'm_fromDate = FormatDate(AsOnIndianDate)
Dim LoanCategary As wisLoanCategories
LoanCategary = wisAgriculural

If LoanCategary = wisAgriculural Then
    OBDate = "7/1/" & IIf(Month(m_FromDate) > 6, Year(m_FromDate), Year(m_FromDate) - 1)
Else
    OBDate = "4/1/" & IIf(Month(m_FromDate) > 3, Year(m_FromDate), Year(m_FromDate) - 1)
End If
'This Report Includes Only agricultural loans so
OBDate = "7/1/" & IIf(Month(m_FromDate) > 6, Year(m_FromDate), Year(m_FromDate) - 1)


'Get The Details Of Agri Loans, and Balance as ondate

SqlStr = "SELECT A.LoanId, AccNum, Balance,  " & _
    " Title + ' ' + FirstName +' ' + MiddleName + ' ' + LastName As Name " & _
    " From BKCCMaster A,BKCCTrans B,NameTab C WHERE A.LoanID = B.LoanID " & _
    " AND C.CustomerID = A.CustomerId AND TransID = (SELECT MAX(TransID) " & _
        " From BKCCTrans D Where D.LoanId = A.LoanId" & _
        " AND D.TransDate <= #" & m_FromDate & "#)"

    
gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(rstLoan, adOpenDynamic) < 1 Then gCancel = True: Exit Function
'Set rstLoan = gDBTrans.Rst.Clone
RaiseEvent Initialise(0, rstLoan.RecordCount * 2)

'Now fix the headers for the shedule
Call Shed1RowCOl

'Now Get the Loan RepayMent Of the during this period
Dim TransType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes
TransType = wDeposit
ContraTrans = wContraDeposit
SqlStr = "SELECT SUM(Amount),LoanID FROM BKCCTrans Where (TransType = " & TransType & _
    " OR TransType = " & ContraTrans & ") AND TransDate >= #" & OBDate & "# " & _
    " AND TransDate <= #" & m_FromDate & "# " & _
    " GROUP BY LoanID"
gDbTrans.SQLStmt = SqlStr
Call gDbTrans.Fetch(RstRepay, adOpenDynamic)

'Now Get the Loan Advance Of the during this period
TransType = wWithdraw
ContraTrans = wContraWithdraw
SqlStr = "SELECT SUM(Amount),LoanID FROM BKCCTrans Where TransType = " & TransType & _
    " OR TransType = " & ContraTrans & " AND TransDate >= #" & OBDate & "# " & _
    " AND TransDate <= #" & m_FromDate & "#" & _
    " GROUP BY LoanID"
gDbTrans.SQLStmt = SqlStr
Call gDbTrans.Fetch(rstAdvance, adOpenDynamic)


ReDim ColAmount(3 To grd.Cols - 1)
ReDim SubTotal(3 To grd.Cols - 1)
ReDim GrandTotal(3 To grd.Cols - 1)

Dim Loanid As Long
Dim L_clsLoan As New clsBkcc
Dim OBIndiandate As String
Dim curRepay As Currency
Dim ODAmount As Currency
Dim PrevOD As Currency
Dim AddRow As Boolean

OBIndiandate = FormatDate(CStr(OBDate))

grd.Row = grd.FixedRows
SlNo = 0
RstRepay.MoveLast
While Not rstLoan.EOF
    'Initialise the Varible
    Loanid = FormatField(rstLoan("LoanID"))
    ColAmount(3) = L_clsLoan.OverDueAmount(Loanid, OBDate)   'Over due as on Opeing date
    ColAmount(4) = L_clsLoan.LoanDemand(Loanid, OBIndiandate, m_FromIndianDate)    'Amount which falls as over due between OB Date & today
    ColAmount(5) = L_clsLoan.OverDueAmount(Loanid, m_FromDate)   'Over due as on today
    
    ColAmount(5) = ColAmount(3) + ColAmount(4)
    
    'The difference between two date is the Loan demand of that period
    'Therefore
    'ColAmount(4) = ColAmount(5) - ColAmount(3)
    
    curRepay = 0: ' curAdvance = 0
    If Not RstRepay Is Nothing Then
        RstRepay.MoveFirst
        RstRepay.Find "LoanID = " & Loanid
        If Not RstRepay.EOF Then
            curRepay = FormatField(RstRepay(0))
        End If
    End If
    Debug.Assert ColAmount(5) - curRepay = L_clsLoan.OverDueAmount(Loanid, m_FromDate)
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
    ODAmount = L_clsLoan.OverDueAmount(Loanid, m_FromDate)
    PrevOD = 0
    ColAmount(16) = L_clsLoan.OverDueSince(5, Loanid, m_FromDate)
    If ColAmount(16) > ODAmount Then ColAmount(16) = ODAmount 'Over due since 5 & above 5 Years
    ODAmount = ODAmount - ColAmount(16)
    
    ColAmount(15) = L_clsLoan.OverDueSince(4, Loanid, m_FromDate) - ColAmount(16)
    If ColAmount(15) > ODAmount Then ColAmount(15) = ODAmount 'Over due since 4 Years
    ODAmount = ODAmount - ColAmount(15)
    
    ColAmount(14) = L_clsLoan.OverDueSince(3, Loanid, m_FromDate) - ColAmount(15)  'Over due since 3 Years
    If ColAmount(14) > ODAmount Then ColAmount(14) = ODAmount 'Over due since 3 Years
    ODAmount = ODAmount - ColAmount(14)
    
    ColAmount(13) = L_clsLoan.OverDueSince(2, Loanid, m_FromDate) - ColAmount(14)  'Over due since 2 Years
    If ColAmount(13) > ODAmount Then ColAmount(13) = ODAmount 'Over due since 2 Years
    ODAmount = ODAmount - ColAmount(13)
    
    ColAmount(12) = L_clsLoan.OverDueSince(1, Loanid, m_FromDate) - ColAmount(13)
    If ColAmount(12) > ODAmount Then ColAmount(12) = ODAmount 'Over due since 1 Year
    ODAmount = ODAmount - ColAmount(12)
    
    ColAmount(11) = ODAmount 'Over due under one year
    AddRow = False
    For Count = 3 To grd.Cols - 1
        If ColAmount(Count) Then
            AddRow = True
            Exit For
        End If
    Next
    If AddRow Then
        If grd.Rows <= grd.Row + 2 Then grd.Rows = grd.Rows + 1
        grd.Row = grd.Row + 1
        SlNo = SlNo + 1
        grd.Col = 0: grd.Text = SlNo
        grd.Col = 1: grd.Text = FormatField(rstLoan("AccNum"))
        grd.Col = 2: grd.Text = FormatField(rstLoan("Name"))
        For Count = 3 To grd.Cols - 1
            If ColAmount(Count) < 0 Then ColAmount(Count) = 0
            grd.Col = Count: grd.Text = FormatCurrency(ColAmount(Count))
            GrandTotal(Count) = GrandTotal(Count) + ColAmount(Count)
        Next
    End If
    DoEvents
    RaiseEvent Processing("Writing the record", (rstLoan.AbsolutePosition / rstLoan.RecordCount))
    rstLoan.MoveNext
Wend

AddRow = False
For Count = 3 To grd.Cols - 1
    If GrandTotal(Count) Then
        AddRow = True
        Exit For
    End If
Next
If AddRow Then
    If grd.Rows <= grd.Row + 2 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
    If grd.Rows <= grd.Row + 2 Then grd.Rows = grd.Rows + 1
    grd.Row = grd.Row + 1
    grd.Col = 2: grd.Text = "Grand Total": grd.CellFontBold = True
    For Count = 3 To grd.Cols - 1
        If GrandTotal(Count) < 0 Then ColAmount(Count) = 0
        grd.Col = Count: grd.CellFontBold = True
        grd.Text = FormatCurrency(GrandTotal(Count))
    Next
End If



Set L_clsLoan = Nothing
lblReportTitle = "Demand, collecion and blance register of fo the month of " & _
    GetMonthString(Month(m_FromDate)) & " as on " & m_FromIndianDate
ShowShed1 = True
Exit Function

ErrLine:
    MsgBox "error Showshed1" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
  'Resume
End Function

Private Function ShowShed2() As Boolean

RaiseEvent Processing("Fetching the record", 0)
m_repType = repShedule_2

On Error GoTo Exitline

'Declarations

Dim rstCurrBalance As Recordset
Dim rstOpBalance As Recordset
Dim rstRepayTill  As Recordset
Dim rstAdvanceTill As Recordset
Dim rstRepayLast  As Recordset
Dim rstAdvanceLast As Recordset

Dim SqlStr As String
Dim Count As Integer

Dim SlNo As Integer
Dim strBranchName As String


'Decalration By SHashi
Dim ColAmount() As Currency
Dim GrandTotal() As Currency
Dim FromDate As Date
Dim LastDate As Date
Dim m_FromDate As Date
Dim OBDate As Date
Dim LoanCategary As wisLoanCategories

Dim LExcel As Boolean ' to be removed later and get the data from outside. - pradeep

'LExcel = True
If LExcel Then
    'Set xlWorkBook = Workbooks.Add
    Set xlWorkSheet = xlWorkBook.Sheets(1)
End If

ShowShed2 = False
''Now Get the Last Friday date
'If AsOnIndianDate = "" Then
'    AsOnIndianDate = FormatDate(gStrDate)
'End If

m_FromDate = FormatDate(m_FromIndianDate)
LoanCategary = wisAgriculural
LastDate = Month(m_FromDate) & "/1/" & Year(m_FromDate)
LastDate = DateAdd("d", -1, LastDate)

'GetLast Friday date
'm_fromDate = DateAdd("M", 1, m_fromDate)
'm_fromDate = Month(m_fromDate) & "/1/" & Year(m_fromDate)
''Get the End Of LastMonth
'LastDate = DateAdd("m", -2, m_fromDate)
'Do
'    m_fromDate = DateAdd("d", -1, m_fromDate)
'    If Format(m_fromDate, "dddd") = "Friday" Then Exit Do
'Loop

If LoanCategary = wisAgriculural Then
    OBDate = "7/1/" & IIf(Month(m_FromDate) > 6, Year(m_FromDate), Year(m_FromDate) - 1)
Else
    OBDate = "4/1/" & IIf(Month(m_FromDate) > 3, Year(m_FromDate), Year(m_FromDate) - 1)
End If


'Me.lblBranchName = strBranchName
RaiseEvent Initialise(0, 15)

'Get The Details Of Loans, and Balance as ondate
    RaiseEvent Processing("Fetching the record", 0.15)
SqlStr = "SELECT A.LoanId, AccNum, Balance, " & _
    " Title + ' ' + FirstName +' ' + MiddleName + ' ' + LastName As Name " & _
    " From BKCCMaster A,BKCCTrans B,NameTab C WHERE A.LoanID = B.LoanID" & _
    " AND C.CustomerID = A.CustomerId AND TransID = (SELECT MAX(TransID) " & _
        " From BKCCTrans D Where D.LoanId = A.LoanId" & _
        " AND D.TransDate <= #" & m_FromDate & "#)" & _
    " ORDER BY A.LoanID"

gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(rstCurrBalance, adOpenDynamic) < 1 Then Exit Function
'Set rstCurrBalance = gDBTrans.Rst.Clone
'RaiseEvent Initialise(0, rstCurrBalance.RecordCount )

RaiseEvent Processing("Fetching the record", 0.55)
'Get The Details Of Loans, and Balance as on 31/M/yyyy
SqlStr = "SELECT A.LoanId, AccNum, Balance,TransDate" & _
    " From BKCCMaster A,BKCCTrans B WHERE A.LoanID = B.LoanID" & _
    " AND A.LoanID = B.LoanID AND TransID = " & _
            " (SELECT MAX(TransID) From BKCCTrans D Where " & _
            " D.LoanId = A.LoanId AND D.TransDate < #" & OBDate & "#)"
gDbTrans.SQLStmt = SqlStr
If gDbTrans.Fetch(rstOpBalance, adOpenDynamic) < 1 Then Exit Function

Dim TransType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes
'Now Get the Loan Repayment From Ob Date to Last Month
TransType = wDeposit
ContraTrans = wContraDeposit
RaiseEvent Processing("Fetching the record", 0.65)
SqlStr = "SELECT SUM(Amount),LoanID FROM BKCCTrans Where " & _
    " TransDate >= #" & OBDate & "# AND TransDate <= #" & LastDate & "#" & _
    " AND (TransType = " & TransType & " OR TransType = " & ContraTrans & ")" & _
    " GROUP BY LoanID"
gDbTrans.SQLStmt = SqlStr
Call gDbTrans.Fetch(rstRepayLast, adOpenDynamic)

'Now Get the Loan Repayment From Lastdate to till date
RaiseEvent Processing("Fetching the record", 0.75)
SqlStr = "SELECT SUM(Amount),LoanID FROM BKCCTrans Where TransDate > #" & LastDate & "# " & _
    " AND TransDate <= #" & m_FromDate & "# AND (TransType = " & TransType & _
    " OR TransType = " & ContraTrans & ") " & _
    " GROUP BY LoanID"
gDbTrans.SQLStmt = SqlStr
Call gDbTrans.Fetch(rstRepayTill, adOpenDynamic)

'Now Get the Loan Advance From Ob date to last month
TransType = wWithdraw
ContraTrans = wContraWithdraw
RaiseEvent Processing("Fetching the record", 0.85)
SqlStr = "SELECT SUM(Amount),LoanID FROM BKCCTrans Where " & _
    " TransDate >= #" & OBDate & "# AND TransDate <= #" & LastDate & "#" & _
    " AND (TransType = " & TransType & " OR TransType = " & ContraTrans & ")" & _
    " GROUP BY LoanID"
gDbTrans.SQLStmt = SqlStr
Call gDbTrans.Fetch(rstAdvanceLast, adOpenDynamic)

'Now Get the Loan Advance during this month
RaiseEvent Processing("Fetching the record", 0.95)
SqlStr = "SELECT SUM(Amount),LoanID FROM BKCCTrans Where " & _
    " TransDate > #" & LastDate & "# AND TransDate <= #" & m_FromDate & "#" & _
    " AND (TransType = " & TransType & " OR TransType = " & ContraTrans & ")" & _
    " GROUP BY LoanID"
gDbTrans.SQLStmt = SqlStr
Call gDbTrans.Fetch(rstAdvanceTill, adOpenDynamic)

SlNo = 0
grd.Visible = True

' now set the shed headers
Call Shed2RowCol

grd.Row = 2
' beginning of the loantrans loop
ReDim ColAmount(4 To grd.Cols - 1)
ReDim SubTotal(4 To grd.Cols - 1)
ReDim GrandTotal(4 To grd.Cols - 1)

Dim Loanid As Long
Dim AddRow As Boolean
Dim L_clsLoan As New clsBkcc

RaiseEvent Initialise(0, rstCurrBalance.RecordCount)

While Not rstCurrBalance.EOF
    Loanid = FormatField(rstCurrBalance("LoanID"))
    'Get the opening balance
    ColAmount(4) = 0
    If Not rstOpBalance Is Nothing Then
        rstOpBalance.MoveFirst
        rstOpBalance.Find "LoanID = " & Loanid
        If Not rstOpBalance.EOF Then
            ColAmount(4) = FormatField(rstOpBalance("Balance")) ' Opening Balance
        End If
    End If
    
    'loan advanced in the last month
    ColAmount(5) = 0
    If Not rstAdvanceLast Is Nothing Then
        rstAdvanceLast.MoveFirst
        rstAdvanceLast.Find "LoanId = " & Loanid
        If Not rstAdvanceLast.EOF Then
            ColAmount(5) = FormatField(rstAdvanceLast(0)) ' Advanced up to Previous month
        End If
    End If
    
    'loan advanced during this month
    ColAmount(6) = 0
    If Not rstAdvanceTill Is Nothing Then
        rstAdvanceTill.MoveFirst
        rstAdvanceTill.Find "LoanId = " & Loanid
        If Not rstAdvanceTill.EOF Then
            ColAmount(6) = FormatField(rstAdvanceTill(0)) ' Advanced during the month
        End If
    End If
    'Toal Loan Advance upto end of this month
    ColAmount(7) = ColAmount(5) + ColAmount(6)
    'Maxmum Loan Balance during the month
    ColAmount(8) = ColAmount(4) + ColAmount(7)
    
    'Loan recoverd up to last month
    ColAmount(9) = 0
    If Not rstRepayLast Is Nothing Then
        rstRepayLast.MoveFirst
        rstRepayLast.Find "LoanId = " & Loanid
        If Not rstRepayLast.EOF Then
            ColAmount(9) = FormatField(rstRepayLast(0)) ' Recovery up to Previous month
        End If
    End If
    
    'Loan recoverd during this month
    ColAmount(10) = 0
    If Not rstRepayTill Is Nothing Then
        rstRepayTill.MoveFirst
        rstRepayTill.Find "LoanId = " & Loanid
        If Not rstRepayTill.EOF Then
            ColAmount(10) = FormatField(rstRepayTill(0)) ' Recovery during this month
        End If
    End If
    
    
    'Total recovery at the end of month
    ColAmount(11) = ColAmount(9) + ColAmount(10)
    'Loan Balance at the end this month
    ColAmount(12) = ColAmount(8) - ColAmount(11)
    
    Debug.Assert ColAmount(12) = FormatField(rstCurrBalance("Balance"))
    'OVER DUE amount as on end of this month
    'i.e. OD of the abave balance
    ColAmount(13) = L_clsLoan.OverDueAmount(Loanid, m_FromDate)
    
    AddRow = False
    For Count = 4 To grd.Cols - 1
        If ColAmount(Count) Then
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
        For Count = 4 To grd.Cols - 1
            grd.Col = Count: grd.Text = FormatCurrency(ColAmount(Count))
            GrandTotal(Count) = GrandTotal(Count) + ColAmount(Count)
        Next
      End With
    End If
    DoEvents
    If gCancel Then Exit Function
    RaiseEvent Processing("Writing the record", rstCurrBalance.AbsolutePosition / rstCurrBalance.RecordCount)
    rstCurrBalance.MoveNext
Wend
AddRow = False
For Count = 4 To grd.Cols - 1
    If GrandTotal(Count) Then
        AddRow = True
        Exit For
    End If
Next
If AddRow Then
  With grd
    If .Rows <= .Row + 3 Then .Rows = .Rows + 2
    .Row = .Row + 2: SlNo = SlNo + 1
    .Col = 2: .Text = "Grand Total": .CellFontBold = True
    For Count = 4 To grd.Cols - 1
        grd.Col = Count: .CellFontBold = True
        grd.Text = FormatCurrency(GrandTotal(Count))
    Next
  End With
End If

Set L_clsLoan = Nothing

lblReportTitle.Caption = "Statement showing the short term loans for the month of " & _
    GetMonthString(Month(m_FromDate)) & " as on " & m_FromIndianDate

grd.Visible = True
         
If LExcel Then
    xlWorkBook.SaveAs App.Path & "|" & "shed2.xls"
    xlWorkBook.Close savechanges:=True
            
    Set xlWorkSheet = Nothing
    Set xlWorkBook = Nothing
End If
         
         
ShowShed2 = True

Exitline:
Screen.MousePointer = vbDefault
grd.Visible = True

If Err Then
    MsgBox "ERROR ShowShed2" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
End If

End Function



Private Sub SetKannadaCaption()
Dim ctrl As Control
On Error Resume Next
For Each ctrl In Me
If Not TypeOf ctrl Is ProgressBar And _
    Not TypeOf ctrl Is Line And _
     Not TypeOf ctrl Is VScrollBar And _
      Not TypeOf ctrl Is Image Then
        ctrl.Font.Name = gFontName
        If Not TypeOf ctrl Is ComboBox Then ctrl.Font.Size = gFontSize
End If
Next
Err.Clear

cmdPrint.Caption = LoadResString(gLangOffSet + 23)  'Print
cmdOk.Caption = LoadResString(gLangOffSet + 1) 'OK


End Sub

Private Sub cmdOk_Click()
Unload Me

End Sub

 
Private Sub cmdPrint_Click()
'Dim my_printClass As New clsPrint
'Set my_printClass.DataSource = grd
'my_printClass.ReportDestination = "PREVIEW"
'my_printClass.ReportTitle = Me.lblReportTitle
'my_printClass.HeaderRectangle = True
'my_printClass.CompanyName = gBankName
'my_printClass.FontName = grd.Font.Name
'frmPrint.picPrint(0).Cls
''my_printClass.ShowPrint
'Set PrintClass = New clsPrint
'PrintClass.ReportTitle = Me.lblReportTitle
'Set PrintClass.DataSource = grd
'PrintClass.CompanyName = gBankName
'PrintClass.ShowPrint

End Sub

Private Sub Form_Load()

'Call CenterMe(Me)
Call SetKannadaCaption

'If m_repType = repConsBalance Then ShowConsoleBalance
'If m_repType = repConsInstOD Then ShowConsoleInstOverDue
'If m_repType = repConsOD Then ShowConsoleODBalance
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
Const MARGIN = 50
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
    cmdOk.Left = fra.Width - cmdOk.Width - (cmdOk.Width / 4)
    cmdPrint.Left = cmdOk.Left - cmdPrint.Width - (cmdPrint.Width / 8)
    cmdPrint.Top = cmdOk.Top
    ' removed the call for personal use - pradeep
    Call InitGrid
End Sub


Private Sub Form_Unload(Cancel As Integer)

Dim Count  As Integer
 For Count = 0 To grd.Cols - 1
            Call SaveSetting(App.EXEName, "LoanReport" & m_repType, "ColWidth", grd.ColWidth(Count))
Next
End Sub


Private Sub grd_Click()
'Dim nwidth As Integer

'nwidth = grd.ColWidth(grd.Col)

'MsgBox grd.Col & " " & nwidth

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

Private Sub PrintClass_ProcessCount(Count As Long)
'm_Count = m_Count + 1
m_Count = Count

End Sub


Private Sub PrintClass_ProcessingMessage(strMessage As String)
On Error Resume Next
RaiseEvent Processing(strMessage, m_Count / m_MaxCount)

End Sub


