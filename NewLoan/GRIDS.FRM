VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRegReport 
   Caption         =   "Grids"
   ClientHeight    =   6495
   ClientLeft      =   945
   ClientTop       =   1695
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   9585
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   285
      Left            =   8640
      TabIndex        =   3
      Top             =   5940
      Width           =   675
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   285
      Left            =   7710
      TabIndex        =   2
      Top             =   5940
      Width           =   705
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   285
      Left            =   6870
      TabIndex        =   1
      Top             =   5940
      Width           =   705
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
   Begin VB.Label lblBankname 
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
      Height          =   315
      Left            =   540
      TabIndex        =   5
      Top             =   480
      Width           =   8085
   End
   Begin VB.Label lblBranchName 
      AutoSize        =   -1  'True
      Caption         =   "..."
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
      Left            =   3270
      TabIndex        =   8
      Top             =   150
      Width           =   450
   End
   Begin VB.Label lblBranch 
      Caption         =   "Branch :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1830
      TabIndex        =   7
      Top             =   150
      Width           =   1365
   End
   Begin VB.Label lblTypeLoan 
      Caption         =   "Type of Loan :"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   1140
      Width           =   1245
   End
   Begin VB.Label lblFrom 
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
      TabIndex        =   4
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

Public p_BankID As Long
Private m_ReportTpe  As Integer

Private Enum wisLoanReportType
    repMonthlyRegister = 1
    repMonthlyRegisterAll = 2
    repShedule_1 = 3
    repShedule_2 = 4
    repShedule_3 = 5
    repShedule_4 = 6
    repShedule_5 = 7
    repShedule_6 = 8
    
    repLoanBalance = 11
    repLoanTransMade = 12
    repLoanIssued = 13
    repLoanInstOD = 14
    repLOanIntCol = 15
    repLoanDailyCash = 16
    repLoanGLedger = 17
    repLoanRepMade = 18
    repLoanOD = 19
    repLoanSanction = 20
    repLoanGuarantor = 21
End Enum

Private m_repType As wisLoanReportType

'Private xlWorkBook As Workbook
'Private xlWorkSheet As Worksheet
Private xlWorkBook As Object
Private xlWorkSheet As Object


Private Sub GridCols(HeadArray() As String, Optional LExcel As Boolean, Optional lSlNo As Boolean)
          
Dim ColNUm As Integer
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
                .Cells(RowNum + 1, ColNUm + 1) = HeadArray(ColNUm)
                .Cells(RowNum + 1, ColNUm + 1).Font.Bold = True
            End With
        End If
    Next
    
    ' if lslno is true
    
    If lSlNo Then
        RowNum = RowNum + 1
        .Row = RowNum
        .Col = 0
        For ColNUm = LBound(HeadArray) To UBound(HeadArray)
            .Col = ColNUm: .Text = ColNUm + 1: .CellFontBold = True: .CellAlignment = 4
            If LExcel Then
                With xlWorkSheet
                    .Cells(RowNum + 1, ColNUm + 1) = ColNUm + 1
                    .Cells(RowNum + 1, ColNUm + 1).Font.Bold = True
                End With
            End If
        Next ColNUm
    End If
    
End With
          

End Sub

Private Sub GridColsKeys(HeadArray() As String, Optional LExcel As Boolean, Optional lSlNo As Boolean)
          
Dim ColNUm As Integer
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
                .Cells(RowNum + 1, ColNUm + 1) = HeadArray(ColNUm)
                .Cells(RowNum + 1, ColNUm + 1).Font.Bold = True
            End With
        End If
    Next
    
    ' if lslno is true
    
    If lSlNo Then
        RowNum = RowNum + 1
        .Row = RowNum
        .Col = 0
        For ColNUm = LBound(HeadArray) To UBound(HeadArray)
            .Col = ColNUm: .Text = ColNUm + 1: .CellFontBold = True: .CellAlignment = 4
            If LExcel Then
                With xlWorkSheet
                    .Cells(RowNum + 1, ColNUm + 1) = ColNUm + 1
                    .Cells(RowNum + 1, ColNUm + 1).Font.Bold = True
                End With
            End If
        Next ColNUm
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
Dim ColNUm As Integer

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
Dim ColNUm As Integer

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

Public Function ShowConsoleBalance(AsOnIndianDate As String) As Boolean
Dim SqlStr As String
Dim Rst As Recordset


SqlStr = "SELECT SUM(Balance),SchemeID,BankID From LoanTrans A,LoanMaste B " & _
    "WHERE A.LoanID = B.LoanID Whe"



End Function

Public Function ShowShed4C(FromDate As String, ToDate As String) As Boolean

' contact pradeep for this function
Dim SlNo As Integer
Dim RowNum As Integer
Dim ColNUm As Integer
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
         " AND a.bankid > " & p_BankID & _
         " AND a.bankid < " & p_BankID + BO_Offset & _
         " AND a.LoanDate >= " & "#" & FromDate & "#" & _
         " AND a.Loandate <= " & "#" & ToDate & "#" & _
         " ORDER BY a.RefId"
         
         
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
Else
    Set rstLoanDetail = gDbTrans.Rst.Clone
End If

' caste wise rst
SqlStr = " SELECT a.RefID,a.CasteID,c.CasteName,a.Male_No,a.Male_Amount,a.Female_No,a.Female_Amount" & _
         " FROM LoanCasteWise a,NewLoans b,Caste c" & _
         " WHERE a.RefId=b.RefId" & _
         " AND a.CasteId=c.CasteId " & _
         " AND b.bankid > " & p_BankID & _
         " AND b.bankid < " & p_BankID + BO_Offset & _
         " AND b.LoanDate >= " & "#" & FromDate & "#" & _
         " AND b.Loandate <= " & "#" & ToDate & "#"
         
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
Else
    Set rstCasteWise = gDbTrans.Rst.Clone
End If


SlNo = 1
RowNum = grd.FixedRows
ColNUm = 0
' start the main loop
While Not rstLoanDetail.EOF

    RefID = FormatField(rstLoanDetail("RefId"))
    BankName = FormatField(rstLoanDetail("BankName"))
    ' find if loan has got data
    rstCasteWise.FindFirst "RefId=" & RefID
    
    If (Not rstCasteWise.NoMatch) And (Not rstCasteWise.EOF) Then  ' if found and not eof
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
                ColNUm = 0
                .Col = ColNUm: .Text = SlNo: .CellAlignment = 4: ColNUm = ColNUm + 1
                .Col = ColNUm: .Text = BankName: ColNUm = ColNUm + 1
                .Col = ColNUm: .Text = CropName: ColNUm = ColNUm + 1
                .Col = ColNUm: .Text = CasteName: ColNUm = ColNUm + 1
                .Col = ColNUm: .Text = Male_No: ColNUm = ColNUm + 1
                .Col = ColNUm: .Text = Male_Amount: ColNUm = ColNUm + 1
                .Col = ColNUm: .Text = Female_No: ColNUm = ColNUm + 1
                .Col = ColNUm: .Text = Female_Amount: ColNUm = ColNUm + 1
                        
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
            
            .Col = ColNUm: .Text = SCF_No: ColNUm = ColNUm + 1
            .Col = ColNUm: .Text = SCF_Amount: ColNUm = ColNUm + 1
            
            totSCF_No = totSCF_No + SCF_No
            totSCF_Amount = totSCF_Amount + SCF_Amount
        End With
        
        RowNum = RowNum + 1
        
        With grd
            ' check the total rows
            MoreRows (2)
            grd.Row = RowNum
            ColNUm = 0
            .Col = ColNUm: .Text = "": ColNUm = ColNUm + 1
            .Col = ColNUm: .Text = "Loan Total": ColNUm = ColNUm + 1
            .Col = ColNUm: .Text = "": ColNUm = ColNUm + 1
            .Col = ColNUm: .Text = "": ColNUm = ColNUm + 1
            .Col = ColNUm: .Text = totMale_No: ColNUm = ColNUm + 1
            .Col = ColNUm: .Text = totMale_Amount: ColNUm = ColNUm + 1
            .Col = ColNUm: .Text = totFemale_No: ColNUm = ColNUm + 1
            .Col = ColNUm: .Text = totFemale_Amount: ColNUm = ColNUm + 1
            .Col = ColNUm: .Text = totSCF_No: ColNUm = ColNUm + 1
            .Col = ColNUm: .Text = totSCF_Amount: ColNUm = ColNUm + 1
                    
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
    ColNUm = 0
    .Col = ColNUm: .Text = "": ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = "Grand Total": .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = "": .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = "": ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = totMale_No: .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = totMale_Amount: .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = totFemale_No: .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = totFemale_Amount: .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = totSCF_No: .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = totSCF_Amount: .CellFontBold = True: ColNUm = ColNUm + 1
            
End With

grd.Visible = True

ShowShed4C = True

End Function


Public Function ShowShed4B(FromDate As String, ToDate As String) As Boolean

Dim RowNum As Integer
Dim ColNUm As Integer
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
         " AND a.bankid > " & p_BankID & _
         " AND a.bankid < " & p_BankID + BO_Offset & _
         " AND a.LoanDate >= " & "#" & FromDate & "#" & _
         " AND a.Loandate <= " & "#" & ToDate & "#" & _
         " ORDER BY a.RefId"
         
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
Else
    Set rstLoanDetail = gDbTrans.Rst.Clone
End If

' this main rst
SqlStr = " SELECT a.RefID,a.BF_No,a.BF_Amount,a.SF_No,a.SF_Amount,a.SCST_No,a.SCST_Amount" & _
         " FROM NewLoanMembers a,NewLoans b " & _
         " WHERE a.RefID=b.RefID" & _
         " AND b.bankid > " & p_BankID & _
         " AND b.bankid < " & p_BankID + BO_Offset & _
         " AND b.LoanDate >= " & "#" & FromDate & "#" & _
         " AND b.Loandate <= " & "#" & ToDate & "#" & _
         " ORDER BY a.RefID"
         
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
Else
    Set rstNewMems = gDbTrans.Rst.Clone
End If

' caste wise rst
SqlStr = " SELECT a.RefID,a.CasteID,c.CasteName,a.Male_No,a.Male_Amount,a.Female_No,a.Female_Amount" & _
         " FROM LoanCasteWise a,NewLoans b,Caste c" & _
         " WHERE a.RefId=b.RefId" & _
         " AND b.bankid > " & p_BankID & _
         " AND b.bankid < " & p_BankID + BO_Offset & _
         " AND b.LoanDate >= " & "#" & FromDate & "#" & _
         " AND b.Loandate <= " & "#" & ToDate & "#" & _
         " AND a.CasteId=c.CasteId "
         
         
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then
    MsgBox "Unable to Fetch the Data"
    Exit Function
Else
    Set rstCasteWise = gDbTrans.Rst.Clone
End If
         
         
RowNum = grd.FixedRows
SlNo = 1
ColNUm = 0

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
    
    rstLoanDetail.FindFirst "RefId=" & RefID
    If Not rstLoanDetail.NoMatch Then
        CropName = FormatField(rstLoanDetail("CropName"))
        BankName = FormatField(rstLoanDetail("BankName"))
    End If
    
       
    With grd
        ' check for the rows
         MoreRows (2)
         ColNUm = 0
         .Row = RowNum
         
        .Col = ColNUm: .Text = str(SlNo): .CellAlignment = 4: ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = BankName: .CellAlignment = 1: ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = CropName: .CellAlignment = 1: ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = BF_No: ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = FormatCurrency(BF_Amount): ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = SF_No: ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = FormatCurrency(SF_Amount): ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = SC_No: ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = FormatCurrency(SC_Amount): ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = totNumber: ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = FormatCurrency(totAmount): ColNUm = ColNUm + 1
        
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
    ColNUm = 0
    .Col = ColNUm: .Text = "": .CellAlignment = 4: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = "Grand Total": .CellFontBold = True: .CellAlignment = 1: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = "": .CellAlignment = 1: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = totBF_No: .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = FormatCurrency(totBF_Amount): .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = totSF_No: .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = FormatCurrency(totSF_Amount): .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = totSC_No: .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = FormatCurrency(totSC_Amount): .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = grdNumber: .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = FormatCurrency(grdAmount): .CellFontBold = True: ColNUm = ColNUm + 1
End With

grd.Visible = True

ShowShed4B = True

End Function


Public Function ShowShed4A(FromDate As String, ToDate As String) As Boolean

Dim rstLoanDetail As Recordset
Dim SqlStr As String
Dim RowNum As Integer
Dim ColNUm As Integer
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
         " AND a.bankid > " & p_BankID & _
         " AND a.bankid < " & p_BankID + BO_Offset & _
         " AND a.LoanDate >= " & "#" & FromDate & "#" & _
         " AND a.Loandate <= " & "#" & ToDate & "#"
         
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then
    MsgBox "Unable to fetch the Data"
    Exit Function
Else
    Set rstLoanDetail = gDbTrans.Rst.Clone
End If


' start the main loop
SlNo = 1
RowNum = grd.FixedRows
ColNUm = 0
While Not rstLoanDetail.EOF
    ColNUm = 0
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
        .Col = ColNUm: .Text = str(SlNo): .CellAlignment = 4: ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = FormatField(rstLoanDetail("BankName")): ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = FormatField(rstLoanDetail("CropName")): ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = FormatField(rstLoanDetail("LoanDate")): ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = FormatField(rstLoanDetail("LoanDueDate")): ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = BF_No: ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = FormatCurrency(BF_Amount): ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = SF_No: ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = FormatCurrency(SF_Amount): ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = SC_No: ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = FormatCurrency(SC_Amount): ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = totNumber: ColNUm = ColNUm + 1
        .Col = ColNUm: .Text = FormatCurrency(totAmount): ColNUm = ColNUm + 1
        
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
    ColNUm = 0
    .Col = ColNUm: .Text = "": .CellAlignment = 4: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = "Grand Total": .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = "": .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = "": .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = "": .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = totBF_No: .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = FormatCurrency(totBF_Amount): .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = totSF_No: .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = FormatCurrency(totSF_Amount): .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = totSC_No: .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = FormatCurrency(totSC_Amount): .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = grdNumber: .CellFontBold = True: ColNUm = ColNUm + 1
    .Col = ColNUm: .Text = FormatCurrency(grdAmount): .CellFontBold = True: ColNUm = ColNUm + 1
End With

ShowShed4A = True

End Function


Private Function Shed1RowCOl() As Boolean

Dim ShedHeader(0 To 16, 1) As String
Dim ShedHeader1(0) As String

Shed1RowCOl = False

    ShedHeader(0, 0) = "Sl No"
    ShedHeader(1, 0) = "Name of the Customer"
    ShedHeader(2, 0) = "Name of the Society"
    ShedHeader(3, 0) = "Arrears"
    ShedHeader(4, 0) = "Current"
    ShedHeader(5, 0) = "Total (3+4)"
    ShedHeader(6, 0) = "Against Arrears Demand"
    ShedHeader(7, 0) = "Against Current Demand"
    ShedHeader(8, 0) = "Advance Recovery If any"
    ShedHeader(9, 0) = "Total Recovery (6 + 7 + 8)"
    ShedHeader(10, 0) = "Balance of Overdue (5-(6+7))"
    ShedHeader(11, 0) = "Less than One year"
    ShedHeader(12, 0) = "1 to 2 years"
    ShedHeader(13, 0) = "2 to 3 years"
    ShedHeader(14, 0) = "3 to 4 years"
    ShedHeader(15, 0) = "4 to 5 years"
    ShedHeader(16, 0) = "Above 5 Years"

    ShedHeader(0, 1) = "500"
    ShedHeader(1, 1) = "3000"
    ShedHeader(2, 1) = "2000"
    ShedHeader(3, 1) = "1200"
    ShedHeader(4, 1) = "1200"
    ShedHeader(5, 1) = "1200"
    ShedHeader(6, 1) = "1200"
    ShedHeader(7, 1) = "1200"
    ShedHeader(8, 1) = "1050"
    ShedHeader(9, 1) = "1095"
    ShedHeader(10, 1) = "1200"
    ShedHeader(11, 1) = "1125"
    ShedHeader(12, 1) = "1125"
    ShedHeader(13, 1) = "1125"
    ShedHeader(14, 1) = "1125"
    ShedHeader(15, 1) = "1125"
    ShedHeader(16, 1) = "1125"
    
    Call shGridCols(1, 2, ShedHeader(), ShedHeader1(), False, True)
    
    With grd
        .RowHeight(0) = 1200
    End With

    
Shed1RowCOl = True

End Function

Private Function Shed2RowCol() As Boolean

Dim ShedHeader(0 To 12, 1) As String
Dim ShedHeader2(0) As String

Shed2RowCol = False

    ShedHeader(0, 0) = "Sl No"
    ShedHeader(1, 0) = "Name of the Customer"
    ShedHeader(2, 0) = "Name of the Society"
    ShedHeader(3, 0) = "Loan Outstanding as on 1 st July"
    ShedHeader(4, 0) = "Loan Advanced Up to the previous month"
    ShedHeader(5, 0) = "Loan Advanced during the month"
    ShedHeader(6, 0) = "Total Loan Advanced upto the end of the month (4 + 5)"
    ShedHeader(7, 0) = "Outstanding at the end of the month 3 + 6"
    ShedHeader(8, 0) = "Recovery Upto Previous month"
    ShedHeader(9, 0) = "Recovery during the month"
    ShedHeader(10, 0) = "Total Recovery upto the end of the the month 8 + 9 "
    ShedHeader(11, 0) = "Balance at the end of the month 7 - 10"
    ShedHeader(12, 0) = "Out of Which overdue as at the end of the month"

    ShedHeader(0, 1) = "500"
    ShedHeader(1, 1) = "3000"
    ShedHeader(2, 1) = "2000"
    ShedHeader(3, 1) = "1200"
    ShedHeader(4, 1) = "1200"
    ShedHeader(5, 1) = "1200"
    ShedHeader(6, 1) = "1200"
    ShedHeader(7, 1) = "1200"
    ShedHeader(8, 1) = "1200"
    ShedHeader(9, 1) = "1200"
    ShedHeader(10, 1) = "1200"
    ShedHeader(11, 1) = "1300"
    ShedHeader(12, 1) = "1300"
        
' populate the grid with the headings

Call shGridCols(1, 2, ShedHeader(), ShedHeader2(), , True)

With grd
    .RowHeight(0) = 1200
End With

  
Shed2RowCol = True

End Function

Private Sub shGridCols(FCol As Integer, FRow As Integer, HeadArray() As String, FirstHead() As String, Optional LExcel As Boolean, Optional lSlNo As Boolean)
          
Dim ColNUm As Integer
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
        For ColNUm = LBound(FirstHead) To UBound(FirstHead)
            .Col = ColNUm: .Text = FirstHead(ColNUm): .CellFontBold = True: .CellAlignment = 4
            If LExcel Then
                With xlWorkSheet
                    .Cells(RowNum + 1, ColNUm + 1) = FirstHead(ColNUm)
                    .Cells(RowNum + 1, ColNUm + 1).Font.Bold = True
                End With
            End If
        Next
        RowNum = RowNum + 1
        .Row = RowNum
    End If
                    
                
    For ColNUm = LBound(HeadArray) To UBound(HeadArray)
        .Col = ColNUm: .Text = HeadArray(ColNUm, 0): .CellFontBold = True: .CellAlignment = 4
        .ColWidth(ColNUm) = Val(HeadArray(ColNUm, 1))
        If LExcel Then
            With xlWorkSheet
                .Cells(RowNum + 1, ColNUm + 1) = HeadArray(ColNUm, 0)
                .Cells(RowNum + 1, ColNUm + 1).Font.Bold = True
            End With
        End If
    Next
    
    Dim SlNo As Integer
    
    If lSlNo Then
        RowNum = RowNum + 1
        .Row = RowNum
        .Col = 0
        SlNo = 0
        For ColNUm = LBound(HeadArray) To UBound(HeadArray)
            If ColNUm <> 2 Then SlNo = SlNo + 1
            .Col = ColNUm: .Text = SlNo: .CellFontBold = True: .CellAlignment = 4
            If LExcel Then
                With xlWorkSheet
                    .Cells(RowNum + 1, ColNUm + 1) = ColNUm + 1
                    .Cells(RowNum + 1, ColNUm + 1).Font.Bold = True
                End With
            End If
        Next
        .Col = 2: .Text = "2a"
    End If
End With
          

End Sub

Private Sub GridRows(HeadArray() As String, Optional LExcel As Boolean)
          
Dim ColNUm As Integer
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
                .Cells(RowNum + 1, ColNUm + 1) = HeadArray(RowNum)
                .Cells(RowNum + 1, ColNUm + 1).Font.Bold = True
            End With
        End If
        RowNum = RowNum + 1
    Next Items
    
End With
          
            


End Sub


Private Sub shGridRows(HeadArray() As String, Optional LExcel As Boolean)
          
Dim ColNUm As Integer
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
                .Cells(RowNum + 1, ColNUm + 1) = HeadArray(RowNum)
                .Cells(RowNum + 1, ColNUm + 1).Font.Bold = True
            End With
        End If
    Next RowNum
    
End With
          
            


End Sub

Private Function Shed6RowCol() As Boolean

Dim ShedHeaders(0 To 16, 1) As String
Dim ShedHeaders2(0) As String


Shed6RowCol = False

ShedHeaders(0, 0) = "slno"
ShedHeaders(1, 0) = "Name of Customer"
ShedHeaders(2, 0) = "Name Society"
ShedHeaders(3, 0) = "Date of Sanction"
ShedHeaders(4, 0) = "Due Date"
ShedHeaders(5, 0) = "Balance at the Begining of the Month"
ShedHeaders(6, 0) = "Advances During the Month"
ShedHeaders(7, 0) = "Recovered During the Month"
ShedHeaders(8, 0) = "Outstanding at the End of the Month"
ShedHeaders(9, 0) = "of which Over Dues"
ShedHeaders(10, 0) = "Maximum Outstanding During the Month"
ShedHeaders(11, 0) = "Under 1 Year"
ShedHeaders(12, 0) = "1 to 2 Years"
ShedHeaders(13, 0) = "2 to 3 Years"
ShedHeaders(14, 0) = "3 to 4 Years"
ShedHeaders(15, 0) = "4 to 5 Years"
ShedHeaders(16, 0) = "Above 5 Years"

ShedHeaders(0, 1) = "1200"
ShedHeaders(1, 1) = "2500"
ShedHeaders(2, 1) = "2000"
ShedHeaders(3, 1) = "1200"
ShedHeaders(4, 1) = "1200"
ShedHeaders(5, 1) = "1500"
ShedHeaders(6, 1) = "1500"
ShedHeaders(7, 1) = "1500"
ShedHeaders(8, 1) = "1500"
ShedHeaders(9, 1) = "1500"
ShedHeaders(10, 1) = "1500"
ShedHeaders(11, 1) = "1500"
ShedHeaders(12, 1) = "1500"
ShedHeaders(13, 1) = "1500"
ShedHeaders(14, 1) = "1500"
ShedHeaders(15, 1) = "1500"
ShedHeaders(16, 1) = "1500"

Call shGridCols(1, 2, ShedHeaders(), ShedHeaders2(), False, True)

grd.RowHeight(0) = 1200

  
Shed6RowCol = True


End Function


Private Function Shed5RowCol() As Boolean

Dim ShedHeaders(0 To 18, 1) As String
Dim ShedHeaders2(0) As String
  
Shed5RowCol = False

ShedHeaders(0, 0) = "slno"
ShedHeaders(1, 0) = "Name of Customer"
ShedHeaders(2, 0) = "Name of Society"
ShedHeaders(3, 0) = "Limit Sanctioned"
ShedHeaders(4, 0) = "Date of Sanction"
ShedHeaders(5, 0) = "Due Date"
ShedHeaders(6, 0) = "Purpose"
ShedHeaders(7, 0) = "Balance at the Begining of the Month"
ShedHeaders(8, 0) = "Advances During the Month"
ShedHeaders(9, 0) = "Recovered During the Month"
ShedHeaders(10, 0) = "Outstanding at the End of the Month"
ShedHeaders(11, 0) = "of which Over Dues"
ShedHeaders(12, 0) = "Maximum Outstanding During the Month"
ShedHeaders(13, 0) = "Under 1 Year"
ShedHeaders(14, 0) = "1 to 2 Years"
ShedHeaders(15, 0) = "2 to 3 Years"
ShedHeaders(16, 0) = "3 to 4 Years"
ShedHeaders(17, 0) = "4 to 5 Years"
ShedHeaders(18, 0) = "Above 5 Years"

ShedHeaders(0, 1) = "1200"
ShedHeaders(1, 1) = "2500"
ShedHeaders(2, 1) = "2000"
ShedHeaders(3, 1) = "1200"
ShedHeaders(4, 1) = "1200"
ShedHeaders(5, 1) = "1200"
ShedHeaders(6, 1) = "1800"
ShedHeaders(7, 1) = "1500"
ShedHeaders(8, 1) = "1500"
ShedHeaders(9, 1) = "1500"
ShedHeaders(10, 1) = "1500"
ShedHeaders(11, 1) = "1500"
ShedHeaders(12, 1) = "1500"
ShedHeaders(13, 1) = "1500"
ShedHeaders(14, 1) = "1500"
ShedHeaders(15, 1) = "1500"
ShedHeaders(16, 1) = "1500"
ShedHeaders(17, 1) = "1500"
ShedHeaders(18, 1) = "1500"

Call shGridCols(1, 2, ShedHeaders(), ShedHeaders2(), False, True)

grd.RowHeight(0) = 1200
  
Shed5RowCol = True


End Function

Private Function Shed4ARowCol() As Boolean

Dim RowNum As Integer
Dim ColNUm As Integer
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
Dim ratio As Double

Select Case choice
    Case "Shedule2"
        ratio = grd.Width / grd.Cols

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



Private Sub InitGrid(Optional Resize As Boolean)

Dim ColCount As Long
    
For ColCount = 0 To grd.Cols - 1
  grd.ColWidth(ColCount) = GetSetting(App.EXEName, "LoanReport" & m_repType, "ColWidth" & ColCount, grd.Width / grd.Cols) * grd.Width
  If Not Resize Then
    If grd.ColWidth(ColCount) > grd.Width / 2 Then grd.ColWidth(ColCount) = grd.Width / grd.Cols
    If grd.ColWidth(ColCount) < 10 Then grd.ColWidth(ColCount) = grd.Width / grd.Cols * 2
  End If
Next ColCount

End Sub


Public Function ShowShed6(AsOnIndianDate As String) As Boolean

m_repType = repShedule_6

ShowShed6 = False
Err.Clear
On Error GoTo ExitLine:


Dim SqlStr As String
Dim TransType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes

Dim rstOpBalance As Recordset
Dim rstClBalance As Recordset
Dim rstAdvance As Recordset
Dim rstRecovery As Recordset

Dim AsOnDate As Date
Dim FirstDate As Date

Dim ColAmount() As Currency
Dim GrandTotal() As Currency


If AsOnIndianDate = "" Then
    AsOnDate = gStrDate
Else
    AsOnDate = FormatDate(AsOnIndianDate)
End If
'Get the First day of the Month
FirstDate = Month(AsOnDate) & "/1/" & Year(AsOnDate)
Dim LoanType As wisLoanType
Dim LoanTerm As wisLoanTerm
Dim LoanCategary As wisLoanCategories

LoanTerm = wisLongTerm
LoanType = wisIndividualLoan
LoanCategary = wisNonAgriculural
'Get The LoanDetails And THier balance as on Date


SqlStr = "SELECT A.LoanID,LoanAccNo,IssueDate,LoanDueDate,LoanAmount,Balance,LoanPurpose," & _
    " Title +' ' + FirstName +' ' + MiddleName + ' ' + LastName As Name," & _
    " CustBankName FROM LoanMaster A,LoanTrans B, CustMaster C " & _
    " WHERE A.SchemeId IN (SELECT SchemeID FROM LoanScheme WHERE " & _
        " Category = " & LoanCategary & " AND (LoanType = " & LoanType & _
        " OR TermType = " & LoanTerm & "))" & _
    " AND B.BankID = A.BankID AND C.BankID = A.BankID AND B.LoanID = A.LoanID" & _
    " AND C.CustomerId =A.CustomerID AND TransID = (SELECT MAX(TransId) FROM " & _
        " LoanTrans D WHERE D.TransDate <= #" & AsOnDate & "# AND D.BankID = A.BankID" & _
        " AND D.LoanId = A.LoanId ) " & _
    " AND (LoanClosed = NULL OR LOanClosed > #" & AsOnDate & "#) AND A.BankID = " & gBankID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then Exit Function
Set rstClBalance = gDbTrans.Rst.Clone

'Get The LoanDetails And THier balance as on first day of the given month
SqlStr = "SELECT A.LoanID,LoanAccNo,IssueDate,LoanDueDate,LoanAmount,Balance," & _
    " CustBankName FROM LoanMaster A,LoanTrans B WHERE B.BankID = A.BankID" & _
    " AND B.LoanID = A.LoanID AND TransID =" & _
        " (SELECT MAX(TransId) FROM LoanTrans D WHERE D.TransDate <" & _
        " #" & FirstDate & "# AND D.BankID = A.BankID AND D.LoanId = A.LoanId)" & _
    " AND A.BankID = " & gBankID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then Exit Function
Set rstOpBalance = gDbTrans.Rst.Clone

'Get The Advances During the MOnth
TransType = wWithdraw
ContraTrans = wContraWithdraw
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans WHERE TransDate >= #" & FirstDate & "#" & _
    " AND TransDate <= #" & AsOnDate & "# AND (TransType = " & TransType & _
    " OR TransType = " & ContraTrans & ") AND BankID = " & gBankID & _
    " GROUP BY LoanID "
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch > 0 Then
    Set rstAdvance = gDbTrans.Rst.Clone
Else
    Set rstAdvance = Nothing
End If


'Get The Recovery During the MOnth
TransType = wDeposit
ContraTrans = wContraDeposit
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans WHERE TransDate >= #" & FirstDate & "#" & _
    " AND TransDate <= #" & AsOnDate & "# AND (TransType = " & TransType & _
    " OR TransType = " & ContraTrans & ") AND BankID = " & gBankID & _
    " GROUP BY LoanID "
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch > 0 Then
    Set rstRecovery = gDbTrans.Rst.Clone
Else
    Set rstRecovery = Nothing
End If

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
Dim Count As Long
Dim SlNo As Long


SlNo = 0
While Not rstClBalance.EOF
    LoanID = FormatField(rstClBalance("LoanId"))
    
    rstOpBalance.FindFirst "LoanID = " & LoanID
    ColAmount(5) = 0
    If Not rstOpBalance.NoMatch Then
        ColAmount(5) = FormatField(rstOpBalance("Balance")) 'Balance as on 31/3/yyyy
    End If
    ColAmount(6) = 0
    If Not rstAdvance Is Nothing Then
        rstAdvance.FindFirst "LoanId = " & LoanID
        If Not rstAdvance.NoMatch Then
            ColAmount(6) = FormatField(rstAdvance(0))  'Advances During the MOnth
        End If
    End If
    ColAmount(7) = 0
    If Not rstRecovery Is Nothing Then
        rstRecovery.FindFirst "LoanId = " & LoanID
        If Not rstRecovery.NoMatch Then
            ColAmount(7) = FormatField(rstRecovery(0)) 'Recovery During the mOnth
        End If
    End If
    ColAmount(8) = FormatField(rstClBalance("Balance")) 'Balance at the end of month
    ODAmount = L_clsLoan.OverDueAmount(LoanID, , AsOnIndianDate) 'Over due
    ColAmount(9) = ODAmount 'Over due amount of the loan as on given date
    
    'Over due amount calssifiacation
    PrevOD = 0
    ODAmount = L_clsLoan.OverDueSince(5, LoanID, , AsOnIndianDate)
    ColAmount(15) = ODAmount - PrevOD 'Over due since & above 5 Years
    PrevOD = ColAmount(15) + PrevOD
    ODAmount = L_clsLoan.OverDueSince(4, LoanID, , AsOnIndianDate)
    ColAmount(14) = ODAmount - PrevOD  'Over due since 4 Years
    PrevOD = ColAmount(14) + PrevOD
    ODAmount = L_clsLoan.OverDueSince(3, LoanID, , AsOnIndianDate)
    ColAmount(13) = ODAmount - PrevOD  'Over due since 3 Years
    PrevOD = ColAmount(13) + PrevOD
    ODAmount = L_clsLoan.OverDueSince(2, LoanID, , AsOnIndianDate)
    ColAmount(12) = ODAmount - PrevOD  'Over due since 2 Years
    PrevOD = ColAmount(12) + PrevOD
    ODAmount = L_clsLoan.OverDueSince(1, LoanID, , AsOnIndianDate)
    ColAmount(11) = ODAmount - PrevOD  'Over due a 1 Year
    PrevOD = ColAmount(11) + PrevOD
    ColAmount(10) = ColAmount(9) - PrevOD  'Over due Less than a year
    
    'Check whther this row has to be write or not
    AddRow = False
    For Count = 7 To grd.Cols - 1
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
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = FormatField(rstClBalance("Name"))
        .Col = 2: .Text = FormatField(rstClBalance("CustBankName"))
        .Col = 3: .Text = FormatField(rstClBalance("IssueDate"))
        .Col = 4: .Text = FormatField(rstClBalance("LoanDueDate"))
        For Count = 5 To grd.Cols - 1
            .Col = Count
            .Text = FormatCurrency(ColAmount(Count))
            GrandTotal(Count) = GrandTotal(Count) + ColAmount(Count)
        Next
      End With
    End If
    rstClBalance.MoveNext
Wend
Set L_clsLoan = Nothing
With grd
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 1
    .Text = "Grand Total"
    For Count = 5 To .Cols - 1
        .Col = Count
        .Text = FormatCurrency(GrandTotal(Count))
    Next
End With

ShowShed6 = True
ExitLine:
If Err Then
    MsgBox "ERROR SHED 5" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
End If

End Function
Public Function ShowMeetingRegistar(Optional SchemeID As Integer = 0, Optional AsOnIndianDate As String, Optional WhichLoans As wisRepSocietyLoans = BothLoans) As Boolean

Dim SqlPrin As String
Dim SqlInt As String
Dim SqlStr As String
Dim Rst As Recordset
Dim rstLoanScheme As Recordset
Dim SchemeName  As String
Dim Date31_3  As Date
Dim DateAsOn As Date
Dim DateLastMonth As Date
Dim TransType As wisTransactionTypes
Dim LoanType As wisLoanType
Dim SchemeStr  As String

On Error GoTo ErrLine
'Get all date in the format of system format

DateAsOn = FormatDate(AsOnIndianDate)
Date31_3 = "31/3/" & Val(Year(DateAsOn) - IIf(Month(DateAsOn) > 3, 0, 1))

DateLastMonth = FormatDate("1/" & Month(DateAsOn) & "/" & Year(DateAsOn))
DateLastMonth = DateAdd("d", -1, DateAsOn)

If SchemeID Then
    m_repType = repMonthlyRegister
   SqlStr = "SELECT * FROM LoanScheme Where SchemeID = " & SchemeID
    gDbTrans.SqlStmt = SqlStr
    Call gDbTrans.SQLFetch
    Set rstLoanScheme = gDbTrans.Rst.Clone
    SchemeName = FormatField(gDbTrans.Rst("SchemeName"))
    SchemeStr = " SchemeID = " & SchemeID & " "
    LoanType = FormatField(gDbTrans.Rst("LoanType"))
Else
    m_repType = repMonthlyRegisterAll
    SchemeStr = " SchemeID  <> " & SchemeID & " "
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
SqlStr = "SELECT * From LoanMaster WHERE " & SchemeStr & " AND BankId = " & gBankID
If WhichLoans = DirectLoans Then
'    SqlStr = SqlStr & " AND CustBankID = 0"
ElseIf WhichLoans = SocietyThruLoans Then
 '   SqlStr = SqlStr & " AND CustBankID <> 0"
End If


gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch > 0 Then
    Set rstMaster = gDbTrans.Rst.Clone
Else
    GoTo ErrLine
End If

'Get the Loan balance on 31/3/yyyy
SqlPrin = "SELECT A.LoanID,Balance  FROM LoanTrans A WHERE BankID = " & gBankID & _
     " AND A.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanTrans C WHERE TransDate <= #" & Date31_3 & "# " & _
        " AND C.LoanID = A.LoanID And C.BankID = A.BankID )"
        
gDbTrans.SqlStmt = SqlPrin
If gDbTrans.SQLFetch > 0 Then
    Set rstPrin31_3 = gDbTrans.Rst.Clone
End If
'Get the Interest Balance 31/3/yyyy
SqlPrin = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B " & _
    " WHERE BankID = " & gBankID & " AND B.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanIntTrans C WHERE TransDate <= #" & DateAsOn & "# " & _
        " AND C.LoanID = B.LoanID And C.BankID = B.BankID )"
        
gDbTrans.SqlStmt = SqlPrin
If gDbTrans.SQLFetch > 0 Then
    Set rstInt31_3 = gDbTrans.Rst.Clone
End If


'Get the Loan balance as on lastMonth
SqlPrin = "SELECT A.LoanID,TransDate,Balance FROM LoanTrans A WHERE " & _
     " BankID = " & gBankID & " AND A.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanTrans C WHERE TransDate <= #" & DateLastMonth & "# " & _
        " AND C.LoanID = A.LoanID And C.BankID = A.BankID )"

'Get the Interest Balance ON  LAST MONTH
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B " & _
    " WHERE BankID = " & gBankID & " AND B.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanIntTrans C WHERE TransDate <= #" & DateLastMonth & "# " & _
        " AND C.LoanID = B.LoanID And C.BankID = B.BankID )"
        
gDbTrans.SqlStmt = SqlPrin
If gDbTrans.SQLFetch > 0 Then
    Set rstPrinLastMonth = gDbTrans.Rst.Clone
End If
gDbTrans.SqlStmt = SqlInt
If gDbTrans.SQLFetch > 0 Then
    Set rstIntLastMonth = gDbTrans.Rst.Clone
End If

'Get the Loan balance as on date
SqlPrin = "SELECT A.LoanID,TransDate,Balance FROM LoanTrans A WHERE " & _
     " BankID = " & gBankID & " AND A.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanTrans C WHERE TransDate <= #" & DateAsOn & "# " & _
        " AND C.LoanID = A.LoanID And C.BankID = A.BankID )"

'Get the Interest Balance ON  Date
SqlInt = "SELECT B.LoanID,IntBalance,TransDate FROM LoanIntTrans B " & _
    " WHERE BankID = " & gBankID & " AND B.TransId = (SELECT MAX(TransID) FROM " & _
        " LoanIntTrans C WHERE TransDate <= #" & DateAsOn & "# " & _
        " AND C.LoanID = B.LoanID And C.BankID = B.BankID )"
        
gDbTrans.SqlStmt = SqlPrin
If gDbTrans.SQLFetch > 0 Then
    Set rstPrinAsOn = gDbTrans.Rst.Clone
End If
gDbTrans.SqlStmt = SqlInt
If gDbTrans.SQLFetch > 0 Then
    Set rstIntAsOn = gDbTrans.Rst.Clone
End If

'GEt the Transacted amount After 31/3/yyyy till last month
SqlPrin = "SELECT SUM(AMOUNT) as SumAmount,LoanID,TransType FROM LoanTrans WHERE " & _
    " TransDate > #" & Date31_3 & "# AND TransDate < #" & DateLastMonth & "# " & _
    " AND BankID = " & gBankID & " GROUP BY LoanId,TransType"
SqlInt = "SELECT SUM(IntAmount) as SumIntAmount,SUM(PenalIntAmount) as SumPenalIntAmount," & _
    " LoanID,TransType FROM LoanIntTrans WHERE " & _
    " TransDate > #" & Date31_3 & "# AND TransDate < #" & DateLastMonth & "# " & _
    " AND BankID = " & gBankID & " GROUP BY LoanId,TransType"

gDbTrans.SqlStmt = SqlPrin
If gDbTrans.SQLFetch > 0 Then
    Set rstPrinTransLast = gDbTrans.Rst.Clone
End If
gDbTrans.SqlStmt = SqlInt
If gDbTrans.SQLFetch > 0 Then
    Set rstIntTransLast = gDbTrans.Rst.Clone
End If

'GEt the Transacted amount From last month to till Today
SqlPrin = "SELECT SUM(AMOUNT)as SumAmount,LoanID,TransType FROM LoanTrans WHERE " & _
    " TransDate > #" & DateLastMonth & "# AND TransDate < #" & DateAsOn & "# " & _
    " AND BankID = " & gBankID & " GROUP BY LoanId,TransType"
SqlInt = "SELECT SUM(IntAmount) as SumIntAmount,SUM(PenalIntAmount) as SumPenalIntAmount," & _
    " LoanID,TransType FROM LoanIntTrans WHERE " & _
    " TransDate > #" & DateLastMonth & "# AND TransDate < #" & DateAsOn & "# " & _
    " AND BankID = " & gBankID & " GROUP BY LoanId,TransType"

gDbTrans.SqlStmt = SqlPrin
If gDbTrans.SQLFetch > 0 Then
    Set rstPrinTransAsOn = gDbTrans.Rst.Clone
End If
gDbTrans.SqlStmt = SqlInt
If gDbTrans.SQLFetch > 0 Then
    Set rstIntTransAsOn = gDbTrans.Rst.Clone
End If

'Now Initialise the grid
grd.Clear
grd.Cols = 23
grd.Rows = 20

Dim SlNo As Integer
Dim LoanID As Long
Dim L_clsCust As New clsCustDet
Dim L_clsLoan As New clsLoan
Dim RetStr As String
Dim StrArr() As String
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

Call SetGrid(SchemeID, IIf(WhichLoans = DirectLoans, False, True), CStr(Year(Date31_3)))

'Call InitGrid
'grd.MergeCells = flexMergeNever

grd.Row = 4
If SchemeID = 0 Then GoTo WithoutSchemeID
grd.Row = 5
Do
    If rstMaster.EOF Then
        Exit Do
    End If
    Balance31_3 = 0: BalanceLastMonth = 0: BalanceNow = 0
    IntBal31_3 = 0: IntBalLastMonth = 0: IntBalNow = 0
    LoanID = FormatField(rstMaster("LoanID"))
    IntRate = FormatField(rstMaster("IntRate"))
    SlNo = SlNo + 1
  With grd
    If .Rows < .Row + 2 Then .Rows = .Rows + 2
    .Row = .Row + 1
    .MergeRow(.Row) = False
    .Col = 0: .Text = Format(SlNo, "00")
    .Col = 1: .Text = FormatField(rstMaster("CustBankName"))
    .Col = 2: .Text = L_clsCust.CustomerName(FormatField(rstMaster("LoanID")))
    RetStr = FormatField(rstMaster("Guarantor1"))
    On Error Resume Next
    Call GetStringArray(RetStr, StrArr, gDelim)
    .Col = 3: .Text = StrArr(0) & vbCrLf & StrArr(1)
    '.Row = .Row + 1: .Text = StrArr(1): .Row = .Row - 1
    RetStr = FormatField(rstMaster("Guarantor2"))
    Call GetStringArray(RetStr, StrArr, gDelim)
    .Col = 4: .Text = StrArr(0) & vbCrLf & StrArr(1)
    '.Row = .Row + 1: .Text = StrArr(1): .Row = .Row - 1
    On Error GoTo ErrLine
    'Loan Advance Date
    .Col = 5: .Text = FormatField(rstMaster("IssueDate"))
    .Col = 6: .Text = FormatField(rstMaster("LoanAmount"))
    
    RetStr = FormatField(rstMaster("otherDets"))
    Call GetStringArray(RetStr, StrArr, gDelim)
    If LoanType = wisVehicleloan Then
        On Error Resume Next
        .Col = 7: .Text = StrArr(0)
        .Col = 8: .Text = StrArr(1)
        On Error GoTo ErrLine
    End If
        
    
    'Out standing Loan Balance as on 31/3/yyyy
    PrevDate = Date31_3
    rstPrin31_3.FindFirst " LoanID = " & LoanID
    If Not rstPrin31_3.NoMatch Then
        rstInt31_3.FindFirst " LoanID = " & LoanID
        Balance31_3 = rstPrin31_3("Balance")
        TransDate = rstInt31_3("TransDate")
        IntBal31_3 = Val(rstInt31_3("IntBalance"))
        IntBal31_3 = IntBal31_3 + DateDiff("d", TransDate, Date31_3) / 365 * IntRate / 100 * Balance
        .Col = 9: .Text = Balance31_3
        PrevDate = TransDate
    End If
    'Over due as on 31/3/yyyy
    ODAmount = L_clsLoan.OverDueAmount(LoanID, , FormatDate(CStr(Date31_3)))
    ODInt = L_clsLoan.OverDueInterest(LoanID, FormatDate(CStr(Date31_3)))
    .Col = 10: .Text = L_clsLoan.DueInstallments(LoanID, FormatDate(CStr(Date31_3)))
    .Col = 11: .Text = FormatCurrency(ODAmount)
    .Col = 12: .Text = FormatCurrency(ODInt)
    .Col = 13: .Text = FormatCurrency(ODAmount + ODInt)
    
    'Loan Repayment From 1/4/yyyy to last month
    TransType = wDeposit
    rstPrinTransLast.FindFirst " LoanID = " & LoanID & " AND TransType = " & TransType
    If Not rstPrinLastMonth.NoMatch Then
        rstIntTransLast.FindFirst " LoanID = " & LoanID & " AND TransType = " & TransType
        Amount = FormatField(rstPrinTransLast("SumAmount"))
        IntAmount = FormatField(rstIntTransLast("SumIntAmount"))
        IntAmount = IntAmount + FormatField(rstIntTransLast("SumPenalIntAmount"))
        .Col = 14: '.Text = IntBalLastMonth
        .Col = 15: .Text = FormatCurrency(Amount)
        .Col = 16: .Text = FormatCurrency(IntAmount)
        .Col = 17: .Text = FormatCurrency(Amount + IntAmount)
    End If
    
    'Out standing Loan Balance as on last month
    rstPrinLastMonth.FindFirst " LoanID = " & LoanID
    If Not rstPrinLastMonth.NoMatch Then
        rstIntLastMonth.FindFirst " LoanID = " & LoanID
        Amount = FormatField(rstPrinLastMonth("Balance"))
        IntAmount = FormatField(rstIntLastMonth("IntBalance"))
        'IntAmount = IntAmount + FormatField(rstIntLastMonth("PenalIntAmount"))
        .Col = 18: '.Text = IntBalLastMonth
        .Col = 19: .Text = FormatCurrency(Amount)
        .Col = 20: .Text = FormatCurrency(IntAmount)
        .Col = 21: .Text = FormatCurrency(Amount + IntAmount)
    End If
    'Over due as on last month
    ODAmount = L_clsLoan.OverDueAmount(LoanID, , FormatDate(CStr(DateLastMonth)))
    ODInt = L_clsLoan.OverDueAmount(LoanID, , FormatDate(CStr(DateLastMonth)))
    .Col = 22: .Text = L_clsLoan.DueInstallments(LoanID, FormatDate(CStr(DateLastMonth)))
    .Col = 23: .Text = FormatCurrency(ODAmount)
    .Col = 24: .Text = FormatCurrency(ODInt)
    .Col = 25: .Text = FormatCurrency(ODAmount + ODInt)
    
    'If Case has field Then Date Of Filing
    .Col = 26: .Text = FormatField(rstMaster("ABNDate"))
    .Col = 27: .Text = FormatField(rstMaster("EpDate"))
  End With
    rstMaster.MoveNext
Loop
Set L_clsLoan = Nothing

ShowMeetingRegistar = True
Screen.MousePointer = vbDefault
Exit Function

WithoutSchemeID:
Do
    If rstMaster.EOF Then Exit Do
    Balance31_3 = 0: BalanceLastMonth = 0: BalanceNow = 0
    IntBal31_3 = 0: IntBalLastMonth = 0: IntBalNow = 0
    LoanID = FormatField(rstMaster("LoanID"))
    IntRate = FormatField(rstMaster("IntRate"))
    SlNo = SlNo + 1
  With grd
    If .Rows < .Row + 2 Then .Rows = .Rows + 2
    .Row = .Row + 1
    .MergeRow(.Row) = False
    .Col = 0: .Text = Format(SlNo, "00")
    .Col = 1: .Text = L_clsCust.CustomerName(FormatField(rstMaster("LoanID")))
    RetStr = FormatField(rstMaster("Guarantor1"))
    On Error Resume Next
    Call GetStringArray(RetStr, StrArr, gDelim)
    .Col = 2: .Text = StrArr(0)
    .Row = .Row + 1: .Text = StrArr(1): .Row = .Row - 1
    RetStr = FormatField(rstMaster("Guarantor2"))
    Call GetStringArray(RetStr, StrArr, gDelim)
    .Col = 3: .Text = StrArr(0)
    .Row = .Row + 1: .Text = StrArr(1): .Row = .Row - 1
    On Error GoTo ErrLine
    'Loan Advance Date
    .Col = 4: .Text = FormatField(rstMaster("IssueDate"))
    .Col = 5: .Text = FormatField(rstMaster("LoanAmount"))
    
    'Loan Balance as on 31/3
    PrevDate = Date31_3
    rstPrin31_3.FindFirst " LoanID = " & LoanID
    If Not rstPrin31_3.NoMatch Then
        rstInt31_3.FindFirst " LoanID = " & LoanID
        Balance31_3 = rstPrin31_3("Balance")
        TransDate = rstInt31_3("TransDate")
        IntBal31_3 = Val(rstInt31_3("IntBalance"))
        IntBal31_3 = IntBal31_3 + DateDiff("d", TransDate, PrevDate) / 365 * IntRate / 100 * Balance
        .Col = 6: .Text = Balance31_3
        .Col = 7: .Text = IntBal31_3
        .Col = 8: .Text = Val(Balance31_3 + IntBal31_3)
        PrevDate = TransDate
    End If
    
    'Loan Balance as on this date
    rstPrinLastMonth.FindFirst " LoanID = " & LoanID
    If Not rstPrinLastMonth.NoMatch Then
        rstIntLastMonth.FindFirst " LoanID = " & LoanID
        BalanceLastMonth = rstPrinLastMonth("Balance")
        TransDate = rstPrinLastMonth("TransDate")
        IntBalLastMonth = Val(rstIntLastMonth("IntBalance"))
        IntBalLastMonth = IntBalLastMonth + DateDiff("d", PrevDate, TransDate) / 365 * IntRate / 100 * Balance
        .Col = 12: .Text = BalanceLastMonth
        .Col = 13: .Text = IntBalLastMonth
        .Col = 14: .Text = Val(BalanceLastMonth + IntBalLastMonth)
        PrevDate = TransDate
        'Recovery as LastMonth
        .Col = 9: .Text = Balance31_3 - BalanceLastMonth
        .Col = 10: .Text = IntBal31_3 - IntBalLastMonth
        .Col = 11: .Text = Val((Balance31_3 - BalanceLastMonth) + (IntBal31_3 - IntBalLastMonth))
        PrevDate = TransDate
    End If
    
    'Recovery as of now
    rstPrinAsOn.FindFirst " LoanID = " & LoanID
    If Not rstPrinAsOn.NoMatch Then
        rstIntAsOn.FindFirst " LoanID = " & LoanID
        BalanceNow = rstPrinLastMonth("Balance")
        TransDate = rstPrinLastMonth("TransDate")
        IntBalNow = Val(rstIntLastMonth("IntBalance"))
        IntBalNow = IntBalNow + DateDiff("d", PrevDate, TransDate) / 365 * IntRate / 100 * Balance
        .Col = 18: .Text = BalanceNow
        .Col = 19: .Text = IntBalNow
        .Col = 20: .Text = Val(BalanceNow + IntBalNow)
        PrevDate = TransDate
        'Recovery as LastMonth
        .Col = 15: .Text = BalanceLastMonth - BalanceNow
        .Col = 16: .Text = IntBalLastMonth - IntBalNow
        .Col = 17: .Text = Val((BalanceLastMonth - BalanceNow) + (IntBalLastMonth - IntBalNow))
        PrevDate = TransDate
    End If
    
  End With
    rstMaster.MoveNext
Loop
Set L_clsLoan = Nothing
ShowMeetingRegistar = True
Screen.MousePointer = vbDefault
ErrLine:
    Screen.MousePointer = vbDefault
    If Err Then
        MsgBox Err.Number & vbCrLf & Err.Description, , wis_MESSAGE_TITLE
       'Resume
        Exit Function
    End If


End Function

Private Sub SetGrid(SchemeID As Integer, Optional PACS As Boolean = False, Optional strYear As String = "YYYY")
Dim Count As Integer
Dim strText As String

With grd
    .Clear
    .AllowUserResizing = flexResizeBoth
    .WordWrap = True
    .FixedCols = 0
    .FixedRows = 0
    .MergeCells = flexMergeNever
    .Col = 0
    .Row = 0: .Text = "Sl No"
    .Row = 1: .Text = "Sl No"
    strText = gBankName '" Name of the BANK "
    .Row = 0
    For Count = 1 To .Cols - 1
        .Col = Count: .Text = strText
        .CellAlignment = 4
    Next
    .Row = 1
    strText = "Meeting Registar"
    For Count = 1 To .Cols - 1
        .Col = Count: .Text = strText
        .CellAlignment = 4
    Next
End With


'Get the Details of the Loan scheme
If SchemeID = 0 Then GoTo CommonSetting

With grd
    .Cols = 28
    .Rows = 10
    .FixedCols = 1
    .FixedRows = 6
    .Row = 2
    .Col = 0: .Text = "Sl No"
    .Col = 1:  .Text = "Name of PACS/PACB"
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
    .Col = 14: .Text = "Repayments From 1/4/" & strYear & " to upto end of last month"
    .Col = 15: .Text = "Repayments From 1/4/" & strYear & " to upto end of last month"
    .Col = 16: .Text = "Repayments From 1/4/" & strYear & " to upto end of last month"
    .Col = 17: .Text = "Repayments From 1/4/" & strYear & " to upto end of last month"
    .Col = 18: .Text = "Loans out standing as on end of last month"
    .Col = 19: .Text = "Loans out standing as on end of last month"
    .Col = 20: .Text = "Loans out standing as on end of last month"
    .Col = 21: .Text = "Loans out standing as on end of last month"
    .Col = 22: .Text = "Of which over due"
    .Col = 23: .Text = "Of which over due"
    .Col = 24: .Text = "Of which over due"
    .Col = 25: .Text = "Of which over due"
    .Col = 26: .Text = "Dates of filing"
    .Col = 27: .Text = "Dates of filing"
    ''3rd Row
    For Count = 0 To .Cols - 1
        .Col = Count
        .Row = 2: strText = .Text
        .Row = 3: .Text = strText
    Next
    .Row = 3
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
    .Col = 14: .Text = "No of Inst"
    .Col = 15: .Text = "Principal"
    .Col = 16: .Text = "Interest"
    .Col = 17: .Text = "Total"
    .Col = 18: .Text = "No of Inst"
    .Col = 19: .Text = "Principal"
    .Col = 20: .Text = "Interest"
    .Col = 21: .Text = "Total"
    .Col = 22: .Text = "No of Inst"
    .Col = 23: .Text = "Principal"
    .Col = 24: .Text = "Interest"
    .Col = 25: .Text = "Total"
    .Col = 26: .Text = "ABN"
    .Col = 27: .Text = "Ep"
    
    
    ''5th Row
    For Count = 0 To .Cols - 1
        .Col = Count
        .Row = 3: strText = .Text
        .Row = 4: .Text = strText
    Next
    .Row = 4
    .Col = 10: .Text = "No of Inst"
    .Col = 11: .Text = "Principal"
    .Col = 12: .Text = "Interest"
    .Col = 13: .Text = "Total"
    ''6th row
    
    .Row = 5
    For Count = 0 To 27
        .Col = Count: .Text = Count
    Next
    .Col = 0: .Text = "1"
    .Col = 1: .Text = "2"
    .Col = 2: .Text = "2a"
    
    .MergeCells = flexMergeRestrictRows
    
    Dim RowCount As Integer
    For RowCount = 0 To 5
        .MergeRow(RowCount) = True
        .Row = RowCount
        For Count = 0 To grd.Cols - 1
            .MergeCol(Count) = True
            .Col = Count
            .CellAlignment = 4: .CellFontBold = True
        Next
    Next
    .Row = 3
    For Count = 14 To 17
        .Col = Count
        .Text = "."
    Next
    For Count = 18 To 21
        .Col = Count
        .Text = ".."
    Next
    For Count = 22 To 25
        .Col = Count
        .Text = "."
    Next
    For Count = 26 To 27
        .Col = Count
        .Text = ".."
    Next
    
End With
Exit Sub

CommonSetting:
With grd
    .Cols = 21
    .Rows = 10
    .FixedCols = 1
    .FixedRows = 5
    .MergeRow(0) = True
    .MergeRow(1) = True
    .MergeRow(2) = True
    .Row = 2: '.MergeRow(2) = True
    .Col = 0: .Text = "Sl No"
    .Col = 1: .Text = "Name of Customer"
    .Col = 2:  .Text = "Guarantor 1"
    .Col = 3:  .Text = "Guarantor 2"
    .Col = 4:  .Text = "Issue Date"
    .Col = 5: .Text = "Loan Amount"
    
    .Col = 6: .Text = "Out standing as 31/3/" & strYear & ""
    .Col = 7: .Text = "Out standing as 31/3/" & strYear & ""
    .Col = 8: .Text = "Out standing as 31/3/" & strYear & ""
    
    '.Col = 6: .Text = "Disbursuments from 1/4/" & strYear & " to upto end of last month"
    .Col = 9: .Text = "Repayments from 1/4/" & strYear & " to up to the end of last month"
    .Col = 10: .Text = "Repayments from 1/4/" & strYear & " to up to the end of last month"
    .Col = 11: .Text = "Repayments from 1/4/" & strYear & " to up to the end of last month"
    .Col = 12: .Text = "Balance OutStanding as on the end of last month"
    .Col = 13: .Text = "Balance OutStanding as on the end of last month"
    .Col = 14: .Text = "Balance OutStanding as on the end of last month"
    
    .Col = 15: .Text = "Repayments during this month"
    .Col = 16: .Text = "Repayments during this month"
    .Col = 17: .Text = "Repayments during this month"
    .Col = 18: .Text = "Balance OutStanding end of this month"
    .Col = 19: .Text = "Balance OutStanding end of this month"
    .Col = 20: .Text = "Balance OutStanding end of this month"
    
    ''2nd row
    .Row = 3: .MergeRow(3) = True
    .Col = 0: .Text = "Sl No"
    .Col = 1: .Text = "Name of Customer"
    .Col = 2:  .Text = "Guarantor 1"
    .Col = 3:  .Text = "Guarantor 2"
    .Col = 4:  .Text = "Issue Date"
    .Col = 5: .Text = "Loan Amount"
    
    '.Col = 6: .Text = "Disbursuments from 1/4/" & strYear & " to upto end of last month"
    .Col = 6: .Text = "Principal"
    .Col = 7: .Text = "Interest"
    .Col = 8: .Text = "Total"
    
    .Col = 9: .Text = "Principal"
    .Col = 10: .Text = "Interest"
    .Col = 11: .Text = "Total"
    .Col = 12: .Text = "Principal"
    .Col = 13: .Text = "Interest"
    .Col = 14: .Text = "Total"
    
    .Col = 15: .Text = "Principal"
    .Col = 16: .Text = "Interest"
    .Col = 17: .Text = "Total"
    .Col = 18: .Text = "Principal"
    .Col = 19: .Text = "Interest"
    .Col = 20: .Text = "Total"
    
    
    .Row = 4: .MergeRow(4) = True
    .MergeCells = flexMergeFree
    For Count = 0 To .Cols - 1
        .Col = Count: .Text = Count
        .CellAlignment = 4
        .MergeCol(Count) = True
        
    Next
    .Col = 2: .Text = "2a"
    .Col = 1: .Text = "2"
    .Col = 0: .Text = "1"
End With
End Sub



Public Function ShowShed5(AsOnIndianDate As String) As Boolean

m_repType = repShedule_5


ShowShed5 = False
Err.Clear
On Error GoTo ExitLine:


Dim SqlStr As String
Dim TransType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes

Dim rstOpBalance As Recordset
Dim rstClBalance As Recordset
Dim rstAdvance As Recordset
Dim rstRecovery As Recordset

Dim AsOnDate As Date
Dim FirstDate As Date

Dim ColAmount() As Currency
Dim GrandTotal() As Currency


If AsOnIndianDate = "" Then
    AsOnDate = gStrDate
Else
    AsOnDate = FormatDate(AsOnIndianDate)
End If
'Get the First day of the Month
FirstDate = Month(AsOnDate) & "/1/" & Year(AsOnDate)
'Fetch Only Cash Credit Loans
Dim LoanType As wisLoanType
LoanType = wisCashCreditLoan

'Get The LoanDetails And THier balance as on Date
SqlStr = "SELECT A.LoanID,LoanAccNo,IssueDate,LoanDueDate,LoanAmount,Balance," & _
    " LoanPurpose, Title+' '+ FirstName+' '+ MiddleName+' '+LastName As Name," & _
    " CustBankName FROM LoanMaster A,LoanTrans B, CustMaster C WHERE A.SchemeId" & _
        " IN (SELECT SchemeId FROM LoanScheme Ls WHERE Ls.LoanType = " & LoanType & ")" & _
    " AND B.BankID = A.BankID AND C.BankID = A.BankID AND B.LoanID = A.LoanID" & _
    " AND C.CustomerId =A.CustomerID AND TransID = (SELECT MAX(TransId) FROM " & _
        " LoanTrans D WHERE D.TransDate <= #" & AsOnDate & "# AND D.BankID = A.BankID" & _
        " AND D.LoanId = A.LoanId ) " & _
    " AND (LoanClosed = NULL OR LOanClosed > #" & AsOnDate & "#) AND A.BankID = " & gBankID

gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then Exit Function
Set rstClBalance = gDbTrans.Rst.Clone

'Get The LoanDetails And THier balance as on first day of the given month
SqlStr = "SELECT A.LoanID,LoanAccNo,IssueDate,LoanDueDate,LoanAmount,Balance," & _
    " CustBankName FROM LoanMaster A,LoanTrans B WHERE B.BankID = A.BankID" & _
    " AND B.LoanID = A.LoanID AND TransID =" & _
        " (SELECT MAX(TransId) FROM LoanTrans D WHERE D.TransDate <" & _
        " #" & FirstDate & "# AND D.BankID = A.BankID AND D.LoanId = A.LoanId)" & _
    " AND A.BankID = " & gBankID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then Exit Function
Set rstOpBalance = gDbTrans.Rst.Clone

'Get The Advances During the MOnth
TransType = wWithdraw
ContraTrans = wContraWithdraw
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans WHERE TransDate >= #" & FirstDate & "#" & _
    " AND TransDate <= #" & AsOnDate & "# AND (TransType = " & TransType & _
    " OR TransType = " & ContraTrans & ") AND BankID = " & gBankID & _
    " GROUP BY LoanID "
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch > 0 Then
    Set rstAdvance = gDbTrans.Rst.Clone
Else
    Set rstAdvance = Nothing
End If


'Get The Recovery During the MOnth
TransType = wDeposit
ContraTrans = wContraDeposit
SqlStr = "SELECT SUM(Amount),LoanID FROM LoanTrans WHERE TransDate >= #" & FirstDate & "#" & _
    " AND TransDate <= #" & AsOnDate & "# AND (TransType = " & TransType & _
    " OR TransType = " & ContraTrans & ") AND BankID = " & gBankID & _
    " GROUP BY LoanID "
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch > 0 Then
    Set rstRecovery = gDbTrans.Rst.Clone
Else
    Set rstRecovery = Nothing
End If

'Now Align the grid
Call Shed5RowCol
ReDim ColAmount(7 To grd.Cols - 1)
ReDim GrandTotal(7 To grd.Cols - 1)
'Now Start to writing to the grid
Dim LoanID As Long
Dim AddRow As Boolean
Dim L_clsLoan As New clsLoan
Dim PrevOD As Currency
Dim ODAmount As Currency
Dim Count As Long
Dim SlNo As Long


SlNo = 0
While Not rstClBalance.EOF
    LoanID = FormatField(rstClBalance("LoanId"))
    
    rstOpBalance.FindFirst "LoanID = " & LoanID
    ColAmount(7) = 0
    If Not rstOpBalance.NoMatch Then
        ColAmount(7) = FormatField(rstOpBalance("Balance")) 'Balance as on 31/3/yyyy
    End If
    ColAmount(8) = 0
    If Not rstAdvance Is Nothing Then
        rstAdvance.FindFirst "LoanId = " & LoanID
        If Not rstAdvance.NoMatch Then
            ColAmount(8) = FormatField(rstAdvance(0))  'Advances During the MOnth
        End If
    End If
    ColAmount(9) = 0
    If Not rstRecovery Is Nothing Then
        rstRecovery.FindFirst "LoanId = " & LoanID
        If Not rstRecovery.NoMatch Then
            ColAmount(9) = FormatField(rstRecovery(0)) 'Recovery During the mOnth
        End If
    End If
    ColAmount(10) = FormatField(rstClBalance("Balance")) 'Balance at the end of month
    ODAmount = L_clsLoan.OverDueAmount(LoanID, , AsOnIndianDate) 'Over due
    ColAmount(11) = ODAmount 'Over due amount of the loan as on given date
    ColAmount(12) = ColAmount(7) + ColAmount(8)  'Maximum O/s Balance
    Debug.Assert SlNo <> 7
    'Over due amount calssifiacation
    PrevOD = 0
    ODAmount = L_clsLoan.OverDueSince(5, LoanID, , AsOnIndianDate)
    ColAmount(18) = ODAmount - PrevOD 'Over due since & above 5 Years
    PrevOD = ColAmount(18) + PrevOD
    ODAmount = L_clsLoan.OverDueSince(4, LoanID, , AsOnIndianDate)
    ColAmount(17) = ODAmount - PrevOD  'Over due since 4 Years
    PrevOD = ColAmount(17) + PrevOD
    ODAmount = L_clsLoan.OverDueSince(3, LoanID, , AsOnIndianDate)
    ColAmount(16) = ODAmount - PrevOD  'Over due since 3 Years
    PrevOD = ColAmount(16) + PrevOD
    ODAmount = L_clsLoan.OverDueSince(2, LoanID, , AsOnIndianDate)
    ColAmount(15) = ODAmount - PrevOD  'Over due since 2 Years
    PrevOD = ColAmount(15) + PrevOD
    ODAmount = L_clsLoan.OverDueSince(1, LoanID, , AsOnIndianDate)
    ColAmount(14) = ODAmount - PrevOD  'Over due a 1 Year
    PrevOD = ColAmount(14) + PrevOD
    ColAmount(13) = ColAmount(11) - PrevOD  'Over due Less than a year
    
    'Check whther this row has to be write or not
    AddRow = False
    For Count = 7 To grd.Cols - 1
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
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = FormatField(rstClBalance("Name"))
        .Col = 2: .Text = FormatField(rstClBalance("CustBankName"))
        .Col = 3: .Text = FormatField(rstClBalance("LoanAmount"))
        .Col = 4: .Text = FormatField(rstClBalance("IssueDate"))
        .Col = 5: .Text = FormatField(rstClBalance("LoanDueDate"))
        .Col = 6: .Text = FormatField(rstClBalance("LoanPurpose"))
        For Count = 7 To grd.Cols - 1
            .Col = Count
            .Text = FormatCurrency(ColAmount(Count))
            GrandTotal(Count) = GrandTotal(Count) + ColAmount(Count)
        Next
      End With
    End If
    rstClBalance.MoveNext
Wend

Set L_clsLoan = Nothing
With grd
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    If .Rows <= .Row + 2 Then .Rows = .Rows + 1
    .Row = .Row + 1
    .Col = 1
    .Text = "Grand Total"
    For Count = 7 To .Cols - 1
        .Col = Count
        .Text = FormatCurrency(GrandTotal(Count))
    Next
End With

ShowShed5 = True
ExitLine:
If Err Then
    MsgBox "ERROR SHED 5" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
End If

End Function

Public Function ShowShed1(AsOnIndianDate As String, Optional RepSocDet As wisRepSocietyLoans = BothLoans) As Boolean



m_repType = repShedule_1

On Error GoTo ErrLine
'Declarations
Dim rstLoan As Recordset
Dim rstRepay  As Recordset
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
Dim ColNUm As Integer

'Decalration By SHashi
Dim ColAmount() As Currency
Dim SubTotal() As Currency
Dim GrandTotal() As Currency
Dim FromDate As Date
Dim LastDate As Date
Dim OBDate As Date
Dim AsOnDate As Date

Dim LExcel As Boolean ' to be removed later and get the data from outside. - pradeep

'LExcel = True
If LExcel Then
    'Set xlWorkBook = Workbooks.Add
    Set xlWorkSheet = xlWorkBook.Sheets(1)
End If

ShowShed1 = False
'Now Get the Last Friday date
If AsOnIndianDate = "" Then
    AsOnIndianDate = FormatDate(gStrDate)
End If

AsOnDate = FormatDate(AsOnIndianDate)
Dim LoanCategary As wisLoanCategories
LoanCategary = wisAgriculural

'GetLast Friday date
'AsOnDate = DateAdd("M", 1, AsOnDate)
'AsOnDate = Month(AsOnDate) & "/1/" & Year(AsOnDate)
''Get the End Of LastMonth
'LastDate = DateAdd("m", -2, AsOnDate)
'Do
'    AsOnDate = DateAdd("d", -1, AsOnDate)
'    If Format(AsOnDate, "dddd") = "Friday" Then Exit Do
'Loop
If LoanCategary = wisAgriculural Then
    OBDate = "7/1/" & IIf(Month(AsOnDate) > 6, Year(AsOnDate), Year(AsOnDate) - 1)
Else
    OBDate = "4/1/" & IIf(Month(AsOnDate) > 3, Year(AsOnDate), Year(AsOnDate) - 1)
End If
'This Report Includes Only agricultural loans so
OBDate = "7/1/" & IIf(Month(AsOnDate) > 6, Year(AsOnDate), Year(AsOnDate) - 1)


' To get the Branch Name to print in the form
SqlStr = " select BankName from BankDet where BankId = " & gBankID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then
    Exit Function
Else
    strBranchName = FormatField(gDbTrans.Rst(0))
End If
frmGrids.lblBranchName = strBranchName

'Get The Details Of Agri Loans, and Balance as ondate

SqlStr = "SELECT A.LoanId, LoanAccNo, Balance, CustBankName, " & _
    " Title + ' ' + FirstName +' ' + MiddleName + ' ' + LastName As Name " & _
    " From LoanMaster A,LoanTrans B,CustMaster C WHERE A.SchemeID IN" & _
        " (SELECT SchemeID From LoanScheme WHERE Category = " & LoanCategary & ")" & _
    " AND A.BankID = B.BankID AND C.BankId = A.BankID AND A.LoanID = B.LoanID " & _
    " AND C.CustomerID = A.CustomerId AND TransID = (SELECT MAX(TransID) " & _
        " From LoanTrans D Where A.BankId = A.BankID AND D.LoanId = A.LoanId" & _
        " AND D.TransDate <= #" & AsOnDate & "#)" & _
    " AND A.BankID = " & gBankID

'Check the whether Loan is Throu society or NOt
If RepSocDet = DirectLoans Then
    SqlStr = SqlStr & " AND CustBankID = 0"
ElseIf RepSocDet = SocietyThruLoans Then
    SqlStr = SqlStr & " AND CustBankID <> 0"
End If
    
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then Exit Function
Set rstLoan = gDbTrans.Rst.Clone

'Now fix the headers for the shedule
If Not Shed1RowCOl Then
    MsgBox "Could Not Set the Headers"
    Exit Function
End If


'Now Get the Loan RepayMent Of the during this period
Dim TransType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes
TransType = wDeposit
ContraTrans = wContraDeposit
SqlStr = "SELECT SUM(Amount),LoanID,BankID FROM LoanTrans Where TransType = " & TransType & _
    " OR TransType = " & ContraTrans & " AND TransDate >= #" & OBDate & "# " & _
    " AND TransDate <= #" & AsOnDate & "# AND BankId = " & gBankID & _
    " GROUP BY LoanID,BankID"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch > 0 Then
    Set rstRepay = gDbTrans.Rst.Clone
Else
    Set rstRepay = Nothing
End If
'Now Get the Loan Advance Of the during this period
TransType = wWithdraw
ContraTrans = wContraWithdraw
SqlStr = "SELECT SUM(Amount),LoanID,BankID FROM LoanTrans Where TransType = " & TransType & _
    " OR TransType = " & ContraTrans & " AND TransDate >= #" & OBDate & "# " & _
    " AND TransDate <= #" & AsOnDate & "# AND BankId = " & gBankID & _
    " GROUP BY LoanID,BankID"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch > 0 Then
    Set rstAdvance = gDbTrans.Rst.Clone
Else
    Set rstAdvance = Nothing
End If


grd.Row = 3
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
OBIndiandate = FormatDate(CStr(OBDate))

SlNo = 0
While Not rstLoan.EOF
    'Initialise the Varible
    LoanID = FormatField(rstLoan("LoanID"))
    'ColAmount(0) = FormatField(rstLoan("Balance"))
    ColAmount(3) = L_clsLoan.OverDueAmount(LoanID, , OBIndiandate)
    ColAmount(4) = L_clsLoan.OverDueDuring(LoanID, , OBIndiandate, AsOnIndianDate)
    ColAmount(5) = ColAmount(3) + ColAmount(4)
    curRepay = 0: ' curAdvance = 0
    If Not rstRepay Is Nothing Then
        rstRepay.FindFirst "LoanID = " & LoanID & " AND BankID = " & gBankID
        If Not rstRepay.NoMatch Then
            curRepay = FormatField(rstRepay(0))
        End If
    End If
    If Not rstAdvance Is Nothing Then
        rstRepay.FindFirst "LoanID = " & LoanID & " AND BankID = " & gBankID
        If Not rstAdvance.NoMatch Then
'            curAdvance = FormatField(rstAdvance(0))
        End If
    End If
    ColAmount(9) = FormatCurrency(curRepay)
    ColAmount(6) = 0: ColAmount(7) = 0: ColAmount(8) = 0
    If curRepay > 0 Then
        If ColAmount(3) >= curRepay Then
            ColAmount(6) = curRepay: curRepay = 0
        Else
            ColAmount(6) = ColAmount(3): curRepay = curRepay - ColAmount(6)
        End If
        If ColAmount(4) >= curRepay Then
            ColAmount(7) = curRepay: curRepay = 0
        Else
            ColAmount(7) = ColAmount(4): curRepay = curRepay - ColAmount(7)
        End If
        ColAmount(8) = curRepay
    End If
    ColAmount(10) = ColAmount(5) - ColAmount(6) - ColAmount(7)
    
    'Over due amount
    PrevOD = 0
    ODAmount = L_clsLoan.OverDueSince(6, LoanID, , AsOnIndianDate)
    ColAmount(16) = ODAmount: PrevOD = ODAmount
    ODAmount = L_clsLoan.OverDueSince(5, LoanID, , AsOnIndianDate)
    ColAmount(15) = ODAmount - PrevOD: PrevOD = ODAmount
    ODAmount = L_clsLoan.OverDueSince(5, LoanID, , AsOnIndianDate)
    ColAmount(14) = ODAmount - PrevOD: PrevOD = ODAmount
    ODAmount = L_clsLoan.OverDueSince(5, LoanID, , AsOnIndianDate)
    ColAmount(13) = ODAmount - PrevOD: PrevOD = ODAmount
    ODAmount = L_clsLoan.OverDueSince(5, LoanID, , AsOnIndianDate)
    ColAmount(12) = ODAmount - PrevOD: PrevOD = ODAmount
    ODAmount = L_clsLoan.OverDueAmount(LoanID, , AsOnIndianDate)
    ColAmount(11) = ODAmount - PrevOD: PrevOD = ODAmount
    AddRow = False
    For Count = 3 To grd.Cols - 1
        If ColAmount(Count) Then
            AddRow = True
            Exit For
        End If
    Next
    If AddRow Then
        If grd.Rows <= grd.Row + 2 Then grd.Rows = grd.Row + 1
        grd.Row = grd.Row + 1
        SlNo = SlNo + 1
        grd.Col = 0: grd.Text = SlNo
        grd.Col = 1: grd.Text = FormatField(rstLoan("name"))
        grd.Col = 2: grd.Text = FormatField(rstLoan("CustBankName"))
        For Count = 3 To grd.Cols - 1
            grd.Col = Count: grd.Text = FormatCurrency(ColAmount(Count))
        Next
    End If
    rstLoan.MoveNext
Wend
Set L_clsLoan = Nothing
frmGrids.lblFrom = "Shedule 1"
ShowShed1 = True
Exit Function

ErrLine:
    MsgBox "error Showshed1" & vbCrLf & Err.Description, vbInformation, wis_MESSAGE_TITLE
  'Resume
End Function

Public Function ShowShed2(AsOnIndianDate As String, Optional RepSocDet As wisRepSocietyLoans = BothLoans) As Boolean


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
Dim Count As Integer

Dim SlNo As Integer
Dim strBranchName As String


'Decalration By SHashi
Dim ColAmount() As Currency
Dim GrandTotal() As Currency
Dim FromDate As Date
Dim LastDate As Date
Dim AsOnDate As Date
Dim OBDate As Date
Dim LoanCategary As wisLoanCategories

Dim LExcel As Boolean ' to be removed later and get the data from outside. - pradeep

'LExcel = True
If LExcel Then
    'Set xlWorkBook = Workbooks.Add
    Set xlWorkSheet = xlWorkBook.Sheets(1)
End If

ShowShed2 = False
'Now Get the Last Friday date
If AsOnIndianDate = "" Then
    AsOnIndianDate = FormatDate(gStrDate)
End If

AsOnDate = FormatDate(AsOnIndianDate)
LoanCategary = wisAgriculural
LastDate = Month(AsOnDate) & "/1/" & Year(AsOnDate)
LastDate = DateAdd("d", -1, LastDate)
'GetLast Friday date
'AsOnDate = DateAdd("M", 1, AsOnDate)
'AsOnDate = Month(AsOnDate) & "/1/" & Year(AsOnDate)
''Get the End Of LastMonth
'LastDate = DateAdd("m", -2, AsOnDate)
'Do
'    AsOnDate = DateAdd("d", -1, AsOnDate)
'    If Format(AsOnDate, "dddd") = "Friday" Then Exit Do
'Loop
If LoanCategary = wisAgriculural Then
    OBDate = "7/1/" & IIf(Month(AsOnDate) > 6, Year(AsOnDate), Year(AsOnDate) - 1)
Else
    OBDate = "4/1/" & IIf(Month(AsOnDate) > 3, Year(AsOnDate), Year(AsOnDate) - 1)
End If

' To get the Branch Name to print in the form
SqlStr = " select BankName from BankDet where BankId = " & gBankID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then
    Exit Function
Else
    strBranchName = FormatField(gDbTrans.Rst(0))
End If
frmGrids.lblBranchName = strBranchName

'Get The Details Of Loans, and Balance as ondate
SqlStr = "SELECT A.LoanId, LoanAccNo, Balance, CustBankName, " & _
    " Title + ' ' + FirstName +' ' + MiddleName + ' ' + LastName As Name " & _
    " From LoanMaster A,LoanTrans B,CustMaster C WHERE A.Schemeid IN " & _
        " (SELECT SchemeID From LoanScheme WHERE Category = " & LoanCategary & ")" & _
    "  AND A.BankID = B.BankID AND C.BankId = A.BankID AND A.LoanID = B.LoanID" & _
    " AND C.CustomerID = A.CustomerId AND TransID = (SELECT MAX(TransID) " & _
        " From LoanTrans D Where D.BankId = A.BankID AND D.LoanId = A.LoanId" & _
        " AND D.TransDate <= #" & AsOnDate & "#)" & _
    " And A.BankID = " & gBankID & " ORDER BY A.LoanID"
'Check the whether Loan is Throu society or NOt
If RepSocDet = DirectLoans Then
    SqlStr = SqlStr & " AND CustBankID = 0"
ElseIf RepSocDet = SocietyThruLoans Then
    SqlStr = SqlStr & " AND CustBankID <> 0"
End If

gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then Exit Function
Set rstCurrBalance = gDbTrans.Rst.Clone

'Get The Details Of Loans, and Balance as on 31/M/yyyy
SqlStr = "SELECT A.LoanId, LoanAccNo, Balance, CustBankName, TransDate," & _
    " Title + ' ' + FirstName +' ' + MiddleName + ' ' + LastName As Name " & _
    " From LoanMaster A,LoanTrans B,CustMaster C WHERE A.BankID = B.BankID" & _
    " AND C.BankId = A.BankID AND A.LoanID = B.LoanID AND C.CustomerID = A.CustomerId" & _
    " AND TransID = (SELECT MAX(TransID) From LoanTrans D Where D.BankId = A.BankID " & _
        " AND D.LoanId = A.LoanId AND D.TransDate < #" & OBDate & "#)" & _
    " And A.BankID = " & gBankID
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch < 1 Then Exit Function
Set rstOpBalance = gDbTrans.Rst.Clone

Dim TransType As wisTransactionTypes
Dim ContraTrans As wisTransactionTypes
'Now Get the Loan Repayment From Ob Date to Last Month
TransType = wDeposit
ContraTrans = wContraDeposit
SqlStr = "SELECT SUM(Amount),LoanID,BankID FROM LoanTrans Where TransDate >= #" & OBDate & "# " & _
    " AND TransDate <= #" & LastDate & "# AND (TransType = " & TransType & _
    " OR TransType = " & ContraTrans & ") AND BankId = " & gBankID & _
    " GROUP BY LoanID,BankID"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch > 0 Then
    Set rstRepayLast = gDbTrans.Rst.Clone
Else
    Set rstRepayLast = Nothing
End If
'Now Get the Loan Repayment From Lastdate to till date
SqlStr = "SELECT SUM(Amount),LoanID,BankID FROM LoanTrans Where TransDate > #" & LastDate & "# " & _
    " AND TransDate <= #" & AsOnDate & "# AND (TransType = " & TransType & _
    " OR TransType = " & ContraTrans & ") AND BankId = " & gBankID & _
    " GROUP BY LoanID,BankID"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch > 0 Then
    Set rstRepayTill = gDbTrans.Rst.Clone
Else
    Set rstRepayTill = Nothing
End If
'Now Get the Loan Advance From Ob date to last month
TransType = wWithdraw
ContraTrans = wContraWithdraw
SqlStr = "SELECT SUM(Amount),LoanID,BankID FROM LoanTrans Where TransDate >= #" & OBDate & "# " & _
    " AND TransDate <= #" & LastDate & "# AND (TransType = " & TransType & _
    " OR TransType = " & ContraTrans & ") AND BankId = " & gBankID & _
    " GROUP BY LoanID,BankID"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch > 0 Then
    Set rstAdvanceLast = gDbTrans.Rst.Clone
Else
    Set rstAdvanceLast = Nothing
End If
'Now Get the Loan Advance during this month
SqlStr = "SELECT SUM(Amount),LoanID,BankID FROM LoanTrans Where TransDate > #" & LastDate & "# " & _
    " AND TransDate <= #" & AsOnDate & "# AND (TransType = " & TransType & _
    " OR TransType = " & ContraTrans & ") AND BankId = " & gBankID & _
    " GROUP BY LoanID,BankID"
gDbTrans.SqlStmt = SqlStr
If gDbTrans.SQLFetch > 0 Then
    Set rstAdvanceTill = gDbTrans.Rst.Clone
Else
    Set rstAdvanceTill = Nothing
End If


SlNo = 0
grd.Visible = True

' now set the shed headers
Call Shed2RowCol

grd.Row = 2
' beginning of the loantrans loop
ReDim ColAmount(3 To grd.Cols - 1)
ReDim SubTotal(3 To grd.Cols - 1)
ReDim GrandTotal(3 To grd.Cols - 1)

' label declarations
Dim LoanID As Long
Dim AddRow As Boolean
Dim L_clsLoan As New clsLoan

While Not rstCurrBalance.EOF
    LoanID = FormatField(rstCurrBalance("LoanID"))
    ColAmount(3) = 0
    If Not rstOpBalance Is Nothing Then
        rstOpBalance.FindFirst "LoanID = " & LoanID
        If Not rstOpBalance.NoMatch Then
            ColAmount(3) = FormatField(rstOpBalance("Balance")) ' Opening Balance
        End If
    End If
    ColAmount(4) = 0
    If Not rstAdvanceLast Is Nothing Then
        rstAdvanceLast.FindFirst "LoanId = " & LoanID
        If Not rstAdvanceLast.NoMatch Then
            ColAmount(4) = FormatField(rstAdvanceLast(0)) ' Advanced up to Previous month
        End If
    End If
    ColAmount(5) = 0
    If Not rstAdvanceTill Is Nothing Then
        rstAdvanceTill.FindFirst "LoanId = " & LoanID
        If Not rstAdvanceTill.NoMatch Then
            ColAmount(5) = FormatField(rstAdvanceTill(0)) ' Advanced during the month
        End If
    End If
    ColAmount(6) = ColAmount(4) + ColAmount(5)  'Toal Loan Advance upto en of month
    ColAmount(7) = ColAmount(3) + ColAmount(6) 'Maxmum LOan BAlance at the End of month
    
    ColAmount(8) = 0
    If Not rstRepayLast Is Nothing Then
        rstRepayLast.FindFirst "LoanId = " & LoanID
        If Not rstRepayLast.NoMatch Then
            ColAmount(8) = FormatField(rstRepayLast(0)) ' Recovery up to Previous month
        End If
    End If
    ColAmount(9) = 0
    If Not rstRepayTill Is Nothing Then
        rstRepayTill.FindFirst "LoanId = " & LoanID
        If Not rstRepayTill.NoMatch Then
            ColAmount(9) = FormatField(rstRepayTill(0)) ' Recovery during this month
        End If
    End If
    ColAmount(10) = ColAmount(8) + ColAmount(9)  'Total recovery at the end of month
    ColAmount(11) = ColAmount(7) - ColAmount(10) 'Balance at the end of month
    
    Debug.Assert ColAmount(11) = FormatField(rstCurrBalance("Balance"))
    ColAmount(12) = L_clsLoan.OverDueAmount(LoanID, , AsOnIndianDate)
    
    AddRow = False
    For Count = 3 To grd.Cols - 1
        If ColAmount(Count) Then
            AddRow = True
            Exit For
        End If
    Next
    If AddRow Then
      With grd
        If .Rows <= .Row + 2 Then .Rows = .Row + 2
        .Row = .Row + 1: SlNo = SlNo + 1
        .Col = 0: .Text = SlNo
        .Col = 1: .Text = FormatField(rstCurrBalance("Name"))
        .Col = 2: .Text = FormatField(rstCurrBalance("CustBankName"))
        For Count = 3 To grd.Cols - 1
            grd.Col = Count: grd.Text = FormatCurrency(ColAmount(Count))
            GrandTotal(Count) = GrandTotal(Count) + ColAmount(Count)
        Next
      End With
    End If
    rstCurrBalance.MoveNext
Wend

Set L_clsLoan = Nothing

lblBranchName.Caption = strBranchName
lblFrom.Caption = "Shedule 2"
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


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Call ShowShed1("")
End Sub

 
Private Sub cmdPrint_Click()
Debug.Assert Me.ActiveControl.Name <> cmdPrint.Name

End Sub

Private Sub Form_Resize()
Const MARGIN = 50
Const CTL_MARGIN = 50
'Const BOTTOM_MARGIN = 600
On Error Resume Next

    lblFrom.Top = 0
    lblFrom.Left = (Me.Width - lblFrom.Width) / 2
    grd.Left = 0
    lblBranch.Top = lblFrom.Top + lblFrom.Height + 50
    lblBankname.Top = lblBranch.Top + lblBranch.Height + 50
    lblBankname.Left = (Me.Width - lblBankname.Width) / 2
    lblBranch.Left = lblBankname.Left
    grd.Top = lblBankname.Top + lblBankname.Height + 200
    grd.Width = Me.Width - 150
    grd.Height = Me.ScaleHeight - (lblFrom.Height + lblBranch.Height + lblBankname.Height + cmdCancel.Height + 370)
    
    With cmdCancel
        .Left = Me.ScaleWidth - MARGIN - .Width
        .Top = grd.Top + grd.Height + CTL_MARGIN 'Me.ScaleHeight - MARGIN - .Height
    End With
    With cmdPrint
        .Left = cmdCancel.Left - CTL_MARGIN - .Width
        .Top = cmdCancel.Top
    End With
   With cmdOK
        .Left = cmdPrint.Left - CTL_MARGIN - .Width
        .Top = cmdPrint.Top
    End With
    
    If Me.Height < cmdCancel.ToolTipText + cmdCancel.Height + CTL_MARGIN Then
        Me.Height = cmdCancel.ToolTipText + cmdCancel.Height + CTL_MARGIN * 2
    End If
    ' removed the call for personal use - pradeep
    'Call InitGrid(True)
End Sub


Private Sub Form_Unload(Cancel As Integer)

Dim Count  As Integer
 For Count = 0 To grd.Cols - 1
            Call SaveSetting(App.EXEName, "ColWidth", "Report" & m_ReportTpe, grd.ColWidth(Count))
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


