VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C7627F52-2756-11D6-9FFE-0080AD7C8DF9}#3.0#0"; "GrdPrint.ocx"
Begin VB.Form wisMainTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "x"
   ClientHeight    =   6525
   ClientLeft      =   1860
   ClientTop       =   1815
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   28
      Top             =   6090
      Width           =   1995
   End
   Begin WIS_GRID_Print.GridPrint grdPrint 
      Left            =   120
      Top             =   6030
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   -30
      Top             =   4530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame"
      Height          =   5805
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   7095
      Begin VB.CommandButton cmdReports 
         Caption         =   "&Reports"
         Height          =   375
         Left            =   4860
         TabIndex        =   29
         Top             =   4800
         Width           =   1995
      End
      Begin VB.CommandButton cmdComp 
         Caption         =   "Company(self) Details"
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   4800
         Width           =   1995
      End
      Begin VB.CommandButton cmdCompany 
         Caption         =   "Create  Company"
         Height          =   375
         Left            =   2640
         TabIndex        =   26
         Top             =   4800
         Width           =   1995
      End
      Begin VB.CommandButton cmdTrnasfer 
         Caption         =   "Stock Transfer"
         Height          =   375
         Left            =   4860
         TabIndex        =   25
         Top             =   3675
         Width           =   1995
      End
      Begin VB.CommandButton cmdInvoiceDet 
         Caption         =   "Invoice Detals"
         Height          =   375
         Left            =   4890
         TabIndex        =   24
         Top             =   4245
         Width           =   1995
      End
      Begin VB.CommandButton cmdGroup 
         Caption         =   "Create Group"
         Height          =   405
         Left            =   4860
         TabIndex        =   23
         Top             =   240
         Width           =   1995
      End
      Begin VB.CommandButton cmdUnit 
         Caption         =   "Create Units"
         Height          =   405
         Left            =   4860
         TabIndex        =   22
         Top             =   1380
         Width           =   1995
      End
      Begin VB.CommandButton cmdProdProp 
         Caption         =   "Produt properties"
         Height          =   405
         Left            =   4860
         TabIndex        =   21
         Top             =   1965
         Width           =   1995
      End
      Begin VB.CommandButton cmdPurchase 
         Caption         =   "Purchase"
         Height          =   405
         Left            =   4860
         TabIndex        =   20
         Top             =   2535
         Width           =   1995
      End
      Begin VB.CommandButton cmdItem 
         Caption         =   "Cteate Item"
         Height          =   405
         Left            =   4860
         TabIndex        =   19
         Top             =   810
         Width           =   1995
      End
      Begin VB.CommandButton cmdSales 
         Caption         =   "Sales"
         Height          =   405
         Left            =   4860
         TabIndex        =   18
         Top             =   3105
         Width           =   1995
      End
      Begin VB.CommandButton cmdCompareDB 
         Caption         =   "Compare DataBase"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   5340
         Width           =   1995
      End
      Begin VB.CommandButton cmdMaterial 
         Caption         =   "Material"
         Height          =   375
         Left            =   2625
         TabIndex        =   16
         Top             =   4245
         Width           =   1995
      End
      Begin VB.CommandButton cmdBankAcc 
         Caption         =   "Bank Accounts"
         Height          =   375
         Left            =   2625
         TabIndex        =   15
         Top             =   3675
         Width           =   1995
      End
      Begin VB.CommandButton cmdPigmyAgent 
         Caption         =   "Add Pigmy Agents"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   3675
         Width           =   1995
      End
      Begin VB.CommandButton cmdContra 
         Caption         =   "Contra"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   4245
         Width           =   1995
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Loan Reports"
         Height          =   405
         Left            =   2625
         TabIndex        =   12
         Top             =   3105
         Width           =   1995
      End
      Begin VB.CommandButton cmdBKCC 
         Caption         =   "B&KCC Loans"
         Height          =   405
         Left            =   2625
         TabIndex        =   11
         Top             =   810
         Width           =   1995
      End
      Begin VB.CommandButton cmdPD 
         Caption         =   "&Pigmy deposits"
         Height          =   405
         Left            =   360
         TabIndex        =   10
         Top             =   3105
         Width           =   1995
      End
      Begin VB.CommandButton cmdRD 
         Caption         =   "RD Deposits"
         Height          =   405
         Left            =   360
         TabIndex        =   9
         Top             =   2535
         Width           =   1995
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Loan Accounts"
         Height          =   405
         Left            =   2625
         TabIndex        =   8
         Top             =   2535
         Width           =   1995
      End
      Begin VB.CommandButton cmdFd 
         Caption         =   "Fixed Deposits"
         Height          =   405
         Left            =   360
         TabIndex        =   7
         Top             =   1965
         Width           =   1995
      End
      Begin VB.CommandButton cmdLnCreate 
         Caption         =   "Loan creation"
         Height          =   405
         Left            =   2625
         TabIndex        =   6
         Top             =   1965
         Width           =   1995
      End
      Begin VB.CommandButton cmdCA 
         Caption         =   "CA module"
         Height          =   405
         Left            =   360
         TabIndex        =   5
         Top             =   1380
         Width           =   1995
      End
      Begin VB.CommandButton cmdLnScheme 
         Caption         =   "LoanScheme"
         Height          =   405
         Left            =   2625
         TabIndex        =   4
         Top             =   1380
         Width           =   1995
      End
      Begin VB.CommandButton cmdSb 
         Caption         =   "&SB Module"
         Height          =   405
         Left            =   360
         TabIndex        =   3
         Top             =   810
         Width           =   1995
      End
      Begin VB.CommandButton cmdDepLoans 
         Caption         =   "&Deposit Loans"
         Height          =   405
         Left            =   2625
         TabIndex        =   2
         Top             =   240
         Width           =   1995
      End
      Begin VB.CommandButton cmdMem 
         Caption         =   "&MemberModule"
         Height          =   405
         Left            =   390
         TabIndex        =   1
         Top             =   240
         Width           =   1995
      End
   End
End
Attribute VB_Name = "wisMainTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_clsobject As Object
Private Sub Command9_Click()

End Sub


Private Sub Command9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub



Private Sub cmdBankAcc_Click()
Dim AccTrans As clsAccTrans

Set AccTrans = New clsAccTrans

AccTrans.ShowAccTrans
Set AccTrans = Nothing

End Sub

Private Sub cmdBKCC_Click()
Set m_clsobject = New clsBkcc
m_clsobject.Show
End Sub

Private Sub cmdCA_Click()
Set m_clsobject = New clsCAAcc
m_clsobject.Show
End Sub

Private Sub cmdClose_Click()
Set m_clsobject = Nothing
Unload Me
End
End Sub

Private Sub cmdComp_Click()
frmCompanyDetails.Show 1

End Sub

Private Sub cmdCompany_Click()
frmCreateCompany.Show 1
End Sub

Private Sub cmdCompareDB_Click()
Dim DBUtilClass As clsDBUtilities

If MsgBox("This will Compare the DataBase" & _
        vbCrLf & "Are you sure do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton2, wis_MESSAGE_TITLE) = vbNo Then Exit Sub

'Compact the database
Set DBUtilClass = New clsDBUtilities

Call DBUtilClass.CompareDBFromDB(App.Path & "\BlankDataBase\" & constDBName, constDBPWD)

Set DBUtilClass = Nothing

MsgBox "Database Compared", vbInformation

End Sub

Private Sub cmdContra_Click()
Set m_clsobject = New clsContra
m_clsobject.Show
End Sub

Private Sub cmdDepLoans_Click()
Set m_clsobject = New clsDepLoan
m_clsobject.Show

End Sub

Private Sub cmdFd_Click()
Set m_clsobject = New clsFDAcc
m_clsobject.Show
End Sub

Private Sub cmdGroup_Click()
Dim Groupclass As New clsAddGroup

Groupclass.ShowAddGroup (grpProduct)

End Sub

Private Sub cmdInvoiceDet_Click()
Dim invClass As clsInvoiceDet
Set invClass = New clsInvoiceDet
invClass.Show
Set invClass = Nothing

End Sub

Private Sub cmdItem_Click()
frmCreateItem.Show
End Sub


Private Sub cmdLnCreate_Click()
Set m_clsobject = New clsLoan
m_clsobject.ShowCreateLoanAccount
End Sub

Private Sub cmdLnScheme_Click()

Set m_clsobject = New clsLoan
m_clsobject.ShowLoanSchemes

End Sub

Private Sub cmdMem_Click()
Set m_clsobject = New clsMMAcc
m_clsobject.Show

End Sub


Private Sub cmdPD_Click()
Set m_clsobject = New clsPDAcc

m_clsobject.Show
End Sub

Private Sub cmdPigmyAgent_Click()

gCurrUser.ShowUserDialog
End Sub

Private Sub cmdProdProp_Click()
Dim MatClass As clsMaterial
'Dim CompType As wis_CompanyType

'CompType = Enum_Manufacturer
Set MatClass = New clsMaterial

Dim HeadID As Long
HeadID = MatClass.GetHeadIDFromHeadsList(Enum_Manufacturer)

If HeadID = 0 Then Exit Sub
With frmProductPropertyNew
    .lblCompanyName = MatClass.GetCompanyName(HeadID)
    .Show vbModal
End With

End Sub

Private Sub cmdPurchase_Click()
Dim PurClass As clsPurchase
Dim MatClass As clsMaterial

Dim HeadID As Long
Set MatClass = New clsMaterial

HeadID = MatClass.GetHeadID
'If HeadID = 0 Then Exit Sub

Set PurClass = New clsPurchase

PurClass.VendorID = HeadID
PurClass.Show

Set PurClass = Nothing
Set MatClass = Nothing

End Sub


'
Private Sub cmdRD_Click()
Set m_clsobject = New clsRDAcc
m_clsobject.Show
End Sub

Private Sub cmdReport_Click()
frmRepTemp.Show 1
End Sub

Private Sub cmdReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Debug.Print "DOWNNN"
End Sub

Private Sub cmdReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Debug.Print "UPPP"

End Sub


Private Sub cmdReports_Click()
Call cmdReport_Click
End Sub

Private Sub cmdSales_Click()
Dim SaleClass As clsSales
Dim MatClass As clsMaterial

Dim HeadID As Long
Set SaleClass = New clsSales
Set MatClass = New clsMaterial

HeadID = MatClass.GetHeadIDFromHeadsList(Enum_Customers)


If HeadID = 0 Then Exit Sub

SaleClass.VendorID = HeadID
SaleClass.Show

Set SaleClass = Nothing
Set MatClass = Nothing

End Sub

'
Private Sub cmdSb_Click()
Set m_clsobject = New clsSBAcc
m_clsobject.Show

End Sub




Private Sub cmdTrnasfer_Click()
Dim TfrClass As clsTransferNew
Set TfrClass = New clsTransferNew

TfrClass.Show

Set TfrClass = Nothing

End Sub

Private Sub cmdUnit_Click()
Dim UnitClass As New clsAddGroup
UnitClass.ShowAddGroup (grpUnit)
Set UnitClass = Nothing

End Sub

'
Private Sub Command1_Click()
Set m_clsobject = Nothing
Set m_clsobject = New clsLoan
m_clsobject.ShowLoanReport
End Sub
'
'
Private Sub Command8_Click()
Set m_clsobject = New clsLoan
m_clsobject.ShowLoanAccountDetail

End Sub


Private Sub frmAbn_Click()
 
End Sub


