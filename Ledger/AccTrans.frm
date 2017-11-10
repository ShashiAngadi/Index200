VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmAccTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ledger Entry"
   ClientHeight    =   6180
   ClientLeft      =   450
   ClientTop       =   2115
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBox 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6720
      TabIndex        =   50
      Top             =   2160
      Width           =   1035
   End
   Begin VB.TextBox txtBefore 
      Height          =   285
      Left            =   5280
      TabIndex        =   49
      Top             =   2580
      Width           =   465
   End
   Begin VB.TextBox txtAfter 
      Height          =   285
      Left            =   6270
      ScrollBars      =   1  'Horizontal
      TabIndex        =   48
      Top             =   2580
      Width           =   435
   End
   Begin MSFlexGridLib.MSFlexGrid grdLedger 
      Height          =   4995
      Index           =   0
      Left            =   5100
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   300
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   8811
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid grdLedger 
      Height          =   4755
      Index           =   2
      Left            =   5160
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   360
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   8387
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   450
      Left            =   8490
      TabIndex        =   39
      Top             =   5520
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Clear"
      Height          =   450
      Index           =   0
      Left            =   9750
      TabIndex        =   21
      Top             =   5520
      Width           =   1155
   End
   Begin VB.TextBox txtParticulars 
      Height          =   825
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4470
      Width           =   4875
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   450
      Index           =   1
      Left            =   11100
      TabIndex        =   22
      Top             =   5520
      Width           =   1065
   End
   Begin MSFlexGridLib.MSFlexGrid grdLedger 
      Height          =   4875
      Index           =   1
      Left            =   5100
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   360
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   8599
      _Version        =   393216
   End
   Begin VB.Frame fraTab 
      Height          =   3825
      Index           =   0
      Left            =   150
      TabIndex        =   25
      Top             =   480
      Width           =   4800
      Begin VB.CommandButton cmdCBStock 
         Caption         =   ".."
         Height          =   405
         Left            =   4330
         TabIndex        =   12
         Top             =   2322
         Visible         =   0   'False
         Width           =   400
      End
      Begin VB.ComboBox cmbTab0Ledger 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1650
         TabIndex        =   8
         Top             =   1266
         Width           =   2955
      End
      Begin VB.ComboBox cmbTab0Ledger 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1650
         TabIndex        =   10
         Top             =   1794
         Width           =   2925
      End
      Begin VB.TextBox txtTab0CurrentDate 
         Height          =   330
         Left            =   1650
         TabIndex        =   5
         Top             =   768
         Width           =   2415
      End
      Begin VB.ComboBox cmbTab0Ledger 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1650
         TabIndex        =   3
         Top             =   240
         Width           =   2985
      End
      Begin VB.OptionButton optTab0From 
         Caption         =   "From"
         Enabled         =   0   'False
         Height          =   435
         Index           =   0
         Left            =   150
         TabIndex        =   16
         Top             =   2250
         Width           =   855
      End
      Begin VB.OptionButton optTab0From 
         Caption         =   "To"
         Enabled         =   0   'False
         Height          =   465
         Index           =   1
         Left            =   1200
         TabIndex        =   17
         Top             =   2250
         Width           =   1185
      End
      Begin VB.CommandButton cmdTab0AddtoGrid 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   400
         Index           =   1
         Left            =   2550
         TabIndex        =   19
         Top             =   3360
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdTab0AddtoGrid 
         Caption         =   "Add"
         Height          =   400
         Index           =   0
         Left            =   1440
         TabIndex        =   18
         Top             =   3360
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Save"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   400
         Left            =   3660
         TabIndex        =   20
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdTab0Date 
         Caption         =   ".."
         Height          =   315
         Left            =   4140
         TabIndex        =   6
         Top             =   768
         Width           =   465
      End
      Begin WIS_Currency_Text_Box.CurrText txtTab0Amount 
         Height          =   345
         Left            =   2550
         TabIndex        =   14
         Top             =   2850
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblTab0Balance 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   360
         Index           =   0
         Left            =   2580
         TabIndex        =   11
         Top             =   2322
         Width           =   1635
      End
      Begin VB.Label lbltab0Amount 
         Caption         =   "Amount"
         Height          =   300
         Index           =   4
         Left            =   930
         TabIndex        =   13
         Top             =   2880
         Width           =   1125
      End
      Begin VB.Label lbltab0Date 
         Caption         =   "Date"
         Height          =   300
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label lblTab0AccountType 
         Caption         =   "Ledger Name"
         Height          =   435
         Index           =   3
         Left            =   60
         TabIndex        =   9
         Top             =   1800
         Width           =   1365
      End
      Begin VB.Label lblTab0HeadType 
         Caption         =   "Ledger Type"
         Height          =   435
         Index           =   2
         Left            =   60
         TabIndex        =   7
         Top             =   1290
         Width           =   1365
      End
      Begin VB.Label lblTab0VoucherType 
         Caption         =   "Voucher Type"
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.Frame fraTab 
      Height          =   3825
      Index           =   1
      Left            =   150
      TabIndex        =   26
      Top             =   480
      Width           =   4800
      Begin VB.CommandButton cmdTab1Ledger 
         Caption         =   ".."
         Height          =   315
         Index           =   2
         Left            =   4350
         TabIndex        =   32
         Top             =   870
         Width           =   405
      End
      Begin VB.CommandButton cmdTab1Ledger 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   4380
         TabIndex        =   29
         Top             =   300
         Width           =   405
      End
      Begin VB.ComboBox cmbTab1Ledger 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1750
         TabIndex        =   28
         Top             =   270
         Width           =   2565
      End
      Begin VB.ComboBox cmbTab1Ledger 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1750
         TabIndex        =   31
         Top             =   840
         Width           =   2565
      End
      Begin VB.CommandButton cmdTab1Show 
         Caption         =   "&Show"
         Height          =   400
         Left            =   3630
         TabIndex        =   15
         Top             =   3120
         Width           =   1035
      End
      Begin VB.TextBox txtTab1Dates 
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1750
         TabIndex        =   35
         Top             =   2040
         Width           =   2685
      End
      Begin VB.TextBox txtTab1Dates 
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   1750
         TabIndex        =   38
         Top             =   2520
         Width           =   2685
      End
      Begin VB.CheckBox chkTab1EnterDates 
         Caption         =   "EnterDates  "
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/d/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   33
         Top             =   1620
         Width           =   3705
      End
      Begin VB.Label lblTab1Ledgers 
         Caption         =   "Ledger Type"
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   27
         Top             =   330
         Width           =   1755
      End
      Begin VB.Label lblTab1Ledgers 
         Caption         =   "Ledger Name"
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   30
         Top             =   900
         Width           =   1755
      End
      Begin VB.Label lblTab1Dates 
         Caption         =   "From Date"
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   34
         Top             =   2130
         Width           =   1425
      End
      Begin VB.Label lblTab1Dates 
         Caption         =   "To Date"
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   36
         Top             =   2580
         Width           =   1575
      End
   End
   Begin ComctlLib.TabStrip tabTrans 
      Height          =   4350
      Left            =   30
      TabIndex        =   37
      Top             =   60
      Width           =   5000
      _ExtentX        =   8811
      _ExtentY        =   7673
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Transaction"
            Key             =   "Trans"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Show Ledger"
            Key             =   "Ledger"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Shows the ledger transctions"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraTab 
      Height          =   3825
      Index           =   2
      Left            =   150
      TabIndex        =   41
      Top             =   480
      Width           =   4800
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   495
         Left            =   480
         TabIndex        =   57
         Top             =   3120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddHeads 
         Caption         =   "->"
         Height          =   375
         Left            =   4080
         TabIndex        =   56
         Top             =   2160
         Width           =   495
      End
      Begin VB.ComboBox cmbAddHeads 
         Height          =   315
         Left            =   1680
         TabIndex        =   54
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cmbTransHeads 
         Height          =   315
         Left            =   1800
         TabIndex        =   52
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton cmdLedgerUpdate 
         Caption         =   "&Save"
         Height          =   525
         Left            =   3360
         TabIndex        =   51
         Top             =   3120
         Width           =   1155
      End
      Begin VB.ComboBox cmbTrans 
         Height          =   315
         Left            =   1800
         TabIndex        =   47
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton cmdLedgerShow 
         Caption         =   "&Show"
         Height          =   400
         Left            =   3600
         TabIndex        =   44
         Top             =   1560
         Width           =   1035
      End
      Begin VB.TextBox txtTransDate 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1800
         TabIndex        =   43
         Top             =   360
         Width           =   2205
      End
      Begin VB.CommandButton cmdTransDate 
         Caption         =   "..."
         Height          =   315
         Left            =   4200
         TabIndex        =   42
         Top             =   360
         Width           =   435
      End
      Begin VB.Label lblAddHeads 
         Caption         =   "Account Heads"
         Height          =   300
         Left            =   120
         TabIndex        =   55
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblAccHeads 
         Caption         =   "Account Heads"
         Height          =   300
         Left            =   240
         TabIndex        =   53
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblTransDate 
         Caption         =   "From Date"
         Height          =   300
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label lblTransactions 
         Caption         =   "Trans ID"
         Height          =   300
         Left            =   240
         TabIndex        =   45
         Top             =   1560
         Width           =   1425
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   240
      X2              =   12060
      Y1              =   5430
      Y2              =   5430
   End
   Begin VB.Label lblLedgerName 
      AutoSize        =   -1  'True
      Caption         =   "Ledger Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4470
      TabIndex        =   1
      Top             =   60
      Width           =   6765
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAccTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' For the Combos
Private Const VoucherType = 0
Private Const LedgerType = 1
Private Const LedgerName = 2

'For the Options
Private Const FromOpt = 0
Private Const ToOpt = 1

' For the Add Commands
Private Const AddToGrid = 0
Private Const DeleteGrid = 1

' To Handle Grid Functions
Private m_GrdFunctions As clsGrdFunctions
Private m_AccTransClass As clsAccTrans

' To show ledgers
Private m_LedgerClass As clsLedger

Private m_dbOperation As wis_DBOperation
Private m_IsAllowTransDate As Boolean
Private m_IsStartParticulars As Boolean
Private m_ActiveTab As Byte
Private m_debitTotal As Currency
Private m_creditTotal As Currency
Private m_PrevAmount As Double
Private m_updating As Boolean
Private m_TransID As Long
Private m_VoucherType As Wis_VoucherTypes

Private Function CheckTransDate() As Boolean

On Error GoTo Hell:

Dim LastTransDate As String
Dim CurrentDate As String

CheckTransDate = True
If m_IsAllowTransDate Then Exit Function


CheckTransDate = False

' Get the Last TransDate for the HeadID
LastTransDate = LoadLastTransDate

If LastTransDate = "" Then Exit Function

CurrentDate = txtTab0CurrentDate.Text

If Not TextBoxDateValidate(txtTab0CurrentDate, "/", True, True) Then Exit Function

If GetSysFormatDate(CurrentDate) < GetSysFormatDate(LastTransDate) Then
    
    If MsgBox("Current Date is Smaller than Last Entered Date!" & _
        vbCrLf & "Do You Want To Continue ?", vbQuestion + vbYesNo) = vbNo Then
        Exit Function
    Else
        m_IsAllowTransDate = True
        CheckTransDate = True
        Exit Function
    End If
Else
    CheckTransDate = True
End If

Exit Function

Hell:
    MsgBox "Check TransDate :" & vbCrLf & Err.Description
    
End Function

Private Sub ClearControls()

If MsgBox("Do You Want To Clear Controls ? ", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
grdLedger(m_ActiveTab).Clear
If m_ActiveTab = 0 Then
    RefreshForTab0
    m_AccTransClass.frmClearClicked
    Call cmbTab0Ledger_Click(0)
Else
    RefreshForTab1
End If

End Sub

'This Procedure  will get the last transacted date from acctrans and load it to
' m_LastTransDate
Private Function LoadLastTransDate() As String

Dim rstTransDate As ADODB.Recordset
Dim headID As Long

LoadLastTransDate = ""

With cmbTab0Ledger(LedgerName)
    If .ListIndex = -1 Then Exit Function
    headID = .ItemData(.ListIndex)
End With
    
gDbTrans.SqlStmt = " SELECT MAX(TransDate) as MaxTransDate " & _
                   " FROM AccTrans WHERE HeadID = " & headID
                 
Call gDbTrans.Fetch(rstTransDate, adOpenForwardOnly)

LoadLastTransDate = FormatField(rstTransDate.Fields("MaxTransDate"))

' If no transaction is made then last trans date will be first day
If LoadLastTransDate = "" Then LoadLastTransDate = FinIndianFromDate

Set rstTransDate = Nothing

End Function

'set the Kannada option here.
Private Sub SetKannadaCaption()

Call SetFontToControlsSkipGrd(Me)

'set the kannada for the Tabs
tabTrans.Tabs(1).Caption = GetResourceString(28)
tabTrans.Tabs(2).Caption = GetResourceString(36, 13)

tabTrans.Font.name = gFontName
tabTrans.Font.Size = gFontSize

'set the Kannada for all controls
lblTab0VoucherType(0).Caption = GetResourceString(41)    'Voucher
lbltab0Date(1).Caption = GetResourceString(37)  'NAme
lblTab0HeadType(2).Caption = GetResourceString(160) & " " & _
            GetResourceString(36)  'Main Account
lblTab0AccountType(3).Caption = GetResourceString(36) & " " & _
                GetResourceString(35)  'Account Name
optTab0From(0).Caption = GetResourceString(107)  'From
optTab0From(1).Caption = GetResourceString(108)   'TO
lbltab0Amount(4).Caption = GetResourceString(40)  'Amount
lblLedgerName.Caption = GetResourceString(36) & " " & _
            GetResourceString(35)  'Account NAme

cmdTab0AddtoGrid(0).Caption = GetResourceString(10)   'Add
cmdTab0AddtoGrid(1).Caption = GetResourceString(14)   ''Delete
cmdOk.Caption = GetResourceString(7)   'Save
cmdCancel(0).Caption = GetResourceString(8)   'Clear
cmdCancel(1).Caption = GetResourceString(11)  'Close
cmdPrint.Caption = GetResourceString(23)  'print

lblTab1Ledgers(0).Caption = GetResourceString(160) & " " & _
            GetResourceString(36)  'Main Account
lblTab1Ledgers(1).Caption = GetResourceString(36) & " " & _
            GetResourceString(35)  'Account NAme
lblTab1Dates(0).Caption = GetResourceString(109) 'After Date
lblTab1Dates(1).Caption = GetResourceString(110)  'Before Date
cmdTab1Show.Caption = GetResourceString(13)   'Show
chkTab1EnterDates.Caption = GetResourceString(106)    'Specify dat range

lblTransDate = GetResourceString(37)    'Date
lblTransactions = GetResourceString(28)     'Transactions
lblAccHeads.Caption = GetResourceString(86, 36)
cmdLedgerShow.Caption = GetResourceString(13)    'Show
cmdLedgerUpdate.Caption = GetResourceString(171)    'Update
'cmdAdHeads.Caption = GetResourceString(10)    'aDD
lblAddHeads.Caption = GetResourceString(36) + GetResourceString(92)  'aDD
End Sub


' Handles total entries made
'Private m_Entries As Integer
Public Sub InitTab0Grid()

With grdLedger(0)
    .Clear
    .Enabled = True
    .AllowUserResizing = flexResizeBoth
    .Rows = 5
    .Cols = 5
    .FixedCols = 1
    .FixedRows = 1
    
    .Row = 0
    
    .Col = 0: .CellFontBold = True: .Text = GetResourceString(33)  '"SlNo"
    .Col = 1: .CellFontBold = True: .Text = GetResourceString(36)  '"Ledger Name"
    .Col = 2: .CellFontBold = True: .Text = GetResourceString(277)  '"Dr "
    .Col = 3: .CellFontBold = True: .Text = GetResourceString(276)  '"Cr "
    .Col = 4: .CellFontBold = True: .Text = GetResourceString(42)  '"Total "
    
    .ColWidth(0) = 435
    .ColWidth(1) = 1900
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .ColWidth(4) = 1500
    
    .Row = 1
End With


End Sub


' Handles total entries made
'Private m_Entries As Integer
Public Sub InitTab2Grid()

m_debitTotal = 0
m_creditTotal = 0
With grdLedger(2)
    .Clear
    .Enabled = True
    .AllowUserResizing = flexResizeBoth
    .Rows = 2
    .Cols = 5
    .FixedCols = 1
    .FixedRows = 1
    
    .Row = 0
    
    .Col = 0: .CellFontBold = True: .Text = GetResourceString(33)  '"SlNo"
    .Col = 1: .CellFontBold = True: .Text = GetResourceString(36)  '"Ledger Name"
    .Col = 2: .CellFontBold = True: .Text = GetResourceString(277)  '"Dr "
    .Col = 3: .CellFontBold = True: .Text = GetResourceString(276)  '"Cr "
    '.Col = 4: .CellFontBold = True: .Text = GetResourceString(42)  '"Total "
    
    .ColWidth(0) = 435
    .ColWidth(1) = 3200
    .ColWidth(2) = 0
    .ColWidth(3) = 1400
    .ColWidth(4) = 1400
    '.ColWidth(4) = 1500
    
    .Row = 1
End With


End Sub


'If will when tab0 is clicked
Public Sub RefreshForTab0()
    
    Call InitTab0Grid
    cmbTab0Ledger(LedgerName).Text = ""
    cmbTab0Ledger(0).Locked = False
    'm_AccTransClass.frmDeleteClicked
    
End Sub

'If will when tab1 is clicked

Public Sub RefreshForTab1()
    
'    grdLedger(m_ActiveTab).Visible = True
'    grdLedger(m_ActiveTab).ZOrder 0
    
    Call LoadParentHeads(cmbTab1Ledger(0))
    
    cmbTab1Ledger(1).Clear
    txtTab1Dates(0).Text = ""
    txtTab1Dates(1).Text = ""
    txtParticulars.Text = ""
    'txtParticulars.Enabled = False
    If Not cmdTab0AddtoGrid(0).Visible Then
        txtParticulars.Enabled = True
        txtParticulars.TabIndex = cmdOk.TabIndex - 1
    End If
    
End Sub

'
Private Sub RefreshOptionButtons()
    
Dim voucherTypes As Wis_VoucherTypes

' This Will Fetch The Current Voucher Type Selected by the User

With cmbTab0Ledger(VoucherType)
    voucherTypes = .ItemData(.ListIndex)
End With

'Please dont change the case creiteria(Why?)
Select Case voucherTypes

    Case payment
        optTab0From(ToOpt).Value = True
        optTab0From(ToOpt).Enabled = True
    Case Receipt
        optTab0From(FromOpt).Value = True
        optTab0From(FromOpt).Enabled = True
    Case Sales
        optTab0From(ToOpt).Value = True
        optTab0From(ToOpt).Enabled = True
    Case Purchase
        optTab0From(FromOpt).Value = True
        optTab0From(FromOpt).Enabled = True
    Case FreePurchase
        optTab0From(FromOpt).Value = True
        optTab0From(FromOpt).Enabled = True
    Case FreeSales
        optTab0From(FromOpt).Value = True
        optTab0From(FromOpt).Enabled = True
    Case CONTRA
        optTab0From(FromOpt).Value = True
        optTab0From(FromOpt).Enabled = True
    Case Journal
        optTab0From(FromOpt).Value = True
        optTab0From(FromOpt).Enabled = True
    Case RejectionsIn
        optTab0From(FromOpt).Value = True
        optTab0From(FromOpt).Enabled = True
    Case RejectionsOut
        optTab0From(ToOpt).Value = True
        optTab0From(ToOpt).Enabled = True
End Select

End Sub

'
Private Sub SwapCheckedOptions()

If chkTab1EnterDates.Value = vbChecked Then
    txtTab1Dates(0).Text = FinIndianFromDate
    txtTab1Dates(1).Text = FinIndianEndDate
    txtTab1Dates(0).Enabled = True
    txtTab1Dates(1).Enabled = True
Else
    txtTab1Dates(0).Text = ""
    txtTab1Dates(1).Text = ""
    txtTab1Dates(0).Enabled = False
    txtTab1Dates(1).Enabled = False
End If

End Sub


'
Private Sub UnLoadME()
    Unload Me
End Sub

'
Private Sub chkTab1EnterDates_Click()
    SwapCheckedOptions
End Sub

Private Sub cmbTab0Ledger_Click(Index As Integer)

Dim VoucherType As Wis_VoucherTypes

If cmbTab0Ledger(Index).ListIndex = -1 Then Exit Sub

'Don't Change this Case creteria
Select Case Index
        
        Case VoucherType
             ' RefreshOptionButtons
             m_AccTransClass.frmVoucherClicked
             
        Case LedgerType
        
            ' This load the Ledgers to combo
            Call LoadLedgersToCombo(cmbTab0Ledger(LedgerName), _
                    cmbTab0Ledger(LedgerType).ItemData(cmbTab0Ledger(LedgerType).ListIndex))
                        
            With cmbTab0Ledger(LedgerName)
                If .ListCount > 0 Then .ListIndex = 0
            End With
            cmbTab0Ledger_Click (LedgerName)
        
        Case LedgerName
        
            If Not DateValidate(txtTab0CurrentDate.Text, "/", True) Then Exit Sub
            
            ' This will Fetch the Balance for the HeadId
            With cmbTab0Ledger(LedgerName)
                m_AccTransClass.frmHeadClicked (.ItemData(.ListIndex))
            End With
End Select

End Sub


Private Sub cmbTab1Ledger_Click(Index As Integer)

If Index = 0 Then

    If cmbTab1Ledger(0).ListIndex = -1 Then Exit Sub
    
    ' This load the Ledgers to combo
    
    Call LoadLedgersToCombo(cmbTab1Ledger(1), _
         cmbTab1Ledger(0).ItemData(cmbTab1Ledger(0).ListIndex))
    
Else
    
End If

End Sub


Private Sub cmbTrans_Change()
        
    Dim rstTrans As Recordset
    Dim negtiveValues As Boolean
    If Len(cmbTrans.Text) > 0 Then m_TransID = CLng(cmbTrans.Text)
    
    gDbTrans.SqlStmt = "Select A.*, iif(IsNull(C.AliasName ), B.HeadName,C.AliasName) as HeadName " & _
            " from AccTrans A inner join " & _
            " (Heads B left join BankHeadIds C on B.HeadID=C.headID)" & _
            " on A.HeadID = B.HeadID where transid = " & m_TransID
    
    cmdLedgerUpdate.Enabled = False
    
    If gDbTrans.Fetch(rstTrans, adOpenDynamic) > 0 Then
        grdLedger(2).Row = grdLedger(2).Rows - 2
        m_updating = True
        Call InitTab2Grid
        grdLedger(2).Row = 0
        m_PrevAmount = 0
        
        m_VoucherType = FormatField(rstTrans("VoucherType"))
        While Not rstTrans.EOF
            With grdLedger(2)
                If .Rows < .Row + 2 Then .Rows = .Row + 2
                .Row = .Row + 1
                .Col = 0: .Text = CStr(.Row)
                .Col = 1: .Text = FormatField(rstTrans("HeadName"))
                .Col = 2: .Text = FormatField(rstTrans("HeadID"))
                If FormatField(rstTrans("Debit")) < 0 Or FormatField(rstTrans("credit")) < 0 Then negtiveValues = True
                .Col = 3: .Text = FormatField(rstTrans("Debit")): m_debitTotal = m_debitTotal + Val(.Text)
                .Col = 4: .Text = FormatField(rstTrans("Credit")): m_creditTotal = m_creditTotal + Val(.Text)
            End With
            rstTrans.MoveNext
        Wend
        With grdLedger(2)
            If .Rows < .Row + 3 Then .Rows = .Row + 3
            .Row = .Row + 2
            .Col = 3: .Text = FormatCurrency(m_debitTotal): .CellFontBold = True
            .Col = 4: .Text = FormatCurrency(m_creditTotal): .CellFontBold = True
            .Col = 1: .Text = GetResourceString(52): .CellFontBold = True
            .Row = 0
        End With
        
        'If m_debitTotal <> m_creditTotal Or negtiveValues Then m_updating = False
        m_updating = False
        
    End If
    
    LoadAllNoTransHeads (m_TransID)
    
End Sub

Private Sub cmbTrans_Click()
Call cmbTrans_Change
End Sub


Private Sub cmbTransHeads_Change()
    Call LedgerTransDetails
    cmbAddHeads.Clear
End Sub

Private Sub cmbTransHeads_Click()
    Call LedgerTransDetails
    On Error Resume Next
    cmbTrans.ListIndex = 0
End Sub

Private Sub cmdAddHeads_Click()
    Call FixHeads
    If cmbAddHeads.ListIndex = -1 Then Exit Sub
    Dim rst As Recordset
    Dim headID As Long
    
    headID = cmbAddHeads.ItemData(cmbAddHeads.ListIndex)
    'Check the Existnace of this ID in the Tablse
    gDbTrans.SqlStmt = "Select * from AccTrans A" & _
        " Where TransID = " & m_TransID & " And headID = " & headID
    
    If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then
        gDbTrans.SqlStmt = "INSERT INTO AccTrans " & _
                     " (HeadID,TransID,VoucherType,TransDate,Debit,Credit)" & _
                     " VALUES (" & headID & "," & m_TransID & "," & m_VoucherType & _
                     ",#" & GetSysFormatDate(txtTransDate.Text) & "#,0, 0 )"
                     
        gDbTrans.BeginTrans
        If Not gDbTrans.SQLExecute Then
            gDbTrans.RollBack
            Exit Sub
        End If
        gDbTrans.CommitTrans
        Call cmbTrans_Change
    End If
    
End Sub

Private Sub cmdCBStock_Click()

Dim ObStr As String

ObStr = InputBox("Enter The Closing Stock Value After this Transction", "Stock Value", FormatCurrency(Val(lblTab0Balance(0))))
If Val(ObStr) = 0 Then Exit Sub

lblTab0Balance(0).Tag = "1"
lblTab0Balance(0) = Format(ObStr)

End Sub

Private Sub cmdLedgerShow_Click()
    Call cmbTrans_Change
End Sub

Private Sub cmdLedgerUpdate_Click()
    With grdLedger(2)
        If Val(.TextMatrix(.Rows - 1, 3)) <> Val(.TextMatrix(.Rows - 1, 4)) Then cmdLedgerUpdate.Enabled = False: Exit Sub
        'If Val(.TextMatrix(.Rows - 1, 3)) <> m_debitTotal Or (.TextMatrix(.Rows - 1, 4)) <> m_creditTotal Then cmdLedgerUpdate.Enabled = False: Exit Sub
        
        ''Now Update each transcation as per the new Value
        Dim loopCount As Integer
        
        gDbTrans.BeginTrans
        For loopCount = 1 To .Rows - 2
            gDbTrans.SqlStmt = "UPdate AccTrans set " & _
                " Debit =  " & Val(.TextMatrix(loopCount, 3)) & ", Credit = " & Val(.TextMatrix(loopCount, 4)) & _
                " Where TransID = " & m_TransID & " And HeadId = " & Val(.TextMatrix(loopCount, 2))
            
            If Not gDbTrans.SQLExecute Then
                gDbTrans.RollBack
                Exit Sub
            End If
        Next loopCount
        
        gDbTrans.CommitTrans
        MsgBox GetResourceString(707), vbOKOnly, wis_MESSAGE_TITLE
        Call cmbTrans_Change
    End With
End Sub

Private Sub cmdPrint_Click()
'Set m_frmCancel = New frmCancel
Dim m_grdPrint  As WISPrint
Set m_grdPrint = wisMain.grdPrint
With m_grdPrint
    .CompanyName = gCompanyName
    .Font.name = gFontName
    .ReportTitle = "Transaction Details"
    '.GridObject = grdLedger(IIf(tabTrans.Tabs(1).Selected, 0, 1))
    .GridObject = grdLedger(tabTrans.SelectedItem.Index - 1)
    'm_frmCancel.Show
    'm_frmCancel.PicStatus.Visible = True
    'UpdateStatus m_frmCancel.PicStatus, 0
    .PrintGrid
 '   Unload m_frmCancel
    
End With

End Sub

Private Sub cmdTab0AddtoGrid_Click(Index As Integer)

If Index = DeleteGrid Then m_AccTransClass.frmDeleteClicked

If Index = AddToGrid Then
    If m_AccTransClass.UpdatingGrid Then
        m_AccTransClass.frmUpdateClicked
    Else
        m_AccTransClass.frmAddClicked
    End If
End If

End Sub
'
Private Sub cmdCancel_Click(Index As Integer)

If Index = 0 Then ClearControls

If Index = 1 Then UnLoadME

End Sub

Private Sub cmdOk_Click()

If m_AccTransClass.frmOKClicked = Success Then m_IsAllowTransDate = False
    
End Sub

Private Sub cmdTab0Date_Click()
With Calendar
    .Top = Me.Top + tabTrans.Top + cmdTab0Date.Top
    .Left = Me.Left + tabTrans.Left + cmdTab0Date.Left
    .selDate = IIf(DateValidate(txtTab0CurrentDate, "/", True), txtTab0CurrentDate, gStrDate)
    .Show vbModal
    txtTab0CurrentDate = .selDate
End With

End Sub

Private Sub cmdTab1Ledger_Click(Index As Integer)
If Index = LedgerName Then
    If m_LedgerClass Is Nothing Then Set m_LedgerClass = New clsLedger
    If cmbTab1Ledger(0).ListIndex >= 0 Then _
        m_LedgerClass.ParentID = cmbTab1Ledger(0).ItemData(cmbTab1Ledger(0).ListIndex)
    
    m_LedgerClass.ShowLedger

ElseIf Index = LedgerType Then
    
    frmSubParent.Show vbModal
    Call LoadParentHeads(cmbTab0Ledger(LedgerType))

End If

'Load the parent heads
Call cmbTab0Ledger_Click(LedgerType)
'LOad the sub heads
Call cmbTab1Ledger_Click(0)

End Sub

Private Sub cmdTab1Show_Click()

On Error GoTo Hell:

' Declarations
Dim lngHeadID As Long
Dim strFromDate As String
Dim strToDate As String
Dim lngCheckedValue As Long

If cmbTab1Ledger(0).ListIndex = -1 Then Err.Raise vbObjectError + 513, , "Parent Ledger Not Selected"
If cmbTab1Ledger(1).ListIndex = -1 Then Err.Raise vbObjectError + 513, , "Ledger Not Selected"

lngCheckedValue = chkTab1EnterDates.Value
     
Set m_LedgerClass = New clsLedger

If lngCheckedValue = vbChecked Then
    
    strFromDate = txtTab1Dates(0).Text
    If Not TextBoxDateValidate(txtTab1Dates(0), "/", True, True) Then Exit Sub
    
    strToDate = txtTab1Dates(1).Text
    If Not TextBoxDateValidate(txtTab1Dates(1), "/", True, True) Then Exit Sub
    
End If

With cmbTab1Ledger(1)

    lngHeadID = .ItemData(.ListIndex)

End With

Me.MousePointer = vbHourglass

m_IsStartParticulars = False

If lngCheckedValue = vbChecked Then
    Call m_AccTransClass.ShowNewLedgerToGrid(lngHeadID, grdLedger(m_ActiveTab), True, _
                            strFromDate, strToDate)
Else
    Call m_AccTransClass.ShowNewLedgerToGrid(lngHeadID, grdLedger(m_ActiveTab), False, "", "")
End If

m_IsStartParticulars = True

Me.MousePointer = vbDefault
    
If lngCheckedValue = vbChecked Then _
    lblLedgerName.Caption = cmbTab1Ledger(1).Text & "           " _
        & strFromDate & " To " & strToDate
        
If lngCheckedValue <> vbChecked Then _
    lblLedgerName.Caption = cmbTab1Ledger(1).Text & "           " _
        & FinIndianFromDate & " To " & FinIndianEndDate
   
   
Exit Sub

Hell:
    
    MsgBox "Voucher Entry : " & vbCrLf & Err.Description
    
End Sub


Private Sub cmdTransDate_Click()
With Calendar
    .Left = Me.Left + Me.fraTab(2).Left + Me.cmdTransDate.Left
    .Top = Me.Top + Me.fraTab(2).Top + Me.cmdTransDate.Top
    .selDate = txtTransDate.Text
    .Show vbModal
    If .selDate <> "" Then txtTransDate.Text = .selDate
End With
End Sub

Private Sub Form_Initialize()
If m_GrdFunctions Is Nothing Then Set m_GrdFunctions = New clsGrdFunctions
If m_AccTransClass Is Nothing Then Set m_AccTransClass = New clsAccTrans: m_AccTransClass.TempTest = "test 2"

Set m_GrdFunctions.fGrd = grdLedger(m_ActiveTab)

End Sub

Private Sub Form_Load()
m_updating = True '
CenterMe Me

Me.Icon = LoadResPicture(147, vbResIcon)
txtTransDate = gStrDate
tabTrans.Tabs.Remove (3)
If (gCurrUser.UserPermissions And perOnlyWaves) Or gCurrUser.IsAdmin Then LedgerTransDetails: LedgerTransHeads


If gLangOffSet <> 0 Then SetKannadaCaption

'LOad The Current Date
txtTab0CurrentDate = gStrDate

'Load the parent heads
Call LoadVouchersToCombo(cmbTab0Ledger(VoucherType))
Call LoadParentHeads(cmbTab0Ledger(LedgerType))

Call LoadParentHeads(cmbTab1Ledger(0))


'Call LoadLedgersToCombo(cmbTab1Ledger(1), _
    cmbTab1Ledger(0).ItemData(cmbTab1Ledger(0).ListIndex))

'RefreshForTab0
Call tabTrans_Click
'Set The Grid
grdLedger(1).Top = grdLedger(0).Top
grdLedger(1).Left = grdLedger(0).Left
grdLedger(1).Width = grdLedger(0).Width
grdLedger(1).Height = grdLedger(0).Height

grdLedger(2).Top = grdLedger(0).Top
grdLedger(2).Left = grdLedger(0).Left
grdLedger(2).Width = grdLedger(0).Width
grdLedger(2).Height = grdLedger(0).Height

Call InitTab0Grid
Call InitTab2Grid

Call m_AccTransClass.InitTab1Grid(grdLedger(1))

grdLedger(0).Visible = True
grdLedger(0).ZOrder 0

Call Form_Resize

If gOnLine Then
    txtTab0CurrentDate.Locked = True
    cmdTab0Date.Enabled = False
End If

'While making transaction we have to cvhange the
'tab setting of this particulars text box
' we are assigning its existing tabindex to it's tag property
txtParticulars.Tag = txtParticulars.TabIndex
m_updating = False

cmdTab1Ledger(1).Visible = CBool((gCurrUser.UserPermissions And perOnlyWaves) = perOnlyWaves)


End Sub
Private Sub FixHeads()
    If Not DateValidate(txtTransDate.Text, "/", True) Then Exit Sub
    
    Dim DateAsOn As Date
    Dim CashDeposits As Currency
    Dim ContraDeposits As Currency
    Dim CashWithdrawls As Currency
    Dim ContraWithdrawls As Currency
    Dim CashDeposits_Head As Currency
    Dim ContraDeposits_head As Currency
    Dim CashWithdrawls_head As Currency
    Dim ContraWithdrawls_head As Currency
    
    Dim rstHeadTrans As Recordset
    Dim ClsBank As New clsBankAcc
    Dim headID As Long
    Dim AccHeadID As Long
    Dim innerSql As String
    DateAsOn = GetSysFormatDate(txtTransDate.Text)
    '' Get the Ledger Trasactions on this date
    innerSql = "Select distinct TransID from AccTrans A" & _
        " Inner Join Heads B On A.HeadID = B.HeadID " & _
        " Where TransDate = #" & DateAsOn & "#" & _
        " And A.HeadID = " & headID & _
        " ANd B.ParentID in ( " & parProfitORLoss & "," & parIncome & "," & parBankAccount & "," & parBankLoanAccount & _
                 "," & parFixedAsset & "," & parShareCapital & "," & parGovtLoanSubsidy & ", " & parOtherLoans & _
                 "," & parPayAble & " ," & parLoanIntProv & " ," & parSuspAcc & _
                 "," & parInvestment & "," & parLoanAdvanceAsset & ", " & parSalaryAdvance & _
                 "," & parSalaryExpense & "," & parReceivable & "," & parReserveFunds & _
                 "," & parBankIncome & "," & parExpense & "," & parTradingExpense & "," & parTradingIncome & _
                 "," & parBankExpense & "," & parSales & "," & parPurchase & _
                ")"
    
    
    gDbTrans.SqlStmt = "Select D.parentID,C.* from AccTrans C" & _
        " Inner Join Heads D On C.HeadID = D.HeadID " & _
        " where TransDate = #" & DateAsOn & "#" & _
        " And Transid Not in (" & innerSql & ")"
    ' Create a view on this Date
    gDbTrans.CreateView ("qryDayBookTrans")
    
    ''Fix the SB Head
    'Find the Transactions Happened in SB.
    Dim SBClass As clsSBAcc
    Set SBClass = New clsSBAcc
    AccHeadID = ClsBank.GetHeadIDCreated(GetResourceString(421))
    'Call HeadTransDetails(AccHeadID, DateAsOn, CashDeposits_Head, ContraDeposits_head, CashWithdrawls_head, ContraWithdrawls_head)
    'Get For the TransDatails from the Individual
    Call SBClass.TotalDepositTransactions(DateAsOn, DateAsOn, CashDeposits, ContraDeposits, CashWithdrawls, ContraWithdrawls)
    gDbTrans.SqlStmt = "Select * from qryDayBookTrans where HeadID = " & AccHeadID
    Dim vType As Wis_VoucherTypes
    If gDbTrans.Fetch(rstHeadTrans, adOpenDynamic) > 0 Then
        While Not rstHeadTrans.EOF
            vType = Val(FormatField(rstHeadTrans("VoucherType")))
            If vType = Receipt Then CashDeposits_Head = CashDeposits_Head + Val(FormatField(rstHeadTrans("Credit")))
            If vType = payment Then CashWithdrawls_head = CashWithdrawls_head + Val(FormatField(rstHeadTrans("Debit")))
            If vType = Journal Then ContraDeposits_head = ContraDeposits_head + Val(FormatField(rstHeadTrans("Credit")))
            If vType = Journal Then ContraWithdrawls_head = ContraWithdrawls_head + Val(FormatField(rstHeadTrans("Debit")))
            
            rstHeadTrans.MoveNext
        Wend
    End If
    If CashDeposits <> CashDeposits_Head Then
        vType = Receipt
    End If
    If CashWithdrawls <> CashWithdrawls_head Then
    vType = payment
    End If
    If ContraDeposits <> ContraDeposits_head Then
        vType = Journal
    End If
    If ContraWithdrawls <> ContraWithdrawls_head Then
        vType = Journal
    End If

End Sub
Private Sub HeadTransDetails(headID As Long, AsOnDate As Date, ByRef CashDeposit As Currency, ByRef ContraDeposit As Currency, ByRef cashWithdraw As Currency, ByRef ContraWithDraw As Currency)
    
CashDeposit = 0
cashWithdraw = 0
ContraDeposit = 0
ContraWithDraw = 0

Dim Deposit As Currency
Dim WithDraw As Currency
Dim isCOntra As Boolean
Dim PrevTransID As Long
Dim rstTrans As Recordset
    

    gDbTrans.SqlStmt = "Select * from AccTrans " & _
        " Where TransDate = #" & AsOnDate & "# And HeadID = " & headID & _
        " Order By TransID"
    
    If gDbTrans.Fetch(rstTrans, adOpenDynamic) > 0 Then
        PrevTransID = 0
        While Not rstTrans.EOF
            If PrevTransID <> rstTrans("TransID") Then
                If PrevTransID > 0 Then
                    If isCOntra Then
                        ContraDeposit = ContraDeposit + Deposit
                        ContraWithDraw = ContraWithDraw + WithDraw
                    Else
                        CashDeposit = CashDeposit + Deposit
                        cashWithdraw = cashWithdraw + WithDraw
                    End If
                End If
                Deposit = 0
                WithDraw = 0
                isCOntra = True
            End If
            If rstTrans("HeadID") = wis_CashHeadID Then isCOntra = False
            If rstTrans("HeadID") = headID Then
                Deposit = Deposit + FormatField(rstTrans("Debit"))
                WithDraw = WithDraw + FormatField(rstTrans("Credit"))
            End If
        rstTrans.MoveNext
        Wend
    End If
    

End Sub

Private Sub LedgerTransDetails()
    
    If tabTrans.Tabs.count < 3 Then _
        Call tabTrans.Tabs.Add(3, "LedgerTrans", GetResourceString(93)) '& " " & GetResourceString(28)
     cmbTrans.Clear
     Dim headID As Long
     'Select the Trnaction ID of the date selected
    If Not DateValidate(txtTransDate, "/", True) Then Exit Sub
    Dim rstTrans As Recordset
    If cmbTransHeads.ListIndex <> -1 Then headID = cmbTransHeads.ItemData(cmbTransHeads.ListIndex)
    gDbTrans.SqlStmt = "Select distinct TransID from AccTrans " & _
        " Where TransDate = #" & GetSysFormatDate(txtTransDate.Text) & "#"
    
    If headID > 0 Then gDbTrans.SqlStmt = gDbTrans.SqlStmt & " And HeadID = " & headID
    
    If gDbTrans.Fetch(rstTrans, adOpenDynamic) > 0 Then
       
        While Not rstTrans.EOF
            cmbTrans.AddItem (rstTrans(0))
        rstTrans.MoveNext
        Wend
    End If
    

End Sub
Private Sub LoadAllNoTransHeads(TranID As Long)
    
    cmbAddHeads.Clear
    Dim ParentID As Long
    Dim rstTrans As Recordset
    
    gDbTrans.SqlStmt = "Select * from Heads Where HeadID" & _
        " not in (Select distinct(HeadID) from AccTrans Where TransID = " & TranID & ")" & _
        " And ParentID not in ( 70000," & parIncome & "," & parBankAccount & "," & parBankLoanAccount & _
                 "," & parFixedAsset & "," & parShareCapital & "," & parGovtLoanSubsidy & ", " & parOtherLoans & _
                 "," & parPayAble & " ," & parLoanIntProv & " ," & parSuspAcc & _
                 "," & parInvestment & "," & parLoanAdvanceAsset & ", " & parSalaryAdvance & _
                 "," & parSalaryExpense & "," & parReceivable & "," & parReserveFunds & _
                 "," & parBankIncome & "," & parExpense & "," & parTradingExpense & "," & parTradingIncome & _
                 "," & parBankExpense & "," & parSales & "," & parPurchase & _
                ")"
        
    If gDbTrans.Fetch(rstTrans, adOpenDynamic) > 0 Then
        
        While Not rstTrans.EOF
            cmbAddHeads.AddItem (rstTrans("HeadName"))
            cmbAddHeads.ItemData(cmbAddHeads.newIndex) = rstTrans("HeadID")
        
            rstTrans.MoveNext
        Wend
    End If
    
End Sub
Private Sub LedgerTransHeads()
    If Not DateValidate(txtTransDate, "/", True) Then Exit Sub
    
    Dim rstTrans As Recordset
    gDbTrans.SqlStmt = "Select distinct A.headID,HeadName from AccTrans A" & _
        " Inner Join Heads B On A.HeadID = B.HeadID" & _
        " Where TransDate = #" & GetSysFormatDate(txtTransDate.Text) & "#" & _
        " ANd B.ParentID not in ( " & parProfitORLoss & "," & parIncome & "," & parBankAccount & "," & parBankLoanAccount & _
                 "," & parFixedAsset & "," & parShareCapital & "," & parGovtLoanSubsidy & ", " & parOtherLoans & _
                 "," & parPayAble & " ," & parLoanIntProv & " ," & parSuspAcc & _
                 "," & parInvestment & "," & parLoanAdvanceAsset & ", " & parSalaryAdvance & _
                 "," & parSalaryExpense & "," & parReceivable & "," & parReserveFunds & _
                 "," & parBankIncome & "," & parExpense & "," & parTradingExpense & "," & parTradingIncome & _
                 "," & parBankExpense & "," & parSales & "," & parPurchase & _
                ")"
    
    cmbTransHeads.Clear
    
    If gDbTrans.Fetch(rstTrans, adOpenDynamic) > 0 Then
        cmbTransHeads.AddItem (GetResourceString(338, 36) & GetResourceString(92))
        cmbTransHeads.ItemData(cmbTransHeads.newIndex) = 0
        While Not rstTrans.EOF
            cmbTransHeads.AddItem (rstTrans(1))
            cmbTransHeads.ItemData(cmbTransHeads.newIndex) = rstTrans(0)
        rstTrans.MoveNext
        Wend
    Else
        cmdUpdate.Visible = True
    End If

End Sub
Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)

Set m_AccTransClass = Nothing
Set m_LedgerClass = Nothing
Set m_GrdFunctions = Nothing

End Sub

Private Sub Form_Resize()
Call m_AccTransClass.SetAccTrans(Me)
Set m_GrdFunctions.fGrd = grdLedger(0)
End Sub


Private Sub Form_Unload(cancel As Integer)

Set frmAccTrans = Nothing

End Sub

Public Property Get DBOperation() As wis_DBOperation

    DBOperation = m_dbOperation
    
End Property

Public Property Let DBOperation(ByVal NewValue As wis_DBOperation)
    m_dbOperation = NewValue
End Property

Private Sub grdLedger_DblClick(Index As Integer)
Dim RowNum As Integer

grdLedger(Index).Visible = False
RowNum = grdLedger(Index).Row
Call m_AccTransClass.frmGridClicked(RowNum, Index)

grdLedger(Index).Visible = True

End Sub


Private Sub grdLedger_EnterCell(Index As Integer)
    If Index <> 2 Then Exit Sub
    If m_updating Then Exit Sub
    'If Not m_Shown Then Exit Sub

    With grdLedger(2)
        If .Row = 0 Then Exit Sub
        If .Row > .Rows - 2 Then Exit Sub
        If .Col < 3 Then Exit Sub
        'If m_PutTotal And .Rows = .Row + 1 Then Exit Sub
        
        txtBox.Text = .Text
        m_PrevAmount = Val(.Text)
        'txtBox.Move grd.Left + grd.CellLeft, grd.Top + grd.CellTop, grd.CellWidth, grd.CellHeight
        txtBox.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
        txtBox.Visible = True
        
        On Error Resume Next
        ActivateTextBox txtBox
        txtBox.ZOrder 0
        Err.Clear
    End With


End Sub


Private Sub grdLedger_GotFocus(Index As Integer)
     Call grdLedger_EnterCell(Index)
End Sub

Private Sub grdLedger_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 4 Then Exit Sub 'CTrl+D
If Index = 1 Then
    DBOperation = DeleteRec
    m_AccTransClass.TransID = grdLedger(Index).RowData(grdLedger(Index).Row)
    m_AccTransClass.frmDeleteIDClicked
    m_IsAllowTransDate = False
End If
End Sub

Private Sub grdLedger_LeaveCell(Index As Integer)

If Index <> 2 Then Exit Sub
If m_updating Then Exit Sub
'If Not m_Shown Then Exit Sub

Static InLoop As Boolean
If InLoop Then Exit Sub
Dim Bal As Currency
Dim Amount As Double
Dim PrevAmount As Double

With grdLedger(2)
    If .Col < 3 Then Exit Sub
    'If m_PutTotal And .Rows = .Row + 1 Then Exit Sub
    'mARK AS CONTROL IS IN LOOP
    InLoop = True
    
    'PrRow = .Row
    'Get the Previous amount
    PrevAmount = m_PrevAmount 'Val(.Text)
    
    'Now update the New Amount to the grid
    .Text = txtBox
    
    txtBox.Visible = False
    Amount = Val(txtBox)
    txtBox = ""
   
    Bal = FormatCurrency(Val(.TextMatrix(.Rows - 1, .Col)) + Amount - PrevAmount)
    .TextMatrix(.Rows - 1, .Col) = Bal
    cmdLedgerUpdate.Enabled = IIf(Val(.TextMatrix(.Rows - 1, 3)) = Val(.TextMatrix(.Rows - 1, 4)), True, False)
End With

'MARK AS CONTROL IS OUT OF LOOP
InLoop = False

End Sub

Private Sub grdLedger_RowColChange(Index As Integer)

If m_IsStartParticulars Then
    txtParticulars.Text = m_LedgerClass.GetTransIDParticulars(grdLedger(Index).RowData(grdLedger(Index).Row))
End If

End Sub

Private Sub tabTrans_Click()

m_ActiveTab = tabTrans.SelectedItem.Index - 1
'fraTab(m_ActiveTab).ZOrder 0

grdLedger(0).Visible = False
grdLedger(1).Visible = False
grdLedger(2).Visible = False
fraTab(0).Visible = False
fraTab(1).Visible = False
fraTab(2).Visible = False
fraTab(m_ActiveTab).ZOrder 0
fraTab(m_ActiveTab).Visible = True
grdLedger(m_ActiveTab).Visible = True
grdLedger(m_ActiveTab).ZOrder 0
    
Set m_GrdFunctions.fGrd = grdLedger(m_ActiveTab)
'Shashi On 16/12/2002

End Sub

Private Sub txtAfter_Change()
'THis COntorl is provided to track the
'Tab Movement the tab key will not be catch by either form or text box
'When control is in text box if the user want to enter
'in to the next amount box which is label
'so this text box will set the txtbox to user's required position
Dim txtNo As Integer
txtNo = Val(txtBox.Tag)
'If txtNo < txtAmount.Count - 1 Then
If txtNo < grdLedger(2).Rows - 1 Then
    txtNo = txtNo + 1
    'Call txtAmount_Click(txtNo)
Else
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtBefore_Change()
'Tab Movement the tab key will not be catch by either form or text box
'When control is in text box if the user want to enter
'in to the previous amount box which is label
'so this text box will set the txtbox to user's required position
Dim txtNo As Integer
txtNo = Val(txtBox.Tag)

If txtNo > 0 Then
    txtNo = txtNo - 1
'    Call txtAmount_Click(txtNo)
Else
    SendKeys "{TAB}"
End If

End Sub


Private Sub txtBox_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 37 Then 'Press Left Arrow
    If txtBox.SelStart = 0 And grdLedger(2).Col > grdLedger(2).FixedCols Then grdLedger(2).Col = grdLedger(2).Col - 1
    
ElseIf KeyCode = 39 Then 'Press Right Arrow
    If txtBox.SelStart = Len(txtBox.Text) And grdLedger(2).Col < grdLedger(2).Cols - 1 Then grdLedger(2).Col = grdLedger(2).Col + 1
    
ElseIf KeyCode = 38 Then  'Press UpArrow
    If grdLedger(2).Row > grdLedger(2).FixedRows Then grdLedger(2).Row = grdLedger(2).Row - 1
    
ElseIf KeyCode = 40 Then ' Press Doun Arroow
    If grdLedger(2).Row < grdLedger(2).Rows - 1 Then grdLedger(2).Row = grdLedger(2).Row + 1
    
ElseIf KeyCode = 33 Then  'Press PageUp
    If grdLedger(2).Row > 0 Then grdLedger(2).Row = grdLedger(2).Row - 1
    
ElseIf KeyCode = 34 Then ' Press PageDown
    If grdLedger(2).Row < grdLedger(2).Rows - 1 Then grdLedger(2).Row = grdLedger(2).Row + 1
    
End If

End Sub


Private Sub txtTransDate_Change()
    Call LedgerTransHeads
    Call LedgerTransDetails
End Sub


