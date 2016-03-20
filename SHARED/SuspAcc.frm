VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F166A15E-AA26-47C4-9C7F-A61A5BECEDFF}#2.0#0"; "CurrText.ocx"
Begin VB.Form frmSuspAcc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suspence Account"
   ClientHeight    =   5955
   ClientLeft      =   2145
   ClientTop       =   1320
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   400
      Left            =   5730
      TabIndex        =   27
      Top             =   5490
      Width           =   1215
   End
   Begin VB.Frame fraPassBook 
      BorderStyle     =   0  'None
      Caption         =   "Frame13"
      Height          =   2265
      Left            =   180
      TabIndex        =   22
      Top             =   3090
      Width           =   6705
      Begin VB.CommandButton cmdPrevTrans 
         Caption         =   "<"
         Height          =   315
         Left            =   6270
         TabIndex        =   24
         Top             =   135
         Width           =   375
      End
      Begin VB.CommandButton cmdNextTrans 
         Caption         =   ">"
         Height          =   315
         Left            =   6270
         TabIndex        =   25
         Top             =   570
         Width           =   375
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   6270
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1740
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   2175
         Left            =   60
         TabIndex        =   23
         Top             =   30
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   5
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame fra1 
      Height          =   2985
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      Begin VB.CommandButton cmdCust 
         Caption         =   ".."
         Height          =   315
         Left            =   6360
         TabIndex        =   15
         Top             =   1050
         Width           =   315
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "&Undo last"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3840
         TabIndex        =   21
         Top             =   2550
         Width           =   1215
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Default         =   -1  'True
         Height          =   375
         Left            =   5250
         TabIndex        =   20
         Top             =   2550
         Width           =   1215
      End
      Begin VB.ComboBox cmbParticulars 
         Height          =   315
         Left            =   1590
         TabIndex        =   16
         Top             =   2010
         Width           =   2025
      End
      Begin VB.ComboBox cmbTrans 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtVoucherNo 
         Height          =   315
         Left            =   5100
         TabIndex        =   19
         Top             =   2010
         Width           =   1305
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   ".."
         Height          =   315
         Left            =   3510
         TabIndex        =   3
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtDate 
         Height          =   345
         Left            =   1590
         TabIndex        =   2
         Top             =   210
         Width           =   1695
      End
      Begin VB.TextBox txtCustName 
         Height          =   345
         Left            =   1590
         TabIndex        =   10
         Top             =   1050
         Width           =   4725
      End
      Begin VB.CommandButton cmdAccNo 
         Caption         =   ".."
         Height          =   315
         Left            =   6360
         TabIndex        =   8
         Top             =   630
         Width           =   315
      End
      Begin VB.TextBox txtAccNo 
         Height          =   345
         Left            =   5460
         TabIndex        =   7
         Top             =   630
         Width           =   855
      End
      Begin VB.ComboBox cmbAccHead 
         Height          =   315
         Left            =   1590
         TabIndex        =   5
         Top             =   630
         Width           =   2325
      End
      Begin WIS_Currency_Text_Box.CurrText txtAmount 
         Height          =   345
         Left            =   5100
         TabIndex        =   14
         Top             =   1530
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         CurrencySymbol  =   ""
         TeenString      =   "eleven,twelwe,thirteen,fourteen,fifteen,sixteen,seventeen,eighteen,ninteen"
         NumberString    =   "zero,one,two,three,four,five,six,seven,eight,nine"
         FontSize        =   8.25
      End
      Begin VB.Label lblCustName 
         Caption         =   "Customer Name"
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Line Line1 
         X1              =   6615
         X2              =   90
         Y1              =   2445
         Y2              =   2445
      End
      Begin VB.Label lblParticular 
         Caption         =   "Particulars : "
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2010
         Width           =   1035
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount (Rs) : "
         Height          =   285
         Left            =   3780
         TabIndex        =   13
         Top             =   1590
         Width           =   1215
      End
      Begin VB.Label lblTrans 
         Caption         =   "Transaction : "
         Height          =   285
         Left            =   90
         TabIndex        =   11
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label lblVoucher 
         Caption         =   "Voucher No: "
         Height          =   255
         Left            =   3810
         TabIndex        =   18
         Top             =   2070
         Width           =   1215
      End
      Begin VB.Label lblDate 
         Caption         =   "Date :"
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblAccNo 
         Caption         =   "Account No :"
         Height          =   255
         Left            =   4110
         TabIndex        =   6
         Top             =   690
         Width           =   1335
      End
      Begin VB.Label lblAccHead 
         Caption         =   "Ledger Head"
         Height          =   285
         Left            =   90
         TabIndex        =   4
         Top             =   630
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmSuspAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_rstTrans As Recordset
Private WithEvents m_frmLookUp As frmLookUp
Attribute m_frmLookUp.VB_VarHelpID = -1
Private m_retVar

Public Event RemoveAmount(ByVal AccHeadID As Long, ByVal AccountId As Long, _
                ByVal name As String, ByVal TransDate As Date, _
                ByVal Amount As Currency, ByVal PrevTrans As Long)
                
Public Event AddAmount(ByVal AccHeadID As Long, ByVal AccountId As Long, _
                ByVal CustomerID As Long, ByVal name As String, _
                ByVal TransDate As Date, ByVal Amount As Currency)
                
Public Event UndoTrans()
Public Event ClearTrans()

Public Event WindowClosed(wHandle As Long)

Private Sub ClearControl()

cmdAccept.Tag = 0
cmdUndo.Caption = GetResourceString(5)

cmbAccHead.ListIndex = -1

txtAccNo.Tag = 0
txtAccNo = ""

txtCustName.Tag = 0
txtCustName = ""

cmbTrans.ListIndex = 0
txtAmount = 0
txtVoucherNo = ""



End Sub

Private Sub cmdAccept_Click()

'Intitally validate the Detals
If Not DateValidate(txtDate, "/", True) Then
    MsgBox GetResourceString(501), wis_MESSAGE_TITLE
    ActivateTextBox txtDate
End If

'CustomerName
If Trim(txtCustName) = "" Then
    MsgBox GetResourceString(516), , wis_MESSAGE_TITLE
    ActivateTextBox txtCustName
    Exit Sub
End If
If txtAmount = 0 Then
    MsgBox GetResourceString(499), , wis_MESSAGE_TITLE
    ActivateTextBox txtAmount
    Exit Sub
End If
If Trim(txtVoucherNo) = "" Then
    MsgBox GetResourceString(755), , wis_MESSAGE_TITLE
    ActivateTextBox txtVoucherNo
    Exit Sub
End If

If Trim(cmbParticulars.Text) = "" Then
    MsgBox GetResourceString(621), , wis_MESSAGE_TITLE
    ActivateTextBox cmbParticulars
    Exit Sub
End If


'Now Transact the amount
Dim AccHeadID As Long
With cmbAccHead
    If .ListIndex >= 0 Then _
        AccHeadID = .ItemData(.ListIndex)
End With

If cmbTrans.ListIndex = 0 Then
    RaiseEvent AddAmount(AccHeadID, Val(txtAccNo.Tag), Val(txtCustName.Tag), _
                 txtCustName, GetSysFormatDate(txtDate), txtAmount)
Else
    RaiseEvent RemoveAmount(AccHeadID, CLng(Val(txtAccNo.Tag)), _
                  txtCustName, GetSysFormatDate(txtDate), txtAmount, CLng(Val(cmdAccept.Tag)))

End If
Call ClearControl

End Sub

Private Sub cmdAccNo_Click()
Dim AccHeadID As Long

With cmbAccHead
    If .ListIndex < 0 Then Exit Sub
    AccHeadID = .ItemData(.ListIndex)
    If Trim$(txtAccNo) = "" Then Exit Sub
End With
Dim rst As Recordset

Set rst = GetAccRecordSet(AccHeadID, txtAccNo)
If rst Is Nothing Then Exit Sub

txtCustName = FormatField(rst("CustName"))
txtCustName.Tag = FormatField(rst("CustomerID"))

End Sub

Private Sub cmdCust_Click()
Dim rst As Recordset
Dim AccHeadID As Long
Dim CustName As String

If cmbAccHead.ListIndex >= 0 Then _
    AccHeadID = cmbAccHead.ItemData(cmbAccHead.ListIndex)


gDbTrans.SqlStmt = "Select CustomerID , FirstName +' ' + MiddleName " & _
            "+' '+LastName as CustNAme From NameTab"
If Len(Trim(txtCustName)) > 0 Then _
    gDbTrans.SqlStmt = gDbTrans.SqlStmt & _
        " Where FirstName like " & AddQuotes(txtCustName & "%")
            
Call gDbTrans.Fetch(rst, adOpenDynamic)

If rst Is Nothing Then Exit Sub
If m_frmLookUp Is Nothing Then Set m_frmLookUp = New frmLookUp

Call FillViewNew(m_frmLookUp.lvwReport, rst, "CustomerID")
m_retVar = ""
m_frmLookUp.Show vbModal
If m_retVar = "" Then Exit Sub
rst.MoveFirst
rst.Find "CustomerID = " & m_retVar

txtCustName.Tag = m_retVar
txtCustName = FormatField(rst("CustName"))


End Sub

Private Sub cmdOk_Click()
Unload Me

End Sub

Private Sub cmdUndo_Click()

If Val(cmdAccept.Tag) Then Call ClearControl: Exit Sub

RaiseEvent UndoTrans

End Sub

Private Sub Form_Load()

Screen.MousePointer = vbHourglass
'set icon for the form caption

Icon = LoadResPicture(161, vbResIcon)

'Show the form at the centre of screen
Call CenterMe(Me)
' Set Kannada caption
Call SetKannadaCaption


'Fill up transaction Types
With cmbTrans
    .AddItem GetResourceString(271) 'Deposit
    .AddItem GetResourceString(272) 'Withdraw
End With

'Fill up particulars with default values from SBAcc.INI
    Dim Particulars As String
    Dim I As Integer
    Do
        Particulars = ReadFromIniFile("Particulars", _
                            "Key" & I, gAppPath & "\SuspAcc.INI")
        If Trim$(Particulars) <> "" Then cmbParticulars.AddItem Particulars
        I = I + 1
    Loop Until Trim$(Particulars) = ""

Call LoadLedgersToCombo(cmbAccHead, parMemberShare)
Call LoadLedgersToCombo(cmbAccHead, parMemberDeposit, False)
Call LoadLedgersToCombo(cmbAccHead, parMemberLoan, False)
Call LoadLedgersToCombo(cmbAccHead, parMemDepLoan, False)

Call PassBookPageInitialize
'gDbTrans.SQLStmt = "Select * From SuspAccount Where Cleared = 0 " & _
            " And (TransType = " & wDeposit & _
                " OR TransType = " & wContraDeposit & ")" & _
            " Order By TransID"
gDbTrans.SqlStmt = "Select * From SuspAccount"
If gDbTrans.Fetch(m_rstTrans, adOpenDynamic) < 1 Then Set m_rstTrans = Nothing
Call PassBookPageShow

Screen.MousePointer = vbDefault

txtDate = DayBeginDate

If gOnLine Then
    txtDate.Locked = True
    cmdDate.Enabled = False
End If

End Sub

'
Private Function PassBookPageInitialize()

With grd
    .Rows = 1
    .Clear: .Rows = 11: .FixedRows = 1
    .Cols = 5: .FixedCols = 0
    
    .Row = 0
    .Col = 0: .Text = GetResourceString(37): .ColWidth(0) = 1150      '"Date"
    .Col = 1: .Text = GetResourceString(39): .ColWidth(1) = 1850       '"Particulars"
    .Col = 2: .Text = GetResourceString(276): .ColWidth(2) = 1000     '"Debit"
    .Col = 3: .Text = GetResourceString(277): .ColWidth(3) = 1000     '"Credit"
    .Col = 4: .Text = GetResourceString(42): .ColWidth(4) = 1050      '"Balance"

End With
     
End Function

'
Private Sub PassBookPageShow()
Const RecorsdToShow = 10

' If no recordset, exit.
If m_rstTrans Is Nothing Then Exit Sub
' If no records, exit.
If m_rstTrans.RecordCount = 0 Then Exit Sub
Dim transType As wisTransactionTypes
Dim TransID  As Long

TransID = m_rstTrans("TransID")
cmdPrevTrans.Enabled = False
If TransID > 10 Then cmdPrevTrans.Enabled = True
Dim CustName As String
Dim CustClass As New clsCustReg

'Show 10 records or till eof of the page being pointed to
With grd
    Call PassBookPageInitialize
    .Visible = False
    .Row = 1
    Do
        .RowData(.Row) = m_rstTrans("TransId")
        .Col = 0: .Text = FormatField(m_rstTrans("TransDate"))
        CustName = Trim$(FormatField(m_rstTrans("CustName")))
        If Len(CustName) = 0 Then _
            CustName = CustClass.CustomerName(m_rstTrans("customerID"))
        .Col = 1: .Text = CustName
        transType = m_rstTrans("TransType")
        .Col = 2
        If transType = wWithdraw Or transType = wContraWithdraw Then .Col = 3
        .Text = FormatField(m_rstTrans("Amount"))
        .Col = 4: .Text = FormatField(m_rstTrans("Balance"))
nextRecord:
        m_rstTrans.MoveNext
        If m_rstTrans.EOF Then Exit Do
        If .Row = .Rows - 1 Then Exit Do
        .Row = .Row + 1
    Loop
    .Visible = True
    .Row = 1
End With

Set CustClass = Nothing

cmdNextTrans.Enabled = Not m_rstTrans.EOF
cmdUndo.Enabled = True

If m_rstTrans.RecordCount < 10 Then
    cmdPrevTrans.Enabled = False
    cmdNextTrans.Enabled = False
End If

End Sub


Private Sub cmdNextTrans_Click()

If m_rstTrans Is Nothing Then Exit Sub
If m_rstTrans.EOF Then cmdNextTrans.Enabled = False: Exit Sub
If m_rstTrans Is Nothing Then
    cmdPrevTrans.Enabled = False
    cmdNextTrans.Enabled = False
    Exit Sub
End If

Call PassBookPageShow

End Sub




Private Sub cmdPrevTrans_Click()

If m_rstTrans Is Nothing Then
    cmdPrevTrans.Enabled = False
    cmdNextTrans.Enabled = False
    Exit Sub
End If

If m_rstTrans.EOF And m_rstTrans.BOF Then Exit Sub
Dim TransID As Long

If m_rstTrans.EOF Then
    m_rstTrans.MoveLast
    TransID = m_rstTrans.AbsolutePosition
    If TransID Mod 10 = 0 Then TransID = TransID - 10
    TransID = TransID - TransID Mod 10
    TransID = TransID - 10
    If TransID < 1 Then TransID = 1
Else
    TransID = m_rstTrans.AbsolutePosition
    TransID = TransID - TransID Mod 10
    TransID = TransID - 20
    If TransID < 1 Then TransID = 1
End If
m_rstTrans.MoveFirst
m_rstTrans.Move TransID

Call PassBookPageShow

End Sub



Private Sub SetKannadaCaption()

Call SetFontToControls(Me)
    
    'Now Assign The Names to the Controls
    'The Below Code load From The the resource file
  
    
' TransCtion Frame
lblAccHead.Caption = GetResourceString(36, 92)
lblAccNO.Caption = GetResourceString(36, 60)
lblDate.Caption = GetResourceString(37)
lblTrans.Caption = GetResourceString(38)
lblParticular.Caption = GetResourceString(39)
lblAmount.Caption = GetResourceString(40)
lblVoucher.Caption = GetResourceString(41)  'Voucher No

cmdAccept.Caption = GetResourceString(4)
cmdUndo.Caption = GetResourceString(19)

cmdOk.Caption = GetResourceString(1)    '"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
RaiseEvent WindowClosed(Me.hwnd)
End Sub

Private Sub grd_DblClick()
Dim TransID As Long
TransID = grd.RowData(grd.Row)
If TransID < 1 Then Exit Sub

'Now Repay this amount to the customer
Dim rst As Recordset
gDbTrans.SqlStmt = "Select * from SuspAccount" & _
                   " Where TransID = " & TransID

If gDbTrans.Fetch(rst, adOpenDynamic) < 1 Then Exit Sub

'if this transaction alredy cleered then need load
If Val(FormatField(rst("Cleared"))) <> 0 Then Exit Sub
'if this transctin is of Payment then also do not load it
If rst("TransType") = wWithdraw Or _
        rst("TransType") = wContraWithdraw Then Exit Sub


Dim CustName As String
Dim CustClass As clsCustReg
Dim AccHeadID As Long
Dim I As Long

CustName = FormatField(rst("CustName"))
AccHeadID = FormatField(rst("AccHeadID"))
txtCustName.Tag = FormatField(rst("CustomerID"))
txtAmount.Value = rst("Amount")

cmbTrans.ListIndex = 1
With cmbAccHead
    .ListIndex = -1
    For I = 0 To .ListCount - 1
        If AccHeadID = .ItemData(I) Then
            .ListIndex = I
            Exit For
        End If
    Next
    If .ListIndex >= 0 Then _
       txtAccNo = GetAccountNumber(AccHeadID, rst("AccID"))
End With

Set CustClass = New clsCustReg
If Len(CustName) = 0 And Val(txtCustName.Tag) > 0 Then _
            txtCustName = CustClass.CustomerName(Val(txtCustName.Tag))
        
Set CustClass = Nothing

cmdAccept.Caption = GetResourceString(20)
cmdAccept.Tag = TransID
cmdUndo.Caption = GetResourceString(8)


End Sub

Private Sub m_frmLookUp_SelectClick(strSelection As String)
m_retVar = strSelection

End Sub


