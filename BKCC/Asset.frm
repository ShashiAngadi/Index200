VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAsset 
   Caption         =   "Member asset details"
   ClientHeight    =   6450
   ClientLeft      =   825
   ClientTop       =   1455
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   8850
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   7380
      TabIndex        =   33
      Top             =   6030
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grd 
      Height          =   2115
      Left            =   120
      TabIndex        =   32
      Top             =   3840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3731
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3285
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   8625
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   400
         Left            =   5910
         TabIndex        =   31
         Top             =   2820
         Width           =   1365
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   400
         Left            =   7410
         TabIndex        =   30
         Top             =   2820
         Width           =   1125
      End
      Begin VB.TextBox txtNoOfIpSet 
         Height          =   345
         Left            =   3690
         TabIndex        =   27
         Top             =   2280
         Width           =   825
      End
      Begin VB.TextBox txtNoOfWell 
         Height          =   315
         Left            =   1620
         TabIndex        =   25
         Top             =   2280
         Width           =   705
      End
      Begin VB.TextBox txtRemarks 
         Height          =   345
         Left            =   6510
         TabIndex        =   29
         Top             =   2310
         Width           =   1995
      End
      Begin VB.TextBox txtRiver 
         Height          =   345
         Left            =   7590
         TabIndex        =   23
         Top             =   1740
         Width           =   915
      End
      Begin VB.TextBox txtCanal 
         Height          =   315
         Left            =   5670
         TabIndex        =   21
         Top             =   1740
         Width           =   885
      End
      Begin VB.TextBox txtValue 
         Height          =   345
         Left            =   6780
         TabIndex        =   14
         Top             =   690
         Width           =   1725
      End
      Begin VB.TextBox txtAccNum 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox txtWell 
         Height          =   345
         Left            =   3600
         TabIndex        =   19
         Top             =   1680
         Width           =   885
      End
      Begin VB.TextBox txtDry 
         Height          =   315
         Left            =   1530
         TabIndex        =   17
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox txtTax 
         Height          =   345
         Left            =   3600
         TabIndex        =   12
         Top             =   690
         Width           =   1725
      End
      Begin VB.TextBox txtArea 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   690
         Width           =   1125
      End
      Begin VB.TextBox txtSurveyNo 
         Height          =   345
         Left            =   6780
         TabIndex        =   8
         Top             =   240
         Width           =   1725
      End
      Begin VB.ComboBox cmbPlace 
         Height          =   315
         Left            =   3600
         TabIndex        =   6
         Top             =   240
         Width           =   1725
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   90
         X2              =   8760
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Label lblNoOfWell 
         Caption         =   "No Of wells"
         Height          =   300
         Left            =   120
         TabIndex        =   24
         Top             =   2310
         Width           =   1155
      End
      Begin VB.Label lblRemarks 
         Caption         =   "Remarks"
         Height          =   300
         Left            =   5040
         TabIndex        =   28
         Top             =   2370
         Width           =   1515
      End
      Begin VB.Label lblNoOfIPSet 
         Caption         =   "No Of pumpset"
         Height          =   300
         Left            =   2460
         TabIndex        =   26
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label lblRiver 
         Caption         =   "By  River"
         Height          =   300
         Left            =   6690
         TabIndex        =   22
         Top             =   1740
         Width           =   825
      End
      Begin VB.Label lblCanal 
         Caption         =   "By Canal"
         Height          =   300
         Left            =   4560
         TabIndex        =   20
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label lblPloughDetails 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plough Land Details"
         Height          =   315
         Left            =   90
         TabIndex        =   15
         Top             =   1260
         Width           =   5235
      End
      Begin VB.Label lblValue 
         Caption         =   "Value"
         Height          =   300
         Left            =   5460
         TabIndex        =   13
         Top             =   690
         Width           =   1245
      End
      Begin VB.Label lblAccNO 
         Caption         =   "Account No"
         Height          =   300
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label lblWell 
         Caption         =   "By Well"
         Height          =   300
         Left            =   2460
         TabIndex        =   18
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label lblDry 
         Caption         =   "By Rain"
         Height          =   300
         Left            =   90
         TabIndex        =   16
         Top             =   1710
         Width           =   1425
      End
      Begin VB.Label lblTax 
         Caption         =   "Tax"
         Height          =   300
         Left            =   2550
         TabIndex        =   11
         Top             =   690
         Width           =   945
      End
      Begin VB.Label lblArea 
         Caption         =   "Area"
         Height          =   300
         Left            =   90
         TabIndex        =   9
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label lblSurveyNO 
         Caption         =   "Survey number"
         Height          =   300
         Left            =   5400
         TabIndex        =   7
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label lblPlace 
         Caption         =   "Place"
         Height          =   300
         Left            =   2550
         TabIndex        =   5
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.Label txtMemName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Name of the member"
      Height          =   345
      Left            =   1890
      TabIndex        =   1
      Top             =   90
      Width           =   6735
   End
   Begin VB.Label lblMemName 
      Caption         =   "Member Name"
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "frmAsset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_FormLoaded As Boolean
Private m_CustomerID As Long
Private m_dbOperation As wis_DBOperation
Private m_AssetId As Integer

Public Event AssetDetailsAdded(DryLand As String, IrrigationLand As String)
Public Event AssetDetailsRemoved(DryLand As String, IrrigationLand As String)

Public Property Let customerID(NewValue As Long)
m_CustomerID = NewValue
If m_FormLoaded Then Call LoadAssetDetails
    
End Property

Private Sub LoadAssetDetails()
If m_CustomerID = 0 Then Exit Sub

''Clear the for
Call ResetUserInteface

'Now Fetch the asset Details of the customer
Dim rstAsset As Recordset
Dim rstTemp As Recordset
gDbTrans.SqlStmt = "Select * From AssetDetails " & _
                " WHERE CustomerID = " & m_CustomerID
                
If gDbTrans.Fetch(rstAsset, adOpenStatic) < 0 Then Exit Sub

'Now Fill the Details
'Now Get the CustomerNAme
gDbTrans.SqlStmt = "Select Title +' '+FirstName+' '+MiddleNAme+' '+LastNAme as CustNAme" & _
                " From NameTab WHERE CustomerID = " & m_CustomerID
Call gDbTrans.Fetch(rstTemp, adOpenDynamic)
txtMemName.Caption = Trim(FormatField(rstTemp("CustName")))
Set rstTemp = Nothing
Dim AssetID As Integer
'now Load the Asset detials to the grid
While Not rstAsset.EOF
    With grd
        AssetID = FormatField(rstAsset("AssetID"))
        If .Rows <= AssetID + 1 Then .Rows = AssetID + 2
        .Row = AssetID + 1
        .RowData(.Row) = AssetID
        .Col = 0: .Text = FormatField(rstAsset("AccNum"))
        .Col = 1: .Text = FormatField(rstAsset("PLace"))
        .Col = 2: .Text = FormatField(rstAsset("SurveyNo"))
        .Col = 3: .Text = FormatField(rstAsset("LandArea"))
        .Col = 4: .Text = FormatField(rstAsset("Tax"))
        .Col = 5: .Text = FormatField(rstAsset("LandValue"))
        .Col = 6: .Text = FormatField(rstAsset("DryLand"))
        .Col = 7: .Text = FormatField(rstAsset("WellLand"))
        .Col = 8: .Text = FormatField(rstAsset("CanalLand"))
        .Col = 9: .Text = FormatField(rstAsset("RiverLand"))
        .Col = 10: .Text = FormatField(rstAsset("NoOfWell"))
        .Col = 11: .Text = FormatField(rstAsset("NoOfIpSet"))
        .Col = 12: .Text = FormatField(rstAsset("Remarks"))
    End With
    rstAsset.MoveNext
    
Wend


End Sub


'This function clears the Asset entry controls
Private Sub ResetAssetDetails()

txtAccNum.Text = ""
cmbPlace.Text = ""
txtSurveyNo.Text = ""
txtArea.Text = ""
txtTax.Text = ""
txtValue.Text = ""
txtDry.Text = ""
txtWell.Text = ""
txtCanal.Text = ""
txtRiver.Text = ""

txtNoOfWell.Text = ""
txtNoOfIpSet.Text = ""
txtRemarks.Text = ""

cmdAdd.Caption = GetResourceString(10)
cmdRemove.Caption = GetResourceString(8)

m_dbOperation = Insert
m_AssetId = 0
End Sub

'This function Reset the user inteface
'This Clears all the Text Field of the form,Grid
'and also set Grid rows /cols
Private Sub ResetUserInteface()

txtMemName.Caption = ""
'Now Clear the Asset entry details
Call ResetAssetDetails

With grd
    .Rows = 2
    .Clear
    .Rows = 3: .Cols = 13
    .FixedCols = 1: .FixedRows = 2
    .MergeCells = flexMergeFree
    .AllowUserResizing = flexResizeBoth
    
    .Row = 0:
'    .Col = 0: .Text = GetResourceString(49,329)
    .Col = 0: .Text = GetResourceString(36, 60)
    .Col = 1: .Text = GetResourceString(49, 329)
    .Col = 2: .Text = GetResourceString(49, 329)
    .Col = 3: .Text = GetResourceString(49, 329)
    .Col = 4: .Text = GetResourceString(49, 329)
    .Col = 5: .Text = GetResourceString(140)
    .Col = 6: .Text = GetResourceString(88, 329)
    .Col = 7: .Text = GetResourceString(88, 329)
    .Col = 8: .Text = GetResourceString(88, 329)
    .Col = 9: .Text = GetResourceString(88, 329)
    .Col = 10: .Text = GetResourceString(400, 60)
    .Col = 11: .Text = "IP Set " & GetResourceString(60)
    .Col = 12: .Text = GetResourceString(261)
    .MergeRow(0) = True
    .MergeCol(0) = True
    Dim I As Integer, MaxI As Integer
    MaxI = .Cols - 1
    For I = 0 To MaxI
        .Col = I
        .CellAlignment = 4
        .CellFontBold = True
    Next

    .Row = 1:
    .Col = 0: .Text = GetResourceString(36, 60)
    .Col = 1: .Text = GetResourceString(112)
    .Col = 2: .Text = GetResourceString(87, 60)
    .Col = 3: .Text = GetResourceString(329)
    .Col = 4: .Text = GetResourceString(173)
    .Col = 5: .Text = GetResourceString(140): .MergeCol(5) = True
    .Col = 6: .Text = GetChangeString(GetResourceString(107), GetResourceString(268))
    .Col = 7: .Text = GetChangeString(GetResourceString(107), GetResourceString(400))
    .Col = 8: .Text = GetChangeString(GetResourceString(107), GetResourceString(234))
    .Col = 9: .Text = GetChangeString(GetResourceString(107), GetResourceString(235))
    .Col = 10: .Text = GetResourceString(400, 60)
    .Col = 11: .Text = "IP Set " & GetResourceString(60)
    .Col = 12: .Text = GetResourceString(261)
    .MergeCol(10) = True
    .MergeCol(11) = True
    .MergeCol(12) = True
    For I = 0 To MaxI
        .Col = 0
        .CellAlignment = 4
        .CellFontBold = True
    Next
    
End With


End Sub


Private Sub SetKannadaCaption()
'First se the Font Name & size to all controls
Call SetFontToControls(Me)

'Now set the Captions to all controls
lblMemName.Caption = GetResourceString(49, 35)
Frame1.Caption = ""
lblAccNO.Caption = GetResourceString(36, 60)
lblPlace.Caption = GetResourceString(112)
lblSurveyNO.Caption = GetResourceString(87, 60)
lblArea.Caption = GetResourceString(329)
lblTax.Caption = GetResourceString(173)
lblValue.Caption = GetResourceString(140)
lblPloughDetails.Caption = GetResourceString(88) & " " & _
        GetResourceString(329, 295)
lblDry.Caption = GetChangeString(GetResourceString(107), GetResourceString(268))
lblWell.Caption = GetChangeString(GetResourceString(107), GetResourceString(400))
lblCanal.Caption = GetChangeString(GetResourceString(107), GetResourceString(234))
lblRiver.Caption = GetChangeString(GetResourceString(107), GetResourceString(235))

lblNoOfWell.Caption = GetResourceString(400, 60)
lblNoOfIPSet.Caption = "IP set " & GetResourceString(60)

lblRemarks.Caption = GetResourceString(261)
cmdRemove.Caption = GetResourceString(12)
cmdAdd.Caption = GetResourceString(10)

cmdOK.Caption = GetResourceString(1)

End Sub



Private Sub cmdAdd_Click()

If m_CustomerID = 0 Then Exit Sub
'FIrst Validate tall the fields

'Check the AccountNo
If Trim(txtAccNum.Text) = "" Then
    'Invalid Account No specified
    MsgBox GetResourceString(36, 60) & " " & _
        GetResourceString(296), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtAccNum
    Exit Sub
End If

If Trim(cmbPlace.Text) = "" Then
    'Invalid place
    MsgBox GetResourceString(112) & " " & _
        GetResourceString(296), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox cmbPlace
    Exit Sub
End If

If Trim(txtSurveyNo.Text) = "" Then
    'Invalid place
    MsgBox GetResourceString(87, 60) & " " & _
        GetResourceString(296), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtSurveyNo
    Exit Sub
End If

If Trim(txtArea.Text) = "" Then
    'Invalid place
    MsgBox GetResourceString(329) & " " & _
        GetResourceString(296), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtArea
    Exit Sub
End If

If Not CurrencyValidate(txtTax.Text, False) Then
    'Invalid place
    MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtTax
    Exit Sub
End If

If Not CurrencyValidate(txtValue.Text, True) Then
    'Invalid place
    MsgBox GetResourceString(499), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtValue
    Exit Sub
End If

'Decale the Variable
Dim strAccNum As String
Dim strPLace As String
Dim strSurveyNo As String
Dim strDry As String
Dim strwell As String
Dim strCanal As String
Dim strRiver As String

strAccNum = txtAccNum.Text
strPLace = cmbPlace.Text
strSurveyNo = txtSurveyNo.Text

strDry = txtDry.Text
strwell = txtWell.Text
strCanal = txtCanal.Text
strRiver = txtRiver.Text

If Val(strDry) + Val(strwell) + Val(strCanal) + Val(strRiver) = 0 Then
    MsgBox GetResourceString(88, 329, 296), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtDry
    Exit Sub
End If

If Not CurrencyValidate(txtNoOfWell, True) Then
    MsgBox GetResourceString(166) & " " & _
        GetResourceString(296), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtNoOfWell
    Exit Sub
End If

If Not CurrencyValidate(txtNoOfIpSet.Text, True) Then
    MsgBox GetResourceString(166) & " " & _
        GetResourceString(296), vbInformation, wis_MESSAGE_TITLE
    ActivateTextBox txtNoOfIpSet
    Exit Sub
End If

Dim AssetID As Integer
Dim rstTemp As Recordset
'First Get the sset ID from th eDB
'now check whether this is updateion or new recor
If m_dbOperation = Insert Then
    gDbTrans.SqlStmt = "SELECT MAx(AssetID) From AssetDetails" & _
                    " WHERE CustomerID = " & m_CustomerID
    If gDbTrans.Fetch(rstTemp, adOpenDynamic) > 0 Then _
            AssetID = FormatField(rstTemp(0))
    AssetID = AssetID + 1
    
    'Now Enter the Detials Into The database
    gDbTrans.SqlStmt = "Insert Into AssetDetails" & _
            " (CustomerID,AssetID,AccNum,Place,SurveyNo," & _
            " LandArea,Tax,LandValue,DryLand,WellLand," & _
            " CanalLand,RiverLand,NoOfWell,NoOfIPSet,Remarks )" & _
                " VALUES (" & m_CustomerID & ", " & AssetID & "," & _
                AddQuotes(strAccNum) & ", " & AddQuotes(strPLace) & "," & _
                AddQuotes(strSurveyNo) & ", " & AddQuotes(Trim(txtArea)) & ", " & _
                Val(txtTax) & ", " & Val(txtValue) & ", " & _
                AddQuotes(Trim(txtDry.Text)) & ", " & AddQuotes(Trim(txtWell)) & ", " & _
                AddQuotes(Trim(txtCanal.Text)) & ", " & AddQuotes(Trim(txtRiver)) & ", " & _
                Val(txtNoOfWell) & ", " & Val(txtNoOfIpSet) & ", " & _
                AddQuotes(Trim(txtRemarks.Text)) & _
            ") "
Else
    AssetID = m_AssetId
    gDbTrans.SqlStmt = "Update AssetDetails" & _
            " Set AccNum = " & AddQuotes(strAccNum) & "," & _
            " Place = " & AddQuotes(strPLace) & "," & _
            " SurveyNo = " & AddQuotes(strSurveyNo) & "," & _
            " LandArea = " & AddQuotes(Trim(txtArea)) & ", " & _
            " Tax = " & Val(txtTax) & ", " & _
            " LandValue = " & Val(txtValue) & ", " & _
            " DryLand = " & AddQuotes(Trim(txtDry.Text)) & ", " & _
            " WellLand = " & AddQuotes(Trim(txtWell)) & ", " & _
            " CanalLand =" & AddQuotes(Trim(txtCanal.Text)) & ", " & _
            " RiverLand = " & AddQuotes(Trim(txtRiver)) & ", " & _
            " NoOfWell = " & Val(txtNoOfWell) & ", " & _
            " NoOfIPSet = " & Val(txtNoOfIpSet) & ", " & _
            " Remarks = " & AddQuotes(Trim(txtRemarks.Text)) & _
            " Where CustomerID = " & m_CustomerID & _
            " AND AssetID = " & AssetID
End If


'Now excute the query
gDbTrans.BeginTrans
If Not gDbTrans.SQLExecute Then
    gDbTrans.RollBack
    Exit Sub
Else
    gDbTrans.CommitTrans
End If

'now Insert the Same InTo grid
With grd
    If .Rows <= AssetID + 1 Then .Rows = AssetID + 2
    .Row = AssetID + 1
    .RowData(.Row) = AssetID
    .Col = 0: .Text = strAccNum
    .Col = 1: .Text = strPLace
    .Col = 2: .Text = strSurveyNo
    .Col = 3: .Text = txtArea
    .Col = 4: .Text = txtTax
    .Col = 5: .Text = txtValue
    .Col = 6: .Text = txtDry
    .Col = 7: .Text = txtWell
    .Col = 8: .Text = txtCanal
    .Col = 9: .Text = txtRiver
    .Col = 10: .Text = txtNoOfWell
    .Col = 11: .Text = txtNoOfIpSet
    .Col = 12: .Text = txtRemarks
End With

'Now Clear the Text boxe
Call ResetAssetDetails

End Sub



Private Sub Command1_Click()
Call ResetUserInteface
End Sub

Private Sub cmdRemove_Click()

If m_dbOperation = Update And m_AssetId Then
    'Delte the Record
    gDbTrans.SqlStmt = "Delete * From THe AssetDetails WHERE " & _
            " CustomerID = " & m_CustomerID & " AND AssetID = " & m_AssetId
    gDbTrans.BeginTrans
    If Not gDbTrans.SQLExecute Then
        gDbTrans.RollBack
        Exit Sub
    End If
    gDbTrans.CommitTrans
    'And Remo the row from grid
    grd.RemoveItem (m_AssetId + 1)
End If

Call ResetAssetDetails

End Sub

Private Sub Form_Load()
'Now set the Captions to the controls
Call SetKannadaCaption

m_FormLoaded = True
If m_CustomerID Then
    Call LoadAssetDetails
Else
    Call ResetUserInteface
End If

'load all the place
Call LoadPlaces(cmbPlace)
End Sub

Private Sub Form_Unload(Cancel As Integer)

m_FormLoaded = False
End Sub


Private Sub grd_DblClick()
If m_CustomerID = 0 Then Exit Sub

'THis Load the Already entered data for updation
Dim rstAsset As Recordset
Dim AssetID As Integer

With grd
    If .Row < .FixedRows Then Exit Sub
    AssetID = .RowData(.Row)
End With

If AssetID = 0 Then Exit Sub

'Now Load the Details To the Text boxes
'now Insert the Same InTo grid
With grd
    
    If .Rows <= AssetID + 1 Then .Rows = AssetID + 2
    .Row = AssetID + 1
    .RowData(.Row) = AssetID
    .Col = 0: txtAccNum = .Text
    .Col = 1: cmbPlace = .Text
    .Col = 2: txtSurveyNo = .Text
    .Col = 3: txtArea = .Text
    .Col = 4: txtTax = .Text
    .Col = 5: txtValue = .Text
    .Col = 6: txtDry = .Text
    .Col = 7: txtWell = .Text
    .Col = 8: txtCanal = .Text
    .Col = 9: txtRiver = .Text
    .Col = 10: txtNoOfWell = .Text
    .Col = 11: txtNoOfIpSet = .Text
    .Col = 12: txtRemarks = .Text
End With

m_dbOperation = Update
m_AssetId = AssetID

cmdAdd.Caption = GetResourceString(171)
cmdRemove.Caption = GetResourceString(12)

End Sub


