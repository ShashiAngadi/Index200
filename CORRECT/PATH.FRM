VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmDbCheck 
   Caption         =   "Form1"
   ClientHeight    =   1410
   ClientLeft      =   2895
   ClientTop       =   2340
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   5280
   Begin MSComDlg.CommonDialog cdb 
      Left            =   420
      Top             =   870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Check"
      Height          =   345
      Left            =   2730
      TabIndex        =   4
      Top             =   720
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Clos&e"
      Height          =   345
      Left            =   3870
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   375
      Left            =   4770
      TabIndex        =   1
      Top             =   240
      Width           =   435
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   1350
      TabIndex        =   0
      Text            =   "c:\indx2000\appmain\index 2000.mdb"
      Top             =   240
      Width           =   3345
   End
   Begin VB.Label lblPath 
      Caption         =   "Data Base Path"
      Height          =   315
      Left            =   30
      TabIndex        =   2
      Top             =   270
      Width           =   1305
   End
End
Attribute VB_Name = "frmDbCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function CorrectdataBase(DbName As String) As Boolean
If Dir(DbName) = "" Then Exit Function

Dim DbObj As Database
Dim TableCount As Integer
Dim Count As Integer
Dim TableName As TableDef
'Dim FieldName As String
Dim FieldCount As Integer
Dim FieldName As Field
Dim Rst As Recordset
Dim rstTarget As Recordset


Set gDBTrans = New clsTransact


If Not gDBTrans.OpenDB(DbName, "WIS!@#") Then Exit Function
Set DbObj = gDBTrans.DbObject


'On Error Resume Next
'**********
Set TableName = DbObj.CreateTableDef("TempTable")

' Create and add the fields.

Dim TheField As Field
With TableName
        ' Create and append the field.
        Set TheField = .CreateField("TransType", dbInteger)
        TheField.Required = False
        .fields.Append TheField
End With

DbObj.TableDefs.Append TableName
'******************

'Close Tjhe Db  & repoen it
gDBTrans.CloseDB
If Not gDBTrans.OpenDB(DbName, "WIS!@#") Then Exit Function
Set DbObj = gDBTrans.DbObject


TableCount = DbObj.TableDefs.Count

For Count = 0 To TableCount - 1
    Set TableName = DbObj.TableDefs(Count)
    'fieldname.CreateProperty (,,,)
    If InStr(1, TableName.Name, "Trans", vbTextCompare) Then
        'DBTrans.SQLStmt = TableName '"select TransType From " & TableName
        For FieldCount = 0 To TableName.fields.Count - 1
            If TableName.fields(FieldCount).Name = "TransType" Then
                If TableName.fields(FieldCount).Type = dbByte Then
                   'TableName.fields(FieldCount).Type = dbInteger
                   gDBTrans.BeginTrans
                   gDBTrans.SQLStmt = "delete * from  TempTable"
                   If Not gDBTrans.SQLExecute Then
                        gDBTrans.RollBack
                        Exit Function
                   End If
                   
                   gDBTrans.SQLStmt = "Insert into Temptable Select TransType from " & TableName.Name
                   If Not gDBTrans.SQLExecute Then
                        gDBTrans.RollBack
                        Exit Function
                   End If
                   gDBTrans.CommitTrans
                   
                   'Now Drop the FIELD tRANTYPE & cREAT IT WITH THE dBINTEGER TYPE
                   TableName.fields.Delete "TransType"
                   'DbObj.TableDefs.Append TableName
                   'Now Create The same field
                    With TableName
                        ' Create and append the field.
                        Set TheField = .CreateField("TransType", dbInteger)
                        TheField.Required = False
                        .fields.Append TheField
                    End With

                    '"now Insert the Values Back to the table
                    gDBTrans.SQLStmt = "Select TransType from TempTable"
                    Call gDBTrans.SQLFetch
                    Set Rst = Nothing
                    Set Rst = gDBTrans.Rst.Clone
                    Set rstTarget = DbObj.OpenRecordset(TableName.Name)
                    gDBTrans.BeginTrans
                    Do
                        If Rst.EOF Then Exit Do
                        If FormatField(Rst(0)) = 255 Then
                            TransType = wWithdraw
                        ElseIf FormatField(Rst(0)) = 254 Then
                            TransType = wCharges
                        Else
                            TransType = FormatField(Rst(0))
                        End If
                        rstTarget.Edit
                        rstTarget!TransType = TransType
                        rstTarget.Update
                        Rst.MoveNext
                        rstTarget.MoveNext
                    Loop
                    gDBTrans.CommitTrans

                End If
            End If
        Next
        
    End If
NextTable:

Next Count
gDBTrans.CloseDB
If Not gDBTrans.OpenDB(DbName, "WIS!@#") Then Exit Function
Set DbObj = gDBTrans.DbObject
DbObj.TableDefs.Delete "TempTable"
Set DbObj = Nothing
CorrectdataBase = True
gDBTrans.CloseDB
End Function

Private Sub cmdCheck_Click()
If Trim(txtPath.Text) = "" Then Exit Sub
'Now CHeck for the existance of mdb
If Dir(txtPath.Text) = "" Then
    MsgBox "File Not existing", vbInformation, "Error - DataBase"
    Exit Sub
End If
If Not CorrectdataBase(txtPath) Then
    MsgBox "unable to rectify the DataBase" & vbCrLf & _
        "Please Inform to theh vendor", vbExclamation, "Database _Error"
Else
    MsgBox "DataBase Recified ", vbInformation, "Coorected"
End If

End Sub

Private Sub cmdPath_Click()
cdb.Filter = "database files(*.mdb)|*.mdb|"
cdb.CancelError = False
cdb.DialogTitle = "Choose the DataBase"
On Error Resume Next
cdb.InitDir = "c:\program Files\Index 2000"
cdb.ShowOpen
If Trim(cdb.FileName) = "" Then Exit Sub
Dim FileName As String
FileName = cdb.FileName
txtPath.Text = FileName
cmdCheck.Enabled = True
'cmdClose.Caption = "Correct"


End Sub

