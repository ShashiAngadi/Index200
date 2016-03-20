VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colTransfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsTransferNew"
Attribute VB_Ext_KEY = "Member0" ,"clsTransferNew"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Dictionary

Public Property Get ObjectName()
    ObjectName = "ColTransferClass"
End Property


Public Function Add(ByVal GodownID As Byte, ByVal GodownName As String, ByVal HeadID As Long, _
            ByVal HeadName As String, ByVal RelationID As Long, ByVal TransID As Long, ByVal FromQuantity As Double, _
            ByVal ToQuantity As Double, ByVal Balance As Double, Optional ByVal sKey As String) As clsTransferNew
    
    'create a new object
    Dim objNewMember As clsTransferNew
    Set objNewMember = New clsTransferNew


    'set the properties passed into the method
    objNewMember.GodownID = GodownID
    objNewMember.GodownName = GodownName
    objNewMember.HeadID = HeadID
    objNewMember.HeadName = HeadName
    objNewMember.RelationID = RelationID
    objNewMember.TransID = TransID
    objNewMember.FromQuantity = FromQuantity
    objNewMember.ToQuantity = ToQuantity
    objNewMember.Balance = Balance
    
    mCol.Add sKey, objNewMember


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsTransferNew
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    'Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Dictionary
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Function GetKeys() As Variant

    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    
  GetKeys = mCol.Keys
  
  
End Function


