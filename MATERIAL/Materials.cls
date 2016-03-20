VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaterials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsMaterial"
Attribute VB_Ext_KEY = "Member0" ,"clsMaterial"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

  Option Explicit

'local variable to hold collection
Private mCol As Dictionary

Public Property Get ObjectName()
    ObjectName = "MaterialCOlClass"
End Property


Public Function Add(ByVal Quantity As Double, ByVal UnitPrice As Currency, _
                    ByVal TradingPrice As Currency, ByVal MRP As Currency, ByVal RelationID As Long, _
                    ByVal Manufacturer As String, ByVal Amount As Currency, ByVal TaxAmount As Currency, _
                    ByVal ProductName As String, ByVal ProductGroupName As String, ByVal ParentID As Long, _
                    Optional Soot As Double, _
                    Optional ByVal sKey As String) As clsMaterial
    'create a new object
    Dim objNewMember As clsMaterial
    
    Set objNewMember = New clsMaterial


    'set the properties passed into the method
    objNewMember.Quantity = Quantity
    objNewMember.UnitPrice = UnitPrice
    objNewMember.TradingPrice = TradingPrice
    objNewMember.MRP = MRP
    objNewMember.RelationID = RelationID
    objNewMember.Manufacturer = Manufacturer
    objNewMember.Amount = Amount
    objNewMember.ProductName = ProductName
    objNewMember.Soot = Soot
    objNewMember.ParentID = ParentID
    objNewMember.TaxAmount = TaxAmount
    objNewMember.ProductGroupName = ProductGroupName
    
    mCol.Add "MAT" & RelationID, objNewMember
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function


Public Property Get Item(vntIndexKey As Variant) As clsMaterial
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    count = mCol.count
End Property


'this function checks if the given key exists in the collection
' if exists returns true
Public Function Exists(KeyToSearch As String) As Boolean

On Error GoTo Hell:

Exists = False

Exists = mCol.Exists(KeyToSearch)
  
Exit Function

Hell:

End Function
Public Function GetKeys() As Variant

    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    
  GetKeys = mCol.Keys
  
  
End Function

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

