VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreenControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection
Private mlngScreenID As Long


Public Function Add(psKey As String) As clsScreenControl
  Dim objNewMember As clsScreenControl
  
  Set objNewMember = New clsScreenControl
  mCol.Add objNewMember, psKey

  'return the object created
  Set Add = objNewMember
  Set objNewMember = Nothing

End Function

Public Property Get Count() As Long
  Count = mCol.Count
    
End Property


Public Property Get Item(vntIndexKey As Variant) As clsScreenControl
  Set Item = mCol(vntIndexKey)
    
End Property


Public Property Get NewEnum() As IUnknown
  Set NewEnum = mCol.[_NewEnum]
    
End Property


Public Sub Remove(vntIndexKey As Variant)
  mCol.Remove vntIndexKey
    
End Sub


Public Property Get ScreenID() As Long
  ' Return the screen's ID.
  ScreenID = mlngScreenID
  
End Property

Public Property Let ScreenID(ByVal plngNewValue As Long)
  ' Set the screen id property.
  mlngScreenID = plngNewValue
  
End Property

Private Sub Class_Initialize()
  Set mCol = New Collection

End Sub


Private Sub Class_Terminate()
  Set mCol = Nothing

End Sub



Public Property Get Collection() As Collection
  Set Collection = mCol
  
End Property

Public Property Let Collection(ByVal pcolNewValue As Collection)
  Set mCol = pcolNewValue
  
End Property
