VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSummaryFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection
Private msTag As String
Private mbManualColumnBreak As Boolean

Public Property Get Tag() As String
  ' Return the object's tag.
  Tag = msTag
  
End Property

Public Property Let Tag(ByVal psNewValue As String)
  ' Set the object's tag property.
  msTag = psNewValue
  
End Property

Private Sub Class_Initialize()
  Set mCol = New Collection

End Sub


Private Sub Class_Terminate()
  Set mCol = Nothing

End Sub


Public Function Add(psKey As String) As clsSummaryField
  Dim objNewMember As clsSummaryField
  
  Set objNewMember = New clsSummaryField
  mCol.Add objNewMember, psKey

  ' Return the object created
  Set Add = objNewMember
  Set objNewMember = Nothing

End Function


Public Property Get Collection() As Collection
  Set Collection = mCol
  
End Property


Public Property Let Collection(ByVal pcolNewValue As Collection)
  Set mCol = pcolNewValue
  
End Property

Public Property Get Count() As Long
  Count = mCol.Count
    
End Property



Public Property Get Item(vntIndexKey As Variant) As clsSummaryField
  Set Item = mCol(vntIndexKey)
    
End Property



Public Property Get NewEnum() As IUnknown
  Set NewEnum = mCol.[_NewEnum]
    
End Property



Public Sub Remove(vntIndexKey As Variant)
  mCol.Remove vntIndexKey
    
End Sub

Public Property Get ManualColumnBreak() As Boolean
  ManualColumnBreak = mbManualColumnBreak
End Property

Public Property Let ManualColumnBreak(ByVal pbNewValue As Boolean)
  mbManualColumnBreak = pbNewValue
End Property

Public Property Get ColumnBreakPoint() As Integer
  
  'Loop through summary fields and get the column break
  Dim objSummaryField As clsSummaryField
  Dim iBreakPoint As Integer

  iBreakPoint = 0
  For Each objSummaryField In mCol
    If objSummaryField.StartOfColumn Then
      ColumnBreakPoint = iBreakPoint
    End If
    iBreakPoint = iBreakPoint + 1
  Next objSummaryField

End Property
