VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOperatorDefParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variable to hold collection.
Private mvar_colParameters As Collection

Public Function Add(piIndex As Integer, _
  piType As Integer) As clsOperatorDefParameter
  Dim objNewMember As clsOperatorDefParameter
  
  ' Create a new object.
  Set objNewMember = New clsOperatorDefParameter

  With objNewMember
    .Index = piIndex
    .ParameterType = piType
  End With
  
  mvar_colParameters.Add objNewMember, "P" & Trim(Str(piIndex))
  
  ' Return the object created.
  Set Add = objNewMember
  
  ' Disassociate object variables.
  Set objNewMember = Nothing

End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mvar_colParameters.[_NewEnum]
  
End Property

Public Property Get Item(pvIndexKey As Variant) As clsOperatorDefParameter
  Set Item = mvar_colParameters(pvIndexKey)
  
End Property





Public Property Get Count() As Long
  Count = mvar_colParameters.Count
  
End Property

Public Sub Remove(pvIndexKey As Variant)
  mvar_colParameters.Remove pvIndexKey
  
End Sub

Private Sub Class_Initialize()
  ' Create the collection when this class is created.
  Set mvar_colParameters = New Collection
  
End Sub

Private Sub Class_Terminate()
  ' Destroys the collection when this class is terminated.
  If Not mvar_colParameters Is Nothing Then
    Set mvar_colParameters = Nothing
  End If
  
End Sub



Public Property Get Collection() As Collection
  Set Collection = mvar_colParameters

End Property






