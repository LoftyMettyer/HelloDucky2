Namespace Things
  <Serializable()> _
  Public Class Relation
    Inherits Things.Base

    Public RelationshipType As ScriptDB.RelationshipType

    Public ParentID As HCMGuid
    Public ChildID As HCMGuid

    'Public Overrides Function Commit() As Boolean
    'End Function

    Public Overrides ReadOnly Property PhysicalName As String
      Get
        Return ScriptDB.Consts.UserTable & MyBase.Name
      End Get
    End Property

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Relation
      End Get
    End Property
  End Class

End Namespace
