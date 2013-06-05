Namespace Things
  <Serializable()>
  Public Class Relation
    Inherits Base

    Public Property RelationshipType As ScriptDB.RelationshipType
    Public Property ParentID As Integer
    Public Property ChildID As Integer

    Public Overrides ReadOnly Property PhysicalName As String
      Get
        Return ScriptDB.Consts.UserTable & MyBase.Name
      End Get
    End Property

  End Class

End Namespace
