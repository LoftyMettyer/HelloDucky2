Imports SystemFramework.Enums

<Serializable()>
Public Class Relation
  Inherits Base

  Public Property RelationshipType As RelationshipType
  Public Property ParentId As Integer
  Public Property ChildId As Integer

  Public Overrides ReadOnly Property PhysicalName As String
    Get
      Return ScriptDB.Consts.UserTable & Name
    End Get
  End Property

End Class
