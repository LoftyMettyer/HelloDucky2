Namespace Things
  <Serializable()> _
  Public Class Relation
    Inherits Things.Base

    Public Property RelationshipType As ScriptDB.RelationshipType
    Public Property ParentID As Integer
    Public Property ChildID As Integer

    Public Property DependantColumns As New Things.Collections.Generic

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
