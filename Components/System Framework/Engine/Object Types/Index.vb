Namespace Things

  <Serializable()>
  Public Class Index
    Inherits Base

    Public Property IsClustered As Boolean
    Public Property Columns As ICollection(Of Column)
    Public Property IncludedColumns As ICollection(Of Column)
    Public Property Relations As ICollection(Of Relation)
    Public Property IsTableIndex As Boolean
    Public Property IncludePrimaryKey As Boolean = True
    Public Property Enabled As Boolean = True
    Public Property IsUnique As Boolean

    Public Sub New()
      Columns = New Collection(Of Column)
      IncludedColumns = New Collection(Of Column)
      Relations = New Collection(Of Relation)
    End Sub

  End Class

End Namespace
