Namespace Things

  <Serializable()>
  Public Class Index
    Inherits Base

    Public Property IsClustered As Boolean
    Public Property Columns As New List(Of Column)
    Public Property IncludedColumns As New List(Of Column)
    Public Property Relations As New List(Of Relation)
    Public Property IsTableIndex As Boolean
    Public Property IncludePrimaryKey As Boolean = True
    Public Property Enabled As Boolean = True
    Public Property IsUnique As Boolean
  End Class

End Namespace
