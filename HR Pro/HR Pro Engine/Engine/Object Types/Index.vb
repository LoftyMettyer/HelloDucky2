Namespace Things

  <Serializable()> _
Public Class Index
    Inherits Things.Base

    Public Property IsClustered As Boolean = False
    Public Property Columns As New List(Of Column)
    Public Property IncludedColumns As New List(Of Column)
    Public Property Relations As New List(Of Relation)
    Public Property IsTableIndex As Boolean
    Public Property IncludePrimaryKey As Boolean = True
    Public Property Enabled As Boolean = True
    Public Property IsUnique As Boolean

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Index
      End Get
    End Property

  End Class

End Namespace
