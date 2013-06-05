Namespace Things

  <Serializable()> _
Public Class Index
    Inherits Things.Base

    Public IsClustered As Boolean = False
    Public Columns As New List(Of Column)
    Public IncludedColumns As New List(Of Column)
    Public Relations As New List(Of Relation)
    Public IsTableIndex As Boolean
    Public IncludePrimaryKey As Boolean = True
    Public Enabled As Boolean = True
    Public IsUnique As Boolean

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Index
      End Get
    End Property

  End Class

End Namespace
