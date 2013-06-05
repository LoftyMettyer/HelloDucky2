Namespace Things

  <Serializable()> _
  Public Class Index
    Inherits Things.Base

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Index
      End Get
    End Property

    Public IsClustered As Boolean = False
    Public Columns As IList(Of Column)
    Public IncludedColumns As IList(Of Column)
    Public Relations As IList(Of Relation)
    Public IsTableIndex As Boolean = False
    Public IncludePrimaryKey As Boolean = True
    Public Enabled As Boolean = True
    Public IsUnique As Boolean = False

    Public Sub New()
      Columns = New List(Of Column)
      IncludedColumns = New List(Of Column)
      Relations = New List(Of Relation)
    End Sub

  End Class

End Namespace
