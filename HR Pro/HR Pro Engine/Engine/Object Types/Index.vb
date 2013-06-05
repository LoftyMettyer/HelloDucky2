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
    Public Columns As Things.Collection
    Public IncludedColumns As Things.Collection
    Public Relations As Things.Collection
    Public IsTableIndex As Boolean = False
    Public IncludePrimaryKey As Boolean = True
    Public Enabled As Boolean = True

    Public Sub New()
      Columns = New Things.Collection
      IncludedColumns = New Things.Collection
      Relations = New Things.Collection
    End Sub

  End Class

End Namespace
