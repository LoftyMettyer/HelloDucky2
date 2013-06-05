Namespace Things
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

    'Private msUniqueName As String

    'Public ReadOnly Property UniqueName As String
    '  Get
    '    Return msUniqueName
    '  End Get
    'End Property

    Public Sub New()
      'msUniqueName = "IDX_" & Guid.NewGuid().ToString("N")
      IncludedColumns = New Things.Collection
      Relations = New Things.Collection
    End Sub

  End Class

End Namespace
