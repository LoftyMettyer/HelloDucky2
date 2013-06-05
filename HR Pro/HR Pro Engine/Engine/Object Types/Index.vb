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
    Public Columns As Things.Collections.Generic
    Public IncludedColumns As Things.Collections.Generic
    Public Relations As Things.Collections.Generic
    Public IsTableIndex As Boolean = False
    Public IncludePrimaryKey As Boolean = True
    Public Enabled As Boolean = True
    Public IsUnique As Boolean = False

    Public Sub New()
      Columns = New Things.Collections.Generic
      IncludedColumns = New Things.Collections.Generic
      Relations = New Things.Collections.Generic
    End Sub

  End Class

End Namespace
