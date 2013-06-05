Namespace Things

  Public Class View
    Inherits Base

    Public Property Filter As Expression
    Public Property Columns As New List(Of Column)

    Public Overrides ReadOnly Property Type As Things.Enums.Type
      Get
        Return Enums.Type.View
      End Get
    End Property

  End Class
End Namespace
