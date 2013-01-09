Namespace Things

  Public Class View
    Inherits Base

    Property Table As Table
    Property Filter As Expression
    Property Columns As ICollection(Of Column)

    Public Sub New()
      Columns = New Collection(Of Column)
    End Sub

  End Class
End Namespace
