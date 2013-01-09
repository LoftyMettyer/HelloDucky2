Namespace Things

  <Serializable()>
  Public Class TableOrder
    Inherits Base

    Public Property Table As Table
    Public Property Items As ICollection(Of TableOrderItem)

    Public Sub New()
      Items = New Collection(Of TableOrderItem)
    End Sub

  End Class

End Namespace
