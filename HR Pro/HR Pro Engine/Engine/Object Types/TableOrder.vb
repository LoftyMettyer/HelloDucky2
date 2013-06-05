Namespace Things

  <Serializable()>
  Public Class TableOrder
    Inherits Base

    Public Property Table As Table
    Public Property Items As New List(Of TableOrderItem)

  End Class

End Namespace
