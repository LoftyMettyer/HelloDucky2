Namespace Things

  <Serializable()>
  Public Class TableOrderItem
    Inherits Base

    Public Property TableOrder As TableOrder
    Public Property ColumnType As String
    Public Property Order As Enums.Order
    Public Property Column As Column
    Public Property Ascending As Enums.Order
    Public Property Sequence As Integer

  End Class
End Namespace
