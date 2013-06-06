Namespace Things

  <Serializable()>
  Public Class TableOrderItem
    Inherits Base

    Public Property TableOrder As TableOrder
    Public Property Column As Column
    Public Property ColumnType As String
    Public Property Sequence As Integer
    Public Property Ascending As Order

  End Class
End Namespace
