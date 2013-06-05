Namespace Things

  <Serializable()> _
  Public Class TableOrderItem
    Inherits Things.Base

    Public Property TableOrder As TableOrder
    Public Property ColumnType As String
    Public Property Order As Enums.Order
    Public Property Column As Things.Column
    Public Property Ascending As Enums.Order
    Public Property Sequence As Integer

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.TableOrderItem
      End Get

    End Property
  End Class
End Namespace
