Namespace Things

  <Serializable()> _
  Public Class TableOrder
    Inherits Things.Base

    Public Property TableOrderItems As New List(Of TableOrderItem)

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.TableOrder
      End Get
    End Property

  End Class

End Namespace
