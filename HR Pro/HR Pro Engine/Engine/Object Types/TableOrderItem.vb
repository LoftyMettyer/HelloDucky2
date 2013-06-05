Namespace Things

  <Serializable()> _
  Public Class TableOrderItem
    Inherits Things.Base

    Public ColumnType As String
    Public Order As Enums.Order
    Public Column As Things.Column
    Public Ascending As Enums.Order
    Public Sequence As Integer

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.TableOrderItem
      End Get

    End Property
  End Class
End Namespace
