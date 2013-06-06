' This is a different name for the global adds/updates/deletes. Chnaged name to avoid confusing with Globals in asp.net

Namespace Things

  <Serializable()>
  Public Class GlobalModifyItem
    Inherits Base

    Public Property ColumnId As Integer
    Public Property Value As String
    Public Property CalculationId As Integer
    Public Property RefColumnId As Integer
    Public Property LookupTableId As Integer
    Public Property LookupColumnId As Integer

    Public ReadOnly Property DisplayValue As String
      Get
        Return Value
      End Get
    End Property

    Public ReadOnly Property DisplayColumn As String
      Get
        Return ColumnId.ToString
      End Get
    End Property

  End Class

End Namespace
