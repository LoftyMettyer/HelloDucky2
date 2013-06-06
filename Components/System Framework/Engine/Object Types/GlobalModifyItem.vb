' This is a different name for the global adds/updates/deletes. Chnaged name to avoid confusing with Globals in asp.net

Namespace Things

  <Serializable()>
  Public Class GlobalModifyItem
    Inherits Base

    Public Property ColumnID As Integer
    Public Property Value As String
    Public Property CalculationID As Integer
    Public Property RefColumnID As Integer
    Public Property LookupTableID As Integer
    Public Property LookupColumnID As Integer

    Public ReadOnly Property DisplayValue As String
      Get
        Return Value
      End Get
    End Property

    Public ReadOnly Property DisplayColumn As String
      Get
        Return ColumnID.ToString
      End Get
    End Property

  End Class

End Namespace
