' This is a different name for the global adds/updates/deletes. Chnaged name to avoid confusing with Globals in asp.net

Namespace Things

  <Serializable()> _
  Public Class GlobalModifyItem
    Inherits Things.Base

    Public ColumnID As HRProEngine.HCMGuid
    Public Value As String
    Public CalculationID As HRProEngine.HCMGuid
    Public RefColumnID As HRProEngine.HCMGuid
    Public LookupTableID As HRProEngine.HCMGuid
    Public LookupColumnID As HRProEngine.HCMGuid

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.GlobalModifyItem
      End Get
    End Property

    Public ReadOnly Property DisplayValue As String
      Get
        Return Value
      End Get
    End Property

    Public ReadOnly Property DisplayColumn As String
      Get
        Return CInt(ColumnID)
      End Get
    End Property

  End Class

End Namespace
