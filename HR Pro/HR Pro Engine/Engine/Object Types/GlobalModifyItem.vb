' This is a different name for the global adds/updates/deletes. Chnaged name to avoid confusing with Globals in asp.net

Namespace Things

  <Serializable()> _
  Public Class GlobalModifyItem
    Inherits Things.Base

    Public ColumnID As SystemFramework.HCMGuid
    Public Value As String
    Public CalculationID As SystemFramework.HCMGuid
    Public RefColumnID As SystemFramework.HCMGuid
    Public LookupTableID As SystemFramework.HCMGuid
    Public LookupColumnID As SystemFramework.HCMGuid

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
        Return CStr(CInt(ColumnID))
      End Get
    End Property

  End Class

End Namespace
