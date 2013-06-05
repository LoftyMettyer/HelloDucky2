' This is a different name for the global adds/updates/deletes. Chnaged name to avoid confusing with Globals in asp.net

Namespace Things

  <Serializable()> _
  Public Class GlobalModify
    Inherits Things.Base

    Public Property GlobalModifyItems As New List(Of GlobalModifyItem)

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.GlobalModify
      End Get
    End Property
  End Class

End Namespace
