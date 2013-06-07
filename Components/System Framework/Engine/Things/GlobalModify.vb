' This is a different name for the global adds/updates/deletes. Chnaged name to avoid confusing with Globals in asp.net

Namespace Things

  <Serializable()>
  Public Class GlobalModify
    Inherits Base

    Public Property Items As ICollection(Of GlobalModifyItem)

    Public Sub New()
      Items = New Collection(Of GlobalModifyItem)
    End Sub

  End Class

End Namespace
