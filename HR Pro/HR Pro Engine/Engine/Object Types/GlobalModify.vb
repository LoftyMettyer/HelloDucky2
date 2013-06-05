' This is a different name for the global adds/updates/deletes. Chnaged name to avoid confusing with Globals in asp.net

Namespace Things

  <Serializable()>
  Public Class GlobalModify
    Inherits Base

    Public Property Items As New List(Of GlobalModifyItem)
  End Class

End Namespace
