Namespace ScriptDB

  Public Class TriggeredUpdate
    Inherits Things.Base

    Public Column As Things.Column
    Public Where As String

    Public Overrides ReadOnly Property Type As Things.Enums.Type
      Get
        Return Things.Type.TriggeredUpdate
      End Get
    End Property
  End Class

End Namespace
