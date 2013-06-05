Namespace Things
  Public Class Screen
    Inherits Things.Base

    Public Property Table As Things.Table

    Public Overrides ReadOnly Property Type As Things.Enums.Type
      Get
        Return Things.Type.Screen
      End Get
    End Property
  End Class

End Namespace