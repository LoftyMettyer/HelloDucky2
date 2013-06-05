Namespace Things

  Public Class View
    Inherits Things.Base

    Public Filter As Things.Expression

    Public Overrides ReadOnly Property Type As Things.Enums.Type
      Get
        Return Enums.Type.View
      End Get
    End Property


  End Class
End Namespace
