Namespace Things
  Public Class Validation
    Inherits Things.Base

    Public ValidationType As Things.ValidationType
    Public Column As Things.Column

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Validation
      End Get
    End Property





  End Class
End Namespace
