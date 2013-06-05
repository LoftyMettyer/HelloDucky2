Namespace Things

  Public Class Datasource
    Inherits Things.Base

    Public Provider As String
    Public Login As Connectivity.Login

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.DataSource
      End Get
    End Property

  End Class
End Namespace
