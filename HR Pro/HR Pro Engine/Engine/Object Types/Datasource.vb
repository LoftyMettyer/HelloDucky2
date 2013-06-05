Namespace Things

  Public Class Datasource
    Inherits Things.Base

    Public Property Provider As String
    Public Property Login As Connectivity.Login

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.DataSource
      End Get
    End Property

  End Class
End Namespace
