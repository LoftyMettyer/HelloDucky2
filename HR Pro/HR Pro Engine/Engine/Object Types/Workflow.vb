Namespace Things
  Public Class Workflow
    Inherits Things.Base

    Public Property Enabled As Boolean
    Public Property InitiationType As Integer
    Public Property Table As Things.Table
    Public Property QueryString As String

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Workflow
      End Get
    End Property

  End Class
End Namespace
