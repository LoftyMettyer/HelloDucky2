Namespace Things
  Public Class Workflow
    Inherits Things.Base

    Public Enabled As Boolean
    Public InitiationType As Integer
    Public Table As Things.Table
    Public QueryString As String

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Workflow
      End Get
    End Property

  End Class
End Namespace
