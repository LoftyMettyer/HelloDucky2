Namespace Things
  Public Class Workflow
    Inherits Things.Base

    Public Enabled As Boolean
    Public InitiationType As Integer
    Public BaseTableID As HCMGuid
    Public QueryString As String

    'Public Overrides Function Commit() As Boolean
    'End Function

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Workflow
      End Get
    End Property
  End Class
End Namespace
