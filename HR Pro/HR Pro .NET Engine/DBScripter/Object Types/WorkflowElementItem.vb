Namespace Things

  Public Class WorkflowElementItem
    Inherits Things.Base

    'Public Overrides Function Commit() As Boolean
    'End Function

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.WorkflowElementItem
      End Get
    End Property


  End Class
End Namespace
