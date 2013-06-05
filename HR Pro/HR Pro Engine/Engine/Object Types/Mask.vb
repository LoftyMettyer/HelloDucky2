Namespace Things
  <Serializable()> _
  Public Class Mask
    Inherits Things.Expression

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Mask
      End Get
    End Property

    Public Overrides Sub GenerateCode()
      Me.ExpressionType = ScriptDB.ExpressionType.Mask
      MyBase.GenerateCode()
    End Sub

  End Class
End Namespace
