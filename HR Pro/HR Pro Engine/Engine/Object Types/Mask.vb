Namespace Things

  <Serializable()>
  Public Class Mask
    Inherits Expression

    Public Overrides Sub GenerateCode()
      Me.ExpressionType = ScriptDB.ExpressionType.Mask
      MyBase.GenerateCode()
    End Sub

  End Class
End Namespace
