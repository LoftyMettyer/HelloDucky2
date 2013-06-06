Namespace Things

  <Serializable()>
  Public Class Mask
    Inherits Expression

    Public Sub GenerateMaskCode()
      ExpressionType = ScriptDB.ExpressionType.Mask
      GenerateCodeForColumn()
    End Sub

  End Class
End Namespace
