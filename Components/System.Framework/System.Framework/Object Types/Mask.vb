Namespace Things

  <Serializable()>
  Public Class Mask
    Inherits Expression

    Public Sub GenerateMaskCode()
      Me.ExpressionType = ScriptDB.ExpressionType.Mask
      MyBase.GenerateCodeForColumn()
    End Sub

  End Class
End Namespace
