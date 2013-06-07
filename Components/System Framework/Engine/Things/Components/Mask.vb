Imports SystemFramework.Enums

<Serializable()>
Public Class Mask
  Inherits Expression

  Public Sub GenerateMaskCode()
    ExpressionType = ExpressionType.Mask
    GenerateCodeForColumn()
  End Sub

End Class
