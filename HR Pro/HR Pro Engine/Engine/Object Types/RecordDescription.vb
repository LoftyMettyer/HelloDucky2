Namespace Things
  <Serializable()> _
  Public Class RecordDescription
    Inherits Things.Expression

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.RecordDescription
      End Get
    End Property

    Public Overrides Sub GenerateCode()
      Me.ExpressionType = ScriptDB.ExpressionType.RecordDescription
      Me.AssociatedColumn = Me.BaseExpression.BaseTable.Columns(0)    'needs to have at least one column.
      MyBase.GenerateCode()
    End Sub

  End Class
End Namespace
