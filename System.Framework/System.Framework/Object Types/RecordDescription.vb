Namespace Things

  Public Class RecordDescription
    Inherits Expression

    Public Sub GenerateRecordDescription()
      Me.ExpressionType = ScriptDB.ExpressionType.RecordDescription
      Me.AssociatedColumn = Me.BaseExpression.BaseTable.Columns(0)    'needs to have at least one column.
      MyBase.GenerateCodeForColumn()
    End Sub

  End Class
End Namespace
