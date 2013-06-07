Imports SystemFramework.Enums

Public Class RecordDescription
  Inherits Expression

  Public Sub GenerateRecordDescription()
    ExpressionType = ExpressionType.RecordDescription
    AssociatedColumn = BaseExpression.BaseTable.Columns(0)    'needs to have at least one column.
    GenerateCodeForColumn()
  End Sub

End Class
