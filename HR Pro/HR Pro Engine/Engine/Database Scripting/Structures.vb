Namespace Things

  <HideModuleName()>
  Public Module Structures

    <Serializable()>
    Public Structure ChildRowDetails
      Property RowSelection As ScriptDB.ColumnRowSelection
      Property RowNumber As Integer

      Property OrderID As Integer   ' Used temporarily while the expressions are loaded. Can tidy up in later release?
      Property FilterID As Integer  ' Used temporarily while the expressions are loaded. Can tidy up in later release?

      Property Column As Things.Column

      Property Order As TableOrder
      Property Filter As Expression
      Property Relation As Relation
    End Structure

  End Module
End Namespace
