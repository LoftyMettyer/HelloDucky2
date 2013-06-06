Namespace Things

  <HideModuleName()>
  Public Module Structures

    <Serializable()>
    Public Structure ChildRowDetails
      Property RowSelection As ScriptDB.ColumnRowSelection
      Property RowNumber As Integer

      Property OrderId As Integer   ' Used temporarily while the expressions are loaded. Can tidy up in later release?
      Property FilterId As Integer  ' Used temporarily while the expressions are loaded. Can tidy up in later release?

      Property Column As Column

      Property Order As TableOrder
      Property Filter As Expression
      Property Relation As Relation
      Property BaseTable As Table
    End Structure

  End Module
End Namespace
