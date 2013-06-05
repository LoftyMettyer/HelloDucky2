Namespace Things

  <HideModuleName()> _
  Public Module Structures

    <Serializable()> _
    Public Structure ChildRowDetails
      Property RowSelection As ScriptDB.ColumnRowSelection
      Property RowNumber As Integer

      Property OrderID As Integer   ' Used temporarily while the expressions are loaded. Can tidy up in later release?
      Property FilterID As Integer  ' Used temporarily while the expressions are loaded. Can tidy up in later release?

      Property Order As Things.TableOrder
      Property Filter As Things.Expression
      Property Relation As Things.Relation
    End Structure

  End Module
End Namespace
