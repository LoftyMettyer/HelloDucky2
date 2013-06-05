Namespace Things

  <HideModuleName()> _
  Public Module Structures

    Public Structure ChildRowDetails
      Property RowSelection As ScriptDB.ColumnRowSelection
      Property RowNumber As Integer

      Property OrderID As HCMGuid   ' Used temporarily while the expressions are loaded. Can tidy up in later release?
      Property FilterID As HCMGuid  ' Used temporarily while the expressions are loaded. Can tidy up in later release?

      Property Order As Things.TableOrder
      Property Filter As Things.Expression
      Property Relation As Things.Relation
    End Structure

  End Module
End Namespace
