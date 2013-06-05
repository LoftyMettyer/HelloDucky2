Namespace Things

  <HideModuleName()> _
  Public Module Interfaces

    'Public Interface iSystemObject
    '  Property ID As HCMGuid
    '  Property Name() As String
    '  Property Description() As String
    '  ReadOnly Property Type() As Things.Type
    '  Sub Edit()
    '  Property ObjectState() As System.Data.DataRowState
    '  Function Commit() As Boolean
    '  Property Objects() As Things.Collection
    '  Property Objects(ByVal Type As Things.Type) As Things.Collection
    '  Property Parent() As iSystemObject
    '  Property Root() As iSystemObject
    '  Property IsSelected As Boolean
    'End Interface

    Public Interface iExpressionComponent
      ReadOnly Property ComponentType() As Things.Type
      Property Name() As String
    End Interface

    'Public Interface iEditObject
    '  Property [Thing]() As Things.iSystemObject
    '  Function Initialise() As Boolean

    'End Interface

  End Module


End Namespace