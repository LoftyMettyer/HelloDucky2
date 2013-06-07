Public Interface ITable
  Inherits IObject
  ' These eventually will be gotten rid of when we port the rest of sysmgr into this framework.
  Property SysMgrInsertTrigger As String
  Property SysMgrUpdateTrigger As String
  Property SysMgrDeleteTrigger As String
End Interface
