
Public Interface ICommitDB
  Function ScriptTables() As Boolean
  Function ScriptTableViews() As Boolean
  Function ScriptObjects() As Boolean
  Function ScriptFunctions() As Boolean
  Function ScriptTriggers() As Boolean
  Function ScriptViews() As Boolean
  Function ScriptIndexes() As Boolean
  Function DropTableViews() As Boolean
  Function DropViews() As Boolean
  Function ApplySecurity() As Boolean
  Function ScriptOvernightStep2() As Boolean
End Interface
