Public Interface IConnection
  Sub Open()
  Sub Close()
  Function ExecStoredProcedure(ByVal queryName As String, ByVal parms As Connectivity.Parameters) As DataSet
  Function ScriptStatement(ByVal statement As String, ByRef isCritical As Boolean) As Boolean
  Property Login As Structures.Login
End Interface
