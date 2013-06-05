Public Module Interfaces

  Public Interface iCommitDB
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
  End Interface

  Public Interface iSystemManager
    Property MetadataDB As Object
    Property CommitDB As Object
    ReadOnly Property ErrorLog As ErrorHandler.Errors
    ReadOnly Property TuningLog As Tuning.Report
    ReadOnly Property Things As Things.Collection
    ReadOnly Property Script As ScriptDB.Script
    Function Initialise() As Boolean
    ReadOnly Property Options As HCMOptions
    Function CloseSafely() As Boolean
  End Interface

  Public Interface iErrors
    Sub OutputToFile(ByRef FileName As String)
  End Interface

  Public Interface iOptions
    Property RefreshObjects As Boolean
    Property DevelopmentMode As Boolean
  End Interface

  Public Interface iConnection
    Sub Open()
    Sub Close()
    Function ExecStoredProcedure(ByVal ProcedureName As String, ByRef Parms As Connectivity.Parameters) As System.Data.DataSet
    Function ScriptStatement(ByVal Statement As String) As Boolean
    Property Login As Connectivity.Login
  End Interface

End Module
