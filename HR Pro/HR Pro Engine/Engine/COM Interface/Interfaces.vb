Public Module COMInterfaces

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
    ReadOnly Property Options As HCMOptions
    Function Initialise() As Boolean
    Function PopulateObjects() As Boolean
    Function CloseSafely() As Boolean
  End Interface

  Public Interface iErrors
    Sub OutputToFile(ByRef FileName As String)
    Sub Show()
    ReadOnly Property ErrorCount As Integer
    ReadOnly Property IsCatastrophic As Boolean
  End Interface

  Public Interface iForm
    Sub Show()
    Sub ShowDialog()
  End Interface

  Public Interface iOptions
    Property RefreshObjects As Boolean
    Property DevelopmentMode As Boolean
    Property OverflowSafety As Boolean
  End Interface

  Public Interface iConnection
    Sub Open()
    Sub Close()
    Function ExecStoredProcedure(ByVal ProcedureName As String, ByRef Parms As Connectivity.Parameters) As System.Data.DataSet
    Function ScriptStatement(ByVal Statement As String) As Boolean
    Property Login As Connectivity.Login
  End Interface

  Public Interface iObjectCollection
    Function Table(ByRef [ID] As HCMGuid) As Things.Table
    Function Setting(ByVal [Module] As String, ByVal [Parameter] As String) As Things.Setting
  End Interface

  Public Interface iObject
    Property Name As String
    ReadOnly Property PhysicalName As String
  End Interface

  Public Interface iTable
    Inherits iObject
    Property CustomTriggers As Things.BaseCollection
    ' These eventually will be gotten rid of when we port the rest of sysmgr into this framework.
    Property SysMgrInsertTrigger As String
    Property SysMgrUpdateTrigger As String
    Property SysMgrDeleteTrigger As String

  End Interface

End Module
