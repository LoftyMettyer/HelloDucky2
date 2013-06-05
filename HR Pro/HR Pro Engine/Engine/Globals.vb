'<HideModuleName()> _
Public Module Globals

  Public MetadataDB As iConnection
  Public CommitDB As iConnection

  'Public Shared Connection As Connectivity.SQL
  Public MetadataProvider As Connectivity.MetadataProvider = Connectivity.MetadataProvider.PhoenixStoredProcs
  '    Public User As Connectivity.User
  Public Things As Things.Collection
  Public Workflows As Things.Collection
  Public Operators As Things.Collection
  Public Functions As Things.Collection
  Public SelectedThings As Things.Collection
  Public ErrorLog As HRProEngine.ErrorHandler.Errors
  Public TuningLog As Tuning.Report
  Public ModuleSetup As Things.Collection
  Public Options As HCMOptions
  Public UniqueCodes As Things.Collection
  Public GetFieldsFromDB As Things.Collection
  Public PerformanceIndexes As Things.Collection

  Public ScriptDB As ScriptDB.Script

  Public Sub Initialise()

    ' Metadata objects
    Things = New Things.Collection
    Workflows = New Things.Collection
    Operators = New Things.Collection
    Functions = New Things.Collection
    ErrorLog = New HRProEngine.ErrorHandler.Errors
    TuningLog = New Tuning.Report
    ModuleSetup = New Things.Collection
    ScriptDB = New ScriptDB.Script
    Options = New HCMOptions

    ' Dependency stack for special objects that will have procedures written for
    UniqueCodes = New Things.Collection
    GetFieldsFromDB = New Things.Collection
    PerformanceIndexes = New Things.Collection

  End Sub

End Module
