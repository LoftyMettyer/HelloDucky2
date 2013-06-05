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
  Public ModuleSetup As Things.Collection
  Public Options As HCMOptions

  Public ScriptDB As ScriptDB.Script

  Public Sub Initialise()

    Things = New Things.Collection
    Workflows = New Things.Collection
    Operators = New Things.Collection
    Functions = New Things.Collection
    ErrorLog = New HRProEngine.ErrorHandler.Errors
    ModuleSetup = New Things.Collection
    ScriptDB = New ScriptDB.Script
    Options = New HCMOptions
  End Sub

End Module
