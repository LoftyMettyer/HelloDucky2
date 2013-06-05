'<HideModuleName()> _
Public Module Globals

  Public MetadataDB As iConnection
  Public CommitDB As iConnection

  'Public Shared Connection As Connectivity.SQL
  Public MetadataProvider As Connectivity.MetadataProvider = Connectivity.MetadataProvider.PhoenixStoredProcs
  '    Public User As Connectivity.User

  Public Tables As New List(Of Things.Table)
  Public Workflows As Things.Collections.Generic
  Public Operators As Things.Collections.Generic
  Public Functions As Things.Collections.Generic
  Public SelectedThings As Things.Collections.Generic
  Public ErrorLog As SystemFramework.ErrorHandler.Errors
  Public TuningLog As Tuning.Report
  Public ModuleSetup As Things.Collections.Generic
  Public SystemSettings As Things.Collections.Generic
  Public Options As HCMOptions
  Public Modifications As Modifications
  '  Public UniqueCodes As Things.Collection
  Public GetFieldsFromDB As Things.Collections.Generic
  Public PerformanceIndexes As Things.Collections.Generic
  Public OnBankHolidayUpdate As Things.Collections.Generic

  Public ScriptDB As ScriptDB.Script

  Public Login As Connectivity.Login

  Public Version As System.Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version

  Public Sub Initialise()

    ' Metadata objects
    Tables = New List(Of Things.Table)
    Workflows = New Things.Collections.Generic
    Operators = New Things.Collections.Generic
    Functions = New Things.Collections.Generic
    ErrorLog = New SystemFramework.ErrorHandler.Errors
    TuningLog = New Tuning.Report
    ModuleSetup = New Things.Collections.Generic
    ScriptDB = New ScriptDB.Script
    Options = New HCMOptions
    Modifications = New Modifications
    SystemSettings = New Things.Collections.Generic

    ' Dependency stack for special objects that will have procedures written for
    GetFieldsFromDB = New Things.Collections.Generic
    OnBankHolidayUpdate = New Things.Collections.Generic
    PerformanceIndexes = New Things.Collections.Generic

  End Sub

End Module
