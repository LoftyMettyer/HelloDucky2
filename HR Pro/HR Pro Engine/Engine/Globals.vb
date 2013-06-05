Imports SystemFramework.Things
Imports SystemFramework.ScriptDB
Imports SystemFramework.ErrorHandler

Public Module Globals

  Public MetadataDB As IConnection
  Public CommitDB As IConnection

  Public MetadataProvider As Connectivity.MetadataProvider = Connectivity.MetadataProvider.PhoenixStoredProcs

  Public Tables As ICollection(Of Table)
  Public Workflows As ICollection(Of Workflow)
  Public Operators As ICollection(Of CodeLibrary)
  Public Functions As ICollection(Of CodeLibrary)
  Public ErrorLog As Errors
  Public TuningLog As Tuning.Report
  Public ModuleSetup As SettingCollection
  Public SystemSettings As SettingCollection
  Public Options As HCMOptions
  Public Modifications As Modifications
  Public GetFieldsFromDB As ICollection(Of Component)
  Public PerformanceIndexes As ICollection(Of Column)
  Public OnBankHolidayUpdate As ICollection(Of TriggeredUpdate)

  Public ScriptDB As ScriptDB.Script
  Public Login As Connectivity.Login

  Public Version As System.Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version

  Public Sub Initialise()

    ' Metadata objects
    Tables = New Collection(Of Table)
    Workflows = New Collection(Of Workflow)
    Operators = New Collection(Of CodeLibrary)
    Functions = New Collection(Of CodeLibrary)
    ErrorLog = New Errors
    TuningLog = New Tuning.Report
    ModuleSetup = New SettingCollection
    ScriptDB = New ScriptDB.Script
    Options = New HCMOptions
    Modifications = New Modifications
    SystemSettings = New SettingCollection

    ' Dependency stack for special objects that will have procedures written for
    GetFieldsFromDB = New Collection(Of Component)
    OnBankHolidayUpdate = New Collection(Of TriggeredUpdate)
    PerformanceIndexes = New Collection(Of Column)

  End Sub

End Module

