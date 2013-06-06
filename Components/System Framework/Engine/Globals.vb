Imports SystemFramework.Things
Imports SystemFramework.ScriptDB
Imports SystemFramework.ErrorHandler

Public Module Globals

  Public MetadataDb As IConnection
  Public CommitDb As IConnection

  Public MetadataProvider As Connectivity.MetadataProvider = Connectivity.MetadataProvider.PhoenixStoredProcs

  Public Tables As ICollection(Of Table)
  Public Workflows As ICollection(Of Workflow)
  Public Operators As ICollection(Of CodeLibrary)
  Public Functions As ICollection(Of CodeLibrary)
  Public Expressions As ICollection(Of Expression)
  Public ErrorLog As Errors
  Public TuningLog As Tuning.Report
  Public ModuleSetup As SettingCollection
  Public SystemSettings As SettingCollection
  Public Options As HCMOptions
  Public Modifications As Modifications
  Public GetFieldsFromDb As ICollection(Of Component)
  Public PerformanceIndexes As ICollection(Of Column)
  Public OnBankHolidayUpdate As ICollection(Of TriggeredUpdate)

  Public ScriptDb As Script
  Public Login As Connectivity.Login

  Public Version As Version = Reflection.Assembly.GetExecutingAssembly().GetName().Version

  Public Sub Initialise()

    ' Metadata objects
    Tables = New Collection(Of Table)
    Workflows = New Collection(Of Workflow)
    Operators = New Collection(Of CodeLibrary)
    Functions = New Collection(Of CodeLibrary)
    Expressions = New Collection(Of Expression)
    ErrorLog = New Errors
    TuningLog = New Tuning.Report
    ModuleSetup = New SettingCollection
    ScriptDb = New Script
    Options = New HCMOptions
    Modifications = New Modifications
    SystemSettings = New SettingCollection

    ' Dependency stack for special objects that will have procedures written for
    GetFieldsFromDb = New Collection(Of Component)
    OnBankHolidayUpdate = New Collection(Of TriggeredUpdate)
    PerformanceIndexes = New Collection(Of Column)

  End Sub

End Module

