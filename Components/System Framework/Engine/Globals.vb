Imports SystemFramework.Things
Imports SystemFramework.ScriptDB
Imports SystemFramework.Enums.Connection

Public Module Globals

  Public MetadataDb As IConnection
  Public CommitDb As IConnection

  Public MetadataProvider As MetadataProvider = MetadataProvider.PhoenixStoredProcs

  Public Tables As ICollection(Of Table)
  Public Operators As ICollection(Of CodeLibrary)
  Public Functions As ICollection(Of CodeLibrary)
  Public Expressions As ICollection(Of Expression)
  Public ErrorLog As Collections.Errors
  Public TuningLog As TuningReport
  Public ModuleSetup As Collections.Settings
  Public SystemSettings As Collections.Settings
  Public Options As [Option]
  Public Modifications As Modifications
  Public GetFieldsFromDb As ICollection(Of Component)
  Public PerformanceIndexes As ICollection(Of Column)
  Public OnBankHolidayUpdate As ICollection(Of TriggeredUpdate)

  Public ScriptDb As Script
  Public Login As Structures.Login

  Public Version As Version = Reflection.Assembly.GetExecutingAssembly().GetName().Version

  Public Sub Initialise()

    ' Metadata objects
    Tables = New Collection(Of Table)
    Operators = New Collection(Of CodeLibrary)
    Functions = New Collection(Of CodeLibrary)
    Expressions = New Collection(Of Expression)
    ErrorLog = New Collections.Errors
    TuningLog = New TuningReport
    ModuleSetup = New Collections.Settings
    ScriptDb = New Script
    Options = New [Option]
    Modifications = New Modifications
    SystemSettings = New Collections.Settings

    ' Dependency stack for special objects that will have procedures written for
    GetFieldsFromDb = New Collection(Of Component)
    OnBankHolidayUpdate = New Collection(Of TriggeredUpdate)
    PerformanceIndexes = New Collection(Of Column)

  End Sub

End Module

