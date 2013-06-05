Imports SystemFramework.Things

Public Module Globals

  Public MetadataDB As IConnection
  Public CommitDB As IConnection

  Public MetadataProvider As Connectivity.MetadataProvider = Connectivity.MetadataProvider.PhoenixStoredProcs

  Public Tables As List(Of Table)
  Public Workflows As Things.Collections.Generic
  Public Operators As List(Of CodeLibrary)
  Public Functions As List(Of CodeLibrary)
  Public SelectedThings As Things.Collections.Generic
  Public ErrorLog As SystemFramework.ErrorHandler.Errors
  Public TuningLog As Tuning.Report
  Public ModuleSetup As SettingsCollection
  Public SystemSettings As SettingsCollection
  Public Options As HCMOptions
  Public Modifications As Modifications
  Public GetFieldsFromDB As Things.Collections.Generic
  Public PerformanceIndexes As Things.Collections.Generic
  Public OnBankHolidayUpdate As Things.Collections.Generic

  Public ScriptDB As ScriptDB.Script
  Public Login As Connectivity.Login

  Public Version As System.Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version

  Public Sub Initialise()

    ' Metadata objects
    Tables = New List(Of Table)
    Workflows = New Things.Collections.Generic
    Operators = New List(Of CodeLibrary)
    Functions = New List(Of CodeLibrary)
    ErrorLog = New SystemFramework.ErrorHandler.Errors
    TuningLog = New Tuning.Report
    ModuleSetup = New SettingsCollection
    ScriptDB = New ScriptDB.Script
    Options = New HCMOptions
    Modifications = New Modifications
    SystemSettings = New SettingsCollection

    ' Dependency stack for special objects that will have procedures written for
    GetFieldsFromDB = New Things.Collections.Generic
    OnBankHolidayUpdate = New Things.Collections.Generic
    PerformanceIndexes = New Things.Collections.Generic

  End Sub

End Module

Namespace Things

  Public Class SettingsCollection
    Inherits ObjectModel.Collection(Of Setting)

    Public Function Setting(ByVal [module] As String, ByVal parameter As String) As Setting

      Dim item = Items.SingleOrDefault(Function(s) s.Module.ToLower = [module].ToLower AndAlso parameter.ToLower = parameter)

      If item IsNot Nothing Then
        Return item
      Else
        Return New Setting
      End If

    End Function

  End Class

End Namespace