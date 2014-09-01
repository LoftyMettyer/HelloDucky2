Public Interface ISystemManager
  Property MetadataDB As Object
  Property CommitDB As Object
  ReadOnly Property ErrorLog As Collections.Errors
  ReadOnly Property TuningLog As TuningReport
	Function GetTable(id As Integer) As Table
  ReadOnly Property Script As ScriptDB.Script
  ReadOnly Property Options As [Option]
  Function Initialise() As Boolean
  Function PopulateObjects() As Boolean
  Function CloseSafely() As Boolean
  ReadOnly Property Version As Version
	ReadOnly Property Modifications As Modifications
	Function UpdateLicence(existingLicence As String) As String
End Interface
