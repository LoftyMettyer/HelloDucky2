' This class temporarily replaces clsOLE whilst we decide on what to do with this functionality. Only the properties that are called externally have had code stumps
' created. Original code is held in clsOLE which is not inclued in the project file as it errors and needs rework to handle encryption and different file access methods.
' Possibly suggest rewriting the class rather than upgrading.

Public Class Ole

  ' Path in which temporary documents are to be created (physical directory on the server)
  Public WriteOnly Property TempLocationPhysical() As String
    Set(ByVal value As String)
    End Set
  End Property

  Public WriteOnly Property Connection() As Object
    Set(ByVal value As Object)
    End Set
  End Property

  Public Sub CleanupOLEFiles()
  End Sub

  Public Function GetPropertiesFromStream(ByRef plngRecordID As Object, ByRef plngColumnID As Object, ByRef pstrRealSource As String) As String
    Return "NOT YET IMPLEMENTED"
  End Function

End Class
