Namespace Things
  Public Class [CodeLibrary]
    Inherits Things.Base

    Public Code As String
    Public OperatorType As ScriptDB.OperatorSubType
    Public SplitIntoCase As Boolean
    Public AppendWildcard As Boolean
    Public AfterCode As String
    Public ReturnType As ScriptDB.ComponentValueTypes
    Public BypassValidation As Boolean

    
    'Public Overrides Function Commit() As Boolean
    'End Function

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.CodeLibrary
      End Get
    End Property

  End Class
End Namespace

