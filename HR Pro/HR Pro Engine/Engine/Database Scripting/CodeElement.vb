Namespace ScriptDB

  <Serializable()>
  Public Structure CodeElement
    Public CaseNumber As Long
    Public Code As String
    Public CodeType As ScriptDB.ComponentTypes
    Public OperatorType As ScriptDB.OperatorSubType
    Public BypassEvaluation As Boolean
  End Structure

End Namespace

