Namespace Things
  Public Class [CodeLibrary]
    Inherits Things.Base

    Public Code As String
    Public OperatorType As ScriptDB.OperatorSubType
    '   Public SplitIntoCase As Boolean
    '    Public AppendWildcard As Boolean
    Public PreCode As String
    Public AfterCode As String
    Public ReturnType As ScriptDB.ComponentValueTypes
    'Public BypassValidation As Boolean
    Public RowNumberRequired As Boolean
    Public CalculatePostAudit As Boolean

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.CodeLibrary
      End Get
    End Property

    Public Property Dependancies As Things.Collection
      Get
        Return Me.Objects
      End Get
      Set(ByVal value As Things.Collection)
        Me.Objects = value
      End Set
    End Property

    Public ReadOnly Property HasDependancies As Boolean
      Get
        Return Me.Objects.Count > 0
      End Get
    End Property

  End Class
End Namespace

