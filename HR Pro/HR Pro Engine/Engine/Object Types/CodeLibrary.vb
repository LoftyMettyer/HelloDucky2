Namespace Things
  Public Class [CodeLibrary]
    Inherits Things.Base

    Public Code As String
    Public OperatorType As ScriptDB.OperatorSubType
    Public PreCode As String
    Public AfterCode As String
    Public ReturnType As ScriptDB.ComponentValueTypes
    Public RowNumberRequired As Boolean
    Public RecordIDRequired As Boolean
    Public CalculatePostAudit As Boolean
    Public IsGetFieldFromDB As Boolean = False
    Public IsUniqueCode As Boolean = False
    Public CaseCount As Integer = 0

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

