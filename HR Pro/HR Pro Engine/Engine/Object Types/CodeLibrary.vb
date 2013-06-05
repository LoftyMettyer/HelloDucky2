Namespace Things

  <Serializable()> _
  Public Class CodeLibrary
    Inherits Things.Base

    Public Property Code As String
    Public Property OperatorType As ScriptDB.OperatorSubType
    Public Property PreCode As String
    Public Property AfterCode As String
    Public Property ReturnType As ScriptDB.ComponentValueTypes
    Public Property RowNumberRequired As Boolean
    Public Property RecordIDRequired As Boolean
    Public Property OvernightOnly As Boolean = False
    Public Property CalculatePostAudit As Boolean
    Public Property IsGetFieldFromDB As Boolean = False
    Public Property IsUniqueCode As Boolean = False
    Public Property CaseCount As Integer = 0
    Public Property MakeTypeSafe As Boolean = False
    Public Property DependsOnBankHoliday As Boolean = False
    Public Property IsTimeDependant As Boolean = False

    Public Property Dependancies As New List(Of Setting)

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.CodeLibrary
      End Get
    End Property

  End Class
End Namespace

