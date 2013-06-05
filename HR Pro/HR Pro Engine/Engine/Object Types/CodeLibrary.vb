Namespace Things

  <Serializable()> _
  Public Class CodeLibrary
    Inherits Things.Base

    Public Code As String
    Public OperatorType As ScriptDB.OperatorSubType
    Public PreCode As String
    Public AfterCode As String
    Public ReturnType As ScriptDB.ComponentValueTypes
    Public RowNumberRequired As Boolean
    Public RecordIDRequired As Boolean
    Public OvernightOnly As Boolean = False
    Public CalculatePostAudit As Boolean
    Public IsGetFieldFromDB As Boolean = False
    Public IsUniqueCode As Boolean = False
    Public CaseCount As Integer = 0
    Public MakeTypeSafe As Boolean = False
    Public DependsOnBankHoliday As Boolean = False
    Public IsTimeDependant As Boolean = False

    Public Property Dependancies As New List(Of Setting)

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.CodeLibrary
      End Get
    End Property

  End Class
End Namespace

