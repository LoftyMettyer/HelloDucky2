Namespace Things

  <Serializable()>
  Public Class CodeLibrary
    Inherits Base

    Public Property Code As String
    Public Property OperatorType As ScriptDB.OperatorSubType
    Public Property PreCode As String
    Public Property AfterCode As String
    Public Property ReturnType As ScriptDB.ComponentValueTypes
    Public Property RowNumberRequired As Boolean
    Public Property RecordIDRequired As Boolean
    Public Property OvernightOnly As Boolean
    Public Property CalculatePostAudit As Boolean
    Public Property IsGetFieldFromDB As Boolean
    Public Property IsUniqueCode As Boolean
    Public Property CaseCount As Integer
    Public Property MakeTypeSafe As Boolean
    Public Property DependsOnBankHoliday As Boolean
    Public Property IsTimeDependant As Boolean

    Public Property Dependancies As ICollection(Of Setting)

    Public Sub New()
      Dependancies = New Collection(Of Setting)
    End Sub

  End Class
End Namespace

