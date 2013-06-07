Imports SystemFramework.Enums

<Serializable()>
Public Class CodeLibrary
  Inherits Base

  Public Property Code As String
  Public Property OperatorType As OperatorSubType
  Public Property PreCode As String
  Public Property AfterCode As String
  Public Property ReturnType As ComponentValueTypes
  Public Property RowNumberRequired As Boolean
  Public Property RecordIdRequired As Boolean
  Public Property OvernightOnly As Boolean
  Public Property CalculatePostAudit As Boolean
  Public Property IsGetFieldFromDb As Boolean
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

