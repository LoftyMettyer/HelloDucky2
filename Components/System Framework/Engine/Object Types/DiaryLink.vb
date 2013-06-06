Namespace Things
  <Serializable()>
  Public Class DiaryLink
    Inherits Base

    Public Property Column As Column
    Public Property Comment As String
    Public Property Offset As Integer
    Public Property OffsetType As DateOffsetType
    Public Property Reminder As Boolean
    Public Property Filter As Expression
    Public Property EffectiveDate As DateTime
    Public Property CheckLeavingDate As Boolean

    Public Udf As ScriptDB.GeneratedUdf

    Public Sub Generate()

    End Sub

  End Class
End Namespace
