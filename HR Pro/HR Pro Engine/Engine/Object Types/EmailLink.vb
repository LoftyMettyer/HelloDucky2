Namespace Things

  Public Class EmailLink
    Inherits Things.Base

    Public Title As String

    Public Filter As Things.Expression
    Public EffectiveDate As Date
    Public Attachment As String
    Public LinkType As EmailLinkType
    Public TableID As Things.Table
    Public DateColumn As Things.Column
    Public DateOffset As Integer
    Public DatePeriod As DateOffsetType
    Public RecordInsert As Boolean
    Public RecordUpdate As Boolean
    Public RecordDelete As Boolean
    Public SubjectContentID As Object
    Public BodyContentID As Object
    Public DateAmendment As Date

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.EmailLink
      End Get
    End Property
  End Class
End Namespace