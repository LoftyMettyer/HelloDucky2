Namespace Things

  Public Class EmailLink
    Inherits Things.Base

    Public Property Title As String

    Public Property Filter As Things.Expression
    Public Property EffectiveDate As Date
    Public Property Attachment As String
    Public Property LinkType As EmailLinkType
    Public Property TableID As Things.Table
    Public Property DateColumn As Things.Column
    Public Property DateOffset As Integer
    Public Property DatePeriod As DateOffsetType
    Public Property RecordInsert As Boolean
    Public Property RecordUpdate As Boolean
    Public Property RecordDelete As Boolean
    Public Property SubjectContentID As Object
    Public Property BodyContentID As Object
    Public Property DateAmendment As Date

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.EmailLink
      End Get
    End Property
  End Class
End Namespace