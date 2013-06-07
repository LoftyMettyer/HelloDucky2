Imports SystemFramework.Enums

Public Class EmailLink
  Inherits Base

  Public Property Title As String

  Public Property Filter As Expression
  Public Property EffectiveDate As Date
  Public Property Attachment As String
  Public Property LinkType As EmailLinkType
  Public Property TableId As Table
  Public Property DateColumn As Column
  Public Property DateOffset As Integer
  Public Property DatePeriod As DateOffsetType
  Public Property RecordInsert As Boolean
  Public Property RecordUpdate As Boolean
  Public Property RecordDelete As Boolean
  Public Property SubjectContentId As Object
  Public Property BodyContentId As Object
  Public Property DateAmendment As Date

End Class
