Option Strict On
Option Explicit On

Imports DMI.NET.Code.Attributes
Imports HR.Intranet.Server.Expressions
Imports HR.Intranet.Server
Namespace Models

   Public Class OrganisationReportFilterItem
      Implements IReportDetail
      Implements IJsonSerialize


      Public Property ReportID As Integer Implements IReportDetail.ReportID
      Public Property ReportType As UtilityType Implements IReportDetail.ReportType

      ''Filter Tab Properties 
      Public Property FieldID As Integer
      Public Property FieldName As String
      Public Property FieldDataType As Integer
      Public Property FieldColumnSize As Integer
      Public Property FieldDecimals As Integer
      Public Property OperatorID As Integer
      Public Property OperatorName As String
      Public Property FilterValue As String

      Public Property ID As Integer Implements IJsonSerialize.ID

   End Class

End Namespace