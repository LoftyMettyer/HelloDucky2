Imports System.ComponentModel
Namespace Models

   Public Class OrganisationReportFilterItem
      Implements IReportDetail

      Public Property ReportID As Integer Implements IReportDetail.ReportID
      Public Property ReportType As UtilityType Implements IReportDetail.ReportType

      ''Filter Tab Properties 
      Public Property FilterFieldID As Integer
      Public Property FilterOperator As Integer
      Public Property FilterOperatorName As String
      Public Property FilterValue As String
   End Class

End Namespace