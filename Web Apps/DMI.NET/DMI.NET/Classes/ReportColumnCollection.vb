Option Strict On
Option Explicit On

Namespace Classes
	Public Class ReportColumnCollection
		Implements IReportDetail

		Public Property ReportID As Integer Implements IReportDetail.ReportID
		Public Property ReportType As UtilityType Implements IReportDetail.ReportType
		Public Property SelectionType As String
		Public Property ColumnsTableID As Integer
		Public Property Columns() As Integer()

	End Class
End Namespace