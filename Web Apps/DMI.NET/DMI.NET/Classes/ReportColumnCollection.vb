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

		''' <summary>
		''' Used to get/set the selected table name.
		''' </summary>
		''' <value>The name of the selected table into the dropdown</value>
		''' <returns>The name of the table</returns>
		''' <remarks></remarks>
		Public Property TableName As String

	End Class
End Namespace