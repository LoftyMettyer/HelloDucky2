Option Explicit On
Option Strict On

Imports HR.Intranet.Server.Enums

Namespace Code.Interfaces
	Public Interface IReportDetail

		Property ReportID As Integer
		Property ReportType As UtilityType

	End Interface
End Namespace