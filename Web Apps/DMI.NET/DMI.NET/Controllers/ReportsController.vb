Option Explicit On
Option Strict On

Namespace Controllers

	Public Class ReportsController
		Inherits Controller

		Function Util_Def_CustomReports() As ActionResult
			Return View()
		End Function

		Function util_def_crosstabs() As ActionResult
			Return View()
		End Function

		Public Function util_def_calendarreport() As ActionResult
			Return View()
		End Function

		Function util_def_mailmerge() As ActionResult
			Return View()
		End Function

	End Class
End Namespace