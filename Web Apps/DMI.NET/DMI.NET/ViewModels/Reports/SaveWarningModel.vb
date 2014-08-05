Option Explicit On
Option Strict On

Namespace ViewModels.Reports
	Public Class SaveWarningModel
		Public Property ReportType() As UtilityType
		Public Property ID As Integer
		Public Property ErrorCode As ReportValidationStatus
		Public Property ErrorMessage As String
	End Class
End Namespace