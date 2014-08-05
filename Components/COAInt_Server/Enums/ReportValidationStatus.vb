Option Strict On
Option Explicit On

Namespace Enums

	Public Enum ReportValidationStatus
		InvalidOnClient = -1
		ServerCheckComplete = 0
		InvalidOnServer = 1
		SaveAsNewDefinition = 2
		Overwrite = 3
		JobsWillBeHidden = 4
	End Enum

End Namespace