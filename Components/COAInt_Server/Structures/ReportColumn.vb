Imports HR.Intranet.Server.Metadata

Namespace Structures
	Friend Class ReportColumn
		Inherits Base
			Public BreakOrPageOnChange As Boolean
			Public HasSummaryLine As Boolean
			Public LastValue As Object
			Friend Sum As Double
			Friend Count As Long
	End Class

End Namespace