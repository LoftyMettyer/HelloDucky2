Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums

Namespace Metadata
	Public Class Table
		Inherits Base
			Public TableType As TableTypes
			Public DefaultOrderID As Integer
			Public RecordDescExprID As Integer
	End Class
End Namespace