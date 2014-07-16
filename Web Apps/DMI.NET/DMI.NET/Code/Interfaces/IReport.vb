Option Explicit On
Option Strict On

Imports HR.Intranet.Server

Namespace Code.Interfaces
	Public Interface IReport

		Property SessionInfo As SessionInfo

		Property BaseTableID As Integer

		Sub SetBaseTable(BaseTableID As Integer)

	End Interface
End Namespace