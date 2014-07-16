Option Explicit On
Option Strict On

Imports HR.Intranet.Server

Namespace Code.Interfaces
	Public Interface IDataAccess

		Property SessionContext As SessionInfo

	End Interface
End Namespace