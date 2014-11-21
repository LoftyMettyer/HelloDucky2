Option Strict On
Option Explicit On

Imports HR.Intranet.Server

Namespace Classes
	Public Class ConfigurationError
		Public Severity As SeverityType
		Public Code As String
		Public Message As String
		Public Detail As String
	End Class
End Namespace