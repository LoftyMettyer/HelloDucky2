Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Structures

Namespace BaseClasses

	Public Class BaseForDMI

		Protected DB As clsDataAccess
		Protected General As clsGeneral

		Private _sessionInfo As SessionInfo
		Private _login As LoginInfo

		Public Property SessionInfo() As SessionInfo
			Set(value As SessionInfo)
				_sessionInfo = value
				_login = _sessionInfo.LoginInfo

				gADOCon = _sessionInfo.Connection
				gsUsername = _sessionInfo.LoginInfo.Username

				DB = New clsDataAccess(_sessionInfo.LoginInfo)
				General = New clsGeneral(_sessionInfo.LoginInfo)

				' Tempry one for expressions as there's a lot of code in module and not classes - yuck!
				dataAccess = New clsDataAccess(_sessionInfo.LoginInfo)


			End Set
			Get
				Return _sessionInfo
			End Get
		End Property


	End Class
End Namespace