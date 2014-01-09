Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Structures

Namespace BaseClasses
	Public Class BaseExpressionComponent

		Private ReadOnly _login As LoginInfo
		Private _datData As New clsDataAccess

		Public Sub New(ByVal Value As LoginInfo)
			_login = Value
			_datData = New clsDataAccess(_login)
		End Sub

		Public Sub New()
			MyBase.New()
		End Sub


	End Class
End Namespace