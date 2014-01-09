Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Structures

Namespace BaseClasses
	Public Class BaseExpressionComponent

		Protected ReadOnly Login As LoginInfo
		Protected General As New clsGeneral
		Private _datData As New clsDataAccess

		Public Sub New(ByVal Value As LoginInfo)
			Login = Value
			_datData = New clsDataAccess(Login)
			General = New clsGeneral(Login)
		End Sub

		'Public Sub New()
		'	MyBase.New()
		'End Sub

		'Friend Function NewExpression() As clsExprExpression
		'	Return New clsExprExpression(Login)
		'End Function

	End Class
End Namespace