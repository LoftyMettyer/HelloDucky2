Option Strict On
Option Explicit On

Namespace Models.ObjectRequests
	Public Class SelectOrderModel
		Inherits GotoOptionBaseModel

		Public Property ScreenID As Integer
		Public Property ViewID As Integer
		Public Property OrderID As Integer

	End Class
End Namespace