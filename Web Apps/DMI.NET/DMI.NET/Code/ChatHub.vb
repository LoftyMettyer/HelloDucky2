Imports System.Web
Imports Microsoft.AspNet.SignalR

Namespace SignalRChat
	Public Class ChatHub
		Inherits Hub

		Public Sub Send(name As String, message As String)
			' Call the broadcastMessage method to update clients.
			Clients.All.broadcastMessage(name, message)
		End Sub
	End Class


	'Public Class LogoutHub
	'	Inherits Hub

	'	Public Sub Send(name As String, message As String)
	'		' Call the broadcastMessage method to update clients.
	'		Clients.All.broadcastMessage(name, message)
	'	End Sub
	'End Class

End Namespace