Namespace Code.Interfaces
	Public Interface ILicenceHub
		Sub CurrentUserList(result As String)
		Sub ActivateLogin()
		Sub SessionTimeout()
		'	Property ConnectionId As String
	End Interface
End Namespace