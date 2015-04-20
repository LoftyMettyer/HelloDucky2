Option Strict On
Option Explicit On

Namespace Code.Interfaces
	Public Interface IDatabaseHub
		Sub SystemAdminMessage(messageFrom As String, message As String, forceLogout As Boolean, loggedInUsersOnly As Boolean)
		Sub NotifyGroup(messageFrom As String, message As String, forceLogout As Boolean)
		Sub ToggleLoginButton(disable As Boolean, message As String)
	End Interface
End Namespace