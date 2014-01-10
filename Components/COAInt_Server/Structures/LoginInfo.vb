Option Explicit On
Option Strict On

Namespace Structures
	Public Class LoginInfo
		Public Server As String
		Public Database As String
		Public Username As String
		Public Password As String
		Public TrustedConnection As Boolean
		Public LoginTime As Date
		Public LastLoginTime As Date
		Public InvalidPasswordAttempts As Integer
		Public LockedOut As Boolean
		Public LockoutTime As Date


		Public UserType As Integer
		Public SelfServiceUserType As Integer
		Public UserGroup As String
		Public LoginFailReason As String = ""

		'Public ReadOnly Property OldConnectionString() As String
		'	Get


		'	End Get
		'End Property

		'Public ReadOnly Property OldConnectionString() As String
		'	Get


		'	End Get
		'End Property


	End Class
End Namespace