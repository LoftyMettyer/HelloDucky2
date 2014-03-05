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
		Public MustChangePassword As Boolean

		Public UserGroup As String
		Public LoginFailReason As String = ""

		Public IsServerRole As Boolean = False
		Public IsSystemOrSecurityAdmin As Boolean = False
		Public IsDMIUser As Boolean = False
		Public IsDMISingle As Boolean = False
		Public IsSSIUser As Boolean = False

	End Class
End Namespace