Option Explicit On
Option Strict On

Imports HR.Intranet.Server.Enums

Namespace Structures
	Public Class LoginInfo
		Public Server As String
		Public Database As String
		Public Username As String
		Public Password As String
		Public TrustedConnection As Boolean
		Public LoginTime As DateTime
		Public LastLoginTime As DateTime
		Public InvalidPasswordAttempts As Integer
		Public LockedOut As Boolean
		Public LockoutTime As Date
		Public MustChangePassword As Boolean

		Public UserGroup As String
		Public LoginFailReason As String = ""

		Public IsServerRole As Boolean = False
		Public IsSystemOrSecurityAdmin As Boolean = False
		Public IsDMIUser As Boolean = False
		Public IsSSIUser As Boolean = False

		Public DefaultWebArea As WebArea = WebArea.SSI

	End Class
End Namespace