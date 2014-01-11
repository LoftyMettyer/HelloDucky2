Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.Enums

Module modUtilityAccess

	Public Const ACCESS_READWRITE As String = "RW"
	Public Const ACCESS_READONLY As String = "RO"
	Public Const ACCESS_HIDDEN As String = "HD"
	Public Const ACCESS_UNKNOWN As String = ""

	Public Const ACCESSDESC_READWRITE As String = "Read / Write"
	Public Const ACCESSDESC_READONLY As String = "Read Only"
	Public Const ACCESSDESC_HIDDEN As String = "Hidden"
	Public Const ACCESSDESC_UNKNOWN As String = "Unknown"


	Public Function AccessCode(ByRef psDescription As String) As String
		' Return the descriptive string associated with the given Access code.
		Select Case psDescription
			Case ACCESSDESC_READWRITE
				AccessCode = ACCESS_READWRITE
			Case ACCESSDESC_READONLY
				AccessCode = ACCESS_READONLY
			Case ACCESSDESC_HIDDEN
				AccessCode = ACCESS_HIDDEN
			Case Else
				AccessCode = ACCESS_UNKNOWN
		End Select

	End Function

	Public Function AccessDescription(ByRef psCode As String) As String
		' Return the descriptive string associated with the given Access code.
		Select Case psCode
			Case ACCESS_READWRITE
				AccessDescription = ACCESSDESC_READWRITE
			Case ACCESS_READONLY
				AccessDescription = ACCESSDESC_READONLY
			Case ACCESS_HIDDEN
				AccessDescription = ACCESSDESC_HIDDEN
			Case Else
				AccessDescription = ACCESSDESC_UNKNOWN
		End Select

	End Function

End Module