Option Strict On
Option Explicit On

Namespace Models.Responses

	Public Class SaveRecordRepsonse
		Inherits PostResponse

		Public TBResultCode As String
		Public CourseOverbooked As String
		Public Warning As Boolean
		Public OK As Boolean

	End Class
End Namespace