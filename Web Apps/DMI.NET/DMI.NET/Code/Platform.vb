' Mobiles. Fix as tiles display.
Namespace Code
	Public Class Platform

		Public Shared Function IsMobileDevice() As Boolean

			Dim ua As String = HttpContext.Current.Request.UserAgent

			If Not ua = Nothing Then
				If ua.Contains("iPhone") Or ua.Contains("iPad") Or ua.Contains("Android") Then
					Return True
				End If
			End If

			Return False

		End Function

		Public Shared Function IsWindowsSupported() As Boolean

			Dim ua As String = HttpContext.Current.Request.UserAgent

			If Not ua = Nothing Then
				If ua.Contains("Windows") Then
					Return True
				End If
			End If

			Return False

		End Function

	End Class
End Namespace
