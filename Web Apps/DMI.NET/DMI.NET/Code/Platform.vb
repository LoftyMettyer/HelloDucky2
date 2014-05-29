' Mobiles. Fix as tiles display.
Imports System.Globalization
Imports HR.Intranet.Server.Structures
Imports System.Threading

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

		Public Shared Function IsWindowsAuthenicatedEnabled() As Boolean

			Dim sUserName = HttpContext.Current.Request.ServerVariables("LOGON_USER").ToString()
			Return sUserName.Length > 0

		End Function

		Public Shared Function PopulateRegionalSettings(CultureName As String) As RegionalSettings

			Dim objCulture = CultureInfo.CreateSpecificCulture(CultureName)
			Dim objSettings As New RegionalSettings
			objSettings.Culture = objCulture

			Return objSettings

		End Function

	End Class
End Namespace
