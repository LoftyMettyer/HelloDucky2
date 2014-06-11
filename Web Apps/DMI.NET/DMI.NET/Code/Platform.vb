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

        ' The stored procedures in SQL are expecting locale date formats to be passed in lowercase and double characters, i.e. mm/dd/yyyy or dd/mm/yyyy.        ' American format return as M/d/yyyy (other languages may get awkward too).        ' Ideally dates should be passed up to SQL either ready formatted or in native date objects, however that would involve changes to some pretty core parts        ' of the stored procs, so I've wrapped up the following areas to do a kind of fettling with the dateformat before sending to SQL.
        Public Shared Function LocaleDateFormatForSQL() As String

            Dim LocaleDateFormat As String = HttpContext.Current.Session("LocaleDateFormat").ToString.ToLower
            If LocaleDateFormat.IndexOf("dd") < 0 Then
                If LocaleDateFormat.IndexOf("d") >= 0 Then
                    LocaleDateFormat = LocaleDateFormat.Replace("d", "dd")
                End If
            End If
            If LocaleDateFormat.IndexOf("mm") < 0 Then
                If LocaleDateFormat.IndexOf("m") >= 0 Then
                    LocaleDateFormat = LocaleDateFormat.Replace("m", "mm")
                End If
            End If
            If LocaleDateFormat.IndexOf("yyyy") < 0 Then
                If LocaleDateFormat.IndexOf("yy") >= 0 Then
                    LocaleDateFormat = LocaleDateFormat.Replace("yy", "yyyy")
                End If
            End If

            Return LocaleDateFormat

        End Function


    End Class
End Namespace
