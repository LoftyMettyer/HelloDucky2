Public Module svrCleanup

	Function CleanString(psString)
		Dim sCleaned

		sCleaned = psString
		'   sCleaned = replace(sCleaned, "<", "&lt;")
		'   sCleaned = replace(sCleaned, ">", "&gt;")
		sCleaned = Replace(sCleaned, "'", "''")
		'	cleanString = "'" & sCleaned & "'"
		cleanString = sCleaned
	End Function

	Function CleanNumeric(pNumber)
		Dim lngCleaned

		lngCleaned = CLng(0)

		If IsNumeric("" & pNumber) Then
			If (CDbl(pNumber) > -2147483647) And (CDbl(pNumber) < 2147483648) Then
				If InStr(1, "" & pNumber, ",") > 0 Then
					lngCleaned = 0
				Else
					lngCleaned = CLng(pNumber)
				End If
			End If
		End If

		cleanNumeric = lngCleaned
	End Function

	Function CleanBoolean(pValue)
		Dim lngCleaned

		lngCleaned = CLng(0)

		If IsNumeric("" & pValue) Then
			If (CDbl(pValue) > -2147483647) And (CDbl(pValue) < 2147483648) Then
				If InStr(1, "" & pValue, ",") > 0 Then
					lngCleaned = 0
				Else
					lngCleaned = CLng(pValue)
				End If
			End If
		Else
			If UCase(Trim(pValue)) = "TRUE" Then
				lngCleaned = 1
			End If
		End If

		If lngCleaned <> 0 Then
			lngCleaned = 1
		End If

		cleanBoolean = lngCleaned
	End Function

	Function CleanStringForJavaScript(psString)
		Dim sCleaned

		sCleaned = psString
		sCleaned = Replace(sCleaned, "\", "\\")
		sCleaned = Replace(sCleaned, "'", "\'")
		sCleaned = Replace(sCleaned, """", "\""")

		cleanStringForJavaScript = sCleaned
	End Function

	Function CleanStringForJavaScript_NotDoubleQuotes(psString)
		Dim sCleaned

		sCleaned = psString
		sCleaned = Replace(sCleaned, "\", "\\")
		sCleaned = Replace(sCleaned, "'", "\'")
		sCleaned = Replace(sCleaned, "\""", """")

		cleanStringForJavaScript_NotDoubleQuotes = sCleaned
	End Function

	Function FormatError(psErrMsg) As String
		Dim iStart
		Dim iFound

		iFound = 0
		Do
			iStart = iFound
			iFound = InStr(iStart + 1, psErrMsg, "]")
		Loop While iFound > 0

		If (iStart > 0) And (iStart < Len(Trim(psErrMsg))) Then
			formatError = Trim(Mid(psErrMsg, iStart + 1))
		Else
			formatError = psErrMsg
		End If
	End Function

	Function ConvertSqlDateToLocale(psDate)
		Dim sLocaleFormat
		Dim iIndex

		If Len(psDate) > 0 Then
			sLocaleFormat = HttpContext.Current.Session("LocaleDateFormat")

			iIndex = InStr(sLocaleFormat, "dd")
			If iIndex > 0 Then
				If Day(psDate) < 10 Then
					sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
					 "0" & Day(psDate) & Mid(sLocaleFormat, iIndex + 2)
				Else
					sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
					 Day(psDate) & Mid(sLocaleFormat, iIndex + 2)
				End If
			End If

			iIndex = InStr(sLocaleFormat, "mm")
			If iIndex > 0 Then
				If Month(psDate) < 10 Then
					sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
					 "0" & Month(psDate) & Mid(sLocaleFormat, iIndex + 2)
				Else
					sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
					 Month(psDate) & Mid(sLocaleFormat, iIndex + 2)
				End If
			End If

			iIndex = InStr(sLocaleFormat, "yyyy")
			If iIndex > 0 Then
				sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
				 Year(psDate) & Mid(sLocaleFormat, iIndex + 4)
			End If

			ConvertSQLDateToLocale = sLocaleFormat
		Else
			ConvertSQLDateToLocale = ""
		End If
	End Function

	Function ConvertSqlDateToTime(psDate) As String
		If Len(psDate) = 0 Then
			ConvertSQLDateToTime = ""
		Else
			ConvertSQLDateToTime = FormatDateTime(psDate, vbShortTime)
		End If
	End Function

End Module
