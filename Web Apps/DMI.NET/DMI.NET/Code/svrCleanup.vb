Imports System.Reflection
Imports System.Globalization

Public Module svrCleanup

	Function CleanStringForHTML(ByVal psString As String) As String

		Dim sCleaned As String = psString

		sCleaned = Replace(sCleaned, "'", "&apos;")
		sCleaned = Replace(sCleaned, """", "&quot;")
		sCleaned = Replace(sCleaned, "<", "&lt;")
		sCleaned = Replace(sCleaned, ">", "&gt;")

		Return sCleaned

	End Function

	Function CleanString(psString) As String
		Dim sCleaned

		sCleaned = psString
		'   sCleaned = replace(sCleaned, "<", "&lt;")
		'   sCleaned = replace(sCleaned, ">", "&gt;")
		sCleaned = Replace(sCleaned, "'", "''")
		'	cleanString = "'" & sCleaned & "'"
		CleanString = sCleaned
	End Function

	Function CleanNumeric(pNumber) As Long
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

		Return lngCleaned
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

    CleanBoolean = lngCleaned
  End Function

	Function CleanStringForJavaScript(psString) As String
		Dim sCleaned

		sCleaned = psString
		sCleaned = Replace(sCleaned, "\", "\\")
		sCleaned = Replace(sCleaned, "'", "\'")
		sCleaned = Replace(sCleaned, """", "\""")

		CleanStringForJavaScript = sCleaned
	End Function

	Function CleanStringForJavaScript_NotDoubleQuotes(ByVal psString As String) As String

		Dim sCleaned As String = psString

		sCleaned = Replace(sCleaned, "\", "\\")
		sCleaned = Replace(sCleaned, "'", "\'")
		sCleaned = Replace(sCleaned, "\""", """")

		Return sCleaned

	End Function

	Function CleanStringSpecialCharacters(psString) As String
		Dim sCleaned

		sCleaned = psString
		sCleaned = Replace(sCleaned, "<", "&lt;")
		sCleaned = Replace(sCleaned, ">", "&gt;")
		CleanStringSpecialCharacters = sCleaned
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
      FormatError = Trim(Mid(psErrMsg, iStart + 1))
    Else
      FormatError = psErrMsg
    End If
  End Function

  Function ConvertSQLDateToLocale(pobjDate As Object) As String

    If IsDate(pobjDate) Then
			'	Return pobjDate.ToShortDateString()
			If Globalization.CultureInfo.CurrentUICulture.ToString() = "en-US" Then
				Return CDate(pobjDate).ToString("MM/dd/yyyy")
			Else
				Return CDate(pobjDate).ToString(HttpContext.Current.Session("sessionContext").RegionalSettings.DateFormat.shortDatePattern)
			End If

		End If

    Return ""

  End Function

  Function ConvertLocaleDateToSQL(psDate As String) As String

		Dim objCulture As CultureInfo = HttpContext.Current.Session("sessionContext").RegionalSettings.Culture

		Try
			Return DateTime.Parse(psDate, objCulture).ToString("MM/dd/yyyy", CultureInfo.CreateSpecificCulture("en-US"))

		Catch ex As Exception
			Return ""

		End Try

  End Function

  Function ConvertSqlDateToTime(psDate) As String
    If Len(psDate) = 0 Then
      ConvertSqlDateToTime = ""
    Else
      ConvertSqlDateToTime = FormatDateTime(psDate, vbShortTime)
    End If
  End Function

  Function GetPageTitle(pageName As String) As String

    With Assembly.GetExecutingAssembly.GetName.Version
      Return String.Format("OpenHR {0} - v{1}.{2}.{3}", pageName, .Major, .Minor, .Build)
    End With

  End Function




End Module
