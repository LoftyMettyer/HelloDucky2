Imports System.Reflection

Public Module svrCleanup

  Function CleanString(psString)
    Dim sCleaned

    sCleaned = psString
    '   sCleaned = replace(sCleaned, "<", "&lt;")
    '   sCleaned = replace(sCleaned, ">", "&gt;")
    sCleaned = Replace(sCleaned, "'", "''")
    '	cleanString = "'" & sCleaned & "'"
    CleanString = sCleaned
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

    CleanNumeric = lngCleaned
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

  Function CleanStringForJavaScript(psString)
    Dim sCleaned

    sCleaned = psString
    sCleaned = Replace(sCleaned, "\", "\\")
    sCleaned = Replace(sCleaned, "'", "\'")
    sCleaned = Replace(sCleaned, """", "\""")

    CleanStringForJavaScript = sCleaned
  End Function

  Function CleanStringForJavaScript_NotDoubleQuotes(psString)
    Dim sCleaned

    sCleaned = psString
    sCleaned = Replace(sCleaned, "\", "\\")
    sCleaned = Replace(sCleaned, "'", "\'")
    sCleaned = Replace(sCleaned, "\""", """")

    CleanStringForJavaScript_NotDoubleQuotes = sCleaned
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

  Function ConvertSQLDateToLocale(psDate As String) As String
    Dim sLocaleFormat As String
    Dim iIndex As Integer

    If Len(psDate) > 0 Then
      sLocaleFormat = HttpContext.Current.Session("LocaleDateFormat")

      iIndex = InStr(sLocaleFormat, "dd")
      If iIndex > 0 Then
        sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
            Mid(psDate, 4, 2) & Mid(sLocaleFormat, iIndex + 2)
      End If

      iIndex = InStr(sLocaleFormat, "MM")
      If iIndex > 0 Then
        sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
            Left(psDate, 2) & Mid(sLocaleFormat, iIndex + 2)
      End If

      iIndex = InStr(sLocaleFormat, "yyyy")
      If iIndex > 0 Then
        sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
            Mid(psDate, 7, 4) & Mid(sLocaleFormat, iIndex + 4)
      End If

      ConvertSQLDateToLocale = sLocaleFormat
    Else
      ConvertSQLDateToLocale = ""
    End If
  End Function

  Function ConvertSQLDateToLocale(pobjDate As Object) As String

    If IsDate(pobjDate) Then
      Return pobjDate.ToShortDateString()
    End If

    Return ""

  End Function

  Function ConvertLocaleDateToSQL(psDate As String) As String
    Dim sLocaleFormat As String
    Dim sSQLFormat As String
    Dim iIndex As Integer

    If Len(psDate) > 0 Then
      sLocaleFormat = HttpContext.Current.Session("LocaleDateFormat")

      iIndex = InStr(sLocaleFormat, "MM")
      If iIndex > 0 Then
        sSQLFormat = Mid(psDate, iIndex, 2) & "/"
      End If

      iIndex = InStr(sLocaleFormat, "dd")
      If iIndex > 0 Then
        sSQLFormat = sSQLFormat & Mid(psDate, iIndex, 2) & "/"
      End If

      iIndex = InStr(sLocaleFormat, "yyyy")
      If iIndex > 0 Then
        sSQLFormat = sSQLFormat & Mid(psDate, iIndex, 4)
      End If

      ConvertLocaleDateToSQL = sSQLFormat
    Else
      ConvertLocaleDateToSQL = ""
    End If
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
