Imports System.Threading

Public Module ASRIntranetFunctions

  'TODO
  Public Function GetRegistrySetting(psAppName As String, psSection As String, psKey As String) As String
    ' Get the required value from the registry with the given registry key values.
    GetRegistrySetting = GetSetting(AppName:=psAppName, Section:=psSection, Key:=psKey)

  End Function

  'TODO
  Function LocaleDateFormat() As String

    Dim sLocaleDateFormat As String = Thread.CurrentThread.CurrentUICulture.DateTimeFormat.ShortDatePattern.ToLower()

    Return sLocaleDateFormat

  End Function

  'TODO
  Function LocaleDecimalSeparator() As String
    Return ""
  End Function

  'TODO
  Function LocaleThousandSeparator() As String
    Return ""
  End Function

  'TODO
  Function LocaleDateSeparator() As String
    Return ""
  End Function

  '****************************************************************
  ' NullSafeString
  '****************************************************************
  Public Function NullSafeString(ByVal arg As Object, _
  Optional ByVal returnIfEmpty As String = "") As String

    Dim returnValue As String

    If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
      OrElse (arg Is String.Empty) Then
      returnValue = returnIfEmpty
    Else
      Try
        returnValue = CStr(arg)
      Catch
        returnValue = returnIfEmpty
      End Try

    End If

    Return returnValue

  End Function

  ' TODO
  Function ValidateDir(ByRef paramType As String) As Boolean
    Return True
  End Function

End Module
