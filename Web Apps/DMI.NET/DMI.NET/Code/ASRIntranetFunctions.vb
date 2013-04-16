Imports System.Threading
Imports Scripting

Public Module ASRIntranetFunctions

  'TODO
  Public Function GetRegistrySetting(psAppName As String, psSection As String, psKey As String) As String
    ' Get the required value from the registry with the given registry key values.
    GetRegistrySetting = GetSetting(AppName:=psAppName, Section:=psSection, Key:=psKey)

  End Function

  'TODO
  Function LocaleDateFormat() As String

    Dim sLocaleDateFormat As String = Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern.ToLower()
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
	'Code from INTCLient 
	'Public Function ValidateDir(psDir As String) As Boolean
	'	Dim fso As New FileSystemObject
	'	On Error Resume Next
	'	ValidateDir = False
	'	ValidateDir = fso.FolderExists(psDir)
	'	fso = Nothing
	'End Function

	'Function ValidateFilePath(psDir As String) As Boolean
	'	'NHRD Based on IntClient but fileSystemObject covers it better and non clienty
	'	'Dim fso As New FileSystemObject
	'	'Dim pathIsGood As Boolean
	'	'pathIsGood = fso.FileExists(psDir)
	'	Return True	'pathIsGood
	'End Function

	Function GeneratePath(filename As String) As String
		Dim currVersion As String = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
		Return String.Format("{0}?v={1}", filename, currVersion)
	End Function

	<System.Runtime.CompilerServices.Extension> _
	Public Function LatestContent(helper As UrlHelper, filename As String)
		Return helper.Content(String.Format("{0}", GeneratePath(filename)))
	End Function

End Module
