Imports System.Threading
Imports System.Drawing
Imports System.IO
Imports System.Drawing.Imaging
Imports ADODB
Imports System.Data.OleDb

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

	'****************************************************************
	' NullSafeInteger
	'****************************************************************
	Public Function NullSafeInteger(ByVal arg As Object, _
	Optional ByVal returnIfEmpty As Integer = 0) As String

		Dim returnValue As Integer

		If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
			OrElse (arg Is String.Empty) Then
			returnValue = returnIfEmpty
		Else
			Try
				returnValue = CInt(arg)
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

	Public Function Base64StringToImage(Base64String As String) As Image
		Dim imageReturn As Image = Nothing

		Dim byteBuffer As Byte() = Convert.FromBase64String(Base64String)
		Dim memStream As New MemoryStream(byteBuffer)

		memStream.Position = 0

		imageReturn = Image.FromStream(memStream)

		memStream.Close()
		memStream = Nothing
		byteBuffer = Nothing

		Return imageReturn
	End Function

	Public Function ImageToBase64String(img As Image) As String
		Using ms As MemoryStream = New MemoryStream()
			'Convert Image to byte()
			Dim qualityParam As New EncoderParameter(Encoder.Quality, 90L)
			Dim encoderParams As New EncoderParameters(1)
			encoderParams.Param(0) = qualityParam
			Dim jgpEncoder As ImageCodecInfo = GetEncoder(ImageFormat.Jpeg)

			img.Save(ms, jgpEncoder, encoderParams)
			Dim imageBytes As Byte() = ms.ToArray()

			'Convert byte() to Base64 String
			Return Convert.ToBase64String(imageBytes)
		End Using
	End Function

	Private Function GetEncoder(format As ImageFormat) As ImageCodecInfo
		Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageDecoders()

		For Each codec As ImageCodecInfo In codecs
			If codec.FormatID = format.Guid Then
				Return codec
			End If
		Next
		Return Nothing
	End Function

	Public Function RecordSetToDataTable(ByVal objRS As Recordset) As DataTable

		Dim objDA As New OleDbDataAdapter()
		Dim objDT As New DataTable()

		' get rid of this if we can implement properly i.e. read sql directly into this datatable
		objRS.Requery()

		objDA.Fill(objDT, objRS)
		Return objDT

	End Function

	Public Function ConvertVb6ColourToArgb(systemColour As Integer) As System.Drawing.Color
		Dim red As String
		Dim green As String
		Dim blue As String

		Try
			Dim hexColour = Hex(CLng(systemColour))

			hexColour = Replace(hexColour, "#", "")
			blue = Val("&H" & Mid(hexColour, 1, 2))
			green = Val("&H" & Mid(hexColour, 3, 2))
			red = Val("&H" & Mid(hexColour, 5, 2))

		Catch ex As Exception
			blue = Val("&H00")
			green = Val("&H00")
			red = Val("&H00")
		End Try

		Return Color.FromArgb(red, green, blue)

	End Function

	Public Function GetReportNameByReportType(ReportType As Integer) As String
		Select Case ReportType
			Case 15
				Return "Absence Breakdown"
			Case 16
				Return "Bradford Factor"
			Case Else	'Does anybody know more of this 'magic numbers' so we can add to this Case? Should we have an enum?
				Return ""
		End Select
	End Function
End Module
