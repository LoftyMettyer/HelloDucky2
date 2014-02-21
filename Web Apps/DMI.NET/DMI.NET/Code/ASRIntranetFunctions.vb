Imports System.Threading
Imports System.Drawing
Imports System.IO
Imports System.Drawing.Imaging
Imports ADODB
Imports System.Data.OleDb
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server
Imports System.Data.SqlClient
Imports DayPilot.Web.Ui

Public Module ASRIntranetFunctions

	Function LocaleDateFormat() As String
		Return Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern.ToLower()
	End Function

	Function LocaleDecimalSeparator() As String
		Return Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator
	End Function

	Function LocaleThousandSeparator() As String
		Return Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator
	End Function

	Function LocaleDateSeparator() As String
		Return Thread.CurrentThread.CurrentCulture.DateTimeFormat.DateSeparator
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

	Public Function GetReportNameByReportType(ReportType As utilityType) As String
		Select Case ReportType
			Case UtilityType.utlAbsenceBreakdown
				Return "Absence Breakdown"
			Case UtilityType.utlBradfordFactor
				Return "Bradford Factor"
			Case Else
				Return ""
		End Select
	End Function

	Public Function Get1000SeparatorFindColumns(TableID As Long, ViewID As Long, OrderID As Long) As String
		Dim objSession As SessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)	'Set session info
		Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
		Dim ThousandColumns As String = ""

		Dim pfError As New SqlParameter("@pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
		Dim piTableID As New SqlParameter("@piTableID", SqlDbType.Int) With {.Value = TableID}
		Dim piViewID As New SqlParameter("@piViewID", SqlDbType.Int) With {.Value = ViewID}
		Dim piOrderID As New SqlParameter("@piOrderID", SqlDbType.Int) With {.Value = OrderID}
		Dim ps1000SeparatorCols As New SqlParameter("@ps1000SeparatorCols", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

		objDataAccess.ExecuteSP("spASRIntGet1000SeparatorFindColumns", _
						pfError, _
						piTableID, _
						piViewID, _
						piOrderID, _
						ps1000SeparatorCols _
		)

		ThousandColumns = ps1000SeparatorCols.Value

		Return ThousandColumns
	End Function

	Public Function GetLookupValues(ColumnID As Integer) As DataTable
		Dim objSession As SessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)	'Set session info
		Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
		Dim piColumnID As New SqlParameter("@piColumnID", SqlDbType.Int) With {.Value = ColumnID}

		Return objDataAccess.GetDataTable("sp_ASRIntGetLookupValues", CommandType.StoredProcedure, piColumnID)
	End Function


	Sub Calendar_BindDataset(ByRef objCalendarControl As DayPilotScheduler)
		Dim dt As New DataTable()
		dt.Columns.Add("startdate", GetType(DateTime))
		dt.Columns.Add("enddate", GetType(DateTime))
		dt.Columns.Add("description", GetType(String))
		dt.Columns.Add("baseid", GetType(String))
		dt.Columns.Add("id", GetType(String))
		dt.Columns.Add("resource", GetType(String))
		dt.Columns.Add("color", GetType(String))
		dt.Columns.Add("eventType", GetType(String))

		Dim dr As DataRow

		Dim objCalendar = CType(HttpContext.Current.Session("objCalendar" & HttpContext.Current.Session("CalRepUtilID")), CalendarReport)

		Dim sDescription As String
		Dim sEventDescription As String

		Dim dStart As Date
		Dim dEnd As Date

		Dim iNextColor As Integer = 0

		If objCalendar.Events Is Nothing Then	'Report contains no records, return empty Data Table
			Exit Sub
		End If

		' Add the resources (i.e. people)
		For Each objRow As DataRow In objCalendar.BaseRecordset.Rows
			sDescription = objCalendar.ConvertDescription(objRow(0).ToString(), objRow(1).ToString(), objRow(2).ToString())
			objCalendarControl.Resources.Add(sDescription, objRow(4).ToString())
		Next

		If Not objCalendar.rsPersonnelBHols Is Nothing Then
			For Each objRow In objCalendar.rsPersonnelBHols.Rows
				dr = dt.NewRow()

				dr("id") = objRow("id")

				Dim objLegend = objCalendar.Legend.Find(Function(n) n.LegendKey = "Bank Holiday")
				If Not objLegend Is Nothing Then
					If objLegend.Count = 0 Then
						objLegend.Count += 1
						objLegend.HTMLColorName = objCalendar.LegendColors(iNextColor).ColDesc
						Dim objColor = Color.FromArgb(objCalendar.LegendColors(iNextColor).ColValue)
						iNextColor += 1
						If iNextColor > objCalendar.LegendColors.Count Then iNextColor = objCalendar.LegendColors.Count - 1
						objLegend.HexColor = String.Format("#{0}{1}{2}", objColor.R.ToString("X").PadLeft(2, "0"), objColor.G.ToString("X").PadLeft(2, "0"), objColor.B.ToString("X").PadLeft(2, "0"))
					End If

					dr("color") = objLegend.HexColor
				End If

				dr("startdate") = CDate(objRow(2))
				dr("enddate") = CDate(objRow(2)).AddDays(1)
				dr("description") = "B"	'Changed to provide short form for bank holiday 
				dr("eventType") = "bank"

				dr("resource") = objRow(0)
				dt.Rows.Add(dr)

			Next
		End If

		For Each objRow As DataRow In objCalendar.Events.Rows

			sEventDescription = objRow("eventdescription1").ToString() & objRow("eventdescription2").ToString()

			If sEventDescription = "" Then
				sEventDescription = objRow(0).ToString()
			End If

			dr = dt.NewRow()
			dr("baseid") = objRow("baseid")
			dr("id") = objRow("id")

			If objRow("startsession") = "AM" Then
				dStart = CDate(objRow("startdate"))
			Else
				dStart = CDate(objRow("startdate")).AddHours(12)
			End If

			If objRow("endsession") = "AM" Then
				dEnd = CDate(objRow("enddate")).AddHours(12)
			Else
				dEnd = CDate(objRow("enddate")).AddDays(1)
			End If

			dr("startdate") = dStart
			dr("enddate") = dEnd
			dr("description") = sEventDescription

			Dim sLegendKey As String = objRow(5).ToString()
			Dim objLegend = objCalendar.Legend.Find(Function(n) n.LegendKey = sLegendKey)

			If Not objLegend Is Nothing Then
				If objLegend.Count = 0 Then
					objLegend.Count += 1
					objLegend.HTMLColorName = objCalendar.LegendColors(iNextColor).ColDesc
					Dim objColor = Color.FromArgb(objCalendar.LegendColors(iNextColor).ColValue)
					iNextColor += 1
					If iNextColor > objCalendar.LegendColors.Count Then iNextColor = objCalendar.LegendColors.Count - 1
					objLegend.HexColor = String.Format("#{0}{1}{2}", objColor.R.ToString("X").PadLeft(2, "0"), objColor.G.ToString("X").PadLeft(2, "0"), objColor.B.ToString("X").PadLeft(2, "0"))
				End If

				dr("color") = objLegend.HexColor
			End If

			dr("resource") = objRow("baseid")
			dt.Rows.Add(dr)

		Next

		objCalendarControl.DataSource = dt

	End Sub


End Module
