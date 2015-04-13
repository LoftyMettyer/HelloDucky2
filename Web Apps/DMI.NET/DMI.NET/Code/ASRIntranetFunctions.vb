Imports System.Threading
Imports System.Drawing
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Data.OleDb
Imports System.Globalization
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server
Imports System.Data.SqlClient
Imports DayPilot.Web.Ui
Imports System.Net.Mail
Imports System.Net.Mime
Imports HR.Intranet.Server.Structures
Imports System.Runtime.CompilerServices

Public Module ASRIntranetFunctions

	Function LocaleDecimalSeparator() As String
		Return Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator
	End Function

	Function LocaleThousandSeparator() As String
		Return Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator
	End Function

	Function LocaleLanguage() As String
		Return Thread.CurrentThread.CurrentCulture.ToString()
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
	Optional ByVal returnIfEmpty As Integer = 0) As Integer

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

	Function GeneratePath(filename As String) As String
#If DEBUG Then
		Return String.Format("{0}?v={1}", filename, System.DateTime.Now.Ticks)
#Else
		Dim currVersion As String = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
		Return String.Format("{0}?v={1}", filename, currVersion)
#End If
	End Function

	<Extension> _
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

	Public Function GetReportNameByReportType(ReportType As UtilityType) As String
		Select Case ReportType
			Case UtilityType.utlAbsenceBreakdown
				Return "Absence Breakdown"
			Case UtilityType.utlBradfordFactor
				Return "Bradford Factor"
			Case Else
				Return ""
		End Select
	End Function

	Public Sub Get1000SeparatorBlankIfZeroFindColumns(TableID As Long, ViewID As Long, OrderID As Long, ByRef ThousandColumns As String, ByRef BlankIfZeroColumns As String)
		Dim objSession As SessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)	'Set session info
		Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class

		Dim pfError As New SqlParameter("@pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
		Dim piTableID As New SqlParameter("@piTableID", SqlDbType.Int) With {.Value = TableID}
		Dim piViewID As New SqlParameter("@piViewID", SqlDbType.Int) With {.Value = ViewID}
		Dim piOrderID As New SqlParameter("@piOrderID", SqlDbType.Int) With {.Value = OrderID}
		Dim ps1000SeparatorCols As New SqlParameter("@ps1000SeparatorCols", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
		Dim psBlankIfZeroCols As New SqlParameter("@psBlankIfZeroCols", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

		objDataAccess.ExecuteSP("spASRIntGet1000SeparatorBlankIfZeroFindColumns", _
						pfError, _
						piTableID, _
						piViewID, _
						piOrderID, _
						ps1000SeparatorCols, _
						psBlankIfZeroCols _
		)

		ThousandColumns = ps1000SeparatorCols.Value
		BlankIfZeroColumns = psBlankIfZeroCols.Value
	End Sub

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
						If iNextColor >= objCalendar.LegendColors.Count Then iNextColor = objCalendar.LegendColors.Count - 1
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

			sEventDescription = objRow("eventdescription1").ToString() & " " & objRow("eventdescription2").ToString()

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
					If iNextColor >= objCalendar.LegendColors.Count Then iNextColor = objCalendar.LegendColors.Count - 1
					objLegend.HexColor = String.Format("#{0}{1}{2}", objColor.R.ToString("X").PadLeft(2, "0"), objColor.G.ToString("X").PadLeft(2, "0"), objColor.B.ToString("X").PadLeft(2, "0"))
				End If

				dr("color") = objLegend.HexColor
			End If

			dr("resource") = objRow("baseid")
			dt.Rows.Add(dr)

		Next

		objCalendarControl.DataSource = dt

	End Sub

	Public Function CompareVersion(Version1 As Version, Version2 As Version, IgnoreBuild As Boolean) As Boolean

		If Version1.Major = Version2.Major And Version1.Minor = Version2.Minor _
			And (Version1.Build = Version2.Build Or IgnoreBuild) Then
			Return True
		End If

		Return False

	End Function

	Public Function GetLoggedInUserRecordID(SingleRecordViewID As Integer) As Integer
		Dim objSession As SessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)	'Set session info
		Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class

		Dim ReturnValue As Integer = -1	'Default value to return in case the SP returns an error or there is more than one personal record

		Dim prmRecordID = New SqlParameter("piRecordID", SqlDbType.Int)
		prmRecordID.Direction = ParameterDirection.Output

		Dim prmRecordCount = New SqlParameter("piRecordCount", SqlDbType.Int)
		prmRecordCount.Direction = ParameterDirection.Output

		Try
			objDataAccess.GetDataSet("spASRIntGetSelfServiceRecordID", prmRecordID, prmRecordCount, New SqlParameter("piViewID", SingleRecordViewID))

			If prmRecordCount.Value = 1 Then
				' Only one record.
				ReturnValue = CInt(prmRecordID.Value)
			Else
				If prmRecordCount.Value = 0 Then
					' No personnel record. 
					ReturnValue = 0
				Else
					' More than one personnel record.
					ReturnValue = -1
				End If
			End If
		Catch ex As Exception
			Throw
		End Try

		Return ReturnValue
	End Function

	Friend Function GetEmailAddressesForGroup(groupID As Integer) As String

		Dim objDataAccess As clsDataAccess = CType(HttpContext.Current.Session("DatabaseAccess"), clsDataAccess)
		Dim sEmailAddresses As String = ""

		Try
			Dim rstEmailAddr = objDataAccess.GetDataTable("spASRIntGetEmailGroupAddresses", CommandType.StoredProcedure _
						, New SqlParameter("EmailGroupID", SqlDbType.Int) With {.Value = groupID})

			If Not rstEmailAddr Is Nothing Then
				sEmailAddresses = rstEmailAddr.Rows.Cast(Of DataRow)().Aggregate(sEmailAddresses, Function(current, objRow) current & (objRow(0).ToString & ";"))
			End If

		Catch ex As Exception
			sEmailAddresses = String.Format("Error getting the email addresses for group.({0})", ex.Message)
		End Try

		Return sEmailAddresses

	End Function

	' The following example sends a binary file as an e-mail attachment.
	Friend Sub SendMailWithAttachment(sSubject As String, objAttachment As Stream, recipientList As String, mstrEmailAttachAs As String)

		Using message As New MailMessage()

			message.Subject = IIf(sSubject.Length = 0, "OpenHR Report", sSubject).ToString()
			message.Body = "Your report is attached."

			Try

				If recipientList.Contains(";") = True Then

					Dim aRecipientList = Split(recipientList, ";")

					For iLoop = 0 To UBound(aRecipientList) - 1
						message.To.Add(aRecipientList(iLoop))
					Next
				Else
					message.To.Add(recipientList)
				End If

				Dim data As New Attachment(objAttachment, New ContentType(MediaTypeNames.Application.Octet))
				Dim disposition As ContentDisposition = data.ContentDisposition
				disposition.FileName = mstrEmailAttachAs
				message.Attachments.Add(data)

				Dim client As New SmtpClient()

				client.Send(message)
				data.Dispose()

			Catch ex As Exception
				Throw
			End Try

		End Using
	End Sub

	Public Function RoundValuesInRange(Value As String) As String
		Dim RetValue As String = ""
		Dim val1 As String = ""
		Dim val2 As String = ""

		If Value.Contains(" - ") Then	'Range (such as A - B), separate in two values
			val1 = Value.Substring(0, Value.IndexOf(" - "))
			val2 = Value.Substring(Value.IndexOf(" - ") + 3)
		ElseIf Value.Contains(".") Then	'Single value, round up and return

			'If value is convertable to decimal then only convert it to two decimal (such as 12.50, 23,456.50) otherwise return as it is (such as I.T.,  > 20.10 , < 12.54)
			If Decimal.TryParse(Value, 2) Then
				Return Decimal.Round(Decimal.Parse(Value, CultureInfo.InvariantCulture), 2).ToString()
			Else
				Return Value
			End If
		Else 'Single value, no round up necessary, return
			Return Value
		End If

		If val1.Contains(".") Then
			RetValue = String.Concat(RetValue, Decimal.Round(Decimal.Parse(val1, CultureInfo.InvariantCulture), 2))
		Else
			RetValue = String.Concat(RetValue, val1)
		End If
		RetValue = String.Concat(RetValue, " - ")
		If val2.Contains(".") Then
			RetValue = String.Concat(RetValue, Decimal.Round(Decimal.Parse(val2, CultureInfo.InvariantCulture), 2))
		Else
			RetValue = String.Concat(RetValue, val2)
		End If

		Return RetValue
	End Function
End Module
