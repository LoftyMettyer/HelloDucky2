Option Explicit On
Option Strict Off

Imports System.IO
Imports Aspose.Cells
Imports System.Drawing
Imports HR.Intranet.Server.Enums
Imports DayPilot.Web.Ui

Namespace Code
	Public Class CalendarOutput

		'Public ReportStart As DateTime
		'Public ReportEnd As DateTime

		Public ReportData As DataTable
		Public Document As MemoryStream
		Public Calendar As HR.Intranet.Server.CalendarReport

		Public GeneratedFile As String
		Public DownloadFileName As String

		Public Function Generate(OutputType As OutputFormats) As Boolean

			GenerateExcel()

			Return True
		End Function

		Public Function GenerateExcel() As Boolean

			Const xOffSet = 1
			Const yOffset = 1

			Dim dCurrentMonth As DateTime
			Dim iRow As Integer

			Dim objDayPilot As New DayPilotScheduler
			Dim iSaveFormat As SaveFormat

			objDayPilot.DataStartField = "startdate"
			objDayPilot.DataEndField = "enddate"
			objDayPilot.DataTextField = "description"
			objDayPilot.DataValueField = "id"
			objDayPilot.DataResourceField = "resource"
			objDayPilot.DataTypeField = "eventtype"

			Calendar_BindDataset(objDayPilot)
			objDayPilot.DataBind()

			Dim objDocument As New Workbook
			Dim objMonthSheet As Worksheet
			Dim objRange As Range

			' Why, oh why, does it create a useless first worksheet?
			objDocument.Worksheets.RemoveAt("Sheet1")

			' Default properties for workbook
			objDocument.DefaultStyle.Font.Name = "Calibri"
			objDocument.DefaultStyle.Font.Size += 1

			' Setup styles
			Dim stlEvent As Style = objDocument.Styles(objDocument.Styles.Add())
			stlEvent.Pattern = BackgroundType.Solid
			stlEvent.Font.Name = objDocument.DefaultStyle.Font.Name
			stlEvent.Font.Size = objDocument.DefaultStyle.Font.Size

			Dim stlCaption As Style = objDocument.Styles(objDocument.Styles.Add())
			stlCaption.IndentLevel = 2
			stlCaption.Font.Size += 5
			stlCaption.Font.IsBold = True

			Dim stlWeekend As Style = objDocument.Styles(objDocument.Styles.Add())
			stlWeekend.Pattern = BackgroundType.Solid
			stlWeekend.ForegroundColor = Color.GhostWhite

			Dim stlWorkingDay As Style = objDocument.Styles(objDocument.Styles.Add())
			stlWorkingDay.Pattern = BackgroundType.Solid
			stlWorkingDay.ForegroundColor = Color.White

			Dim stlHeading As Style = objDocument.Styles(objDocument.Styles.Add())
			stlHeading.Pattern = BackgroundType.Solid
			stlHeading.Font.IsBold = True
			stlHeading.Font.Name = objDocument.DefaultStyle.Font.Name
			stlHeading.Font.Size = objDocument.DefaultStyle.Font.Size


			Dim objEnd As DateTime = Calendar.ReportEndDate

			Dim dMonthStart = Calendar.ReportStartDate.AddDays(-Calendar.ReportStartDate.Day + 1)
			Dim dMonthEnd = New DateTime(objEnd.Year, objEnd.Month, DateTime.DaysInMonth(objEnd.Year, objEnd.Month))
			Dim iMonthsBetweenStartAndEnd = CInt(DateDiff(DateInterval.Month, dMonthStart, dMonthEnd))

			For iMonth = 0 To iMonthsBetweenStartAndEnd
				dCurrentMonth = dMonthStart.AddMonths(iMonth)

				objDayPilot.StartDate = dCurrentMonth
				objDayPilot.Days = DateTime.DaysInMonth(dCurrentMonth.Year, dCurrentMonth.Month)

				objDayPilot.DataBind()
				objDayPilot.LoadEventsToDays()

				objMonthSheet = objDocument.Worksheets.Add(String.Format("{0} {1}", MonthName(dCurrentMonth.Month, True), dCurrentMonth.Year))
				objMonthSheet.IsRowColumnHeadersVisible = False
				objMonthSheet.IsGridlinesVisible = False


				' Nice display of month and year
				Dim strWorksheetTitle = (String.Format("{0} ({1} {2})", Calendar.CalendarReportName, MonthName(dCurrentMonth.Month), dCurrentMonth.Year))
				objMonthSheet.Cells(yOffset, 0).PutValue(strWorksheetTitle)
				objMonthSheet.Cells(yOffset, 0).SetStyle(stlCaption)
				objMonthSheet.AutoFitRow(yOffset)

				' Display day numbers
				For iCount = 1 To objDayPilot.Days
					objMonthSheet.Cells(yOffset + 2, iCount + xOffSet).PutValue(iCount)
					objMonthSheet.Cells(yOffset + 2, iCount + xOffSet).SetStyle(stlHeading)
					objMonthSheet.Cells.SetColumnWidth(iCount + yOffset, 2.75)
				Next

				' Style the calendar background
				For iCount = 1 To objDayPilot.Days
					objRange = objMonthSheet.Cells.CreateRange(yOffset + 3, iCount + xOffSet + 2, 20, 1)

					If dCurrentMonth.AddDays(iCount).DayOfWeek = DayOfWeek.Saturday Or dCurrentMonth.AddDays(iCount).DayOfWeek = DayOfWeek.Sunday Then
						objRange.SetStyle(stlWeekend)
						'Else
						'	objRange.SetStyle(stlWorkingDay)
					End If

				Next


				iRow = yOffset + 3
				For Each objDay As Day In objDayPilot.EventDays

					objMonthSheet.Cells(iRow, xOffSet).PutValue(objDay.Name)
					objMonthSheet.Cells(iRow, xOffSet).SetStyle(stlHeading)

					For Each objEvent As DayPilot.Web.Ui.Event In objDay.events

						' Event Name
						objMonthSheet.Cells(iRow, objEvent.Start.Day + xOffSet).PutValue(objEvent.Name)
						'objMonthSheet.Cells(iRow, objEvent.Start.Day + xOffSet).SetStyle(stlHeading)

						' Event itself
						stlEvent.ForegroundColor = ColorTranslator.FromHtml(objEvent.Source.Row(6).ToString())
						objRange = objMonthSheet.Cells.CreateRange(iRow, objEvent.Start.Day + xOffSet, 1, objEvent.End.Day - objEvent.Start.Day)

						'						objRange.SetOutlineBorder(BorderType.TopBorder, BorderStyle.Solid, Color.LightGray)
						objRange.SetOutlineBorders(BorderStyle.Solid, Color.LightGray)


						objRange.SetStyle(stlEvent)

						' Blank text afetr event to stop over lapping
						objMonthSheet.Cells(iRow, objEvent.End.Day + xOffSet).PutValue(" ")
						objMonthSheet.AutoFitRow(iRow)


						'objMonthSheet.Cells(iRow, objEvent.Start.Day + xOffSet).PutValue(objEvent.Name)
						'For iDayStyle = objEvent.Start.Day To objEvent.End.Day - 1
						'	stlEvent.ForegroundColor = ColorTranslator.FromHtml(objEvent.Source.Row(6).ToString())
						'	objMonthSheet.Cells(iRow, iDayStyle + xOffSet).SetStyle(stlEvent)
						'Next

					Next

						iRow += 1
					Next


				' Display legend
				iRow = yOffset + 2
				objMonthSheet.Cells(iRow, 34).PutValue("Legend")
				iRow += 1
				For Each objLegend In Calendar.Legend
					If objLegend.Count > 0 Then
						stlEvent.ForegroundColor = ColorTranslator.FromHtml(objLegend.HexColor)
						objMonthSheet.Cells(iRow, 34).SetStyle(stlEvent)
						objMonthSheet.Cells(iRow, 34).PutValue(objLegend.LegendDescription)

						iRow += 2
					End If
				Next

				' Autosize some columns
				objMonthSheet.AutoFitColumn(34)
				objMonthSheet.AutoFitColumn(xOffSet)

			Next

			iSaveFormat = SaveAsFormat(Path.GetExtension(DownloadFileName))
			GeneratedFile = Path.GetTempFileName().Replace(".tmp", ".xlsx")

			objDocument.Worksheets.ActiveSheetIndex = 0
			objDocument.Save(GeneratedFile, iSaveFormat)


			Return True

		End Function

		Private Shared Function SaveAsFormat(strExtension As String) As SaveFormat

			strExtension = strExtension.Replace(".", "")

			Select Case UCase(strExtension)
				Case "XLS"
					Return SaveFormat.Excel97To2003
				Case "HTML"
					Return SaveFormat.Html
				Case "PDF"
					Return SaveFormat.Pdf
				Case "CSV"
					Return SaveFormat.CSV
				Case "TIFF"
					Return SaveFormat.TIFF
				Case Else
					Return SaveFormat.Xlsx

			End Select

		End Function

	End Class


End Namespace