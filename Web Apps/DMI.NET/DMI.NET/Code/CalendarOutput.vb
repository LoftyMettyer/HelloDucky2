Option Explicit On
Option Strict On

Imports System.IO
Imports Aspose.Cells
Imports System.Drawing
Imports HR.Intranet.Server.Enums
Imports DayPilot.Web.Ui

Namespace Code
	Public Class CalendarOutput

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

			Dim dCurrentMonthStart As DateTime
			Dim dCurrentMonthEnd As DateTime

			Dim iRow As Integer
			Dim iLegendColumn As Integer

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
			Dim iStartRange As Integer
			Dim iRangeLength As Integer
			Dim iLastRangeEnd As Integer

			Dim objCellsLicense As New License
			objCellsLicense.SetLicense("Aspose.Cells.lic")

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

			Dim stlLegend As Style = objDocument.Styles(objDocument.Styles.Add())
			stlLegend.Pattern = BackgroundType.Solid
			stlLegend.Font.Name = objDocument.DefaultStyle.Font.Name
			stlLegend.Font.Size = objDocument.DefaultStyle.Font.Size

			Dim stlCaption As Style = objDocument.Styles(objDocument.Styles.Add())
			stlCaption.IndentLevel = 2
			stlCaption.Font.Size += 5
			stlCaption.Font.IsBold = True

			Dim stlWeekend As Style = objDocument.Styles(objDocument.Styles.Add())
			stlWeekend.Pattern = BackgroundType.Solid
			stlWeekend.ForegroundColor = Color.LightGray

			Dim stlWorkingDay As Style = objDocument.Styles(objDocument.Styles.Add())
			stlWorkingDay.Pattern = BackgroundType.Solid
			stlWorkingDay.ForegroundColor = Color.White

			Dim stlHeading As Style = objDocument.Styles(objDocument.Styles.Add())
			stlHeading.Pattern = BackgroundType.Solid
			stlHeading.Font.IsBold = True
			stlHeading.Font.Name = objDocument.DefaultStyle.Font.Name
			stlHeading.Font.Size = objDocument.DefaultStyle.Font.Size
			stlHeading.HorizontalAlignment = TextAlignmentType.Left
			stlHeading.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.LightGray)
			stlHeading.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.LightGray)

			Dim objEnd As DateTime = Calendar.ReportEndDate

			Dim dMonthStart = Calendar.ReportStartDate.AddDays(-Calendar.ReportStartDate.Day + 1)
			Dim dMonthEnd = New DateTime(objEnd.Year, objEnd.Month, DateTime.DaysInMonth(objEnd.Year, objEnd.Month))
			Dim iMonthsBetweenStartAndEnd = CInt(DateDiff(DateInterval.Month, dMonthStart, dMonthEnd))

			For iMonth = 0 To iMonthsBetweenStartAndEnd
				dCurrentMonthStart = dMonthStart.AddMonths(iMonth)
				dCurrentMonthEnd = dCurrentMonthStart.AddMonths(1)

				objDayPilot.StartDate = dCurrentMonthStart
				objDayPilot.Days = DateTime.DaysInMonth(dCurrentMonthStart.Year, dCurrentMonthStart.Month)

				objDayPilot.DataBind()
				objDayPilot.LoadEventsToDays()

				objMonthSheet = objDocument.Worksheets.Add(String.Format("{0} {1}", MonthName(dCurrentMonthStart.Month, True), dCurrentMonthStart.Year))
				objMonthSheet.IsRowColumnHeadersVisible = False
				objMonthSheet.IsGridlinesVisible = False

				' Nice display of month and year
				Dim strWorksheetTitle = (String.Format("{0} ({1} {2})", Calendar.Name, MonthName(dCurrentMonthStart.Month), dCurrentMonthStart.Year))
				objMonthSheet.Cells(yOffset, 0).PutValue(strWorksheetTitle)
				objMonthSheet.Cells(yOffset, 0).SetStyle(stlCaption)
				objMonthSheet.AutoFitRow(yOffset)

				' Display day numbers
				Dim iCellColumn As Integer = xOffSet + 1
				For iCount = 1 To objDayPilot.Days

					objMonthSheet.Cells(yOffset + 2, iCellColumn).PutValue(iCount)
					objMonthSheet.Cells(yOffset + 2, iCellColumn).SetStyle(stlHeading)
					objMonthSheet.Cells(yOffset + 2, iCellColumn + 1).SetStyle(stlHeading)
					objMonthSheet.Cells.SetColumnWidth(iCellColumn, 1.5)
					objMonthSheet.Cells.SetColumnWidth(iCellColumn + 1, 1.5)

					objMonthSheet.Cells.Merge(yOffset + 2, iCellColumn, 1, 2)
					iCellColumn += 2
				Next

				' Style the calendar background (Apply weekends)
				iCellColumn = xOffSet + 1
				Dim iEventCount = 500

				' Weekend styles
				For iCount = 0 To objDayPilot.Days - 1
					If dCurrentMonthStart.AddDays(iCount).DayOfWeek = DayOfWeek.Saturday Or dCurrentMonthStart.AddDays(iCount).DayOfWeek = DayOfWeek.Sunday Then
						objRange = objMonthSheet.Cells.CreateRange(yOffset + 3, iCellColumn, iEventCount, 2)
						objRange.SetStyle(stlWeekend)
					End If
					iCellColumn += 2
				Next


				' Day styles
				iCellColumn = xOffSet + 1
				For iCount = 0 To objDayPilot.Days - 1
					If Not dCurrentMonthStart.AddDays(iCount).DayOfWeek = DayOfWeek.Saturday Or dCurrentMonthStart.AddDays(iCount).DayOfWeek = DayOfWeek.Sunday Then
						objRange = objMonthSheet.Cells.CreateRange(yOffset + 3, iCellColumn, iEventCount, 2)
						objRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.LightGray)
						objRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.LightGray)
					End If
					iCellColumn += 2
				Next


				iRow = yOffset + 3
				For Each objDay As Day In objDayPilot.EventDays

					objMonthSheet.Cells(iRow, xOffSet).PutValue(objDay.Name)
					objMonthSheet.Cells(iRow, xOffSet).SetStyle(stlHeading)

					For Each objEvent As DayPilot.Web.Ui.Event In objDay.events

						Dim dEventStart = Max(objEvent.Start, dCurrentMonthStart)
						Dim dEventEnd = Min(objEvent.End, dCurrentMonthEnd)

						iStartRange = dEventStart.Day * 2
						'			iRangeLength = (dEventEnd.Day - dEventStart.Day) * 2

						iRangeLength = CInt(CInt(DateDiff(DateInterval.Hour, dEventStart, dEventEnd)) / 12)

						' Cater for AM
						If dEventStart.Hour > 0 Then
							iStartRange += 1
							'		iRangeLength -= 1
						End If

						' Cater for PM
						'If dEventEnd.Hour >= 12 Then
						'	iRangeLength += 1
						'End If

						' Event itself
						stlEvent.ForegroundColor = ColorTranslator.FromHtml(CType(objEvent.Source, DataRowView).Row(6).ToString())

						'		objMonthSheet.Shapes.AddShape(msoShapeRoundedRectangle)

						' If overlapping event jump to next row
						If iStartRange < iLastRangeEnd Then
							iRow += 1
						End If

						objRange = objMonthSheet.Cells.CreateRange(iRow, iStartRange, 1, iRangeLength)
						objRange.PutValue(objEvent.Name, False, False)
						objRange.SetStyle(stlEvent)
						objRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.LightGray)
						objRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.LightGray)
						objRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.LightGray)
						objRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.LightGray)

						objRange.Merge()

						' Blank text after event to stop over lapping
						objMonthSheet.Cells(iRow, iStartRange + iRangeLength).PutValue(" ")
						objMonthSheet.AutoFitRow(iRow)

						iLastRangeEnd = iStartRange + iRangeLength
					Next

					' Underline current event
					objRange = objMonthSheet.Cells.CreateRange(iRow, yOffset, 1, (objDayPilot.Days * 2) + 1)
					objRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.LightGray)

					iRow += 1
					iLastRangeEnd = 0
				Next


				' Clear styling of unused rows
				objRange = objMonthSheet.Cells.CreateRange(iRow, yOffset, 500, (objDayPilot.Days * 2) + 1)
				objRange.SetStyle(objDocument.DefaultStyle)

				' Place borders
				objRange = objMonthSheet.Cells.CreateRange(xOffSet + 2, yOffset, iRow - xOffSet - 2, (objDayPilot.Days * 2) + 1)
				objRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.LightGray)
				objRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.LightGray)
				objRange.SetOutlineBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.LightGray)
				objRange.SetOutlineBorder(BorderType.RightBorder, CellBorderType.Thin, Color.LightGray)

				objRange = objMonthSheet.Cells.CreateRange(xOffSet + 2, yOffset, 1, 1)
				objRange.SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.LightGray)


				' Display legend
				iLegendColumn = (objDayPilot.Days * 2) + 3
				iRow = yOffset + 2
				objMonthSheet.Cells(iRow, iLegendColumn).PutValue("Legend")
				iRow += 1
				For Each objLegend In Calendar.Legend
					If objLegend.Count > 0 Then
						stlLegend.ForegroundColor = ColorTranslator.FromHtml(objLegend.HexColor)
						objMonthSheet.Cells(iRow, iLegendColumn).SetStyle(stlLegend)
						objMonthSheet.Cells(iRow, iLegendColumn).PutValue(objLegend.LegendDescription)

						iRow += 2
					End If
				Next


				' Autosize some columns
				objMonthSheet.AutoFitColumn(iLegendColumn)
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