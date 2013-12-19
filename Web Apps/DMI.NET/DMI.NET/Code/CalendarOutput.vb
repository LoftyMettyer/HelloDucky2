Imports System.IO
Imports HR.Intranet.Server.Enums
Imports Aspose.Words
Imports Aspose.Words.Tables

Namespace Code
	Public Class CalendarOutput

		Public ReportData As DataTable
		Public Document As MemoryStream
		Public Calendar As HR.Intranet.Server.CalendarReport

		Public Function Generate(ByRef OutputType As OutputFormats) As Boolean

			Dim objDayPilot As New DayPilot.Web.Ui.DayPilotScheduler

			Dim objDocument As New Aspose.Words.Document
			Dim builder As New DocumentBuilder(objDocument)
			Dim table As Aspose.Words.Tables.Table = builder.StartTable()

'			objDayPilot.DataStartField = "startdate"
'			objDayPilot.DataEndField = "enddate"
'			objDayPilot.DataTextField = "description"
'			objDayPilot.DataValueField = "id"
'			objDayPilot.DataResourceField = "resource"

'			BindDataToSchedule(objDayPilot)
'			objDayPilot.DataBind()

'			For Each objEvent As Object In objDayPilot.Resources.ToArrayList()
'				Debug.Print(objEvent.ToString())
'			Next

			builder.PageSetup.Orientation = Orientation.Landscape

			For Each dataRow As DataRow In ReportData.Rows
					For Each item As Object In dataRow.ItemArray
							' Insert a new cell for each object.
							builder.InsertCell()

							Select Case item.GetType().Name
									'Case "Byte[]"
									'		' Assume a byte array is an image. Other data types can be added here.
									'		builder.InsertImage(GetImageFromByteArray(CType(item, Byte())), 50, 50)
									Case "DateTime"
											' Define a custom format for dates and times.
											Dim dateTime As DateTime = CDate(item)
											builder.Write(dateTime.ToString("MMMM d, yyyy"))
									Case Else
											' By default any other item will be inserted as text.
											builder.Write(item.ToString())
							End Select

					Next item

					' After we insert all the data from the current record we can end the table row.
					builder.EndRow()
			Next dataRow

			' We have finished inserting all the data from the DataTable, we can end the table.
			builder.EndTable()

			table.StyleIdentifier = StyleIdentifier.MediumList2Accent1
			table.StyleOptions = TableStyleOptions.FirstRow Or TableStyleOptions.RowBands Or TableStyleOptions.LastColumn
			table.FirstRow.LastCell.RemoveAllChildren()

			Document = New MemoryStream()

			objDocument.Save(Document, SaveFormat.Docx)
			Document.Position = 0

			Return True
		End Function

	Protected Sub BindDataToSchedule(objSchedule As DayPilot.Web.Ui.DayPilotScheduler)

		Dim dt As New DataTable()
		dt.Columns.Add("startdate", GetType(DateTime))
		dt.Columns.Add("enddate", GetType(DateTime))
		dt.Columns.Add("description", GetType(String))
		dt.Columns.Add("baseid", GetType(String))
		dt.Columns.Add("id", GetType(String))
		dt.Columns.Add("resource", GetType(String))
		dt.Columns.Add("color", GetType(String))

		Dim dr As DataRow


		Dim sDescription As String
		Dim sPreviousDescription As String = ""
		Dim sEventDescription As String

		Dim dStart As Date
		Dim dEnd As Date

		For Each objRow In ReportData.Rows

			sEventDescription = objRow("eventdescription1").ToString() & objRow("eventdescription2").ToString()
			sDescription = Calendar.ConvertDescription(objRow("description1").ToString(), objRow("description2").ToString(), objRow("descriptionExpr").ToString())

			' Add to resource collection
			If Not sPreviousDescription = sDescription Then
				objSchedule.Resources.Add(sDescription, objRow("baseid").ToString())
				sPreviousDescription = sDescription
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
				dr("resource") = objRow("baseid")
			dt.Rows.Add(dr)

		Next

	End Sub

	End Class
End Namespace