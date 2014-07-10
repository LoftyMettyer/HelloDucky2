Option Strict On
Option Explicit On

Imports HR.Intranet.Server
Imports DMI.NET.Models
Imports System.Data.SqlClient
Imports HR.Intranet.Server.Metadata
Imports System.Collections.ObjectModel
Imports DMI.NET.Classes
Imports HR.Intranet.Server.Enums
Imports Dapper
Imports DMI.NET.Enums
Imports DMI.NET.Classses

Namespace Repository
	Public Class ReportRepository

		Private _customreports As New Collection(Of CustomReportModel)
		Private _crosstabs As New Collection(Of CrossTabModel)
		Private _calendarreports As New Collection(Of CalendarReportModel)
		Private _mailmerges As New Collection(Of MailMergeModel)

		Private _objSessionInfo As SessionInfo
		Private _objDataAccess As clsDataAccess
		Private _username As String

		Public Sub New()

			MyBase.New()
			_objSessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			_objDataAccess = CType(HttpContext.Current.Session("DatabaseAccess"), clsDataAccess)
			_username = HttpContext.Current.Session("username").ToString

		End Sub

		Public Function LoadCustomReport(ID As Integer, bIsCopy As Boolean, Action As String) As CustomReportModel

			Dim objModel As New CustomReportModel

			Try

				'' TODO -- tidy these up.
				Dim lngAction = HttpContext.Current.Session("action")
				Dim sUserName = HttpContext.Current.Session("username")

				'' TODO - tidy up this proc to return a dataset instead of millions of bloody parameters!
				Dim dsDefinition As DataSet = _objDataAccess.GetDataSet("spASRIntGetCustomReportDefinition" _
					, New SqlParameter("piReportID", SqlDbType.Int) With {.Value = CInt(ID)} _
					, New SqlParameter("psCurrentUser", SqlDbType.VarChar, 255) With {.Value = sUserName} _
					, New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = lngAction})
				'	
				PopulateDefintion(objModel, dsDefinition.Tables(0))

				objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCustomReport, ID, bIsCopy)

				If dsDefinition.Tables(0).Rows.Count = 1 Then

					Dim row As DataRow = dsDefinition.Tables(0).Rows(0)

					objModel.Parent1.ID = CInt(row("Parent1ID"))
					objModel.Parent1.SelectionType = CType(row("Parent1SelectionType"), RecordSelectionType)
					objModel.Parent1.Name = row("Parent1Name").ToString
					objModel.Parent1.PicklistID = CInt(row("Parent1PicklistID"))
					objModel.Parent1.PicklistName = row("Parent1PicklistName").ToString
					objModel.Parent1.FilterID = CInt(row("Parent1FilterID"))
					objModel.Parent1.FilterName = row("Parent1FilterName").ToString

					objModel.Parent2.ID = CInt(row("Parent2ID"))
					objModel.Parent2.SelectionType = CType(row("Parent2SelectionType"), RecordSelectionType)
					objModel.Parent2.Name = row("Parent2Name").ToString
					objModel.Parent2.PicklistID = CInt(row("Parent2PicklistID"))
					objModel.Parent2.PicklistName = row("Parent2PicklistName").ToString
					objModel.Parent2.FilterID = CInt(row("Parent2FilterID"))
					objModel.Parent2.FilterName = row("Parent2FilterName").ToString

				End If

				objModel.Columns.BaseTableID = objModel.BaseTableID

				'' todo replace with dapper
				objModel.Columns.Selected = New Collection(Of ReportColumnItem)
				objModel.Columns.AvailableTables = GetTables()

				For Each objRow As DataRow In dsDefinition.Tables(1).Rows
					Dim objItem As New ReportColumnItem() With {
						.CustomReportId = ID,
						.Heading = objRow("Heading").ToString,
						.id = CInt(objRow("id")),
						.Name = objRow("Name").ToString,
						.Sequence = CInt(objRow("Sequence")),
						.Size = CInt(objRow("Size")),
						.Decimals = CInt(objRow("Decimals")),
						.IsAverage = CBool(objRow("IsAverage")),
						.IsCount = CBool(objRow("IsCount")),
						.IsTotal = CBool(objRow("IsTotal")),
						.IsHidden = CBool(objRow("IsHidden")),
						.IsGroupWithNext = CBool(objRow("IsGroupWithNext"))}
					objModel.Columns.Selected.Add(objItem)

				Next

				' Repetition
				For Each objRow As DataRow In dsDefinition.Tables(3).Rows
					Dim objRepeatItem As New ReportRepetition() With {
							.ID = CInt(objRow("id")),
							.Name = objRow("Name").ToString,
							.IsExpression = CBool(objRow("IsExpression")),
							.IsRepeated = CBool(objRow("IsRepeated"))}
					objModel.Repetition.Add(objRepeatItem)
				Next

				PopulateSortOrder(objModel, dsDefinition.Tables(2))
				PopulateOutput(objModel.Output, dsDefinition.Tables(0))

				' Populate the child tables
				For Each objRow As DataRow In dsDefinition.Tables(4).Rows
					objModel.ChildTables.Add(New ReportChildTables() With {
									.TableName = objRow("tablename").ToString,
									.FilterName = objRow("filtername").ToString,
									.OrderName = objRow("ordername").ToString,
									.TableID = CInt(objRow("tableid")),
									.FilterID = CInt(objRow("filterid")),
									.OrderID = CInt(objRow("orderid")),
									.Records = CInt(objRow("Records"))})
				Next

				If bIsCopy Then
					objModel.ID = 0
				Else
					objModel.ID = ID
				End If

				_customreports.Add(objModel)

			Catch ex As Exception
				Throw

			End Try

			Return objModel

		End Function

		Public Function LoadMailMerge(ID As Integer, bIsCopy As Boolean, Action As String) As MailMergeModel

			Dim objModel As New MailMergeModel
			Dim objItem As ReportColumnItem

			Dim dsDefinition = _objDataAccess.GetDataSet("spASRIntGetMailMergeDefinition" _
				, New SqlParameter("@piReportID", SqlDbType.Int) With {.Value = ID} _
				, New SqlParameter("@psCurrentUser", SqlDbType.VarChar, 255) With {.Value = _username} _
				, New SqlParameter("@psAction", SqlDbType.VarChar, 255) With {.Value = Action})

			PopulateDefintion(objModel, dsDefinition.Tables(0))
			objModel.GroupAccess = GetUtilityAccess(UtilityType.utlMailMerge, ID, bIsCopy)

			' Columns
			For Each objRow As DataRow In dsDefinition.Tables(1).Rows
				objItem = New ReportColumnItem
				objItem.IsExpression = False
				objItem.IsHidden = False
				objItem.id = CInt(objRow("columnId"))
				objItem.Name = objRow("name").ToString
				objItem.Heading = objRow("heading").ToString
				objItem.DataType = CType(objRow("datatype"), SQLDataType)
				objItem.Size = CInt(objRow("size"))
				objItem.Decimals = CInt(objRow("decimals"))
				objModel.Columns.Selected.Add(objItem)
			Next

			' Eexpressions
			For Each objRow As DataRow In dsDefinition.Tables(2).Rows
				objItem = New ReportColumnItem
				objItem.IsExpression = True
				objItem.IsHidden = CBool(objRow("ishidden"))
				objItem.id = CInt(objRow("columnId"))
				objItem.Name = objRow("name").ToString
				objItem.Heading = objRow("heading").ToString
				objItem.DataType = SQLDataType.sqlUnknown
				objItem.Size = CInt(objRow("size"))
				objItem.Decimals = CInt(objRow("decimals"))
				objModel.Columns.Selected.Add(objItem)
			Next

			' Orders (expressions)
			PopulateSortOrder(objModel, dsDefinition.Tables(3))

			objModel.Columns.BaseTableID = objModel.BaseTableID

			If dsDefinition.Tables(0).Rows.Count = 1 Then

				Dim row As DataRow = dsDefinition.Tables(0).Rows(0)

				objModel.TemplateFileName = row("TemplateFileName").ToString()
				objModel.OutputFormat = CType(row("Format"), MailMergeOutputTypes)
				objModel.DisplayOutputOnScreen = CBool(row("DisplayOutputOnScreen"))
				objModel.SendToPrinter = CBool(row("SendToPrinter"))
				objModel.PrinterName = row("PrinterName").ToString()
				objModel.SaveTofile = CBool(row("SaveTofile"))
				objModel.Filename = row("FileName").ToString
				objModel.EmailGroupID = CInt(row("EmailGroupID"))
				objModel.EmailSubject = row("EmailSubject").ToString()
				objModel.EmailAsAttachment = CBool(row("EmailAsAttachment"))
				objModel.EmailAttachmentName = row("EmailAttachmentName").ToString()

				objModel.SuppressBlankLines = CBool(row("SuppressBlankLines"))
				objModel.PauseBeforeMerge = CBool(row("PauseBeforeMerge"))

			End If

			If bIsCopy Then
				objModel.ID = 0
			Else
				objModel.ID = ID
			End If

			_mailmerges.Add(objModel)

			Return objModel

		End Function

		Public Function NewCrossTab() As CrossTabModel

			Dim objModel As New CrossTabModel

			objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCrossTab, 0, False)

			Return objModel

		End Function

		Public Function NewCalendarReport() As CalendarReportModel

			Dim objModel As New CalendarReportModel

			objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCalendarReport, 0, False)

			Return objModel

		End Function

		Public Function NewCustomReport() As CustomReportModel

			Dim objModel As New CustomReportModel

			objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCustomReport, 0, False)

			Return objModel

		End Function

		Public Function NewMailMerge() As MailMergeModel

			Dim objModel As New MailMergeModel

			objModel.GroupAccess = GetUtilityAccess(UtilityType.utlMailMerge, 0, False)

			Return objModel

		End Function

		Public Function LoadCrossTab(ID As Integer, bIsCopy As Boolean, Action As String) As CrossTabModel

			Dim objModel As New CrossTabModel

			Try

				Dim dtDefinition = _objDataAccess.GetFromSP("spASRIntGetCrossTabDefinition", _
						New SqlParameter("piReportID", SqlDbType.Int) With {.Value = ID}, _
						New SqlParameter("psCurrentUser", SqlDbType.VarChar, 255) With {.Value = _username}, _
						New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = Action})

				PopulateDefintion(objModel, dtDefinition)

				If dtDefinition.Rows.Count = 1 Then
					Dim objRow As DataRow = dtDefinition.Rows(0)

					objModel.HorizontalID = CInt(objRow("HorizontalID"))
					objModel.HorizontalDataType = _objSessionInfo.GetColumn(objModel.HorizontalID).DataType
					objModel.HorizontalStart = CInt(objRow("HorizontalStart"))
					objModel.HorizontalStop = CInt(objRow("HorizontalStop"))
					objModel.HorizontalIncrement = CInt(objRow("HorizontalIncrement"))

					objModel.VerticalID = CInt(objRow("VerticalID"))
					objModel.VerticalDataType = _objSessionInfo.GetColumn(objModel.VerticalID).DataType
					objModel.VerticalStart = CInt(objRow("VerticalStart"))
					objModel.VerticalStop = CInt(objRow("VerticalStop"))
					objModel.VerticalIncrement = CInt(objRow("VerticalIncrement"))

					objModel.PageBreakID = CInt(objRow("PageBreakID"))
					objModel.PageBreakDataType = _objSessionInfo.GetColumn(objModel.PageBreakID).DataType
					objModel.PageBreakStart = CInt(objRow("PageBreakStart"))
					objModel.PageBreakStop = CInt(objRow("PageBreakStop"))
					objModel.PageBreakIncrement = CInt(objRow("PageBreakIncrement"))

					objModel.IntersectionID = CInt(objRow("IntersectionID"))
					objModel.IntersectionType = CType(objRow("IntersectionType"), IntersectionType)
					objModel.PercentageOfType = CBool(objRow("PercentageOfType"))
					objModel.PercentageOfPage = CBool(objRow("PercentageOfPage"))
					objModel.SuppressZeros = CBool(objRow("SuppressZeros"))
					objModel.UseThousandSeparators = CBool(objRow("UseThousandSeparators"))

				End If

				objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCrossTab, ID, bIsCopy)

				' Columns tab
				' '' TODO - Load columns tab (needs dynamic based on table selection)
				objModel.AvailableColumns = GetColumnsForTable(objModel.BaseTableID)

				' Output Tab
				PopulateOutput(objModel.Output, dtDefinition)

			Catch ex As Exception
				Throw

			End Try

			If bIsCopy Then
				objModel.ID = 0
			Else
				objModel.ID = ID
			End If

			Return objModel

		End Function

		Public Function LoadCalendarReport(ID As Integer, bIsCopy As Boolean, Action As String) As CalendarReportModel

			Dim objModel As New CalendarReportModel
			Dim objEvent As CalendarEventDetail

			Dim dsDefinition = _objDataAccess.GetDataSet("spASRIntGetCalendarReportDefinition", _
					New SqlParameter("@piCalendarReportID", SqlDbType.Int) With {.Value = ID}, _
					New SqlParameter("psCurrentUser", SqlDbType.VarChar, 255) With {.Value = _username}, _
					New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = Action})

			PopulateDefintion(objModel, dsDefinition.Tables(0))
			If dsDefinition.Tables(0).Rows.Count = 1 Then

				Dim row As DataRow = dsDefinition.Tables(0).Rows(0)

				objModel.Description1Id = CInt(row("Description1Id"))
				objModel.Description2Id = CInt(row("Description2Id"))
				objModel.Description3Id = CInt(row("Description3Id"))
				objModel.Description3Name = row("Description3Name").ToString

				objModel.RegionID = CInt(row("RegionID"))
				objModel.GroupByDescription = CBool(row("GroupByDescription"))
				objModel.Separator = row("Separator").ToString

				objModel.StartType = CType(row("StartType"), CalendarDataType)
				objModel.StartFixedDate = CDate(row("StartFixedDate"))
				objModel.StartOffset = CInt(row("StartOffset"))
				objModel.StartOffsetPeriod = CType(row("StartOffsetPeriod"), DatePeriod)
				objModel.StartCustomId = CInt(row("StartCustomId"))
				objModel.StartCustomName = row("StartCustomName").ToString

				objModel.EndType = CType(row("EndType"), CalendarDataType)
				objModel.EndFixedDate = CDate(row("EndFixedDate"))
				objModel.EndOffset = CInt(row("EndOffset"))
				objModel.EndOffsetPeriod = CType(row("EndOffsetPeriod"), DatePeriod)
				objModel.EndCustomId = CInt(row("EndCustomId"))
				objModel.EndCustomName = row("EndCustomName").ToString

				objModel.IncludeBankHolidays = CBool(row("IncludeBankHolidays"))
				objModel.WorkingDaysOnly = CBool(row("WorkingDaysOnly"))
				objModel.ShowBankHolidays = CBool(row("ShowBankHolidays"))
				objModel.ShowCaptions = CBool(row("ShowCaptions"))
				objModel.ShowWeekends = CBool(row("ShowWeekends"))
				objModel.StartOnCurrentMonth = CBool(row("StartOnCurrentMonth"))

			End If


			' Replace with Automapper?
			For Each objRow As DataRow In dsDefinition.Tables(1).Rows
				objEvent = New CalendarEventDetail

				objEvent.ID = CInt(objRow("ID"))
				objEvent.Name = objRow("Name").ToString
				objEvent.EventKey = objRow("EventKey").ToString
				objEvent.CalendarReportID = ID
				objEvent.TableID = CInt(objRow("TableID"))
				objEvent.FilterID = CInt(objRow("FilterID"))
				objEvent.FilterName = objRow("FilterName").ToString
				objEvent.EventStartDateID = CInt(objRow("EventStartDateID"))
				objEvent.EventStartSessionID = CInt(objRow("EventStartSessionID"))
				objEvent.EventStartSessionName = objRow("EventStartSessionName").ToString
				objEvent.EventEndDateID = CInt(objRow("EventEndDateID"))
				objEvent.EventEndDateName = objRow("EventEndDateName").ToString
				objEvent.EventEndSessionID = CInt(objRow("EventEndSessionID"))
				objEvent.EventDurationName = objRow("EventDurationName").ToString
				objEvent.EventEndSessionName = objRow("EventEndSessionName").ToString
				objEvent.EventDurationID = CInt(objRow("EventDurationID"))
				objEvent.LegendType = objRow("LegendType").ToString
				objEvent.LegendTypeName = objRow("LegendTypeName").ToString
				objEvent.LegendCharacter = objRow("LegendCharacter").ToString
				objEvent.LegendLookupTableID = CInt(objRow("LegendLookupTableID"))
				objEvent.LegendLookupColumnID = CInt(objRow("LegendLookupColumnID"))
				objEvent.LegendLookupCodeID = CInt(objRow("LegendLookupCodeID"))
				objEvent.LegendEventColumnID = CInt(objRow("LegendEventColumnID"))
				objEvent.EventDesc1ColumnID = CInt(objRow("EventDesc1ColumnID"))
				objEvent.EventDesc1ColumnName = objRow("EventDesc1ColumnName").ToString
				objEvent.EventDesc2ColumnID = CInt(objRow("EventDesc2ColumnID"))
				objEvent.EventDesc2ColumnName = objRow("EventDesc2ColumnName").ToString
				objEvent.FilterHidden = objRow("FilterHidden").ToString

				objModel.Events.Add(objEvent)

			Next

			PopulateSortOrder(objModel, dsDefinition.Tables(2))

			objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCalendarReport, ID, bIsCopy)

			If bIsCopy Then
				objModel.ID = 0
			Else
				objModel.ID = ID
			End If

			Return objModel

		End Function

		Public Function SaveReportDefinition(objModel As MailMergeModel) As Boolean

			Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = objModel.ID}

			' TODO old access stuff - needs updating
			Dim psJobsToHide As String = ""	' Request.Form("txtSend_jobsToHide")
			Dim psJobsToHideGroups As String = ""	' Request.Form("txtSend_jobsToHideGroups")}

			Dim sAccess = UtilityAccessAsString(objModel.GroupAccess)
			Dim sColumns = ReportColumnsAsString(objModel.Columns.Selected, objModel.SortOrderColumns)

			_objDataAccess.ExecuteSP("sp_ASRIntSaveMailMerge" _
				, New SqlParameter("@psName", SqlDbType.VarChar, 255) With {.Value = objModel.Name} _
				, New SqlParameter("@psDescription", SqlDbType.VarChar, -1) With {.Value = objModel.Description} _
				, New SqlParameter("@piTableID", SqlDbType.Int) With {.Value = objModel.BaseTableID} _
				, New SqlParameter("@piSelection", SqlDbType.Int) With {.Value = objModel.SelectionType} _
				, New SqlParameter("@piPicklistID", SqlDbType.Int) With {.Value = objModel.PicklistID} _
				, New SqlParameter("@piFilterID", SqlDbType.Int) With {.Value = objModel.FilterID} _
				, New SqlParameter("@piOutputFormat", SqlDbType.Int) With {.Value = objModel.OutputFormat} _
				, New SqlParameter("@pfOutputSave", SqlDbType.Bit) With {.Value = True} _
				, New SqlParameter("@psOutputFilename", SqlDbType.VarChar, -1) With {.Value = objModel.Filename} _
				, New SqlParameter("@piEmailAddrID", SqlDbType.Int) With {.Value = objModel.EmailGroupID} _
				, New SqlParameter("@psEmailSubject", SqlDbType.VarChar, -1) With {.Value = objModel.EmailSubject} _
				, New SqlParameter("@psTemplateFileName", SqlDbType.VarChar, -1) With {.Value = objModel.TemplateFileName} _
				, New SqlParameter("@pfOutputScreen", SqlDbType.Bit) With {.Value = objModel.DisplayOutputOnScreen} _
				, New SqlParameter("@psUserName", SqlDbType.VarChar, 255) With {.Value = objModel.Owner} _
				, New SqlParameter("@pfEmailAsAttachment", SqlDbType.Bit) With {.Value = objModel.EmailAsAttachment} _
				, New SqlParameter("@psEmailAttachmentName", SqlDbType.VarChar, -1) With {.Value = objModel.EmailAttachmentName} _
				, New SqlParameter("@pfSuppressBlanks", SqlDbType.Bit) With {.Value = objModel.SuppressBlankLines} _
				, New SqlParameter("@pfPauseBeforeMerge", SqlDbType.Bit) With {.Value = objModel.PauseBeforeMerge} _
				, New SqlParameter("@pfOutputPrinter", SqlDbType.Bit) With {.Value = objModel.SendToPrinter} _
				, New SqlParameter("@psOutputPrinterName", SqlDbType.VarChar, 255) With {.Value = objModel.PrinterName} _
				, New SqlParameter("@piDocumentMapID", SqlDbType.Int) With {.Value = 0} _
				, New SqlParameter("@pfManualDocManHeader", SqlDbType.Bit) With {.Value = False} _
				, New SqlParameter("@psAccess", SqlDbType.VarChar, -1) With {.Value = sAccess} _
				, New SqlParameter("@psJobsToHide", SqlDbType.VarChar, -1) With {.Value = ""} _
				, New SqlParameter("@psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = ""} _
				, New SqlParameter("@psColumns", SqlDbType.VarChar, -1) With {.Value = sColumns} _
				, New SqlParameter("@psColumns2", SqlDbType.VarChar, -1) With {.Value = ""} _
			, prmID)

			_mailmerges.Remove(objModel)

			Return True

		End Function

		Public Function SaveReportDefinition(objModel As CrossTabModel) As Boolean

			Try

				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = objModel.ID}

				Dim sAccess As String = UtilityAccessAsString(objModel.GroupAccess)

				Dim psJobsToHide As String = ""	' Request.Form("txtSend_jobsToHide")
				Dim psJobsToHideGroups As String = ""	' Request.Form("txtSend_jobsToHideGroups")}

				_objDataAccess.ExecuteSP("sp_ASRIntSaveCrossTab", _
						New SqlParameter("psName", SqlDbType.VarChar, 255) With {.Value = objModel.Name}, _
						New SqlParameter("psDescription", SqlDbType.VarChar, -1) With {.Value = objModel.Description}, _
						New SqlParameter("piTableID", SqlDbType.Int) With {.Value = objModel.BaseTableID}, _
						New SqlParameter("piSelection", SqlDbType.Int) With {.Value = objModel.SelectionType}, _
						New SqlParameter("piPicklistID", SqlDbType.Int) With {.Value = objModel.PicklistID}, _
						New SqlParameter("piFilterID", SqlDbType.Int) With {.Value = objModel.FilterID}, _
						New SqlParameter("pfPrintFilter", SqlDbType.Bit) With {.Value = objModel.DisplayTitleInReportHeader}, _
						New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = objModel.Owner}, _
						New SqlParameter("piHColID", SqlDbType.Int) With {.Value = objModel.HorizontalID}, _
						New SqlParameter("psHStart", SqlDbType.VarChar, 100) With {.Value = objModel.HorizontalStart}, _
						New SqlParameter("psHStop", SqlDbType.VarChar, 100) With {.Value = objModel.HorizontalStop}, _
						New SqlParameter("psHStep", SqlDbType.VarChar, 100) With {.Value = objModel.HorizontalIncrement}, _
						New SqlParameter("piVColID", SqlDbType.Int) With {.Value = objModel.VerticalID}, _
						New SqlParameter("psVStart", SqlDbType.VarChar, 100) With {.Value = objModel.VerticalStart}, _
						New SqlParameter("psVStop", SqlDbType.VarChar, 100) With {.Value = objModel.VerticalStop}, _
						New SqlParameter("psVStep", SqlDbType.VarChar, 100) With {.Value = objModel.VerticalIncrement}, _
						New SqlParameter("piPColID", SqlDbType.Int) With {.Value = objModel.PageBreakID}, _
						New SqlParameter("psPStart", SqlDbType.VarChar, 100) With {.Value = objModel.PageBreakStart}, _
						New SqlParameter("psPStop", SqlDbType.VarChar, 100) With {.Value = objModel.PageBreakStop}, _
						New SqlParameter("psPStep", SqlDbType.VarChar, 100) With {.Value = objModel.PageBreakIncrement}, _
						New SqlParameter("piIType", SqlDbType.Int) With {.Value = objModel.IntersectionType}, _
						New SqlParameter("piIColID", SqlDbType.Int) With {.Value = objModel.IntersectionID}, _
						New SqlParameter("pfPercentage", SqlDbType.Bit) With {.Value = objModel.PercentageOfType}, _
						New SqlParameter("pfPerPage", SqlDbType.Bit) With {.Value = objModel.PercentageOfPage}, _
						New SqlParameter("pfSuppress", SqlDbType.Bit) With {.Value = objModel.SuppressZeros}, _
						New SqlParameter("pfUse1000Separator", SqlDbType.Bit) With {.Value = objModel.UseThousandSeparators}, _
						New SqlParameter("pfOutputPreview", SqlDbType.Bit) With {.Value = objModel.Output.IsPreview}, _
						New SqlParameter("piOutputFormat", SqlDbType.Int) With {.Value = objModel.Output.Format}, _
						New SqlParameter("pfOutputScreen", SqlDbType.Bit) With {.Value = objModel.Output.ToScreen}, _
						New SqlParameter("pfOutputPrinter", SqlDbType.Bit) With {.Value = objModel.Output.ToPrinter}, _
						New SqlParameter("psOutputPrinterName", SqlDbType.VarChar, -1) With {.Value = objModel.Output.PrinterName}, _
						New SqlParameter("pfOutputSave", SqlDbType.Bit) With {.Value = objModel.Output.SaveToFile}, _
						New SqlParameter("piOutputSaveExisting", SqlDbType.Int) With {.Value = objModel.Output.SaveExisting}, _
						New SqlParameter("pfOutputEmail", SqlDbType.Bit) With {.Value = objModel.Output.SendToEmail}, _
						New SqlParameter("piOutputEmailAddr", SqlDbType.Int) With {.Value = objModel.Output.EmailGroupID}, _
						New SqlParameter("psOutputEmailSubject", SqlDbType.VarChar, -1) With {.Value = objModel.Output.EmailSubject}, _
						New SqlParameter("psOutputEmailAttachAs", SqlDbType.VarChar, -1) With {.Value = objModel.Output.EmailAttachmentName}, _
						New SqlParameter("psOutputFilename", SqlDbType.VarChar, -1) With {.Value = objModel.Output.Filename}, _
						New SqlParameter("psAccess", SqlDbType.VarChar, -1) With {.Value = sAccess}, _
						New SqlParameter("psJobsToHide", SqlDbType.VarChar, -1) With {.Value = psJobsToHide}, _
						New SqlParameter("psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = psJobsToHideGroups}, _
						prmID)

			Catch
				Throw

			End Try

			Return True
		End Function

		Public Function SaveReportDefinition(objModel As CustomReportModel) As Boolean

			_customreports.Remove(objModel.ID)

			Return True	' TODO


			Try

				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = objModel.ID}

				Dim sAccess As String = UtilityAccessAsString(objModel.GroupAccess)
				Dim sJobsToHide = JobsToHideAsString(objModel.JobsToHide)
				Dim sJobsToHideGroups As String = "" ' TODO?
				Dim sColumns = ReportColumnsAsString(objModel.Columns.Selected, objModel.SortOrderColumns)
				Dim sChildren As String = ReportChildTablesAsString(objModel.ChildTables)

				_objDataAccess.ExecuteSP("sp_ASRIntSaveCustomReport", _
						New SqlParameter("psName", SqlDbType.VarChar, 255) With {.Value = objModel.Name}, _
						New SqlParameter("psDescription", SqlDbType.VarChar, -1) With {.Value = objModel.Description}, _
						New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Value = objModel.BaseTableID}, _
						New SqlParameter("pfAllRecords", SqlDbType.Bit) With {.Value = (objModel.PicklistID = 0 And objModel.FilterID = 0)}, _
						New SqlParameter("piPicklistID", SqlDbType.Int) With {.Value = objModel.PicklistID}, _
						New SqlParameter("piFilterID", SqlDbType.Int) With {.Value = objModel.FilterID}, _
						New SqlParameter("piParent1TableID", SqlDbType.Int) With {.Value = objModel.Parent1.ID}, _
						New SqlParameter("piParent1FilterID", SqlDbType.Int) With {.Value = objModel.Parent1.FilterID}, _
						New SqlParameter("piParent2TableID", SqlDbType.Int) With {.Value = objModel.Parent2.ID}, _
						New SqlParameter("piParent2FilterID", SqlDbType.Int) With {.Value = objModel.Parent2.FilterID}, _
						New SqlParameter("pfSummary", SqlDbType.Bit) With {.Value = objModel.IsSummary}, _
						New SqlParameter("pfPrintFilterHeader", SqlDbType.Bit) With {.Value = objModel.DisplayTitleInReportHeader}, _
						New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = objModel.Owner}, _
						New SqlParameter("pfOutputPreview", SqlDbType.Bit) With {.Value = objModel.Output.IsPreview}, _
						New SqlParameter("piOutputFormat", SqlDbType.Int) With {.Value = objModel.Output.Format}, _
						New SqlParameter("pfOutputScreen", SqlDbType.Bit) With {.Value = objModel.Output.ToScreen}, _
						New SqlParameter("pfOutputPrinter", SqlDbType.Bit) With {.Value = objModel.Output.ToPrinter}, _
						New SqlParameter("psOutputPrinterName", SqlDbType.VarChar, -1) With {.Value = objModel.Output.PrinterName}, _
						New SqlParameter("pfOutputSave", SqlDbType.Bit) With {.Value = objModel.Output.SaveToFile}, _
						New SqlParameter("piOutputSaveExisting", SqlDbType.Int) With {.Value = objModel.Output.SaveExisting}, _
						New SqlParameter("pfOutputEmail", SqlDbType.Bit) With {.Value = objModel.Output.SendToEmail}, _
						New SqlParameter("piOutputEmailAddr", SqlDbType.Int) With {.Value = objModel.Output.EmailGroupID}, _
						New SqlParameter("psOutputEmailSubject", SqlDbType.VarChar, -1) With {.Value = objModel.Output.EmailSubject}, _
						New SqlParameter("psOutputEmailAttachAs", SqlDbType.VarChar, -1) With {.Value = objModel.Output.EmailAttachmentName}, _
						New SqlParameter("psOutputFilename", SqlDbType.VarChar, -1) With {.Value = objModel.Output.Filename}, _
						New SqlParameter("pfParent1AllRecords", SqlDbType.Bit) With {.Value = (objModel.Parent1.PicklistID = 0 And objModel.Parent1.FilterID = 0)}, _
						New SqlParameter("piParent1Picklist", SqlDbType.Int) With {.Value = objModel.Parent1.PicklistID}, _
						New SqlParameter("pfParent2AllRecords", SqlDbType.Bit) With {.Value = (objModel.Parent2.PicklistID = 0 And objModel.Parent2.FilterID = 0)}, _
						New SqlParameter("piParent2Picklist", SqlDbType.Int) With {.Value = objModel.Parent2.PicklistID}, _
						New SqlParameter("psAccess", SqlDbType.VarChar, -1) With {.Value = sAccess}, _
						New SqlParameter("psJobsToHide", SqlDbType.VarChar, -1) With {.Value = sJobsToHide}, _
						New SqlParameter("psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = sJobsToHideGroups}, _
						New SqlParameter("psColumns", SqlDbType.VarChar, -1) With {.Value = sColumns}, _
						New SqlParameter("psColumns2", SqlDbType.VarChar, -1) With {.Value = ""}, _
						New SqlParameter("psChildString", SqlDbType.VarChar, -1) With {.Value = sChildren}, _
						prmID,
						New SqlParameter("pfIgnoreZeros", SqlDbType.Bit) With {.Value = objModel.IgnoreZerosForAggregates})


				'Dim sSQL As String = String.Format("UPDATE ASRSysCustomReportsName SET Name = '{1}', OutputEmailAttachAs = '{2}' WHERE ID = {0}" _
				'															, objModel.ID, objModel.Name, objModel.Output.EmailAttachmentName)
				'objDataAccess.ExecuteSql(sSQL)

				'sSQL = String.Format("DELETE ASRSysCustomReportsChildDetails WHERE CustomReportID = {0}", objModel.ID)
				'objDataAccess.ExecuteSql(sSQL)

				'For Each objChild In objModel.ChildTables
				'	sSQL = String.Format("INSERT ASRSysCustomReportsChildDetails (CustomReportID, ChildTable, ChildFilter, ChildMaxRecords, ChildOrder) VALUES ({0}, {1}, {2}, {3}, {4})", _
				'											objModel.ID, objChild.TableID, objChild.FilterID, objChild.Records, objChild.OrderID)
				'	objDataAccess.ExecuteSql(sSQL)
				'Next


				'sSQL = String.Format("DELETE ASRSysCustomReportAccess WHERE ID = {0}", objModel.ID)
				'objDataAccess.ExecuteSql(sSQL)

				'For Each objChild In objModel.GroupAccess
				'	sSQL = String.Format("INSERT ASRSysCustomReportAccess (id, groupname, access) VALUES ({0}, '{1}', '{2}')", _
				'											objModel.ID, objChild.Name, objChild.Access)
				'	objDataAccess.ExecuteSql(sSQL)
				'Next



			Catch ex As Exception
				Throw

			End Try

			Return True

		End Function

		Public Function SaveReportDefinition(objModel As CalendarReportModel) As Boolean

			Return True	' TODO

			Try

				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = objModel.ID}

				Dim sAccess = UtilityAccessAsString(objModel.GroupAccess)
				Dim sJobsToHide = JobsToHideAsString(objModel.JobsToHide)
				Dim sJobsToHideGroups As String = "" ' TODO?
				Dim sEvents As String = "" 'TODO
				Dim sReportOrder As String = ""	'TODO
				Dim bAllRecords As Boolean

				' Calendar reports don't save the selection type - instead they have a boolean allrecords flag
				If objModel.PicklistID = 0 And objModel.FilterID = 0 Then bAllRecords = True

				_objDataAccess.ExecuteSP("spASRIntSaveCalendarReport", _
				New SqlParameter("psName", SqlDbType.VarChar, 255) With {.Value = objModel.Name}, _
					New SqlParameter("psDescription", SqlDbType.VarChar, -1) With {.Value = objModel.Description}, _
					New SqlParameter("piBaseTable", SqlDbType.Int) With {.Value = objModel.BaseTableID}, _
					New SqlParameter("pfAllRecords", SqlDbType.Bit) With {.Value = bAllRecords}, _
					New SqlParameter("piPicklist", SqlDbType.Int) With {.Value = objModel.PicklistID}, _
					New SqlParameter("piFilter", SqlDbType.Int) With {.Value = objModel.FilterID}, _
					New SqlParameter("pfPrintFilterHeader", SqlDbType.Bit) With {.Value = objModel.DisplayTitleInReportHeader}, _
					New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = objModel.Owner}, _
					New SqlParameter("piDescription1", SqlDbType.Int) With {.Value = objModel.Description1Id}, _
					New SqlParameter("piDescription2", SqlDbType.Int) With {.Value = objModel.Description2Id}, _
					New SqlParameter("piDescriptionExpr", SqlDbType.Int) With {.Value = objModel.Description3Id}, _
					New SqlParameter("piRegion", SqlDbType.Int) With {.Value = objModel.RegionID}, _
					New SqlParameter("pfGroupByDesc", SqlDbType.Bit) With {.Value = objModel.GroupByDescription}, _
					New SqlParameter("psDescSeparator", SqlDbType.VarChar, 100) With {.Value = objModel.Separator}, _
					New SqlParameter("piStartType", SqlDbType.Int) With {.Value = objModel.StartType}, _
					New SqlParameter("psFixedStart", SqlDbType.VarChar) With {.Value = objModel.StartFixedDate}, _
					New SqlParameter("piStartFrequency", SqlDbType.Int) With {.Value = objModel.StartOffsetPeriod}, _
					New SqlParameter("piStartPeriod", SqlDbType.Int) With {.Value = objModel.StartOffsetPeriod}, _
					New SqlParameter("piStartDateExpr", SqlDbType.Int) With {.Value = objModel.StartCustomId}, _
					New SqlParameter("piEndType", SqlDbType.Int) With {.Value = objModel.EndType}, _
					New SqlParameter("psFixedEnd", SqlDbType.VarChar) With {.Value = objModel.EndFixedDate}, _
					New SqlParameter("piEndFrequency", SqlDbType.Int) With {.Value = objModel.EndOffset}, _
					New SqlParameter("piEndPeriod", SqlDbType.Int) With {.Value = objModel.EndOffsetPeriod}, _
					New SqlParameter("piEndDateExpr", SqlDbType.Int) With {.Value = objModel.EndCustomId}, _
					New SqlParameter("pfShowBankHols", SqlDbType.Bit) With {.Value = objModel.ShowBankHolidays}, _
					New SqlParameter("pfShowCaptions", SqlDbType.Bit) With {.Value = objModel.ShowCaptions}, _
					New SqlParameter("pfShowWeekends", SqlDbType.Bit) With {.Value = objModel.ShowWeekends}, _
					New SqlParameter("pfStartOnCurrentMonth", SqlDbType.Bit) With {.Value = objModel.StartOnCurrentMonth}, _
					New SqlParameter("pfIncludeWorkdays", SqlDbType.Bit) With {.Value = objModel.WorkingDaysOnly}, _
					New SqlParameter("pfIncludeBankHols", SqlDbType.Bit) With {.Value = objModel.IncludeBankHolidays}, _
					New SqlParameter("pfOutputPreview", SqlDbType.Bit) With {.Value = objModel.Output.IsPreview}, _
					New SqlParameter("piOutputFormat", SqlDbType.Int) With {.Value = objModel.Output.Format}, _
					New SqlParameter("pfOutputScreen", SqlDbType.Bit) With {.Value = objModel.Output.ToScreen}, _
					New SqlParameter("pfOutputPrinter", SqlDbType.Bit) With {.Value = objModel.Output.ToPrinter}, _
					New SqlParameter("psOutputPrinterName", SqlDbType.VarChar, -1) With {.Value = objModel.Output.PrinterName}, _
					New SqlParameter("pfOutputSave", SqlDbType.Bit) With {.Value = objModel.Output.SaveToFile}, _
					New SqlParameter("piOutputSaveExisting", SqlDbType.Int) With {.Value = objModel.Output.SaveExisting}, _
					New SqlParameter("pfOutputEmail", SqlDbType.Bit) With {.Value = objModel.Output.SendToEmail}, _
					New SqlParameter("piOutputEmailAddr", SqlDbType.Int) With {.Value = objModel.Output.EmailGroupID}, _
					New SqlParameter("psOutputEmailSubject", SqlDbType.VarChar, -1) With {.Value = objModel.Output.EmailSubject}, _
					New SqlParameter("psOutputEmailAttachAs", SqlDbType.VarChar, -1) With {.Value = objModel.Output.EmailAttachmentName}, _
					New SqlParameter("psOutputFilename", SqlDbType.VarChar, -1) With {.Value = objModel.Output.Filename}, _
					New SqlParameter("psAccess", SqlDbType.VarChar, -1) With {.Value = sAccess}, _
					New SqlParameter("psJobsToHide", SqlDbType.VarChar, -1) With {.Value = sJobsToHide}, _
					New SqlParameter("psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = sJobsToHideGroups}, _
					New SqlParameter("psEvents", SqlDbType.VarChar, -1) With {.Value = sEvents}, _
					New SqlParameter("psEvents2", SqlDbType.VarChar, -1) With {.Value = ""}, _
					New SqlParameter("psOrderString", SqlDbType.VarChar, -1) With {.Value = sReportOrder}, _
					prmID)

			Catch
				Throw

			End Try

			Return True
		End Function

		Private Function GetUtilityAccess(utilType As UtilityType, ID As Integer, IsCopy As Boolean) As Collection(Of GroupAccess)

			Dim objAccess As New Collection(Of GroupAccess)

			Try

				'Dim con = objDataAccess.Connection

				'objAccess = con.Query(Of GroupAccess)("spASRIntGetUtilityAccessRecords")
				'objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCalendarReport, ID, bIsCopy)


				Dim rstAccessInfo As DataTable = _objDataAccess.GetDataTable("spASRIntGetUtilityAccessRecords", CommandType.StoredProcedure _
					, New SqlParameter("piUtilityType", SqlDbType.Int) With {.Value = CInt(utilType)} _
					, New SqlParameter("piID", SqlDbType.Int) With {.Value = ID} _
					, New SqlParameter("piFromCopy", SqlDbType.Int) With {.Value = IsCopy})

				' TODO - replace with dapper
				For Each objRow As DataRow In rstAccessInfo.Rows
					objAccess.Add(New GroupAccess() With {
									.Access = objRow("access").ToString,
									.Name = objRow("name").ToString})
				Next

			Catch ex As Exception
				Throw

			End Try

			Return objAccess

		End Function

		' Old style update of the column selection stuff
		' could be dapperised, but the rest of our stored procs need updating too as everything has different column names and the IDs are not currently returned.
		Private Function ReportChildTablesAsString(objSortColumns As Collection(Of ReportChildTables)) As String

			Dim sOrderString As String = ""

			For Each objItem In objSortColumns
				sOrderString += String.Format("{0}||{1}||{2}||{3}**" _
													, objItem.TableID, objItem.FilterID, objItem.OrderID, objItem.Records)
			Next

			Return sOrderString

		End Function


		' Old style update of the column selection stuff
		' could be dapperised, but the rest of our stored procs need updating too as everything has different column names and the IDs are not currently returned.
		Private Function ReportColumnsAsString(objColumns As Collection(Of ReportColumnItem), objSortColumns As Collection(Of ReportSortItem)) As String

			Dim sColumns As String = ""
			Dim sOrderString As String

			Dim iCount As Integer = 1
			For Each objItem In objColumns

				' this could be improve with some linq or whatever! No panic because the whole function could be tidied up
				sOrderString = "||0||"
				For Each objSortItem In objSortColumns
					If objSortItem.ID = objItem.id Then
						sOrderString = "||1||" & objSortItem.Order & "||"
					End If
				Next

				sColumns += String.Format("{0}||{1}||{2}||{3}||{4}||{5}**" _
																	, iCount, IIf(objItem.IsExpression, "E", "C"), objItem.id, objItem.Size, objItem.Decimals, sOrderString)
				iCount += 1
			Next

			Return sColumns

		End Function

		' Old style update of the utility access grid
		' could be dapperised, but the rest of our stored procs need updating too as everything has different column names and the IDs are not currently returned.
		Private Function UtilityAccessAsString(objAccess As Collection(Of GroupAccess)) As String

			Dim sAccess As String = ""
			For Each group In objAccess
				sAccess += group.Name + Chr(9) + group.Access + Chr(9)
			Next

			Return sAccess

		End Function

		' TODO - Sometimes we may need to hide dependant objects
		Private Function JobsToHideAsString(objJobs As Collection(Of Integer)) As String
			Return ""
		End Function

		Public Function GetTables() As List(Of ReportTableItem)

			Dim objSessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			Dim objItems As New List(Of ReportTableItem)

			For Each objTable In objSessionInfo.Tables.OrderBy(Function(n) n.Name)
				Dim objItem As New ReportTableItem() With {.id = objTable.ID, .Name = objTable.Name}
				objItems.Add(objItem)
			Next

			Return objItems

		End Function

		Public Function GetColumnsForTable(id As Integer) As List(Of ReportColumnItem)

			Dim objSessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			Dim objReturnData As New List(Of ReportColumnItem)

			Try

				Dim objToAdd As New ReportColumnItem With {
							.id = 0,
							.Name = "None",
							.DataType = SQLDataType.sqlUnknown,
							.Size = 0,
							.Decimals = 0}
				objReturnData.Add(objToAdd)

				For Each objColumn In objSessionInfo.Columns.OrderBy(Function(n) n.Name)
					If objColumn.TableID = id And objColumn.Name <> "ID" Then

						objToAdd = New ReportColumnItem With {
							.id = objColumn.ID,
							.Name = objColumn.Name,
							.Heading = objColumn.Name,
							.DataType = objColumn.DataType,
							.Size = objColumn.Size,
							.Decimals = objColumn.Decimals}

						objReturnData.Add(objToAdd)
					End If

				Next

			Catch ex As Exception
				Throw

			End Try

			Return objReturnData

		End Function


		' can be done with dapper?
		Private Sub PopulateDefintion(outputModel As ReportBaseModel, data As DataTable)

			Try

				If data.Rows.Count = 1 Then

					Dim row As DataRow = data.Rows(0)

					outputModel.BaseTableID = CInt(row("BaseTableID"))

					outputModel.Name = row("name").ToString
					outputModel.Description = row("description").ToString
					outputModel.Owner = row("owner").ToString

					outputModel.FilterID = CInt(row("FilterID"))
					outputModel.FilterName = row("filtername").ToString
					outputModel.PicklistID = CInt(row("PicklistID"))
					outputModel.PicklistName = row("picklistname").ToString

					outputModel.SelectionType = CType(row("SelectionType"), RecordSelectionType)

					If data.Columns.Contains("PrintFilterHeader") Then
						outputModel.DisplayTitleInReportHeader = CBool(row("PrintFilterHeader"))
					End If

				End If

			Catch ex As Exception
				Throw

			End Try

		End Sub

		Private Sub PopulateSortOrder(outputModel As ReportBaseModel, data As DataTable)

			Dim objSort As ReportSortItem

			For Each objRow As DataRow In data.Rows
				objSort = New ReportSortItem
				objSort.TableID = CInt(objRow("tableid"))
				objSort.ID = CInt(objRow("Id"))
				objSort.Name = objRow("name").ToString
				objSort.Order = objRow("order").ToString
				objSort.Sequence = CInt(objRow("sequence"))

				If data.Columns.Contains("PageOnChange") Then
					objSort.PageOnChange = CBool(objRow("PageOnChange"))
				End If

				If data.Columns.Contains("ValueOnChange") Then
					objSort.PageOnChange = CBool(objRow("ValueOnChange"))
				End If

				If data.Columns.Contains("BreakOnChange") Then
					objSort.PageOnChange = CBool(objRow("BreakOnChange"))
				End If

				If data.Columns.Contains("SuppressRepeated") Then
					objSort.PageOnChange = CBool(objRow("SuppressRepeated"))
				End If

				outputModel.SortOrderColumns.Add(objSort)
			Next

		End Sub



		' can be done with dapper?
		Private Sub PopulateOutput(outputModel As ReportOutputModel, data As DataTable)

			Try

				If data.Rows.Count = 1 Then

					Dim row As DataRow = data.Rows(0)

					outputModel.IsPreview = CBool(row("IsPreview"))
					outputModel.Format = CType(row("Format"), OutputFormats)
					outputModel.ToScreen = CBool(row("ToScreen"))
					outputModel.ToPrinter = CBool(row("ToPrinter"))
					outputModel.PrinterName = row("PrinterName").ToString()
					outputModel.SaveToFile = CBool(row("SaveToFile"))
					outputModel.Filename = row("FileName").ToString
					outputModel.SaveExisting = CType(row("SaveExisting"), ExistingFile)
					outputModel.SendToEmail = CBool(row("SendToEmail"))
					outputModel.EmailGroupID = CInt(row("EmailGroupID"))
					outputModel.EmailGroupName = row("EmailGroupName").ToString
					outputModel.EmailSubject = row("EmailSubject").ToString()
					outputModel.EmailAttachmentName = row("EmailAttachmentName").ToString()

				End If

			Catch ex As Exception
				Throw

			End Try

		End Sub


		'Public Function getModel(id As Integer) As ReportBaseModel

		'	Return _reports.Item(id)

		'End Function





	End Class
End Namespace