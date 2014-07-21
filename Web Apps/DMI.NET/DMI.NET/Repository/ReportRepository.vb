Option Strict On
Option Explicit On

Imports HR.Intranet.Server
Imports DMI.NET.Models
Imports System.Data.SqlClient
Imports HR.Intranet.Server.Metadata
Imports System.Collections.ObjectModel
Imports DMI.NET.Classes
Imports HR.Intranet.Server.Enums
Imports DMI.NET.Enums
Imports DMI.NET.ViewModels
Imports DMI.NET.ViewModels.Reports
Imports DMI.NET.Code.Extensions

Namespace Repository
	Public Class ReportRepository

		Private _customreports As New Collection(Of CustomReportModel)
		Private _crosstabs As New Collection(Of CrossTabModel)
		Private _calendarreports As New Collection(Of CalendarReportModel)
		Private _mailmerges As New Collection(Of MailMergeModel)

		Private _objSessionInfo As SessionInfo
		Private _objDataAccess As clsDataAccess
		Private _username As String
		Private _defaultBaseTableID As Integer

		Public Sub New()

			MyBase.New()
			_objSessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			_objDataAccess = CType(HttpContext.Current.Session("DatabaseAccess"), clsDataAccess)
			_username = HttpContext.Current.Session("username").ToString
			_defaultBaseTableID = CInt(HttpContext.Current.Session("Personnel_EmpTableID"))

		End Sub

		Public Function LoadCustomReport(ID As Integer, action As UtilityActionType) As CustomReportModel

			Dim objModel As New CustomReportModel
			Try

				objModel.Attach(_objSessionInfo)

				If action = UtilityActionType.New Then
					objModel.BaseTableID = _defaultBaseTableID
					objModel.Owner = _username
				Else

					objModel.ID = ID

					Dim dsDefinition As DataSet = _objDataAccess.GetDataSet("spASRIntGetCustomReportDefinition" _
					, New SqlParameter("piReportID", SqlDbType.Int) With {.Value = objModel.ID} _
					, New SqlParameter("psCurrentUser", SqlDbType.VarChar, 255) With {.Value = _username} _
					, New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = action})

					PopulateDefintion(objModel, dsDefinition.Tables(0))

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

					' Selected columns
					PopulateColumns(objModel, dsDefinition.Tables(1))

					' Repetition
					For Each objRow As DataRow In dsDefinition.Tables(3).Rows
						Dim objRepeatItem As New ReportRepetition() With {
								.ID = CInt(objRow("id")),
								.Name = objRow("Name").ToString,
								.IsExpression = False,
								.IsRepeated = CBool(objRow("IsRepeated"))}
						objModel.Repetition.Add(objRepeatItem)
					Next

					'.IsExpression = CBool(objRow("IsExpression")), // TODO

					PopulateSortOrder(objModel, dsDefinition.Tables(2))
					PopulateOutput(objModel.Output, dsDefinition.Tables(0))

					' Populate the child tables
					For Each objRow As DataRow In dsDefinition.Tables(4).Rows
						objModel.ChildTables.Add(New ChildTableViewModel() With {
										.ReportID = objModel.ID,
										.TableName = objRow("tablename").ToString,
										.FilterName = objRow("filtername").ToString,
										.OrderName = objRow("ordername").ToString,
										.TableID = CInt(objRow("tableid")),
										.FilterID = CInt(objRow("filterid")),
										.OrderID = CInt(objRow("orderid")),
										.Records = CInt(objRow("Records"))})
					Next

				End If

				objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCustomReport, objModel.ID, action)
				objModel.IsReadOnly = (action = UtilityActionType.View)

				_customreports.Remove(objModel.ID)
				_customreports.Add(objModel)

			Catch ex As Exception
				Throw

			End Try

			Return objModel

		End Function

		Public Function LoadMailMerge(ID As Integer, action As UtilityActionType) As MailMergeModel

			Dim objModel As New MailMergeModel

			Try

				objModel.Attach(_objSessionInfo)

				If action = UtilityActionType.New Then
					objModel.BaseTableID = _defaultBaseTableID
					objModel.Owner = _username
				Else

					objModel.ID = ID

					Dim dsDefinition = _objDataAccess.GetDataSet("spASRIntGetMailMergeDefinition" _
						, New SqlParameter("@piReportID", SqlDbType.Int) With {.Value = objModel.ID} _
						, New SqlParameter("@psCurrentUser", SqlDbType.VarChar, 255) With {.Value = _username} _
						, New SqlParameter("@psAction", SqlDbType.VarChar, 255) With {.Value = action})

					PopulateDefintion(objModel, dsDefinition.Tables(0))

					' Selected columns and expressions
					PopulateColumns(objModel, dsDefinition.Tables(1))

					' Orders
					PopulateSortOrder(objModel, dsDefinition.Tables(2))

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

				End If

				objModel.GroupAccess = GetUtilityAccess(UtilityType.utlMailMerge, ID, action)
				objModel.IsReadOnly = (action = UtilityActionType.View)

				_mailmerges.Remove(objModel.ID)
				_mailmerges.Add(objModel)

			Catch ex As Exception
				Throw

			End Try

			Return objModel

		End Function

		Public Function LoadCrossTab(ID As Integer, action As UtilityActionType) As CrossTabModel

			Dim objModel As New CrossTabModel

			Try
				objModel.Attach(_objSessionInfo)

				If action = UtilityActionType.New Then
					objModel.BaseTableID = _defaultBaseTableID
					objModel.Owner = _username
				Else

					objModel.ID = ID

					Dim dtDefinition = _objDataAccess.GetFromSP("spASRIntGetCrossTabDefinition", _
							New SqlParameter("piReportID", SqlDbType.Int) With {.Value = objModel.ID}, _
							New SqlParameter("psCurrentUser", SqlDbType.VarChar, 255) With {.Value = _username}, _
							New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = action})

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

					' Output Tab
					PopulateOutput(objModel.Output, dtDefinition)

				End If

				objModel.AvailableColumns = GetColumnsForTable(objModel.BaseTableID)
				objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCrossTab, ID, action)
				objModel.IsReadOnly = (action = UtilityActionType.View)

				_crosstabs.Remove(objModel.ID)
				_crosstabs.Add(objModel)

			Catch ex As Exception
				Throw

			End Try

			Return objModel

		End Function

		Public Function LoadCalendarReport(ID As Integer, action As UtilityActionType) As CalendarReportModel

			Dim objModel As New CalendarReportModel
			Dim objEvent As CalendarEventDetailViewModel

			Try
				objModel.Attach(_objSessionInfo)

				If action = UtilityActionType.New Then
					objModel.BaseTableID = _defaultBaseTableID
					objModel.Owner = _username
				Else

					objModel.ID = ID

					Dim dsDefinition = _objDataAccess.GetDataSet("spASRIntGetCalendarReportDefinition", _
							New SqlParameter("@piCalendarReportID", SqlDbType.Int) With {.Value = objModel.ID}, _
							New SqlParameter("psCurrentUser", SqlDbType.VarChar, 255) With {.Value = _username}, _
							New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = action})

					PopulateDefintion(objModel, dsDefinition.Tables(0))
					If dsDefinition.Tables(0).Rows.Count = 1 Then

						Dim row As DataRow = dsDefinition.Tables(0).Rows(0)

						objModel.Description1ID = CInt(row("Description1Id"))
						objModel.Description2ID = CInt(row("Description2Id"))
						objModel.Description3ID = CInt(row("Description3Id"))
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
						objEvent = New CalendarEventDetailViewModel

						objEvent.ID = CInt(objRow("ID"))
						objEvent.Name = objRow("Name").ToString
						objEvent.EventKey = objRow("EventKey").ToString
						objEvent.ReportID = objModel.ID
						objEvent.TableID = CInt(objRow("TableID"))
						objEvent.FilterID = CInt(objRow("FilterID"))
						objEvent.FilterName = objRow("FilterName").ToString
						objEvent.EventStartDateID = CInt(objRow("EventStartDateID"))
						objEvent.EventStartSessionID = CInt(objRow("EventStartSessionID"))
						objEvent.EventStartSessionName = objRow("EventStartSessionName").ToString
						objEvent.EventEndType = CType(objRow("EventEndType"), CalendarEventEndType)
						objEvent.EventEndDateID = CInt(objRow("EventEndDateID"))
						objEvent.EventEndDateName = objRow("EventEndDateName").ToString
						objEvent.EventEndSessionID = CInt(objRow("EventEndSessionID"))
						objEvent.EventDurationName = objRow("EventDurationName").ToString
						objEvent.EventEndSessionName = objRow("EventEndSessionName").ToString
						objEvent.EventDurationID = CInt(objRow("EventDurationID"))
						objEvent.LegendType = CType(objRow("LegendType"), CalendarLegendType)
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

				End If

				objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCalendarReport, ID, action)
				objModel.IsReadOnly = (action = UtilityActionType.View)

				_calendarreports.Remove(objModel.ID)
				_calendarreports.Add(objModel)

			Catch ex As Exception
				Throw

			End Try

			Return objModel

		End Function

		Public Function SaveReportDefinition(objModel As MailMergeModel) As Boolean

			Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = objModel.ID}

			' TODO old access stuff - needs updating
			Dim psJobsToHide As String = ""	' Request.Form("txtSend_jobsToHide")
			Dim psJobsToHideGroups As String = ""	' Request.Form("txtSend_jobsToHideGroups")}

			Try

				Dim sAccess = UtilityAccessAsString(objModel.GroupAccess)
				Dim sColumns = MailMergeColumnsAsString(objModel.Columns, objModel.SortOrders)

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

			Catch ex As Exception
				Throw

			End Try

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

				_crosstabs.Remove(objModel.ID)

			Catch
				Throw

			End Try

			Return True
		End Function

		Public Function SaveReportDefinition(objModel As CustomReportModel) As Boolean

			Try

				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = objModel.ID}

				Dim sAccess As String = UtilityAccessAsString(objModel.GroupAccess)
				Dim sJobsToHide = JobsToHideAsString(objModel.JobsToHide)
				Dim sJobsToHideGroups As String = "" ' TODO?
				Dim sColumns = CustomReportColumnsAsString(objModel.Columns, objModel.SortOrders)
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

				_customreports.Remove(objModel.ID)

			Catch ex As Exception
				Throw

			End Try

			Return True

		End Function

		Public Function SaveReportDefinition(objModel As CalendarReportModel) As Boolean

			Try

				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = objModel.ID}

				Dim sAccess = UtilityAccessAsString(objModel.GroupAccess)
				Dim sJobsToHide = JobsToHideAsString(objModel.JobsToHide)
				Dim sJobsToHideGroups As String = "" ' TODO?
				Dim sEvents As String = EventsAsString(objModel.Events)

				Dim sReportOrder As String = SortOrderAsString(objModel.SortOrders)
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
					New SqlParameter("piDescription1", SqlDbType.Int) With {.Value = objModel.Description1ID}, _
					New SqlParameter("piDescription2", SqlDbType.Int) With {.Value = objModel.Description2ID}, _
					New SqlParameter("piDescriptionExpr", SqlDbType.Int) With {.Value = objModel.Description3ID}, _
					New SqlParameter("piRegion", SqlDbType.Int) With {.Value = objModel.RegionID}, _
					New SqlParameter("pfGroupByDesc", SqlDbType.Bit) With {.Value = objModel.GroupByDescription}, _
					New SqlParameter("psDescSeparator", SqlDbType.VarChar, 100) With {.Value = objModel.Separator}, _
					New SqlParameter("piStartType", SqlDbType.Int) With {.Value = objModel.StartType}, _
					New SqlParameter("psFixedStart", SqlDbType.VarChar) With {.Value = If(objModel.StartFixedDate.HasValue, objModel.StartFixedDate.Value.ToString("yyyy-MM-dd hh:mm:ss"), "")}, _
					New SqlParameter("piStartFrequency", SqlDbType.Int) With {.Value = objModel.StartOffsetPeriod}, _
					New SqlParameter("piStartPeriod", SqlDbType.Int) With {.Value = objModel.StartOffsetPeriod}, _
					New SqlParameter("piStartDateExpr", SqlDbType.Int) With {.Value = objModel.StartCustomId}, _
					New SqlParameter("piEndType", SqlDbType.Int) With {.Value = objModel.EndType}, _
					New SqlParameter("psFixedEnd", SqlDbType.VarChar) With {.Value = If(objModel.EndFixedDate.HasValue, objModel.EndFixedDate.Value.ToString("yyyy-MM-dd hh:mm:ss"), "")}, _
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
					New SqlParameter("pfOutputEmailAddr", SqlDbType.Int) With {.Value = objModel.Output.EmailGroupID}, _
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

				_calendarreports.Remove(objModel.ID)

			Catch ex As Exception
				Throw

			End Try

			Return True
		End Function

		Private Function GetUtilityAccess(utilType As UtilityType, ID As Integer, action As UtilityActionType) As Collection(Of GroupAccess)

			Dim objAccess As New Collection(Of GroupAccess)
			Dim isCopy = (action = UtilityActionType.Copy)

			Try

				Dim rstAccessInfo As DataTable = _objDataAccess.GetDataTable("spASRIntGetUtilityAccessRecords", CommandType.StoredProcedure _
					, New SqlParameter("piUtilityType", SqlDbType.Int) With {.Value = CInt(utilType)} _
					, New SqlParameter("piID", SqlDbType.Int) With {.Value = ID} _
					, New SqlParameter("piFromCopy", SqlDbType.Int) With {.Value = isCopy})

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
		Private Function ReportChildTablesAsString(objSortColumns As List(Of ChildTableViewModel)) As String

			Dim sOrderString As String = ""

			For Each objItem In objSortColumns
				sOrderString += String.Format("{0}||{1}||{2}||{3}**" _
													, objItem.TableID, objItem.FilterID, objItem.OrderID, objItem.Records)
			Next

			Return sOrderString

		End Function


		' Old style update of the column selection stuff
		' could be dapperised, but the rest of our stored procs need updating too as everything has different column names and the IDs are not currently returned.
		Private Function MailMergeColumnsAsString(objColumns As IEnumerable(Of ReportColumnItem), objSortColumns As Collection(Of SortOrderViewModel)) As String

			Dim sColumns As String = ""
			Dim sOrderString As String
			Dim iSortSequence As Integer = 1

			Dim iCount As Integer = 1
			For Each objItem In objColumns

				' this could be improve with some linq or whatever! No panic because the whole function could be tidied up
				sOrderString = "0||"
				For Each objSortItem In objSortColumns
					If objSortItem.ColumnID = objItem.ID Then
						sOrderString = String.Format("{0}||{1}||", iSortSequence, IIf(objSortItem.Order = OrderType.Ascending, "Asc", "Desc").ToString)
						iSortSequence += 1
					End If
				Next

				sColumns += String.Format("{0}||{1}||{2}||{3}||{4}||{5}||{6}**" _
													, iCount, IIf(objItem.IsExpression, "E", "C"), objItem.ID, objItem.Size, objItem.Decimals, objItem.IsNumeric, sOrderString)

				iCount += 1
			Next

			Return sColumns

		End Function

		' Old style update of the column selection stuff
		' could be dapperised, but the rest of our stored procs need updating too as everything has different column names and the IDs are not currently returned.
		Private Function CustomReportColumnsAsString(objColumns As IEnumerable(Of ReportColumnItem), objSortColumns As Collection(Of SortOrderViewModel)) As String

			Dim sColumns As String = ""
			Dim sOrderString As String

			Dim iCount As Integer = 1
			Dim iSortSequence As Integer
			For Each objItem In objColumns

				' this could be improve with some linq or whatever! No panic because the whole function could be tidied up
				sOrderString = "||0||"
				iSortSequence = 1
				For Each objSortItem In objSortColumns
					If objSortItem.ColumnID = objItem.ID Then
						sOrderString = String.Format("{0}||{1}||{2}||{3}||{4}||{5}" _
							, iSortSequence, IIf(objSortItem.Order = OrderType.Ascending, "Asc", "Desc").ToString _
							, If(objSortItem.BreakOnChange, 1, 0), If(objSortItem.PageOnChange, 1, 0) _
							, If(objSortItem.ValueOnChange, 1, 0), If(objSortItem.SuppressRepeated, 1, 0))

						iSortSequence += 1
						Exit For
					End If
				Next

				sColumns += String.Format("{0}||{1}||{2}||{3}||{4}||{5}||{6}||{7}||{8}||{9}||{10}||{11}||{12}||{13}**" _
													, iCount, IIf(objItem.IsExpression, "E", "C"), objItem.ID, objItem.Heading, objItem.Size, objItem.Decimals _
													, If(objItem.IsNumeric, 1, 0), If(objItem.IsAverage, 1, 0), If(objItem.IsCount, 1, 0) _
													, If(objItem.IsTotal, 1, 0), If(objItem.IsHidden, 1, 0), If(objItem.IsGroupWithNext, 1, 0) _
													, sOrderString, "0")
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

		' Old style update of the events selection stuff
		Public Function EventsAsString(objEvents As Collection(Of CalendarEventDetailViewModel)) As String

			Dim sEvents As String = ""
			Dim sLegend As String

			For Each objItem In objEvents
				If objItem.LegendType = CalendarLegendType.LookupTable Then
					sLegend = String.Format("1||||{0}||{1}||{2}||{3}" _
																			, objItem.LegendLookupTableID, objItem.LegendLookupColumnID, objItem.LegendLookupCodeID, objItem.LegendEventColumnID)
				Else
					sLegend = String.Format("0||{0}||||||||", objItem.LegendCharacter)
				End If

				sEvents += String.Format("{0}||{1}||{2}||{3}||{4}||{5}||{6}||{7}||{8}||{9}||{10}||{11}||**" _
																 , objItem.EventKey, objItem.Name, objItem.TableID, objItem.FilterID _
																 , objItem.EventStartDateID, objItem.EventStartSessionID, objItem.EventEndDateID, objItem.EventEndSessionID _
																 , objItem.EventDurationID, sLegend, objItem.EventDesc1ColumnID, objItem.EventDesc2ColumnID)
			Next

			Return sEvents

		End Function

		' Old style update of the events selection stuff
		Public Function SortOrderAsString(objSortOrders As Collection(Of SortOrderViewModel)) As String

			Dim sOrders As String = ""
			Dim iCount As Integer = 1
			For Each objItem In objSortOrders
				iCount += 1
				sOrders += String.Format("{0}||{1}||{2}||**", objItem.ID, iCount, IIf(objItem.Order = OrderType.Ascending, "Asc", "Desc").ToString)
			Next

			Return sOrders

		End Function


		Public Function GetTables() As List(Of ReportTableItem)

			Dim objSessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			Dim objItems As New List(Of ReportTableItem)

			For Each objTable In objSessionInfo.Tables.OrderBy(Function(n) n.Name)
				Dim objItem As New ReportTableItem() With {.id = objTable.ID, .Name = objTable.Name}
				objItems.Add(objItem)
			Next

			Return objItems.OrderBy(Function(m) m.Name).ToList

		End Function

		Public Function GetChildTables(BaseTableID As Integer, IncludeSelf As Boolean) As List(Of ReportTableItem)

			Dim objSessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			Dim objItems As New List(Of ReportTableItem)
			Dim objTable As Table
			Dim objItem As ReportTableItem

			For Each objRelation In objSessionInfo.Relations.Where(Function(n) n.ParentID = BaseTableID)
				objTable = objSessionInfo.Tables.Where(Function(m) m.ID = objRelation.ChildID).FirstOrDefault
				objItem = New ReportTableItem() With {.id = objRelation.ChildID, .Name = objTable.Name}
				objItems.Add(objItem)
			Next

			If IncludeSelf Then
				objTable = objSessionInfo.Tables.Where(Function(m) m.ID = BaseTableID).FirstOrDefault
				objItem = New ReportTableItem() With {.id = objTable.ID, .Name = objTable.Name}
				objItems.Add(objItem)
			End If

			Return objItems.OrderBy(Function(m) m.Name).ToList

		End Function

		Public Function GetColumnsForTable(id As Integer) As List(Of ReportColumnItem)

			Dim objSessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			Dim objReturnData As New List(Of ReportColumnItem)

			Dim objToAdd As New ReportColumnItem

			Try

				For Each objColumn In objSessionInfo.Columns.Where(Function(m) m.TableID = id And m.IsVisible).OrderBy(Function(n) n.Name)

					objToAdd = New ReportColumnItem With {
						.ID = objColumn.ID,
						.Name = objColumn.Name,
						.IsExpression = False,
						.Heading = objColumn.Name,
						.DataType = objColumn.DataType,
						.Size = objColumn.Size,
						.Decimals = objColumn.Decimals}

					objReturnData.Add(objToAdd)

				Next

			Catch ex As Exception
				Throw

			End Try

			Return objReturnData

		End Function

		' Improve with Dapper?
		Public Function GetCalculationsForTable(tableId As Integer) As List(Of ReportColumnItem)

			Dim objReturnData As New List(Of ReportColumnItem)

			Try

				Dim dtDefinition As DataTable = _objDataAccess.GetDataTable("spASRGetCalculationsForTable", CommandType.StoredProcedure _
				, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = tableId})

				For Each objRow As DataRow In dtDefinition.Rows

					Dim objToAdd = New ReportColumnItem With {
						.ID = CInt(objRow("ID")),
						.Name = objRow("Name").ToString,
						.IsExpression = True,
						.Heading = "",
						.DataType = CType(objRow("DataType"), SQLDataType),
						.Size = CInt(objRow("Size")),
						.Decimals = CInt(objRow("Decimals"))}

					objReturnData.Add(objToAdd)

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

			Dim objSort As SortOrderViewModel
			Dim iSequence As Integer = 1

			Try

				For Each objRow As DataRow In data.Rows
					objSort = New SortOrderViewModel

					objSort.ReportID = outputModel.ID
					objSort.ReportType = outputModel.ReportType

					objSort.TableID = CInt(objRow("tableid"))
					objSort.ID = iSequence
					objSort.ColumnID = CInt(objRow("Id"))

					objSort.Name = objRow("name").ToString
					objSort.Order = CType(IIf(objRow("order").ToString.ToUpper = "ASC", OrderType.Ascending, OrderType.Descending), OrderType)
					objSort.Sequence = iSequence

					If data.Columns.Contains("PageOnChange") Then
						objSort.PageOnChange = CBool(objRow("PageOnChange"))
					End If

					If data.Columns.Contains("ValueOnChange") Then
						objSort.ValueOnChange = CBool(objRow("ValueOnChange"))
					End If

					If data.Columns.Contains("BreakOnChange") Then
						objSort.BreakOnChange = CBool(objRow("BreakOnChange"))
					End If

					If data.Columns.Contains("SuppressRepeated") Then
						objSort.SuppressRepeated = CBool(objRow("SuppressRepeated"))
					End If

					outputModel.SortOrders.Add(objSort)
					iSequence += 1
				Next

			Catch ex As Exception
				Throw

			End Try

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

		' can be done with dapper?
		Private Sub PopulateColumns(outputModel As ReportBaseModel, data As DataTable)

			Try

				outputModel.Columns = New List(Of ReportColumnItem)

				For Each objRow As DataRow In data.Rows
					Dim objItem As New ReportColumnItem() With {
						.ReportType = outputModel.ReportType,
						.ReportID = outputModel.ID,
						.Heading = objRow("Heading").ToString,
						.IsExpression = CBool(objRow("IsExpression")),
						.ID = CInt(objRow("id")),
						.Name = objRow("Name").ToString,
						.DataType = CType(objRow("DataType"), SQLDataType),
						.Sequence = CInt(objRow("Sequence")),
						.Size = CInt(objRow("Size")),
						.Decimals = CInt(objRow("Decimals")),
						.IsAverage = CBool(objRow("IsAverage")),
						.IsCount = CBool(objRow("IsCount")),
						.IsTotal = CBool(objRow("IsTotal")),
						.IsHidden = CBool(objRow("IsHidden")),
						.IsGroupWithNext = CBool(objRow("IsGroupWithNext"))}
					outputModel.Columns.Add(objItem)

				Next

				outputModel.Columns = outputModel.Columns.OrderBy(Function(x) x.Sequence).ToList()

			Catch ex As Exception
				Throw

			End Try

		End Sub

		Public Function RetrieveCustomReport(id As Integer) As CustomReportModel

			Try
				Return _customreports.Where(Function(m) m.ID = id).FirstOrDefault

			Catch ex As Exception
				Return New CustomReportModel

			End Try

		End Function

		Public Function RetrieveCalendarReport(id As Integer) As CalendarReportModel

			Try
				Return _calendarreports.Where(Function(m) m.ID = id).FirstOrDefault

			Catch ex As Exception
				Return New CalendarReportModel

			End Try

		End Function

		Public Function RetrieveParent(reportID As Integer, reportType As UtilityType) As IReport

			Try

				Select Case reportType
					Case UtilityType.utlCalendarReport
						Return _calendarreports.Where(Function(m) m.ID = reportID).FirstOrDefault()

					Case UtilityType.utlMailMerge
						Return _mailmerges.Where(Function(m) m.ID = reportID).FirstOrDefault()

					Case UtilityType.utlCrossTab
						Return _crosstabs.Where(Function(m) m.ID = reportID).FirstOrDefault()

					Case Else
						Return _customreports.Where(Function(m) m.ID = reportID).FirstOrDefault

				End Select

			Catch ex As Exception
				Throw
			End Try

		End Function

		Public Function RetrieveParent(model As IReportDetail) As IReport
			Return RetrieveParent(model.ReportID, model.ReportType)
		End Function



		Function GetAllTablesInReport(reportID As Integer) As List(Of ReportTableItem)

			Dim objReport = RetrieveCustomReport(reportID)
			Dim objItems As New List(Of ReportTableItem)

			If objReport.Parent1.ID > 0 Then
				objItems.Add(New ReportTableItem() With {.id = objReport.Parent1.ID, .Name = objReport.Parent1.Name})
			End If

			If objReport.Parent2.ID > 0 Then
				objItems.Add(New ReportTableItem() With {.id = objReport.Parent2.ID, .Name = objReport.Parent2.Name})
			End If

			For Each objTable In objReport.ChildTables
				objItems.Add(New ReportTableItem() With {.id = objTable.TableID, .Name = objTable.TableName})
			Next

			Dim objBaseTable = _objSessionInfo.Tables.Where(Function(m) m.ID = objReport.BaseTableID).FirstOrDefault
			objItems.Add(New ReportTableItem() With {.id = objBaseTable.ID, .Name = objBaseTable.Name})

			Return objItems

		End Function

		Sub SetBaseTable(objModel As IReport)

			objModel.SessionInfo = _objSessionInfo
			objModel.SetBaseTable(objModel.BaseTableID)

		End Sub

	End Class
End Namespace