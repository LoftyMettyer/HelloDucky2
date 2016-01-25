Option Strict On
Option Explicit On

Imports HR.Intranet.Server
Imports DMI.NET.Models
Imports System.Data.SqlClient
Imports HR.Intranet.Server.Metadata
Imports System.Collections.ObjectModel
Imports DMI.NET.Classes
Imports DMI.NET.ViewModels.Reports
Imports DMI.NET.Code.Extensions
Imports HR.Intranet.Server.Expressions

Namespace Repository
	Public Class ReportRepository

		Private ReadOnly _customreports As New Collection(Of CustomReportModel)
		Private ReadOnly _crosstabs As New Collection(Of CrossTabModel)
		Private ReadOnly _calendarreports As New Collection(Of CalendarReportModel)
		Private ReadOnly _mailmerges As New Collection(Of MailMergeModel)
		Private ReadOnly _talentreports As New Collection(Of TalentReportModel)
		Private ReadOnly _nineboxgrids As New Collection(Of NineBoxGridModel)

		Private ReadOnly Property _username As String
			Get
				Return HttpContext.Current.Session("username").ToString
			End Get
		End Property

		Private ReadOnly Property _objSessionInfo As SessionInfo
			Get
				Return CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			End Get
		End Property

		Private ReadOnly Property _objDataAccess As clsDataAccess
			Get
				Return CType(HttpContext.Current.Session("DatabaseAccess"), clsDataAccess)
			End Get
		End Property

		Public Function LoadCustomReport(ID As Integer, action As UtilityActionType) As CustomReportModel

			Dim objModel As New CustomReportModel
			Try

				objModel.Attach(_objSessionInfo)
				objModel.ActionType = action

				If action = UtilityActionType.New Then
					objModel.BaseTableID = SettingsConfig.Personnel_EmpTableID
					objModel.Owner = _username
					objModel.Output.ToScreen = True
				Else

					objModel.ID = ID

					Dim dsDefinition As DataSet = _objDataAccess.GetDataSet("spASRIntGetCustomReportDefinition" _
					, New SqlParameter("piReportID", SqlDbType.Int) With {.Value = objModel.ID} _
					, New SqlParameter("psCurrentUser", SqlDbType.VarChar, 255) With {.Value = _username} _
					, New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = action})

					objModel.ID = If(action = UtilityActionType.Copy, 0, objModel.ID)

					PopulateDefintion(objModel, dsDefinition.Tables(0))

					If dsDefinition.Tables(0).Rows.Count = 1 Then

						Dim row As DataRow = dsDefinition.Tables(0).Rows(0)

						objModel.IsSummary = CBool(row("IsSummary"))
						objModel.IgnoreZerosForAggregates = CBool(row("IgnoreZerosForAggregates"))

						objModel.Parent1.ID = CInt(row("Parent1ID"))
						objModel.Parent1.SelectionType = CType(row("Parent1SelectionType"), RecordSelectionType)
						objModel.Parent1.Name = row("Parent1Name").ToString
						objModel.Parent1.PicklistID = CInt(row("Parent1PicklistID"))
						objModel.Parent1.PicklistName = row("Parent1PicklistName").ToString
						objModel.Parent1.FilterID = CInt(row("Parent1FilterID"))
						objModel.Parent1.FilterName = row("Parent1FilterName").ToString
						objModel.Parent1ViewAccess = row("Parent1ViewAccess").ToString

						objModel.Parent2.ID = CInt(row("Parent2ID"))
						objModel.Parent2.SelectionType = CType(row("Parent2SelectionType"), RecordSelectionType)
						objModel.Parent2.Name = row("Parent2Name").ToString
						objModel.Parent2.PicklistID = CInt(row("Parent2PicklistID"))
						objModel.Parent2.PicklistName = row("Parent2PicklistName").ToString
						objModel.Parent2.FilterID = CInt(row("Parent2FilterID"))
						objModel.Parent2.FilterName = row("Parent2FilterName").ToString
						objModel.Parent2ViewAccess = row("Parent2ViewAccess").ToString

					End If

					' Selected columns
					PopulateColumns(objModel, dsDefinition.Tables(1))

					PopulateSortOrder(objModel, dsDefinition.Tables(2))
					PopulateOutput(objModel.ReportType, objModel.Output, dsDefinition.Tables(0))

					' Populate the child tables
					Dim iChildIndex As Integer = 1
					For Each objRow As DataRow In dsDefinition.Tables(3).Rows
						objModel.ChildTables.Add(New ChildTableViewModel() With {
										.ID = iChildIndex,
										.ReportID = objModel.ID,
										.TableName = objRow("tablename").ToString,
										.FilterName = objRow("filtername").ToString,
										.FilterViewAccess = objRow("FilterViewAccess").ToString(),
										.OrderName = objRow("ordername").ToString,
										.TableID = CInt(objRow("tableid")),
										.FilterID = CInt(objRow("filterid")),
										.OrderID = CInt(objRow("orderid")),
										.Records = CInt(objRow("Records"))})
						iChildIndex += 1
					Next

				End If

				' if copy the defination then check if group access needs to be hidden
				If objModel.ActionType = UtilityActionType.Copy Then
					If objModel.BaseViewAccess = "HD" Or
						objModel.DefinitionAccessBasedOnSelectedCalculationColumns = "HD" Or
						objModel.Parent1ViewAccess = "HD" Or
						objModel.Parent2ViewAccess = "HD" Or
						(Not objModel.ChildTables.Find(Function(n) n.FilterViewAccess = "HD") Is Nothing) Then
						objModel.IsGroupAccessHiddenWhenCopyTheDefinition = True
					End If
				End If

				objModel.ChildTablesAvailable = CInt(_objSessionInfo.Relations.LongCount(Function(m) m.ParentID = objModel.BaseTableID))
				objModel.GroupAccess = GetUtilityAccess(objModel, action)
				objModel.IsReadOnly = (action = UtilityActionType.View)
				objModel.Owner = If(action = UtilityActionType.Copy, _username, objModel.Owner)
				objModel.CategoryList = GetCategoryList()
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
				objModel.ActionType = action

				If action = UtilityActionType.New Then
					objModel.BaseTableID = SettingsConfig.Personnel_EmpTableID
					objModel.Owner = _username
				Else

					objModel.ID = ID

					Dim dsDefinition = _objDataAccess.GetDataSet("spASRIntGetMailMergeDefinition" _
						, New SqlParameter("@piReportID", SqlDbType.Int) With {.Value = objModel.ID} _
						, New SqlParameter("@psCurrentUser", SqlDbType.VarChar, 255) With {.Value = _username} _
						, New SqlParameter("@psAction", SqlDbType.VarChar, 255) With {.Value = action})

					objModel.ID = If(action = UtilityActionType.Copy, 0, objModel.ID)

					PopulateDefintion(objModel, dsDefinition.Tables(0))

					' Selected columns and expressions
					PopulateColumns(objModel, dsDefinition.Tables(1))

					' Orders
					PopulateSortOrder(objModel, dsDefinition.Tables(2))

					If dsDefinition.Tables(0).Rows.Count = 1 Then

						Dim row As DataRow = dsDefinition.Tables(0).Rows(0)

						objModel.OutputFormat = CType(row("Format"), MailMergeOutputTypes)
						If (objModel.OutputFormat = MailMergeOutputTypes.WordDocument) Then
							objModel.WordDocumentPrinter = row("PrinterName").ToString()
						ElseIf (objModel.OutputFormat = MailMergeOutputTypes.DocumentManagement) Then
							objModel.DocumentManagementPrinter = row("PrinterName").ToString()
						End If
						objModel.DisplayOutputOnScreen = CBool(row("DisplayOutputOnScreen"))
						objModel.SendToPrinter = CBool(row("SendToPrinter"))
						objModel.SaveToFile = CBool(row("SaveToFile"))
						objModel.Filename = row("FileName").ToString
						objModel.EmailGroupID = CInt(row("EmailGroupID"))
						objModel.EmailSubject = row("EmailSubject").ToString()
						objModel.EmailAsAttachment = CBool(row("EmailAsAttachment"))
						objModel.EmailAttachmentName = row("EmailAttachmentName").ToString()

						objModel.SuppressBlankLines = CBool(row("SuppressBlankLines"))
						objModel.PauseBeforeMerge = CBool(row("PauseBeforeMerge"))

						If Not (TypeOf row.Item("UploadTemplate") Is DBNull) Then
							objModel.UploadTemplate = CType(row.Item("UploadTemplate"), Byte())
							objModel.UploadTemplateName = row.Item("UploadTemplateName").ToString
						End If

					End If

				End If

				' if copy the defination then check if group access needs to be hidden
				If objModel.ActionType = UtilityActionType.Copy Then
					If objModel.BaseViewAccess = "HD" Or
						objModel.DefinitionAccessBasedOnSelectedCalculationColumns = "HD" Then
						objModel.IsGroupAccessHiddenWhenCopyTheDefinition = True
					End If
				End If

				objModel.AvailableEmails = GetAvailableEmails(objModel.BaseTableID)

				objModel.GroupAccess = GetUtilityAccess(objModel, action)
				objModel.IsReadOnly = (action = UtilityActionType.View)
				objModel.Owner = If(action = UtilityActionType.Copy, _username, objModel.Owner)
				objModel.CategoryList = GetCategoryList()
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
				objModel.ActionType = action

				If action = UtilityActionType.New Then
					objModel.BaseTableID = SettingsConfig.Personnel_EmpTableID
					objModel.Owner = _username
					objModel.Output.ToScreen = True
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
						objModel.HorizontalStart = CDbl(objRow("HorizontalStart"))
						objModel.HorizontalStop = CDbl(objRow("HorizontalStop"))
						objModel.HorizontalIncrement = CDbl(objRow("HorizontalIncrement"))

						objModel.VerticalID = CInt(objRow("VerticalID"))
						objModel.VerticalDataType = _objSessionInfo.GetColumn(objModel.VerticalID).DataType
						objModel.VerticalStart = CDbl(objRow("VerticalStart"))
						objModel.VerticalStop = CDbl(objRow("VerticalStop"))
						objModel.VerticalIncrement = CDbl(objRow("VerticalIncrement"))

						objModel.PageBreakID = CInt(objRow("PageBreakID"))
						objModel.PageBreakDataType = _objSessionInfo.GetColumn(objModel.PageBreakID).DataType
						objModel.PageBreakStart = CDbl(objRow("PageBreakStart"))
						objModel.PageBreakStop = CDbl(objRow("PageBreakStop"))
						objModel.PageBreakIncrement = CDbl(objRow("PageBreakIncrement"))

						objModel.IntersectionID = CInt(objRow("IntersectionID"))
						objModel.IntersectionType = CType(objRow("IntersectionType"), IntersectionType)
						objModel.PercentageOfType = CBool(objRow("PercentageOfType"))
						objModel.PercentageOfPage = CBool(objRow("PercentageOfPage"))
						objModel.SuppressZeros = CBool(objRow("SuppressZeros"))
						objModel.UseThousandSeparators = CBool(objRow("UseThousandSeparators"))

					End If

					' Output Tab
					PopulateOutput(objModel.ReportType, objModel.Output, dtDefinition)

				End If

				' if copy the defination then check if group access needs to be hidden
				If objModel.ActionType = UtilityActionType.Copy Then
					If objModel.BaseViewAccess = "HD" Then
						objModel.IsGroupAccessHiddenWhenCopyTheDefinition = True
					End If
				End If

				objModel.AvailableColumns = GetColumnsForTable(objModel.BaseTableID)
				objModel.GroupAccess = GetUtilityAccess(objModel, action)
				objModel.IsReadOnly = (action = UtilityActionType.View)
				objModel.ID = If(action = UtilityActionType.Copy, 0, objModel.ID)
				objModel.Owner = If(action = UtilityActionType.Copy, _username, objModel.Owner)
				objModel.CategoryList = GetCategoryList()
				_crosstabs.Remove(objModel.ID)
				_crosstabs.Add(objModel)

			Catch ex As Exception
				Throw

			End Try

			Return objModel

		End Function

		Public Function LoadNineBoxGrid(ID As Integer, action As UtilityActionType) As NineBoxGridModel

			Dim objModel As New NineBoxGridModel

			Try
				objModel.Attach(_objSessionInfo)
				objModel.ActionType = action

				If action = UtilityActionType.New Then
					objModel.BaseTableID = SettingsConfig.Personnel_EmpTableID
					objModel.Owner = _username
					objModel.Output.ToScreen = True
				Else
					objModel.ID = ID

					Dim dtDefinition = _objDataAccess.GetFromSP("spASRIntGetNineBoxGridDefinition", _
							New SqlParameter("piReportID", SqlDbType.Int) With {.Value = objModel.ID}, _
							New SqlParameter("psCurrentUser", SqlDbType.VarChar, 255) With {.Value = _username}, _
							New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = action})

					PopulateDefintion(objModel, dtDefinition)

					If dtDefinition.Rows.Count = 1 Then
						Dim objRow As DataRow = dtDefinition.Rows(0)

						objModel.HorizontalID = CInt(objRow("HorizontalID"))
						objModel.HorizontalDataType = _objSessionInfo.GetColumn(objModel.HorizontalID).DataType
						objModel.HorizontalStart = CDbl(objRow("HorizontalStart"))
						objModel.HorizontalStop = CDbl(objRow("HorizontalStop"))

						objModel.VerticalID = CInt(objRow("VerticalID"))
						objModel.VerticalDataType = _objSessionInfo.GetColumn(objModel.VerticalID).DataType
						objModel.VerticalStart = CDbl(objRow("VerticalStart"))
						objModel.VerticalStop = CDbl(objRow("VerticalStop"))

						objModel.PageBreakID = CInt(objRow("PageBreakID"))
						objModel.PageBreakDataType = _objSessionInfo.GetColumn(objModel.PageBreakID).DataType

						objModel.PercentageOfType = CBool(objRow("PercentageOfType"))
						objModel.PercentageOfPage = CBool(objRow("PercentageOfPage"))
						objModel.SuppressZeros = CBool(objRow("SuppressZeros"))
						objModel.UseThousandSeparators = CBool(objRow("UseThousandSeparators"))

						objModel.XAxisLabel = objRow("XAxisLabel").ToString()
						objModel.XAxisSubLabel1 = objRow("XAxisSubLabel1").ToString()
						objModel.XAxisSubLabel2 = objRow("XAxisSubLabel2").ToString()
						objModel.XAxisSubLabel3 = objRow("XAxisSubLabel3").ToString()
						objModel.YAxisLabel = objRow("YAxisLabel").ToString()
						objModel.YAxisSubLabel1 = objRow("YAxisSubLabel1").ToString()
						objModel.YAxisSubLabel2 = objRow("YAxisSubLabel2").ToString()
						objModel.YAxisSubLabel3 = objRow("YAxisSubLabel3").ToString()
						objModel.Description1 = objRow("Description1").ToString()
						objModel.ColorDesc1 = objRow("ColorDesc1").ToString()
						objModel.Description2 = objRow("Description2").ToString()
						objModel.ColorDesc2 = objRow("ColorDesc2").ToString()
						objModel.Description3 = objRow("Description3").ToString()
						objModel.ColorDesc3 = objRow("ColorDesc3").ToString()
						objModel.Description4 = objRow("Description4").ToString()
						objModel.ColorDesc4 = objRow("ColorDesc4").ToString()
						objModel.Description5 = objRow("Description5").ToString()
						objModel.ColorDesc5 = objRow("ColorDesc5").ToString()
						objModel.Description6 = objRow("Description6").ToString()
						objModel.ColorDesc6 = objRow("ColorDesc6").ToString()
						objModel.Description7 = objRow("Description7").ToString()
						objModel.ColorDesc7 = objRow("ColorDesc7").ToString()
						objModel.Description8 = objRow("Description8").ToString()
						objModel.ColorDesc8 = objRow("ColorDesc8").ToString()
						objModel.Description9 = objRow("Description9").ToString()
						objModel.ColorDesc9 = objRow("ColorDesc9").ToString()
					End If

					' Output Tab
					PopulateOutput(objModel.ReportType, objModel.Output, dtDefinition)

				End If

				' if copy the defination then check if group access needs to be hidden
				If objModel.ActionType = UtilityActionType.Copy Then
					If objModel.BaseViewAccess = "HD" Then
						objModel.IsGroupAccessHiddenWhenCopyTheDefinition = True
					End If
				End If

				objModel.AvailableColumns = GetColumnsForTable(objModel.BaseTableID)
				objModel.GroupAccess = GetUtilityAccess(objModel, action)
				objModel.IsReadOnly = (action = UtilityActionType.View)
				objModel.ID = If(action = UtilityActionType.Copy, 0, objModel.ID)
				objModel.Owner = If(action = UtilityActionType.Copy, _username, objModel.Owner)
				objModel.CategoryList = GetCategoryList()
				_nineboxgrids.Remove(objModel.ID)
				_nineboxgrids.Add(objModel)

			Catch ex As Exception
				Throw

			End Try

			Return objModel

		End Function

		Public Function LoadTalentReport(ID As Integer, action As UtilityActionType) As TalentReportModel


			Dim objModel As New TalentReportModel

			Try

				objModel.Attach(_objSessionInfo)
				objModel.ActionType = action

				If action = UtilityActionType.New Then
					objModel.BaseTableID = SettingsConfig.Post_TableID
					objModel.MatchTableID = SettingsConfig.Personnel_EmpTableID
					objModel.Owner = _username
					objModel.Output.Format = OutputFormats.DataOnly
					objModel.Output.IsPreview = True
					objModel.Output.ToScreen = True
				Else

					objModel.ID = ID

					Dim dsDefinition = _objDataAccess.GetDataSet("spASRIntGetTalentReportDefinition" _
						, New SqlParameter("@piReportID", SqlDbType.Int) With {.Value = objModel.ID} _
						, New SqlParameter("@psCurrentUser", SqlDbType.VarChar, 255) With {.Value = _username} _
						, New SqlParameter("@psAction", SqlDbType.VarChar, 255) With {.Value = action})

					objModel.ID = If(action = UtilityActionType.Copy, 0, objModel.ID)

					PopulateDefintion(objModel, dsDefinition.Tables(0))

					' Selected columns and expressions
					PopulateColumns(objModel, dsDefinition.Tables(1))

					' Orders
					PopulateSortOrder(objModel, dsDefinition.Tables(2))

					If dsDefinition.Tables(0).Rows.Count = 1 Then

						Dim row As DataRow = dsDefinition.Tables(0).Rows(0)

						objModel.BaseChildTableID = CInt(row("BaseChildTableID"))
						objModel.BaseChildColumnID = CInt(row("BaseChildColumnID"))
						objModel.BaseMinimumRatingColumnID = CInt(row("BaseMinimumRatingColumnID"))
						objModel.BasePreferredRatingColumnID = CInt(row("BasePreferredRatingColumnID"))
						objModel.MatchTableID = CInt(row("MatchTableID"))
						objModel.MatchSelectionType = CType(row("MatchSelection"), RecordSelectionType)
						objModel.MatchPicklistID = CInt(row("MatchPicklistID"))
						objModel.MatchFilterID = CInt(row("MatchFilterID"))
						objModel.MatchChildTableID = CInt(row("MatchChildTableID"))
						objModel.MatchChildColumnID = CInt(row("MatchChildColumnID"))
						objModel.MatchChildRatingColumnID = CInt(row("MatchChildRatingColumnID"))
						objModel.MatchAgainstType = CType(row("MatchAgainstType"), MatchAgainstType)

						' Output Tab
						objModel.Output.Format = CType(row("Format"), OutputFormats)
						objModel.Output.ToScreen = CBool(row("DisplayOutputOnScreen"))
						objModel.Output.SendToEmail = CBool(row("SendToEmail"))

						If objModel.Output.Format = OutputFormats.ExcelWorksheet Then
							objModel.Output.IsPreview = False
							If objModel.Output.SendToEmail = False Then
								objModel.Output.ToScreen = True
							End If
						Else
							objModel.Output.IsPreview = True
						End If

						objModel.Output.EmailGroupID = CInt(row("EmailGroupID"))
						objModel.Output.EmailGroupName = row("EmailGroupName").ToString()
						objModel.Output.EmailSubject = row("EmailSubject").ToString()
						objModel.Output.EmailAttachmentName = row("EmailAttachmentName").ToString()

					End If

				End If

				' if copy the defination then check if group access needs to be hidden
				If objModel.ActionType = UtilityActionType.Copy Then
					If objModel.BaseViewAccess = "HD" Or
						objModel.DefinitionAccessBasedOnSelectedCalculationColumns = "HD" Then
						objModel.IsGroupAccessHiddenWhenCopyTheDefinition = True
					End If
				End If

				objModel.GroupAccess = GetUtilityAccess(objModel, action)
				objModel.IsReadOnly = (action = UtilityActionType.View)
				objModel.Owner = If(action = UtilityActionType.Copy, _username, objModel.Owner)
				objModel.CategoryList = GetCategoryList()
				_talentreports.Remove(objModel.ID)
				_talentreports.Add(objModel)

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
				objModel.ActionType = action

				If action = UtilityActionType.New Then
					objModel.BaseTableID = SettingsConfig.Personnel_EmpTableID
					objModel.Owner = _username
					objModel.Output.ToScreen = True
				Else

					objModel.ID = ID

					Dim dsDefinition = _objDataAccess.GetDataSet("spASRIntGetCalendarReportDefinition", _
							New SqlParameter("@piCalendarReportID", SqlDbType.Int) With {.Value = objModel.ID}, _
							New SqlParameter("psCurrentUser", SqlDbType.VarChar, 255) With {.Value = _username}, _
							New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = action})

					objModel.ID = If(action = UtilityActionType.Copy, 0, objModel.ID)

					PopulateDefintion(objModel, dsDefinition.Tables(0))
					If dsDefinition.Tables(0).Rows.Count = 1 Then

						Dim row As DataRow = dsDefinition.Tables(0).Rows(0)

						objModel.Description1ID = CInt(row("Description1Id"))
						objModel.Description2ID = CInt(row("Description2Id"))
						objModel.Description3ID = CInt(row("Description3Id"))
						objModel.Description3Name = row("Description3Name").ToString
						objModel.Description3ViewAccess = row("Description3ViewAccess").ToString

						objModel.RegionID = CInt(row("RegionID"))
						objModel.GroupByDescription = CBool(row("GroupByDescription"))

						Select Case row("Separator").ToString
							Case ""
								objModel.Separator = "None"
							Case " "
								objModel.Separator = "Space"
							Case Else
								objModel.Separator = row("Separator").ToString.Trim
						End Select

						objModel.StartType = CType(row("StartType"), CalendarDataType)
						objModel.StartFixedDate = CDate(row("StartFixedDate"))
						objModel.StartOffset = CInt(row("StartOffset"))
						objModel.StartOffsetPeriod = CType(row("StartOffsetPeriod"), DatePeriod)
						objModel.StartCustomId = CInt(row("StartCustomId"))
						objModel.StartCustomName = row("StartCustomName").ToString
						objModel.StartCustomViewAccess = row("StartCustomViewAccess").ToString

						objModel.EndType = CType(row("EndType"), CalendarDataType)
						objModel.EndFixedDate = CDate(row("EndFixedDate"))
						objModel.EndOffset = CInt(row("EndOffset"))
						objModel.EndOffsetPeriod = CType(row("EndOffsetPeriod"), DatePeriod)
						objModel.EndCustomId = CInt(row("EndCustomId"))
						objModel.EndCustomName = row("EndCustomName").ToString
						objModel.EndCustomViewAccess = row("EndCustomViewAccess").ToString

						objModel.IncludeBankHolidays = CBool(row("IncludeBankHolidays"))
						objModel.WorkingDaysOnly = CBool(row("WorkingDaysOnly"))
						objModel.ShowBankHolidays = CBool(row("ShowBankHolidays"))
						objModel.ShowCaptions = CBool(row("ShowCaptions"))
						objModel.ShowWeekends = CBool(row("ShowWeekends"))
						objModel.StartOnCurrentMonth = CBool(row("StartOnCurrentMonth"))

					End If

					Dim iEventCount As Integer = 0
					For Each objRow As DataRow In dsDefinition.Tables(1).Rows
						objEvent = New CalendarEventDetailViewModel

						iEventCount += 1
						objEvent.ID = CInt(objRow("ID"))
						objEvent.Name = objRow("Name").ToString
						objEvent.EventKey = iEventCount
						objEvent.ReportID = objModel.ID
						objEvent.TableID = CInt(objRow("TableID"))
						objEvent.TableName = objRow("TableName").ToString
						objEvent.FilterID = CInt(objRow("FilterID"))
						objEvent.FilterName = objRow("FilterName").ToString
						objEvent.FilterViewAccess = objRow("FilterViewAccess").ToString()
						objEvent.EventStartDateID = CInt(objRow("EventStartDateID"))
						objEvent.EventStartDateName = objRow("EventStartDateName").ToString
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

						objModel.Events.Add(objEvent)

					Next

					PopulateSortOrder(objModel, dsDefinition.Tables(2))
					PopulateOutput(objModel.ReportType, objModel.Output, dsDefinition.Tables(0))

				End If

				' if copy the defination then check if group access needs to be hidden
				If objModel.ActionType = UtilityActionType.Copy Then
					If objModel.BaseViewAccess = "HD" Or
						objModel.DefinitionAccessBasedOnSelectedCalculationColumns = "HD" Or
						(Not objModel.Events.ToList().Find(Function(n) n.FilterViewAccess = "HD") Is Nothing) Then
						objModel.IsGroupAccessHiddenWhenCopyTheDefinition = True
					End If
				End If

				objModel.GroupAccess = GetUtilityAccess(objModel, action)
				objModel.IsReadOnly = (action = UtilityActionType.View)
				objModel.Owner = If(action = UtilityActionType.Copy, _username, objModel.Owner)
				objModel.CategoryList = GetCategoryList()
				_calendarreports.Remove(objModel.ID)
				_calendarreports.Add(objModel)

			Catch ex As Exception
				Throw

			End Try

			Return objModel

		End Function

		Public Function SaveReportDefinition(objModel As MailMergeModel) As Boolean

			Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = objModel.ID}

			Try

				Dim sAccess = UtilityAccessAsString(objModel.GroupAccess)
				Dim sColumns = MailMergeColumnsAsString(objModel.Columns, objModel.SortOrders)

				'objModel.SendToPrinter
				If (objModel.OutputFormat = MailMergeOutputTypes.WordDocument) Then
					objModel.PrinterName = objModel.WordDocumentPrinter
				ElseIf (objModel.OutputFormat = MailMergeOutputTypes.DocumentManagement) Then
					objModel.PrinterName = objModel.DocumentManagementPrinter
				End If

				objModel.EmailAttachmentName = If(objModel.EmailAttachmentName Is Nothing, "", objModel.EmailAttachmentName)
				objModel.EmailSubject = If(objModel.EmailSubject Is Nothing, "", objModel.EmailSubject)
				objModel.PrinterName = If(objModel.PrinterName Is Nothing, "", objModel.PrinterName)
				objModel.Filename = If(objModel.Filename Is Nothing, "", objModel.Filename)
				objModel.UploadTemplateName = If(objModel.UploadTemplateName Is Nothing, "", objModel.UploadTemplateName)


				_objDataAccess.ExecuteSP("spASRIntSaveMailMerge" _
						, New SqlParameter("@psName", SqlDbType.VarChar, 255) With {.Value = objModel.Name} _
						, New SqlParameter("@psDescription", SqlDbType.VarChar, -1) With {.Value = objModel.Description} _
						, New SqlParameter("@piTableID", SqlDbType.Int) With {.Value = objModel.BaseTableID} _
						, New SqlParameter("@piSelection", SqlDbType.Int) With {.Value = objModel.SelectionType} _
						, New SqlParameter("@piPicklistID", SqlDbType.Int) With {.Value = objModel.PicklistID} _
						, New SqlParameter("@piFilterID", SqlDbType.Int) With {.Value = objModel.FilterID} _
						, New SqlParameter("@UploadTemplate", SqlDbType.Image) With {.Value = objModel.UploadTemplate} _
						, New SqlParameter("@UploadTemplateName", SqlDbType.VarChar, 255) With {.Value = objModel.UploadTemplateName} _
						, New SqlParameter("@piOutputFormat", SqlDbType.Int) With {.Value = objModel.OutputFormat} _
						, New SqlParameter("@pfOutputSave", SqlDbType.Bit) With {.Value = objModel.SaveToFile} _
						, New SqlParameter("@psOutputFilename", SqlDbType.VarChar, -1) With {.Value = objModel.Filename} _
						, New SqlParameter("@piEmailAddrID", SqlDbType.Int) With {.Value = objModel.EmailGroupID} _
						, New SqlParameter("@psEmailSubject", SqlDbType.VarChar, -1) With {.Value = objModel.EmailSubject} _
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
						, New SqlParameter("@psJobsToHide", SqlDbType.VarChar, -1) With {.Value = objModel.Dependencies.JobIDsToHide} _
						, New SqlParameter("@psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = objModel.GroupAccess.HiddenGroups()} _
						, New SqlParameter("@psColumns", SqlDbType.VarChar, -1) With {.Value = sColumns} _
						, New SqlParameter("@psColumns2", SqlDbType.VarChar, -1) With {.Value = ""} _
						, New SqlParameter("@piCategoryID", SqlDbType.Int) With {.Value = objModel.CategoryID} _
				, prmID)

				_mailmerges.Remove(objModel)
				objModel.ID = CInt(prmID.Value)

			Catch ex As Exception
				Throw

			End Try

			Return True

		End Function

		Public Function SaveReportDefinition(objModel As CrossTabModel) As Boolean

			Try

				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = objModel.ID}

				Dim sAccess As String = UtilityAccessAsString(objModel.GroupAccess)

				_objDataAccess.ExecuteSP("spASRIntSaveCrossTab", _
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
								New SqlParameter("psJobsToHide", SqlDbType.VarChar, -1) With {.Value = objModel.Dependencies.JobIDsToHide}, _
								New SqlParameter("psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = objModel.GroupAccess.HiddenGroups()}, _
								New SqlParameter("piCategoryID", SqlDbType.Int) With {.Value = objModel.CategoryID}, _
								prmID)

				_crosstabs.Remove(objModel.ID)
				objModel.ID = CInt(prmID.Value)

			Catch
				Throw

			End Try

			Return True
		End Function

		Public Function SaveReportDefinition(objModel As NineBoxGridModel) As Boolean

			Try

				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = objModel.ID}

				Dim sAccess As String = UtilityAccessAsString(objModel.GroupAccess)

				'Parameters "psOutputPrinterName" and "psOutputFilename" below are not needed for 9-Box Grids and are not present in the Output tab; passing empty values for them

				_objDataAccess.ExecuteSP("spASRIntSaveNineBoxGrid", _
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
								New SqlParameter("piVColID", SqlDbType.Int) With {.Value = objModel.VerticalID}, _
								New SqlParameter("psVStart", SqlDbType.VarChar, 100) With {.Value = objModel.VerticalStart}, _
								New SqlParameter("psVStop", SqlDbType.VarChar, 100) With {.Value = objModel.VerticalStop}, _
								New SqlParameter("piPColID", SqlDbType.Int) With {.Value = objModel.PageBreakID}, _
								New SqlParameter("psPStart", SqlDbType.VarChar, 100) With {.Value = objModel.PageBreakStart}, _
								New SqlParameter("psPStop", SqlDbType.VarChar, 100) With {.Value = objModel.PageBreakStop}, _
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
								New SqlParameter("psOutputPrinterName", SqlDbType.VarChar, -1) With {.Value = ""}, _
								New SqlParameter("pfOutputSave", SqlDbType.Bit) With {.Value = objModel.Output.SaveToFile}, _
								New SqlParameter("piOutputSaveExisting", SqlDbType.Int) With {.Value = objModel.Output.SaveExisting}, _
								New SqlParameter("pfOutputEmail", SqlDbType.Bit) With {.Value = objModel.Output.SendToEmail}, _
								New SqlParameter("piOutputEmailAddr", SqlDbType.Int) With {.Value = objModel.Output.EmailGroupID}, _
								New SqlParameter("psOutputEmailSubject", SqlDbType.VarChar, -1) With {.Value = objModel.Output.EmailSubject}, _
								New SqlParameter("psOutputEmailAttachAs", SqlDbType.VarChar, -1) With {.Value = objModel.Output.EmailAttachmentName}, _
								New SqlParameter("psOutputFilename", SqlDbType.VarChar, -1) With {.Value = ""}, _
								New SqlParameter("psAccess", SqlDbType.VarChar, -1) With {.Value = sAccess}, _
								New SqlParameter("psJobsToHide", SqlDbType.VarChar, -1) With {.Value = objModel.Dependencies.JobIDsToHide}, _
								New SqlParameter("psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = objModel.GroupAccess.HiddenGroups()}, _
								New SqlParameter("XAxisLabel", SqlDbType.VarChar, 255) With {.Value = objModel.XAxisLabel}, _
								New SqlParameter("XAxisSubLabel1", SqlDbType.VarChar, 255) With {.Value = objModel.XAxisSubLabel1}, _
								New SqlParameter("XAxisSubLabel2", SqlDbType.VarChar, 255) With {.Value = objModel.XAxisSubLabel2}, _
								New SqlParameter("XAxisSubLabel3", SqlDbType.VarChar, 255) With {.Value = objModel.XAxisSubLabel3}, _
								New SqlParameter("YAxisLabel", SqlDbType.VarChar, 255) With {.Value = objModel.YAxisLabel}, _
								New SqlParameter("YAxisSubLabel1", SqlDbType.VarChar, 255) With {.Value = objModel.YAxisSubLabel1}, _
								New SqlParameter("YAxisSubLabel2", SqlDbType.VarChar, 255) With {.Value = objModel.YAxisSubLabel2}, _
								New SqlParameter("YAxisSubLabel3", SqlDbType.VarChar, 255) With {.Value = objModel.YAxisSubLabel3}, _
								New SqlParameter("Description1", SqlDbType.VarChar, 255) With {.Value = IIf(objModel.Description1 Is Nothing, "", objModel.Description1)}, _
								New SqlParameter("ColorDesc1", SqlDbType.VarChar, 6) With {.Value = objModel.ColorDesc1}, _
								New SqlParameter("Description2", SqlDbType.VarChar, 255) With {.Value = IIf(objModel.Description2 Is Nothing, "", objModel.Description2)}, _
								New SqlParameter("ColorDesc2", SqlDbType.VarChar, 6) With {.Value = objModel.ColorDesc2}, _
								New SqlParameter("Description3", SqlDbType.VarChar, 255) With {.Value = IIf(objModel.Description3 Is Nothing, "", objModel.Description3)}, _
								New SqlParameter("ColorDesc3", SqlDbType.VarChar, 6) With {.Value = objModel.ColorDesc3}, _
								New SqlParameter("Description4", SqlDbType.VarChar, 255) With {.Value = IIf(objModel.Description4 Is Nothing, "", objModel.Description4)}, _
								New SqlParameter("ColorDesc4", SqlDbType.VarChar, 6) With {.Value = objModel.ColorDesc4}, _
								New SqlParameter("Description5", SqlDbType.VarChar, 255) With {.Value = IIf(objModel.Description5 Is Nothing, "", objModel.Description5)}, _
								New SqlParameter("ColorDesc5", SqlDbType.VarChar, 6) With {.Value = objModel.ColorDesc5}, _
								New SqlParameter("Description6", SqlDbType.VarChar, 255) With {.Value = IIf(objModel.Description6 Is Nothing, "", objModel.Description6)}, _
								New SqlParameter("ColorDesc6", SqlDbType.VarChar, 6) With {.Value = objModel.ColorDesc6}, _
								New SqlParameter("Description7", SqlDbType.VarChar, 255) With {.Value = IIf(objModel.Description7 Is Nothing, "", objModel.Description7)}, _
								New SqlParameter("ColorDesc7", SqlDbType.VarChar, 6) With {.Value = objModel.ColorDesc7}, _
								New SqlParameter("Description8", SqlDbType.VarChar, 255) With {.Value = IIf(objModel.Description8 Is Nothing, "", objModel.Description8)}, _
								New SqlParameter("ColorDesc8", SqlDbType.VarChar, 6) With {.Value = objModel.ColorDesc8}, _
								New SqlParameter("Description9", SqlDbType.VarChar, 255) With {.Value = IIf(objModel.Description9 Is Nothing, "", objModel.Description9)}, _
								New SqlParameter("ColorDesc9", SqlDbType.VarChar, 6) With {.Value = objModel.ColorDesc9}, _
								New SqlParameter("piCategoryID", SqlDbType.Int) With {.Value = objModel.CategoryID}, _
								prmID)

				_nineboxgrids.Remove(objModel.ID)
				objModel.ID = CInt(prmID.Value)

			Catch
				Throw

			End Try

			Return True
		End Function

		Public Function SaveReportDefinition(objModel As CustomReportModel) As Boolean

			Try

				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = objModel.ID}
				For Each columnItem As ReportColumnItem In objModel.Columns
					If columnItem.IsExpression AndAlso columnItem.DataType = ColumnDataType.sqlUnknown Then
						columnItem.ValidateExpressionDataType = True
					End If
				Next

				Dim sAccess As String = UtilityAccessAsString(objModel.GroupAccess)
				Dim sColumns = CustomReportColumnsAsString(objModel.BaseTableID, objModel.Columns, objModel.SortOrders)
				Dim sChildren As String = ReportChildTablesAsString(objModel.ChildTables)

				_objDataAccess.ExecuteSP("spASRIntSaveCustomReport", _
								New SqlParameter("psName", SqlDbType.VarChar, 255) With {.Value = objModel.Name}, _
								New SqlParameter("psDescription", SqlDbType.VarChar, -1) With {.Value = objModel.Description}, _
								New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Value = objModel.BaseTableID}, _
								New SqlParameter("pfAllRecords", SqlDbType.Bit) With {.Value = (objModel.PicklistID = 0 AndAlso objModel.FilterID = 0)}, _
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
								New SqlParameter("pfParent1AllRecords", SqlDbType.Bit) With {.Value = (objModel.Parent1.PicklistID = 0 AndAlso objModel.Parent1.FilterID = 0)}, _
								New SqlParameter("piParent1Picklist", SqlDbType.Int) With {.Value = objModel.Parent1.PicklistID}, _
								New SqlParameter("pfParent2AllRecords", SqlDbType.Bit) With {.Value = (objModel.Parent2.PicklistID = 0 AndAlso objModel.Parent2.FilterID = 0)}, _
								New SqlParameter("piParent2Picklist", SqlDbType.Int) With {.Value = objModel.Parent2.PicklistID}, _
								New SqlParameter("psAccess", SqlDbType.VarChar, -1) With {.Value = sAccess}, _
								New SqlParameter("psJobsToHide", SqlDbType.VarChar, -1) With {.Value = objModel.Dependencies.JobIDsToHide}, _
								New SqlParameter("psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = objModel.GroupAccess.HiddenGroups()}, _
								New SqlParameter("psColumns", SqlDbType.VarChar, -1) With {.Value = sColumns}, _
								New SqlParameter("psColumns2", SqlDbType.VarChar, -1) With {.Value = ""}, _
								New SqlParameter("psChildString", SqlDbType.VarChar, -1) With {.Value = sChildren}, _
								New SqlParameter("piCategoryID", SqlDbType.Int) With {.Value = objModel.CategoryID}, _
								prmID,
								New SqlParameter("pfIgnoreZeros", SqlDbType.Bit) With {.Value = objModel.IgnoreZerosForAggregates})

				_customreports.Remove(objModel.ID)
				objModel.ID = CInt(prmID.Value)

			Catch ex As Exception
				Throw

			End Try

			Return True

		End Function

		Public Function SaveReportDefinition(objModel As CalendarReportModel) As Boolean

			Try

				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = objModel.ID}

				Dim sAccess = UtilityAccessAsString(objModel.GroupAccess)
				Dim sEvents As String = EventsAsString(objModel.Events)

				Dim sSeparator As String
				Select Case objModel.Separator
					Case "None"
						' No space when seperator is None
						sSeparator = ""
					Case "Space"
						'Add single space when seperator is Space.
						sSeparator = " "
					Case ","
						' Add single space when seperator is comma (,)
						sSeparator = objModel.Separator + " "
					Case "."
						' Add two space when seperator is dot (.)
						sSeparator = objModel.Separator + "  "
					Case Else
						' Add right and left both side single space when seperator is not above case
						sSeparator = " " + objModel.Separator + " "
				End Select

				Dim sReportOrder As String = SortOrderAsString(objModel.SortOrders)
				Dim bAllRecords As Boolean

				' Calendar reports don't save the selection type - instead they have a boolean allrecords flag
				If objModel.PicklistID = 0 AndAlso objModel.FilterID = 0 Then bAllRecords = True

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
						New SqlParameter("psDescSeparator", SqlDbType.VarChar, 100) With {.Value = sSeparator}, _
						New SqlParameter("piStartType", SqlDbType.Int) With {.Value = objModel.StartType}, _
						New SqlParameter("psFixedStart", SqlDbType.VarChar) With {.Value = If(objModel.StartFixedDate.HasValue, objModel.StartFixedDate.Value.ToString("yyyy-MM-dd"), "")}, _
						New SqlParameter("piStartFrequency", SqlDbType.Int) With {.Value = objModel.StartOffset}, _
						New SqlParameter("piStartPeriod", SqlDbType.Int) With {.Value = objModel.StartOffsetPeriod}, _
						New SqlParameter("piStartDateExpr", SqlDbType.Int) With {.Value = objModel.StartCustomId}, _
						New SqlParameter("piEndType", SqlDbType.Int) With {.Value = objModel.EndType}, _
						New SqlParameter("psFixedEnd", SqlDbType.VarChar) With {.Value = If(objModel.EndFixedDate.HasValue, objModel.EndFixedDate.Value.ToString("yyyy-MM-dd"), "")}, _
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
						New SqlParameter("psJobsToHide", SqlDbType.VarChar, -1) With {.Value = objModel.Dependencies.JobIDsToHide}, _
						New SqlParameter("psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = objModel.GroupAccess.HiddenGroups()}, _
						New SqlParameter("psEvents", SqlDbType.VarChar, -1) With {.Value = sEvents}, _
						New SqlParameter("psEvents2", SqlDbType.VarChar, -1) With {.Value = ""}, _
						New SqlParameter("psOrderString", SqlDbType.VarChar, -1) With {.Value = sReportOrder}, _
						New SqlParameter("piCategoryID", SqlDbType.Int) With {.Value = objModel.CategoryID}, _
						prmID)

				_calendarreports.Remove(objModel.ID)
				objModel.ID = CInt(prmID.Value)

			Catch ex As Exception
				Throw

			End Try

			Return True
		End Function


		Public Function SaveReportDefinition(objModel As TalentReportModel) As Boolean

			Try

				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = objModel.ID}
				Dim sAccess = UtilityAccessAsString(objModel.GroupAccess)
				Dim sColumns = MatchReportColumnsAsString(objModel.Columns, objModel.SortOrders)

				_objDataAccess.ExecuteSP("spASRIntSaveTalentReport", _
								New SqlParameter("psName", SqlDbType.VarChar, 255) With {.Value = objModel.Name}, _
								New SqlParameter("psDescription", SqlDbType.VarChar, -1) With {.Value = objModel.Description}, _
								New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Value = objModel.BaseTableID}, _
								New SqlParameter("piBaseSelection", SqlDbType.Int) With {.Value = objModel.SelectionType}, _
								New SqlParameter("piBasePicklistID", SqlDbType.Int) With {.Value = objModel.PicklistID}, _
								New SqlParameter("piBaseFilterID", SqlDbType.Int) With {.Value = objModel.FilterID}, _
								New SqlParameter("piBaseChildTableID", SqlDbType.Int) With {.Value = objModel.BaseChildTableID}, _
								New SqlParameter("piBaseChildColumnID", SqlDbType.Int) With {.Value = objModel.BaseChildColumnID}, _
								New SqlParameter("piBaseMinimumRatingColumnID", SqlDbType.Int) With {.Value = objModel.BaseMinimumRatingColumnID}, _
								New SqlParameter("piBasePreferredRatingColumnID", SqlDbType.Int) With {.Value = objModel.BasePreferredRatingColumnID}, _
								New SqlParameter("piMatchTableID", SqlDbType.Int) With {.Value = objModel.MatchTableID}, _
								New SqlParameter("piMatchSelection", SqlDbType.Int) With {.Value = CInt(objModel.MatchSelectionType)}, _
								New SqlParameter("piMatchPicklistID", SqlDbType.Int) With {.Value = objModel.MatchPicklistID}, _
								New SqlParameter("piMatchFilterID", SqlDbType.Int) With {.Value = objModel.MatchFilterID}, _
								New SqlParameter("piMatchChildTableID", SqlDbType.Int) With {.Value = objModel.MatchChildTableID}, _
								New SqlParameter("piMatchChildColumnID", SqlDbType.Int) With {.Value = objModel.MatchChildColumnID}, _
								New SqlParameter("piMatchChildRatingColumnID", SqlDbType.Int) With {.Value = objModel.MatchChildRatingColumnID}, _
								New SqlParameter("piMatchAgainstType", SqlDbType.Int) With {.Value = objModel.MatchAgainstType}, _
								New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = objModel.Owner}, _
								New SqlParameter("psAccess", SqlDbType.VarChar, -1) With {.Value = sAccess}, _
								New SqlParameter("psJobsToHide", SqlDbType.VarChar, -1) With {.Value = objModel.Dependencies.JobIDsToHide}, _
								New SqlParameter("psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = objModel.GroupAccess.HiddenGroups()}, _
								New SqlParameter("psColumns", SqlDbType.VarChar, -1) With {.Value = sColumns}, _
								prmID,
								New SqlParameter("piCategoryID", SqlDbType.Int) With {.Value = objModel.CategoryID}, _
								New SqlParameter("piOutputFormat", SqlDbType.Int) With {.Value = objModel.Output.Format}, _
								New SqlParameter("pfOutputScreen", SqlDbType.Bit) With {.Value = objModel.Output.ToScreen}, _
								New SqlParameter("pfOutputEmail", SqlDbType.Bit) With {.Value = objModel.Output.SendToEmail}, _
								New SqlParameter("piOutputEmailAddr", SqlDbType.Int) With {.Value = objModel.Output.EmailGroupID}, _
								New SqlParameter("psOutputEmailSubject", SqlDbType.VarChar, -1) With {.Value = objModel.Output.EmailSubject}, _
								New SqlParameter("psOutputEmailAttachAs", SqlDbType.VarChar, -1) With {.Value = objModel.Output.EmailAttachmentName})

				_talentreports.Remove(objModel.ID)
				objModel.ID = CInt(prmID.Value)

			Catch ex As Exception
				Throw

			End Try

		End Function


		Private Function GetUtilityAccess(objModel As IReport, action As UtilityActionType) As Collection(Of GroupAccess)

			Dim objAccess As New Collection(Of GroupAccess)
			Dim isCopy = (action = UtilityActionType.Copy)

			Try

				Dim rstAccessInfo As DataTable = _objDataAccess.GetDataTable("spASRIntGetUtilityAccessRecords", CommandType.StoredProcedure _
					, New SqlParameter("piUtilityType", SqlDbType.Int) With {.Value = objModel.ReportType} _
					, New SqlParameter("piID", SqlDbType.Int) With {.Value = objModel.ID} _
					, New SqlParameter("piFromCopy", SqlDbType.Int) With {.Value = isCopy})

				For Each objRow As DataRow In rstAccessInfo.Rows

					Dim bIsOwnerGroup = CBool(objRow("isOwner"))
					Dim bIsReportOwner = (objModel.Owner.ToLower() = _username.ToLower())

					objAccess.Add(New GroupAccess() With {
									.Access = objRow("access").ToString,
									.Name = objRow("name").ToString,
									.IsReadOnly = bIsOwnerGroup OrElse Not bIsReportOwner,
									.DefinitionOwner = objModel.Owner.ToString,
									.LoggedInUser = _username.ToString
									})
				Next

			Catch ex As Exception
				Throw

			End Try

			Return objAccess

		End Function

		' Old style update of the column selection stuff
		' could be dapperised, but the rest of our stored procs need updating too as everything has different column names and the IDs are not currently returned.
		Private Function ReportChildTablesAsString(objSortColumns As IEnumerable(Of ChildTableViewModel)) As String

			Dim sOrderString As String = ""

			For Each objItem In objSortColumns
				sOrderString += String.Format("{0}||{1}||{2}||{3}**" _
													, objItem.TableID, objItem.FilterID, objItem.OrderID, objItem.Records)
			Next

			Return sOrderString

		End Function


		' Old style update of the column selection stuff
		' could be dapperised, but the rest of our stored procs need updating too as everything has different column names and the IDs are not currently returned.
		Private Function MailMergeColumnsAsString(objColumns As IEnumerable(Of ReportColumnItem), objSortColumns As List(Of SortOrderViewModel)) As String

			Dim sColumns As String = ""
			Dim sOrderString As String

			Dim iCount As Integer = 1
			For Each objItem In objColumns

				' this could be improve with some linq or whatever! No panic because the whole function could be tidied up
				sOrderString = "0||"
				For Each objSortItem In objSortColumns.OrderBy((Function(m) m.Sequence))
					If objSortItem.ColumnID = objItem.ID Then
						sOrderString = String.Format("{0}||{1}||", objSortItem.Sequence, IIf(objSortItem.Order = OrderType.Ascending, "Asc", "Desc").ToString)
					End If
				Next

				sColumns += String.Format("{0}||{1}||{2}||{3}||{4}||{5}||{6}**" _
													, iCount, IIf(objItem.IsExpression, "E", "C"), objItem.ID, objItem.Size, objItem.Decimals, objItem.IsNumeric, sOrderString)

				iCount += 1
			Next

			Return sColumns

		End Function

		Private Function MatchReportColumnsAsString(objColumns As IEnumerable(Of ReportColumnItem), objSortColumns As List(Of SortOrderViewModel)) As String

			Dim sColumns As String = ""
			Dim sOrderString As String

			Dim iCount As Integer = 1
			For Each objItem In objColumns

				' this could be improve with some linq or whatever! No panic because the whole function could be tidied up
				sOrderString = "0||0||"
				Dim itemId = objItem.ID
				For Each objSortItem In objSortColumns.Where(Function(m) m.ColumnID = itemId)
					sOrderString = String.Format("{0}||{1}||", objSortItem.Sequence, IIf(objSortItem.Order = OrderType.Ascending, "A", "").ToString)
				Next

				sColumns += String.Format("{0}||{1}||{2}||{3}||{4}||{5}||{6}||{7}**" _
													, iCount, IIf(objItem.IsExpression, "E", "C"), objItem.ID, objItem.Size, objItem.Decimals, objItem.IsNumeric, sOrderString, objItem.Heading)

				iCount += 1
			Next

			Return sColumns

		End Function

		' Old style update of the column selection stuff
		' could be dapperised, but the rest of our stored procs need updating too as everything has different column names and the IDs are not currently returned.
		Private Function CustomReportColumnsAsString(baseTableID As Integer, objColumns As IEnumerable(Of ReportColumnItem), objSortColumns As List(Of SortOrderViewModel)) As String

			Dim sColumns As String = ""
			Dim sOrderString As String
			Dim iRepeated As Integer

			Dim iCount As Integer
			For Each objItem In objColumns

				' this could be improve with some linq or whatever! No panic because the whole function could be tidied up
				sOrderString = "0||0||0||0||0||0"

				For Each objSortItem In objSortColumns.OrderBy((Function(m) m.Sequence))
					If objSortItem.ColumnID = objItem.ID Then
						sOrderString = String.Format("{0}||{1}||{2}||{3}||{4}||{5}" _
							, objSortItem.Sequence, IIf(objSortItem.Order = OrderType.Ascending, "Asc", "Desc").ToString _
							, If(objSortItem.BreakOnChange, 1, 0), If(objSortItem.PageOnChange, 1, 0) _
							, If(objSortItem.ValueOnChange, 1, 0), If(objSortItem.SuppressRepeated, 1, 0))

						Exit For
					End If
				Next

				iRepeated = -1
				If objItem.IsRepeated Then
					iRepeated = 1
				Else
					Dim objColumn = _objSessionInfo.Columns.Where(Function(m) m.ID = objItem.ID And m.TableID = baseTableID).FirstOrDefault
					If objColumn IsNot Nothing Then
						If objColumn.TableID = baseTableID Then
							iRepeated = 0
						End If
					End If
				End If

				sColumns += String.Format("{0}||{1}||{2}||{3}||{4}||{5}||{6}||{7}||{8}||{9}||{10}||{11}||{12}||{13}**" _
													, iCount, IIf(objItem.IsExpression, "E", "C"), objItem.ID, HttpUtility.UrlDecode(objItem.Heading), objItem.Size, objItem.Decimals _
													, If(objItem.IsNumeric, 1, 0), If(objItem.IsAverage, 1, 0), If(objItem.IsCount, 1, 0) _
													, If(objItem.IsTotal, 1, 0), If(objItem.IsHidden, 1, 0), If(objItem.IsGroupWithNext, 1, 0) _
													, sOrderString, iRepeated)
				iCount += 1
			Next

			Return sColumns

		End Function


		' Old style update of the utility access grid
		' could be dapperised, but the rest of our stored procs need updating too as everything has different column names and the IDs are not currently returned.
		Private Function UtilityAccessAsString(objAccess As IEnumerable(Of GroupAccess)) As String

			Dim sAccess As String = ""
			For Each group In objAccess
				sAccess += group.Name + Chr(9) + group.Access + Chr(9)
			Next

			Return sAccess

		End Function

		' Old style update of the events selection stuff
		Public Function EventsAsString(objEvents As Collection(Of CalendarEventDetailViewModel)) As String

			Dim sEvents As String = ""
			Dim sLegend As String
			Dim iEventCount As Integer = 0

			For Each objItem In objEvents
				iEventCount += 1
				If objItem.LegendType = CalendarLegendType.LookupTable Then
					sLegend = String.Format("1||||{0}||{1}||{2}||{3}" _
																			, objItem.LegendLookupTableID, objItem.LegendLookupColumnID, objItem.LegendLookupCodeID, objItem.LegendEventColumnID)
				Else
					sLegend = String.Format("0||{0}||||||||", objItem.LegendCharacter)
				End If

				sEvents += String.Format("EV_{0}||{1}||{2}||{3}||{4}||{5}||{6}||{7}||{8}||{9}||{10}||{11}||**" _
																 , iEventCount, objItem.Name, objItem.TableID, objItem.FilterID _
																 , objItem.EventStartDateID, objItem.EventStartSessionID, objItem.EventEndDateID, objItem.EventEndSessionID _
																 , objItem.EventDurationID, sLegend, objItem.EventDesc1ColumnID, objItem.EventDesc2ColumnID)
			Next

			Return sEvents

		End Function

		' Old style update of the events selection stuff
		Public Function SortOrderAsString(objSortOrders As List(Of SortOrderViewModel)) As String

			Dim sOrders As String = ""
			For Each objItem In objSortOrders.OrderBy(Function(m) m.Sequence)
				sOrders += String.Format("{0}||{1}||{2}||**", objItem.ColumnID, objItem.Sequence, IIf(objItem.Order = OrderType.Ascending, "Asc", "Desc").ToString)
			Next

			Return sOrders

		End Function


		Public Function GetTables(reportType As UtilityType) As List(Of ReportTableItem)

			Dim objSessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			Dim objItems As New List(Of ReportTableItem)

			For Each objTable In objSessionInfo.Tables.OrderBy(Function(n) n.Name)
				Dim objItem As New ReportTableItem() With {.id = objTable.ID, .Name = objTable.Name}

				If reportType = UtilityType.utlCrossTab Then
					If objSessionInfo.Columns.LongCount(Function(m) m.TableID = objTable.ID AndAlso m.IsVisible = True) > 1 Then
						objItems.Add(objItem)
					End If
				ElseIf reportType = UtilityType.utlNineBoxGrid Then	' for 9-box grid, get only numeric columns
					If objSessionInfo.Columns.LongCount(Function(m) m.TableID = objTable.ID AndAlso m.IsNumeric = True AndAlso m.IsVisible = True) > 1 Then
						objItems.Add(objItem)
					End If
				ElseIf reportType = UtilityType.utlCalendarReport Then ' for calendar report fetch only those tables which have date columns
					If objSessionInfo.Columns.LongCount(Function(m) m.TableID = objTable.ID AndAlso m.DataType = ColumnDataType.sqlDate) > 0 Then
						objItems.Add(objItem)
					End If
				Else
					objItems.Add(objItem)
				End If

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

		Public Function GetTablesWithEvents(BaseTableID As Integer) As List(Of ReportTableItem)

			Dim objSessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			Dim objItems As New List(Of ReportTableItem)
			Dim objTable As Table
			Dim objItem As ReportTableItem

			For Each objRelation In objSessionInfo.Relations.Where(Function(n) n.ParentID = BaseTableID)
				objTable = objSessionInfo.Tables.Where(Function(m) m.ID = objRelation.ChildID).FirstOrDefault
				If objSessionInfo.Columns.LongCount(Function(m) m.TableID = objTable.ID AndAlso m.DataType = ColumnDataType.sqlDate) > 0 Then
					objItem = New ReportTableItem() With {.id = objRelation.ChildID, .Name = objTable.Name}
					objItems.Add(objItem)
				End If
			Next

			objTable = objSessionInfo.Tables.Where(Function(m) m.ID = BaseTableID).FirstOrDefault
			objItem = New ReportTableItem() With {.id = objTable.ID, .Name = objTable.Name}
			objItems.Add(objItem)

			Return objItems.OrderBy(Function(m) m.Name).ToList

		End Function

		Public Function GetColumnsForTable(id As Integer) As List(Of ReportColumnItem)

			Dim objSessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			Dim objReturnData As New List(Of ReportColumnItem)

			Dim objToAdd As ReportColumnItem

			Try

				For Each objColumn In objSessionInfo.Columns.Where(Function(m) m.TableID = id And m.IsVisible).OrderBy(Function(n) n.Name)

					objToAdd = New ReportColumnItem With {
						.ID = objColumn.ID,
						.Name = objColumn.Name,
						.IsExpression = False,
						.Heading = objColumn.Name,
						.DataType = objColumn.DataType,
						.Size = objColumn.Size,
						.ColumnSize = objColumn.ColumnSize,
						.Decimals = objColumn.Decimals}

					objReturnData.Add(objToAdd)

				Next

			Catch ex As Exception
				Throw

			End Try

			Return objReturnData

		End Function

		Public Function GetAvailableCharacterLookupsForTable(id As Integer) As List(Of ReportColumnItem)

			Dim objSessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			Dim objReturnData As New List(Of ReportColumnItem)

			Dim objToAdd As ReportColumnItem

			Try

				For Each objColumn In objSessionInfo.Columns.Where(Function(m) m.TableID = id AndAlso m.IsVisible AndAlso m.DataType = ColumnDataType.sqlVarChar).OrderBy(Function(n) n.Name)

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

					' If the hidden calculation is not created by the Loggedin user, then don't add it to "Columns / Calculations Available" grid
					If (objRow("Access").ToString().ToUpper() <> "HD" AndAlso _username.ToUpper() <> objRow("Username").ToString().ToUpper()) Or
						_username.ToUpper() = objRow("Username").ToString().ToUpper() Then

						Dim returnType = CalculationAsColumnType(tableId, CInt(objRow("ID")))

						Dim objToAdd = New ReportColumnItem With {
						.ID = CInt(objRow("ID")),
						.Name = objRow("Name").ToString,
						.IsExpression = True,
						.Heading = "",
						.DataType = returnType,
						.Size = CInt(objRow("Size")),
						.Decimals = CInt(objRow("Decimals")),
						.Access = objRow("Access").ToString}

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
					outputModel.CategoryID = CInt(row("CategoryID"))
					outputModel.Name = row("name").ToString
					outputModel.Description = row("description").ToString
					outputModel.Owner = row("owner").ToString

					outputModel.SelectionType = CType(row("SelectionType"), RecordSelectionType)

					If outputModel.SelectionType <> RecordSelectionType.AllRecords Then
						outputModel.FilterID = CInt(row("FilterID"))
						outputModel.FilterName = row("filtername").ToString
						outputModel.PicklistID = CInt(row("PicklistID"))
						outputModel.PicklistName = row("picklistname").ToString
					End If

					If data.Columns.Contains("PrintFilterHeader") Then
						outputModel.DisplayTitleInReportHeader = CBool(row("PrintFilterHeader"))
					End If

					outputModel.Timestamp = CLng(row("Timestamp"))
					outputModel.BaseViewAccess = row("BaseViewAccess").ToString()

				End If

			Catch ex As Exception
				Throw

			End Try

		End Sub

		Private Sub PopulateSortOrder(outputModel As ReportBaseModel, data As DataTable)

			Dim objSort As SortOrderViewModel
			Dim iSortOrderID As Integer = 0

			Try

				For Each objRow As DataRow In data.Rows
					objSort = New SortOrderViewModel

					iSortOrderID += 1

					objSort.ReportID = outputModel.ID
					objSort.ReportType = outputModel.ReportType

					objSort.TableID = CInt(objRow("tableid"))
					objSort.ID = iSortOrderID
					objSort.ColumnID = CInt(objRow("Id"))

					objSort.Name = objRow("name").ToString
					objSort.Order = CType(IIf(objRow("order").ToString.ToUpper = "ASC", OrderType.Ascending, OrderType.Descending), OrderType)
					objSort.Sequence = CInt(objRow("Sequence"))

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
				Next

			Catch ex As Exception
				Throw

			End Try

		End Sub

		' can be done with dapper?
		Private Sub PopulateOutput(ReportType As UtilityType, outputModel As ReportOutputModel, data As DataTable)

			Try

				If data.Rows.Count = 1 Then

					Dim row As DataRow = data.Rows(0)

					outputModel.ReportType = ReportType
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

				Dim bContainsHiddenObject As Boolean

				outputModel.Columns = New List(Of ReportColumnItem)

				For Each objRow As DataRow In data.Rows

					Dim objItem As New ReportColumnItem() With {
						.ReportType = outputModel.ReportType,
						.ReportID = outputModel.ID,
						.Heading = HttpUtility.HtmlEncode(objRow("Heading").ToString),
						.IsExpression = CBool(objRow("IsExpression")),
						.ID = CInt(objRow("id")),
						.Name = objRow("Name").ToString,
						.TableID = CInt(objRow("Tableid")),
						.Sequence = CInt(objRow("Sequence")),
						.Size = CInt(objRow("Size")),
						.Decimals = CInt(objRow("Decimals")),
						.IsAverage = CBool(objRow("IsAverage")),
						.IsCount = CBool(objRow("IsCount")),
						.IsTotal = CBool(objRow("IsTotal")),
						.IsHidden = CBool(objRow("IsHidden")),
						.IsGroupWithNext = CBool(objRow("IsGroupWithNext")),
						.IsRepeated = CBool(objRow("IsRepeated"))}

					If objItem.IsExpression Then
						objItem.DataType = CalculationAsColumnType(objItem.TableID, objItem.ID)
					Else
						objItem.DataType = CType(objRow("DataType"), ColumnDataType)
					End If

					bContainsHiddenObject = bContainsHiddenObject OrElse CBool(objRow("AccessHidden"))

					If (objItem.IsExpression = True AndAlso bContainsHiddenObject = True) Then
						objItem.Access = "HD"
					End If

					outputModel.Columns.Add(objItem)

				Next

				' If found any selected calculation column hidden then set defination access rights to HD
				If (bContainsHiddenObject) Then
					outputModel.DefinitionAccessBasedOnSelectedCalculationColumns = "HD"
				End If

				'	outputModel.ContainsHiddenObjects = outputModel.ContainsHiddenObjects OrElse bContainsHiddenObject
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
						Return _calendarreports.Where(Function(m) m.ID = reportID).FirstOrDefault

					Case UtilityType.utlMailMerge
						Return _mailmerges.Where(Function(m) m.ID = reportID).FirstOrDefault

					Case UtilityType.utlCrossTab
						Return _crosstabs.Where(Function(m) m.ID = reportID).FirstOrDefault

					Case UtilityType.utlNineBoxGrid
						Return _nineboxgrids.Where(Function(m) m.ID = reportID).FirstOrDefault

					Case UtilityType.TalentReport
						Return _talentreports.Where(Function(m) m.ID = reportID).FirstOrDefault

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

		Public Function RetrieveDependencies(reportID As Integer, reportType As UtilityType) As ReportDependencies
			Return RetrieveParent(reportID, reportType).Dependencies
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


		Public Function GetExpressionListForTable(type As String, tableId As Integer) As List(Of ExpressionSelectionItem)

			Dim objReturnData As New List(Of ExpressionSelectionItem)

			Try

				Dim dtDefinition As DataTable = _objDataAccess.GetDataTable("spASRIntGetRecordSelection", CommandType.StoredProcedure _
				, New SqlParameter("@psType", SqlDbType.VarChar, 255) With {.Value = type} _
				, New SqlParameter("@piTableID", SqlDbType.Int) With {.Value = tableId})

				For Each objRow As DataRow In dtDefinition.Rows

					Dim objToAdd = New ExpressionSelectionItem With {
						.ID = CInt(objRow("ID")),
						.Name = objRow("Name").ToString,
						.Description = objRow("Description").ToString,
						.UserName = objRow("Username").ToString,
						.Access = objRow("Access").ToString}

					objReturnData.Add(objToAdd)

				Next

			Catch ex As Exception
				Throw

			End Try

			Return objReturnData

		End Function


		Public Function GetAvailableEmails(baseTableID As Integer) As Collection(Of ReportTableItem)

			Dim rstReportColumns = _objDataAccess.GetDataTable("spASRIntGetEmailAddresses", CommandType.StoredProcedure _
			, New SqlParameter("baseTableID", SqlDbType.Int) With {.Value = baseTableID})
			Dim items = New Collection(Of ReportTableItem)()
			For Each objRow As DataRow In rstReportColumns.Rows
				Dim objItem As New ReportTableItem() With {.id = CInt(objRow("id")), .Name = objRow("Name").ToString}
				items.Add(objItem)
			Next

			Return items

		End Function

		Public Function ServerValidate(objModel As CalendarReportModel) As SaveWarningModel

			Dim objSaveMessage As SaveWarningModel

			Try

				Dim prmErrorMsg As New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmErrorCode As New SqlParameter("piErrorCode", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmDeletedFilters As New SqlParameter("psDeletedFilters", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmHiddenFilters As New SqlParameter("psHiddenFilters", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmDeletedCalcs As New SqlParameter("psDeletedCalcs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmHiddenCalcs As New SqlParameter("psHiddenCalcs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmDeletedPicklists As New SqlParameter("psDeletedPicklists", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmHiddenPicklists As New SqlParameter("psHiddenPicklists", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmJobIDsToHide As New SqlParameter("psJobIDsToHide", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

				_objDataAccess.ExecuteSP("spASRIntValidateCalendarReport", _
								New SqlParameter("psUtilName", SqlDbType.VarChar, 255) With {.Value = objModel.Name}, _
								New SqlParameter("piUtilID", SqlDbType.Int) With {.Value = objModel.ID}, _
								New SqlParameter("piTimestamp", SqlDbType.Int) With {.Value = objModel.Timestamp}, _
								New SqlParameter("piBasePicklistID", SqlDbType.Int) With {.Value = objModel.PicklistID}, _
								New SqlParameter("piBaseFilterID", SqlDbType.Int) With {.Value = objModel.FilterID}, _
								New SqlParameter("piEmailGroupID", SqlDbType.Int) With {.Value = objModel.Output.EmailGroupID}, _
								New SqlParameter("piDescExprID", SqlDbType.Int) With {.Value = objModel.Description3ID}, _
								New SqlParameter("psEventFilterIDs", SqlDbType.VarChar, -1) With {.Value = objModel.Dependencies.EventFilters}, _
								New SqlParameter("piCustomStartID", SqlDbType.Int) With {.Value = objModel.StartCustomId}, _
								New SqlParameter("piCustomEndID", SqlDbType.Int) With {.Value = objModel.EndCustomId}, _
								New SqlParameter("piCategoryID", SqlDbType.Int) With {.Value = objModel.CategoryID}, _
								New SqlParameter("psHiddenGroups ", SqlDbType.VarChar, -1) With {.Value = objModel.GroupAccess.HiddenGroups()}, _
								prmErrorMsg, prmErrorCode, prmDeletedFilters, prmHiddenFilters, _
								prmDeletedCalcs, prmHiddenCalcs, prmDeletedPicklists, prmHiddenPicklists, prmJobIDsToHide)

				If prmJobIDsToHide.Value.ToString().Length > 0 Then
					objModel.Dependencies.JobIDsToHide = vbTab + prmJobIDsToHide.Value.ToString() + vbTab
				End If

				objSaveMessage = New SaveWarningModel With {
					.ReportType = objModel.ReportType,
					.ID = objModel.ID,
					.ErrorCode = CType(prmErrorCode.Value, ReportValidationStatus),
					.ErrorMessage = prmErrorMsg.Value.ToString()}

			Catch ex As Exception
				Throw

			End Try

			Return objSaveMessage

		End Function

		Public Function ServerValidate(objModel As CrossTabModel) As SaveWarningModel

			Dim objSaveMessage As SaveWarningModel

			Try

				Dim prmErrorMsg As New SqlParameter("@psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmErrorCode As New SqlParameter("@piErrorCode", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmDeletedFilters As New SqlParameter("@psDeletedFilters", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmHiddenFilters As New SqlParameter("@psHiddenFilters", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmJobIDsToHide As New SqlParameter("@psJobIDsToHide", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

				_objDataAccess.ExecuteSP("spASRIntValidateCrossTab", _
												New SqlParameter("psUtilName", SqlDbType.VarChar, 255) With {.Value = objModel.Name}, _
												New SqlParameter("piUtilID", SqlDbType.Int) With {.Value = objModel.ID}, _
												New SqlParameter("piTimestamp", SqlDbType.Int) With {.Value = objModel.Timestamp}, _
												New SqlParameter("piBasePicklistID", SqlDbType.Int) With {.Value = objModel.PicklistID}, _
												New SqlParameter("piBaseFilterID", SqlDbType.Int) With {.Value = objModel.FilterID}, _
												New SqlParameter("piEmailGroupID", SqlDbType.Int) With {.Value = objModel.Output.EmailGroupID}, _
												New SqlParameter("piCategoryID", SqlDbType.Int) With {.Value = objModel.CategoryID}, _
												New SqlParameter("psHiddenGroups ", SqlDbType.VarChar, -1) With {.Value = objModel.GroupAccess.HiddenGroups()}, _
												prmErrorMsg, prmErrorCode, prmDeletedFilters, prmHiddenFilters, prmJobIDsToHide)

				If prmJobIDsToHide.Value.ToString().Length > 0 Then
					objModel.Dependencies.JobIDsToHide = vbTab + prmJobIDsToHide.Value.ToString() + vbTab
				End If

				objSaveMessage = New SaveWarningModel With {
					.ReportType = objModel.ReportType,
					.ID = objModel.ID,
					.ErrorCode = CType(prmErrorCode.Value, ReportValidationStatus),
					.ErrorMessage = prmErrorMsg.Value.ToString()}

			Catch ex As Exception
				Throw

			End Try

			Return objSaveMessage

		End Function

		Public Function ServerValidate(objModel As NineBoxGridModel) As SaveWarningModel

			Dim objSaveMessage As SaveWarningModel

			Try

				Dim prmErrorMsg As New SqlParameter("@psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmErrorCode As New SqlParameter("@piErrorCode", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmDeletedFilters As New SqlParameter("@psDeletedFilters", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmHiddenFilters As New SqlParameter("@psHiddenFilters", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmJobIDsToHide As New SqlParameter("@psJobIDsToHide", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

				_objDataAccess.ExecuteSP("spASRIntValidateNineBoxGrid", _
												New SqlParameter("psUtilName", SqlDbType.VarChar, 255) With {.Value = objModel.Name}, _
												New SqlParameter("piUtilID", SqlDbType.Int) With {.Value = objModel.ID}, _
												New SqlParameter("piTimestamp", SqlDbType.Int) With {.Value = objModel.Timestamp}, _
												New SqlParameter("piBasePicklistID", SqlDbType.Int) With {.Value = objModel.PicklistID}, _
												New SqlParameter("piBaseFilterID", SqlDbType.Int) With {.Value = objModel.FilterID}, _
												New SqlParameter("piEmailGroupID", SqlDbType.Int) With {.Value = objModel.Output.EmailGroupID}, _
												New SqlParameter("piCategoryID", SqlDbType.Int) With {.Value = objModel.CategoryID}, _
												New SqlParameter("psHiddenGroups ", SqlDbType.VarChar, -1) With {.Value = objModel.GroupAccess.HiddenGroups()}, _
												prmErrorMsg, prmErrorCode, prmDeletedFilters, prmHiddenFilters, prmJobIDsToHide)

				If prmJobIDsToHide.Value.ToString().Length > 0 Then
					objModel.Dependencies.JobIDsToHide = vbTab + prmJobIDsToHide.Value.ToString() + vbTab
				End If

				objSaveMessage = New SaveWarningModel With {
					.ReportType = objModel.ReportType,
					.ID = objModel.ID,
					.ErrorCode = CType(prmErrorCode.Value, ReportValidationStatus),
					.ErrorMessage = prmErrorMsg.Value.ToString()}

			Catch ex As Exception
				Throw

			End Try

			Return objSaveMessage

		End Function

		Public Function ServerValidate(objModel As CustomReportModel) As SaveWarningModel

			Dim objSaveMessage As SaveWarningModel

			Try

				Dim prmErrorCode As New SqlParameter("@piErrorCode", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
				Dim prmErrorMsg As New SqlParameter("@psErrorMsg", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
				Dim prmDeletedCalcs As New SqlParameter("@psDeletedCalcs", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
				Dim prmHiddenCalcs As New SqlParameter("@psHiddenCalcs", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
				Dim prmDeletedFilters As New SqlParameter("@psDeletedFilters", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
				Dim prmHiddenFilters As New SqlParameter("@psHiddenFilters", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
				Dim prmDeletedOrders As New SqlParameter("@psDeletedOrders", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
				Dim prmJobIDsToHide As New SqlParameter("@psJobIDsToHide", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
				Dim prmDeletedPicklists As New SqlParameter("@psDeletedPicklists", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
				Dim prmHiddenPicklists As New SqlParameter("@psHiddenPicklists", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}

				_objDataAccess.ExecuteSP("spASRIntValidateCustomReport", _
												New SqlParameter("@psUtilName", SqlDbType.VarChar) With {.Value = objModel.Name, .Size = 255}, _
												New SqlParameter("@piUtilID", SqlDbType.Int) With {.Value = objModel.ID}, _
												New SqlParameter("@piTimestamp", SqlDbType.Int) With {.Value = objModel.Timestamp}, _
												New SqlParameter("@piBasePicklistID", SqlDbType.Int) With {.Value = objModel.PicklistID}, _
												New SqlParameter("@piBaseFilterID", SqlDbType.Int) With {.Value = objModel.FilterID}, _
												New SqlParameter("@piEmailGroupID", SqlDbType.Int) With {.Value = objModel.Output.EmailGroupID}, _
												New SqlParameter("@piParent1PicklistID", SqlDbType.Int) With {.Value = objModel.Parent1.PicklistID}, _
												New SqlParameter("@piParent1FilterID", SqlDbType.Int) With {.Value = objModel.Parent1.FilterID}, _
												New SqlParameter("@piParent2PicklistID", SqlDbType.Int) With {.Value = objModel.Parent2.PicklistID}, _
												New SqlParameter("@piParent2FilterID", SqlDbType.Int) With {.Value = objModel.Parent2.FilterID},
												New SqlParameter("@piChildFilterID", SqlDbType.VarChar) With {.Value = objModel.Dependencies.ChildFilters, .Size = 100}, _
												New SqlParameter("@psCalculations", SqlDbType.VarChar) With {.Value = objModel.Dependencies.Calculations, .Size = -1}, _
												New SqlParameter("@piCategoryID", SqlDbType.Int) With {.Value = objModel.CategoryID}, _
												New SqlParameter("@psHiddenGroups ", SqlDbType.VarChar) With {.Value = objModel.GroupAccess.HiddenGroups(), .Size = -1}, _
												prmErrorMsg, prmErrorCode, prmDeletedCalcs, prmHiddenCalcs, prmDeletedFilters, prmHiddenFilters, prmDeletedOrders, _
												prmJobIDsToHide, prmDeletedPicklists, prmHiddenPicklists)

				If prmJobIDsToHide.Value.ToString().Length > 0 Then
					objModel.Dependencies.JobIDsToHide = vbTab + prmJobIDsToHide.Value.ToString() + vbTab
				End If

				objSaveMessage = New SaveWarningModel With {
					.ReportType = objModel.ReportType,
					.ID = objModel.ID,
					.ErrorCode = CType(prmErrorCode.Value, ReportValidationStatus),
					.ErrorMessage = prmErrorMsg.Value.ToString()}

			Catch
				Throw

			End Try

			Return objSaveMessage

		End Function

		Public Function ServerValidate(objModel As MailMergeModel) As SaveWarningModel

			Dim objSaveMessage As SaveWarningModel

			Try

				Dim prmErrorMsg = New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmErrorCode = New SqlParameter("piErrorCode", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmDeletedCalcs = New SqlParameter("psDeletedCalcs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmHiddenCalcs = New SqlParameter("psHiddenCalcs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmJobIDsToHide = New SqlParameter("psJobIDsToHide", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

				_objDataAccess.ExecuteSP("spASRIntValidateMailMerge" _
						, New SqlParameter("@psUtilName", SqlDbType.VarChar, 255) With {.Value = objModel.Name} _
						, New SqlParameter("@piUtilID", SqlDbType.Int) With {.Value = objModel.ID} _
						, New SqlParameter("@piTimestamp", SqlDbType.Int) With {.Value = objModel.Timestamp} _
						, New SqlParameter("@piBasePicklistID", SqlDbType.Int) With {.Value = objModel.PicklistID} _
						, New SqlParameter("@piBaseFilterID", SqlDbType.Int) With {.Value = objModel.FilterID} _
						, New SqlParameter("@piCategoryID", SqlDbType.Int) With {.Value = objModel.CategoryID} _
						, New SqlParameter("@psCalculations", SqlDbType.VarChar, -1) With {.Value = objModel.Dependencies.Calculations} _
						, New SqlParameter("@psHiddenGroups", SqlDbType.VarChar, -1) With {.Value = objModel.GroupAccess.HiddenGroups()} _
						, prmErrorMsg, prmErrorCode, prmDeletedCalcs, prmHiddenCalcs, prmJobIDsToHide)

				If prmJobIDsToHide.Value.ToString().Length > 0 Then
					objModel.Dependencies.JobIDsToHide = vbTab + prmJobIDsToHide.Value.ToString() + vbTab
				End If

				objSaveMessage = New SaveWarningModel With {
					.ReportType = objModel.ReportType,
					.ID = objModel.ID,
					.ErrorCode = CType(prmErrorCode.Value, ReportValidationStatus),
					.ErrorMessage = prmErrorMsg.Value.ToString()}

			Catch ex As Exception
				Throw

			End Try

			Return objSaveMessage

		End Function

		Public Function ServerValidate(objModel As TalentReportModel) As SaveWarningModel

			Dim objSaveMessage As SaveWarningModel

			Try

				Dim prmErrorMsg = New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmErrorCode = New SqlParameter("piErrorCode", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmDeletedCalcs = New SqlParameter("psDeletedCalcs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmHiddenCalcs = New SqlParameter("psHiddenCalcs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmJobIDsToHide = New SqlParameter("psJobIDsToHide", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

				' We need to add the match table picklist and filter to this
				_objDataAccess.ExecuteSP("spASRIntValidateTalentReport" _
						, New SqlParameter("@psUtilName", SqlDbType.VarChar, 255) With {.Value = objModel.Name} _
						, New SqlParameter("@piUtilID", SqlDbType.Int) With {.Value = objModel.ID} _
						, New SqlParameter("@piTimestamp", SqlDbType.Int) With {.Value = objModel.Timestamp} _
						, New SqlParameter("@piBasePicklistID", SqlDbType.Int) With {.Value = objModel.PicklistID} _
						, New SqlParameter("@piBaseFilterID", SqlDbType.Int) With {.Value = objModel.FilterID} _
						, New SqlParameter("@piMatchPicklistID", SqlDbType.Int) With {.Value = objModel.MatchPicklistID} _
						, New SqlParameter("@piMatchFilterID", SqlDbType.Int) With {.Value = objModel.MatchFilterID} _
						, New SqlParameter("@piCategoryID", SqlDbType.Int) With {.Value = objModel.CategoryID} _
						, New SqlParameter("@psCalculations", SqlDbType.VarChar, -1) With {.Value = objModel.Dependencies.Calculations} _
						, New SqlParameter("@psHiddenGroups", SqlDbType.VarChar, -1) With {.Value = objModel.GroupAccess.HiddenGroups()} _
						, prmErrorMsg, prmErrorCode, prmDeletedCalcs, prmHiddenCalcs, prmJobIDsToHide)

				If prmJobIDsToHide.Value.ToString().Length > 0 Then
					objModel.Dependencies.JobIDsToHide = vbTab + prmJobIDsToHide.Value.ToString() + vbTab
				End If

				objSaveMessage = New SaveWarningModel With {
					.ReportType = objModel.ReportType,
					.ID = objModel.ID,
					.ErrorCode = CType(prmErrorCode.Value, ReportValidationStatus),
					.ErrorMessage = prmErrorMsg.Value.ToString()}

			Catch ex As Exception
				Throw

			End Try

			Return objSaveMessage

		End Function


		' Approximation of the calculation as a column type (not a direct match because we cannot handle table lookups, and expressions don't return integers
		Private Function CalculationAsColumnType(tableID As Integer, objectID As Integer) As ColumnDataType

			Dim returnType = ColumnDataType.sqlUnknown
			Dim objCalcExpr = New clsExprExpression(_objSessionInfo)
			objCalcExpr.Initialise(tableID, objectID, ExpressionTypes.giEXPR_RUNTIMECALCULATION, ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
			objCalcExpr.ConstructExpression()
			objCalcExpr.ValidateExpression(True)

			Select Case objCalcExpr.ReturnType
				Case ExpressionValueTypes.giEXPRVALUE_CHARACTER, ExpressionValueTypes.giEXPRVALUE_BYREF_CHARACTER
					returnType = ColumnDataType.sqlVarChar

				Case ExpressionValueTypes.giEXPRVALUE_NUMERIC, ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC
					returnType = ColumnDataType.sqlNumeric

				Case ExpressionValueTypes.giEXPRVALUE_DATE, ExpressionValueTypes.giEXPRVALUE_BYREF_DATE
					returnType = ColumnDataType.sqlDate

				Case ExpressionValueTypes.giEXPRVALUE_TABLEVALUE
					returnType = ColumnDataType.sqlUnknown

				Case ExpressionValueTypes.giEXPRVALUE_LOGIC, ExpressionValueTypes.giEXPRVALUE_BYREF_LOGIC
					returnType = ColumnDataType.sqlBoolean

				Case ExpressionValueTypes.giEXPRVALUE_PHOTO, ExpressionValueTypes.giEXPRVALUE_BYREF_PHOTO, ExpressionValueTypes.giEXPRVALUE_OLE, ExpressionValueTypes.giEXPRVALUE_BYREF_OLE
					returnType = ColumnDataType.sqlVarBinary

			End Select

			Return returnType

		End Function

		''' <summary>
		''' Gets category list
		''' </summary>
		''' <returns>List of category of type selectlistitem</returns>
		''' <remarks></remarks>
		Private Function GetCategoryList() As Collection(Of SelectListItem)

			Dim objItem As New Collection(Of SelectListItem)

			Try

				Dim dsDefinition As DataSet = _objDataAccess.GetDataSet("spsys_getobjectcategories" _
							 , New SqlParameter("utilityType", SqlDbType.Int) With {.Value = 0} _
							 , New SqlParameter("UtilityID", SqlDbType.Int) With {.Value = 0} _
							 , New SqlParameter("tableID", SqlDbType.Int) With {.Value = 0})

				Dim objRowDefaultItem As New SelectListItem() With {
																.Value = "0",
																.Text = "None"}

				objItem.Add(objRowDefaultItem)

				If dsDefinition.Tables(0).Rows.Count > 0 Then
					For Each objRow As DataRow In dsDefinition.Tables(0).Rows

						Dim objRowItem As New SelectListItem() With {
								.Value = CStr(objRow(0)),
								.Text = objRow(1).ToString()}

						objItem.Add(objRowItem)

					Next
				End If

			Catch ex As Exception
				Throw
			End Try

			Return objItem

		End Function

	End Class
End Namespace