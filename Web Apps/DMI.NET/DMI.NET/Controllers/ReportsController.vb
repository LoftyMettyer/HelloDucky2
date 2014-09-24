Option Explicit On
Option Strict On

Imports System.Collections.ObjectModel
Imports DMI.NET.Models
Imports DMI.NET.Classes
Imports DMI.NET.Repository
Imports System.Web.Script.Serialization
Imports DMI.NET.ViewModels.Reports
Imports DMI.NET.Code
Imports DMI.NET.Code.Extensions

Namespace Controllers

	Public Class ReportsController
		Inherits Controller

		Private objReportRepository As ReportRepository

		Public Sub New()
			objReportRepository = New ReportRepository
		End Sub

		Protected Overrides Sub Initialize(requestContext As RequestContext)
			MyBase.Initialize(requestContext)

			If requestContext.HttpContext.Session("reportrepository") Is Nothing Then
				requestContext.HttpContext.Session("reportrepository") = objReportRepository
			Else
				objReportRepository = CType(requestContext.HttpContext.Session("reportrepository"), ReportRepository)
			End If

		End Sub

		<HttpGet>
		Function util_def_customreport() As ActionResult

			Dim iReportID As Integer = CInt(Session("utilid"))
			Dim iAction = ActionToUtilityAction(Session("action").ToString)
			Dim objModel = objReportRepository.LoadCustomReport(iReportID, iAction)

			Return View(objModel)

		End Function

		<HttpGet>
		Function util_def_mailmerge() As ActionResult

			Dim iReportID As Integer = CInt(Session("utilid"))
			Dim iAction = ActionToUtilityAction(Session("action").ToString)

			Dim objModel = objReportRepository.LoadMailMerge(iReportID, iAction)

			Return View(objModel)

		End Function

		<HttpGet>
		Function util_def_crosstab() As ActionResult

			Dim iReportID As Integer = CInt(Session("utilid"))
			Dim iAction = ActionToUtilityAction(Session("action").ToString)
			Dim objModel = objReportRepository.LoadCrossTab(iReportID, iAction)

			Return View(objModel)

		End Function

		<HttpGet>
		Function util_def_calendarreport() As ActionResult

			Dim iReportID As Integer = CInt(Session("utilid"))
			Dim iAction = ActionToUtilityAction(Session("action").ToString)
			Dim objModel = objReportRepository.LoadCalendarReport(iReportID, iAction)

			Return View(objModel)

		End Function

		<HttpPost, ValidateInput(False)>
	 Function util_def_customreport(objModel As CustomReportModel) As ActionResult

			Dim objSaveWarning As SaveWarningModel
			Dim deserializer = New JavaScriptSerializer()

			objModel.Dependencies = objReportRepository.RetrieveDependencies(objModel.ID, UtilityType.utlCustomReport)

			If objModel.ColumnsAsString IsNot Nothing Then
				If objModel.ColumnsAsString.Length > 0 Then
					objModel.Columns = deserializer.Deserialize(Of List(Of ReportColumnItem))(objModel.ColumnsAsString)
				End If

				If objModel.IsSummary AndAlso objModel.Columns.Where(Function(m) m.IsAverage OrElse m.IsCount OrElse m.IsTotal).LongCount() = 0 Then
					ModelState.AddModelError("IsSummaryOK", "There are no columns defined as aggregates for this summary report.")
				End If

				If objModel.IgnoreZerosForAggregates AndAlso objModel.Columns.Where( _
					Function(m) (m.DataType = ColumnDataType.sqlInteger OrElse m.DataType = ColumnDataType.sqlNumeric) AndAlso (m.IsAverage OrElse m.IsCount OrElse m.IsTotal) _
						).LongCount() = 0 Then
					ModelState.AddModelError("IsIgnoreZerosOK", "You have chosen to ignore zeros when calculating aggregates, but have not selected to show aggregates for any numeric columns.")
				End If

			End If

			If objModel.ChildTablesString IsNot Nothing Then
				If objModel.ChildTablesString.Length > 0 Then
					objModel.ChildTables = deserializer.Deserialize(Of List(Of ChildTableViewModel))(objModel.ChildTablesString)
				End If
			End If

			If objModel.SortOrdersString IsNot Nothing Then
				If objModel.SortOrdersString.Length > 0 Then
					objModel.SortOrders = deserializer.Deserialize(Of List(Of SortOrderViewModel))(objModel.SortOrdersString)
				End If

				If objModel.IsSummary AndAlso objModel.SortOrders.Where(Function(m) m.ValueOnChange).LongCount() = 0 Then
					ModelState.AddModelError("IsValueOnChangeOK", "There are no columns defined as 'value on change' for this summary report.")
				End If

			End If

			If objModel.ValidityStatus = ReportValidationStatus.ServerCheckComplete Then
				objReportRepository.SaveReportDefinition(objModel)
				Session("reaction") = "CUSTOMREPORTS"
				Session("utilid") = objModel.ID
				Return RedirectToAction("confirmok", "home")

			Else
				If ModelState.IsValid Then
					objSaveWarning = objReportRepository.ServerValidate(objModel)
				Else
					objSaveWarning = ModelState.ToWebMessage
				End If

				Return Json(objSaveWarning, JsonRequestBehavior.AllowGet)

			End If

		End Function

		<HttpPost, ValidateInput(False)>
	 Function util_def_mailmerge(objModel As MailMergeModel) As ActionResult

			Dim objSaveWarning As SaveWarningModel
			Dim deserializer = New JavaScriptSerializer()

			objModel.Dependencies = objReportRepository.RetrieveDependencies(objModel.ID, UtilityType.utlMailMerge)

			If objModel.ColumnsAsString IsNot Nothing Then
				If objModel.ColumnsAsString.Length > 0 Then
					objModel.Columns = deserializer.Deserialize(Of List(Of ReportColumnItem))(objModel.ColumnsAsString)
				End If
			End If

			If objModel.SortOrdersString IsNot Nothing Then
				If objModel.SortOrdersString.Length > 0 Then
					objModel.SortOrders = deserializer.Deserialize(Of List(Of SortOrderViewModel))(objModel.SortOrdersString)
				End If
			End If

			If objModel.ValidityStatus = ReportValidationStatus.ServerCheckComplete Then
				objReportRepository.SaveReportDefinition(objModel)
				Session("reaction") = "MAILMERGE"
				Session("utilid") = objModel.ID
				Return RedirectToAction("confirmok", "home")

			Else

				If ModelState.IsValid Then
					objSaveWarning = objReportRepository.ServerValidate(objModel)
				Else
					objSaveWarning = ModelState.ToWebMessage
				End If

				Return Json(objSaveWarning, JsonRequestBehavior.AllowGet)

			End If

		End Function

		<HttpPost, ValidateInput(False)>
		Function util_def_crosstab(objModel As CrossTabModel) As ActionResult

			Dim objSaveWarning As SaveWarningModel
			objModel.Dependencies = objReportRepository.RetrieveDependencies(objModel.ID, UtilityType.utlCrossTab)

			If objModel.ValidityStatus = ReportValidationStatus.ServerCheckComplete Then
				objReportRepository.SaveReportDefinition(objModel)
				Session("reaction") = "CROSSTABS"
				Session("utilid") = objModel.ID
				Return RedirectToAction("confirmok", "home")

			Else

				If ModelState.IsValid Then
					objSaveWarning = objReportRepository.ServerValidate(objModel)
				Else
					objSaveWarning = ModelState.ToWebMessage
				End If

				Return Json(objSaveWarning, JsonRequestBehavior.AllowGet)

			End If

		End Function

		<HttpPost, ValidateInput(False)>
		Function util_def_calendarreport(objModel As CalendarReportModel) As ActionResult

			Dim objSaveWarning As SaveWarningModel
			Dim deserializer = New JavaScriptSerializer()

			objModel.Dependencies = objReportRepository.RetrieveDependencies(objModel.ID, UtilityType.utlCalendarReport)

			If objModel.EventsString IsNot Nothing Then
				If objModel.EventsString.Length > 0 Then
					objModel.Events = deserializer.Deserialize(Of Collection(Of CalendarEventDetailViewModel))(objModel.EventsString)
				End If
			End If

			If objModel.SortOrdersString IsNot Nothing Then
				If objModel.SortOrdersString.Length > 0 Then
					objModel.SortOrders = deserializer.Deserialize(Of List(Of SortOrderViewModel))(objModel.SortOrdersString)
				End If
			End If

			If objModel.ValidityStatus = ReportValidationStatus.ServerCheckComplete Then
				objReportRepository.SaveReportDefinition(objModel)
				Session("reaction") = "CALENDARREPORTS"
				Session("utilid") = objModel.ID
				Return RedirectToAction("confirmok", "home")

			Else

				If ModelState.IsValid Then
					objSaveWarning = objReportRepository.ServerValidate(objModel)
				Else
					objSaveWarning = ModelState.ToWebMessage
				End If

				Return Json(objSaveWarning, JsonRequestBehavior.AllowGet)
			End If

		End Function

		<HttpGet>
		Function GetAvailableColumnsForTable(TableID As Integer) As JsonResult

			Dim objResults = objReportRepository.GetColumnsForTable(TableID)
			Return Json(objResults, JsonRequestBehavior.AllowGet)

		End Function

		<HttpGet>
		Function GetAvailableItemsForTable(TableID As Integer, reportID As Integer, reportType As UtilityType, selectionType As String) As JsonResult

			Dim objReport = objReportRepository.RetrieveParent(reportID, reportType)
			Dim objAvailable As List(Of ReportColumnItem)

			If selectionType = "C" Then
				objAvailable = objReportRepository.GetColumnsForTable(TableID)
				For Each objItem In objReport.Columns.Where(Function(m) Not m.IsExpression)
					objAvailable.RemoveAll(Function(m) m.ID = objItem.ID)
				Next

			Else
				objAvailable = objReportRepository.GetCalculationsForTable(TableID)
				For Each objItem In objReport.Columns.Where(Function(m) m.IsExpression)
					objAvailable.RemoveAll(Function(m) m.ID = objItem.ID)
				Next

			End If

			Dim results = New With {.total = 1, .page = 1, .records = 0, .rows = objAvailable}
			Return Json(results, JsonRequestBehavior.AllowGet)

		End Function

		<HttpGet>
		Function GetBaseTables(reportType As UtilityType) As JsonResult

			Dim objTables = objReportRepository.GetTables(reportType)
			Return Json(objTables, JsonRequestBehavior.AllowGet)

		End Function

		<HttpPost>
		Function AddChildTable(ReportID As Integer) As ActionResult

			Dim objModel As New ChildTableViewModel With {.ReportID = ReportID, .ReportType = UtilityType.utlCustomReport}
			Dim objReport = CType(objReportRepository.RetrieveParent(objModel), CustomReportModel)

			objModel.AvailableTables = objReportRepository.GetChildTables(objReport.BaseTableID, False)

			For Each objTable In objReport.ChildTables
				objModel.AvailableTables.RemoveAll(Function(m) m.id = objTable.TableID)
			Next

			objModel.ID = objReport.ChildTables.Count + 1
			objModel.IsAdd = True

			Return PartialView("EditorTemplates\ReportChildTable", objModel)


		End Function

		<HttpPost>
		Function EditChildTable(objModel As ChildTableViewModel) As ActionResult

			Dim objReport = CType(objReportRepository.RetrieveParent(objModel), CustomReportModel)
			objModel.AvailableTables = objReportRepository.GetChildTables(objReport.BaseTableID, False)

			For Each objTable In objReport.ChildTables
				objModel.AvailableTables.RemoveAll(Function(m) m.id = objTable.TableID AndAlso objModel.TableID <> m.id)
			Next

			objModel.IsAdd = False

			Return PartialView("EditorTemplates\ReportChildTable", objModel)
		End Function

		<HttpPost>
		Sub PostChildTable(objModel As ChildTableViewModel)

			Try

				Dim objReport = CType(objReportRepository.RetrieveParent(objModel), CustomReportModel)

				' Remove original
				objReport.ChildTables.RemoveAll(Function(m) m.TableID = objModel.TableID)
				objReport.ChildTables.Add(objModel)

			Catch ex As Exception
				Throw

			End Try

		End Sub

		<HttpPost>
		Function AddCalendarEvent(ReportID As Integer) As ActionResult

			Dim objReport = objReportRepository.RetrieveCalendarReport(ReportID)

			Dim objModel As New CalendarEventDetailViewModel

			objModel.ID = 0
			objModel.TableID = objReport.BaseTableID
			objModel.ReportID = ReportID
			objModel.EventKey = String.Format("EV_{0}", objReport.Events.Count + 1)
			objModel.AvailableTables = objReportRepository.GetTablesWithEvents(objReport.BaseTableID)

			ModelState.Clear()
			Return PartialView("EditorTemplates\CalendarEventDetail", objModel)


		End Function

		<HttpPost>
		Function EditCalendarEvent(objModel As CalendarEventDetailViewModel) As ActionResult

			Dim objReport = objReportRepository.RetrieveCalendarReport(objModel.ReportID)
			objModel.AvailableTables = objReportRepository.GetTablesWithEvents(objReport.BaseTableID)

			ModelState.Clear()
			Return PartialView("EditorTemplates\CalendarEventDetail", objModel)
		End Function

		<HttpPost>
		Sub PostCalendarEvent(objModel As CalendarEventDetailViewModel)

			Dim objReport = objReportRepository.RetrieveCalendarReport(objModel.ReportID)
			Dim original = objReport.Events.Where(Function(m) m.EventKey = objModel.EventKey).FirstOrDefault

			If original IsNot Nothing Then
				objReport.Events.Remove(original)
			End If

			objReport.Events.Add(objModel)

		End Sub

		<HttpPost, ValidateInput(False)>
	 Function ChangeEventBaseTable(objModel As CalendarEventDetailViewModel) As ActionResult

			Dim objReport = objReportRepository.RetrieveCalendarReport(objModel.ReportID)

			objModel.ChangeBaseTable()
			objModel.AvailableTables = objReportRepository.GetChildTables(objReport.BaseTableID, True)

			ModelState.Clear()
			Return PartialView("EditorTemplates\CalendarEventDetail", objModel)

		End Function

		<HttpPost>
	 Function ChangeEventLookupTable(objModel As CalendarEventDetailViewModel) As ActionResult

			Dim objReport = objReportRepository.RetrieveCalendarReport(objModel.ReportID)
			objModel.AvailableTables = objReportRepository.GetChildTables(objReport.BaseTableID, True)
			Return PartialView("EditorTemplates\CalendarEventDetail", objModel)

		End Function

		<HttpPost>
		Sub RemoveCalendarEvent(objModel As CalendarEventDetailViewModel)

			Dim objReport = objReportRepository.RetrieveCalendarReport(objModel.ReportID)
			Dim original = objReport.Events.Where(Function(m) m.EventKey = objModel.EventKey).FirstOrDefault

			If original IsNot Nothing Then
				objReport.Events.Remove(original)
			End If

		End Sub

		<HttpGet>
		Function GetAllTablesInReport(reportID As Integer, reportType As UtilityType) As JsonResult

			Dim objReport = objReportRepository.RetrieveParent(reportID, reportType)
			Return Json(objReport.GetAvailableTables(), JsonRequestBehavior.AllowGet)

		End Function

		<HttpPost>
		Function ChangeBaseTable(ReportID As Integer, ReportType As UtilityType, BaseTableID As Integer) As JsonResult

			Dim objDetail As New ReportColumnItem
			Dim bChildTablesAvailable As Integer
			objDetail.ReportID = ReportID
			objDetail.ReportType = ReportType

			Dim objReport = objReportRepository.RetrieveParent(objDetail)
			objReport.BaseTableID = BaseTableID

			objReport.SetBaseTable(BaseTableID)

			If ReportType = UtilityType.utlCustomReport Then
				bChildTablesAvailable = CType(objReport, CustomReportModel).ChildTablesAvailable
			End If

			Dim result = New With {.childTablesAvailable = bChildTablesAvailable}
			Return Json(result, JsonRequestBehavior.AllowGet)

		End Function

		<HttpPost>
		Function AddSortOrder(ReportID As Integer, ReportType As UtilityType) As ActionResult

			Dim objModel As New SortOrderViewModel

			objModel.ReportID = ReportID
			objModel.ReportType = ReportType

			Dim objReport = objReportRepository.RetrieveParent(objModel)

			If objReport.SortOrders.Count = 0 Then
				objModel.ID = 1
				objModel.Sequence = 1
			Else
				objModel.ID = objReport.SortOrders.Max(Function(m) m.ID) + 1
				objModel.Sequence = objReport.SortOrders.Max(Function(m) m.Sequence) + 1
			End If

			objModel.AvailableColumns = objReport.GetAvailableSortColumns(objModel)

			ModelState.Clear()
			Return PartialView("EditorTemplates\SortOrder", objModel)

		End Function

		<HttpPost>
		Function EditSortOrder(objModel As SortOrderViewModel) As ActionResult

			Dim objReport = objReportRepository.RetrieveParent(objModel)
			objModel.AvailableColumns = objReport.GetAvailableSortColumns(objModel)

			ModelState.Clear()
			Return PartialView("EditorTemplates\SortOrder", objModel)
		End Function

		<HttpPost>
		Sub PostSortOrder(objModel As SortOrderViewModel)

			Dim objReport As IReport
			objReport = objReportRepository.RetrieveParent(objModel)

			Dim original = objReport.SortOrders.Where(Function(m) m.ID = objModel.ID).FirstOrDefault

			If original IsNot Nothing Then
				objReport.SortOrders.Remove(original)
			End If

			objReport.SortOrders.Add(objModel)

		End Sub

		<HttpPost>
		Sub RemoveSortOrder(objModel As SortOrderViewModel)

			Dim objReport As IReport
			objReport = objReportRepository.RetrieveParent(objModel)

			Dim original = objReport.SortOrders.Where(Function(m) m.ID = objModel.ID).FirstOrDefault

			If original IsNot Nothing Then
				objReport.SortOrders.Remove(original)
			End If

		End Sub

		<HttpPost>
		Sub AddReportColumn(objModel As ReportColumnItem)

			Dim objReport As ReportBaseModel
			objReport = CType(objReportRepository.RetrieveParent(objModel), ReportBaseModel)

			objReport.Columns.Add(objModel)

		End Sub

		<HttpPost>
		Sub RemoveAllChildTables(objModel As ReportColumnItem)

			Dim objReport As CustomReportModel
			objReport = CType(objReportRepository.RetrieveParent(objModel), CustomReportModel)

			For Each objChildTable In objReport.ChildTables
				objReport.Columns.RemoveAll(Function(m) m.TableID = objChildTable.TableID)
			Next

			objReport.ChildTables.Clear()

		End Sub

		<HttpPost>
		Sub RemoveChildTable(objModel As ReportColumnItem)

			Dim objReport As CustomReportModel
			objReport = CType(objReportRepository.RetrieveParent(objModel), CustomReportModel)

			objReport.ChildTables.RemoveAll(Function(m) m.ID = objModel.ID)
			objReport.Columns.RemoveAll(Function(m) m.TableID = objModel.TableID)

		End Sub

		<HttpPost, ValidateInput(False)>
		Sub RemoveReportColumn(objModel As ReportColumnCollection)

			Dim objReport As ReportBaseModel
			objReport = CType(objReportRepository.RetrieveParent(objModel), ReportBaseModel)

			For Each iColumnID In objModel.Columns
				objReport.Columns.RemoveAll(Function(m) m.ID = iColumnID)
				objReport.SortOrders.RemoveAll(Function(m) m.ColumnID = iColumnID)
			Next

		End Sub

		<HttpPost>
		Sub RemoveAllReportColumns(objModel As ReportColumnItem)

			Dim objReport As ReportBaseModel
			objReport = CType(objReportRepository.RetrieveParent(objModel), ReportBaseModel)

			objReport.Columns.Clear()
			objReport.SortOrders.Clear()

		End Sub


		<HttpGet>
		Function GetExpressionsForTable(TableID As Integer, SelectionType As String) As JsonResult

			Dim objAvailable As List(Of ExpressionSelectionItem)

			objAvailable = objReportRepository.GetExpressionListForTable(SelectionType, TableID)

			Dim results = New With {.total = 1, .page = 1, .records = 0, .rows = objAvailable}
			Return Json(results, JsonRequestBehavior.AllowGet)

		End Function

	End Class

End Namespace