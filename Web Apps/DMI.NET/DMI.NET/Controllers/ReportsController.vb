Option Explicit On
Option Strict On

Imports System.Net.Mail
Imports System.Net
Imports HR.Intranet.Server
Imports System.Data.SqlClient
Imports HR.Intranet.Server.Metadata
Imports System.Collections.ObjectModel
Imports DMI.NET.Models
Imports DMI.NET.Classes
Imports DMI.NET.Repository
Imports System.Web.Script.Serialization
Imports Newtonsoft.Json
Imports DMI.NET.ViewModels
Imports DMI.NET.ViewModels.Reports

Namespace Controllers

	Public Class ReportsController
		Inherits Controller

		Private objReportRepository As ReportRepository

		Public Sub New()
			objReportRepository = New ReportRepository
		End Sub

		' TODO (code beautification) - Replace with some kind of dependency injection (structuremap maybe?)
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

			Dim objModel As CustomReportModel
			Dim iReportID As Integer = CInt(Session("utilid"))
			Dim sAction = Session("action").ToString

			Select Case Session("action").ToString.ToUpper
				Case "NEW"
					objModel = objReportRepository.NewCustomReport()

				Case "COPY"
					objModel = objReportRepository.LoadCustomReport(iReportID, True, sAction)

				Case "VIEW"
					objModel = objReportRepository.LoadCustomReport(iReportID, False, sAction)
					objModel.IsReadOnly = True

				Case Else
					objModel = objReportRepository.LoadCustomReport(iReportID, False, sAction)

			End Select

			Return View(objModel)

		End Function

		<HttpPost, ValidateInput(False)>
	 Function util_def_customreport(objModel As CustomReportModel) As ActionResult

			Dim deserializer = New JavaScriptSerializer()

			If objModel.ChildTablesString.Length > 0 Then
				objModel.ChildTables = deserializer.Deserialize(Of Collection(Of ChildTableViewModel))(objModel.ChildTablesString)
			End If

			If objModel.SortOrderColumnsString.Length > 0 Then
				objModel.SortOrderColumns = deserializer.Deserialize(Of Collection(Of ReportSortItem))(objModel.SortOrderColumnsString)
			End If

			If ModelState.IsValid Then
				objReportRepository.SaveReportDefinition(objModel)
				Session("reaction") = "CUSTOMREPORTS"
				Return RedirectToAction("confirmok", "home")
			Else

				Dim allErrors = ModelState.Values.SelectMany(Function(v) v.Errors)

				Return View(objModel)
			End If

		End Function

		<HttpGet>
		Function util_def_mailmerge() As ActionResult

			Dim iReportID As Integer = CInt(Session("utilid"))
			Dim sAction = Session("action").ToString

			Dim objModel As New MailMergeModel

			Select Case Session("action").ToString.ToUpper
				Case "NEW"
					objModel = objReportRepository.NewMailMerge()

				Case "COPY"
					objModel = objReportRepository.LoadMailMerge(iReportID, True, sAction)

				Case "VIEW"
					objModel = objReportRepository.LoadMailMerge(iReportID, False, sAction)
					objModel.IsReadOnly = True

				Case Else
					objModel = objReportRepository.LoadMailMerge(iReportID, False, sAction)

			End Select

			Return View(objModel)

		End Function

		<HttpPost, ValidateInput(False)>
	 Function util_def_mailmerge(objModel As MailMergeModel) As ActionResult

			If ModelState.IsValid Then
				objReportRepository.SaveReportDefinition(objModel)
				Session("reaction") = "MAILMERGE"
				Return RedirectToAction("confirmok", "home")
			Else
				objModel.BaseTables = objReportRepository.GetTables()
				Return View(objModel)
			End If

		End Function

		<HttpGet>
		Function util_def_crosstab() As ActionResult

			Dim objModel As CrossTabModel
			Dim iReportID As Integer = CInt(Session("utilid"))
			Dim sAction = Session("action").ToString

			Select Case Session("action").ToString.ToUpper
				Case "NEW"
					objModel = objReportRepository.NewCrossTab()

				Case "COPY"
					objModel = objReportRepository.LoadCrossTab(iReportID, True, sAction)

				Case "VIEW"
					objModel = objReportRepository.LoadCrossTab(iReportID, False, sAction)
					objModel.IsReadOnly = True

				Case Else
					objModel = objReportRepository.LoadCrossTab(iReportID, False, sAction)

			End Select

			Return View(objModel)

		End Function

		<HttpPost, ValidateInput(False)>
		Function util_def_crosstab(objModel As CrossTabModel) As ActionResult

			If ModelState.IsValid Then
				objReportRepository.SaveReportDefinition(objModel)
				Session("reaction") = "CROSSTABS"
				Return RedirectToAction("confirmok", "home")
			Else
				objModel.BaseTables = objReportRepository.GetTables()
				objModel.AvailableColumns = objReportRepository.GetColumnsForTable(objModel.BaseTableID)

				Return View(objModel)
			End If

		End Function

		<HttpGet>
		Function util_def_calendarreport() As ActionResult

			Dim objModel As CalendarReportModel
			Dim iReportID As Integer = CInt(Session("utilid"))
			Dim sAction = Session("action").ToString

			Select Case Session("action").ToString.ToUpper
				Case "NEW"
					objModel = objReportRepository.NewCalendarReport()

				Case "COPY"
					objModel = objReportRepository.LoadCalendarReport(iReportID, True, sAction)

				Case "VIEW"
					objModel = objReportRepository.LoadCalendarReport(iReportID, False, sAction)
					objModel.IsReadOnly = True

				Case Else
					objModel = objReportRepository.LoadCalendarReport(iReportID, False, sAction)

			End Select

			Return View(objModel)

		End Function

		<HttpPost, ValidateInput(False)>
		Function util_def_calendarreport(objModel As CalendarReportModel) As ActionResult

			Dim deserializer = New JavaScriptSerializer()

			If objModel.EventsString.Length > 0 Then
				objModel.Events = deserializer.Deserialize(Of Collection(Of CalendarEventDetailViewModel))(objModel.EventsString)
			End If

			If objModel.SortOrderColumnsString.Length > 0 Then
				objModel.SortOrderColumns = deserializer.Deserialize(Of Collection(Of ReportSortItem))(objModel.SortOrderColumnsString)
			End If


			If ModelState.IsValid Then
				objReportRepository.SaveReportDefinition(objModel)
				Session("reaction") = "CALENDARREPORTS"
				Return RedirectToAction("confirmok", "home")
			Else

				Dim allErrors = ModelState.Values.SelectMany(Function(v) v.Errors)

				Return View(objModel)
			End If

		End Function

		<HttpGet>
		Function GetColumnsForTable(TableID As Integer) As JsonResult

			Dim objColumns = objReportRepository.GetColumnsForTable(TableID)
			Dim results = New With {.total = 1, .page = 1, .records = 0, .rows = objColumns}
			Return Json(results, JsonRequestBehavior.AllowGet)

		End Function

		<HttpGet>
		Function GetAvailableTablesForReport(ID As Integer) As JsonResult

			Dim objReport = objReportRepository.RetrieveReport(ID)

			Dim objTables = objReportRepository.GetTables()

			' subtract the currently selected ones

			Return Json(objTables, JsonRequestBehavior.AllowGet)

		End Function

		<HttpGet>
		Function GetBaseTables() As JsonResult

			Dim objTables = objReportRepository.GetTables()
			Return Json(objTables, JsonRequestBehavior.AllowGet)

		End Function

		<HttpPost>
		Function AddChildTable(ReportID As Integer) As ActionResult

			Dim objReport = objReportRepository.RetrieveReport(ReportID)

			Dim objModel As New ChildTableViewModel
			objModel.AvailableTables = objReportRepository.GetChildTables(objReport.BaseTableID, False)

			For Each objTable In objReport.ChildTables
				objModel.AvailableTables.RemoveAll(Function(m) m.id = objTable.TableID)
			Next

			objModel.ReportID = ReportID


			Return PartialView("EditorTemplates\ReportChildTable", objModel)


		End Function

		<HttpPost>
		Function EditChildTable(objModel As ChildTableViewModel) As ActionResult

			Dim objReport = objReportRepository.RetrieveReport(objModel.ReportID)

			objModel.AvailableTables = objReportRepository.GetTables()

			Return PartialView("EditorTemplates\ReportChildTable", objModel)
		End Function

		<HttpPost>
		Sub PostChildTable(objModel As ChildTableViewModel)

			Try

				If ModelState.IsValid Then

					Dim objReport = objReportRepository.RetrieveReport(objModel.ReportID)
					Dim original = objReport.ChildTables.Where(Function(m) m.TableID = objModel.TableID).FirstOrDefault

					If Not original Is Nothing Then
						objReport.ChildTables.Remove(original)
					End If

					objReport.ChildTables.Add(objModel)

				End If

			Catch ex As Exception
				Throw

			End Try

		End Sub


		<HttpPost>
		Function AddCalendarEvent(ReportID As Integer) As ActionResult

			Dim objReport = objReportRepository.RetrieveCalendarReport(ReportID)

			Dim objModel As New CalendarEventDetailViewModel

			objModel.TableID = objReport.BaseTableID
			objModel.CalendarReportID = ReportID
			objModel.EventKey = String.Format("EV_{0}", objReport.Events.Count + 1)

			ModelState.Clear()
			Return PartialView("EditorTemplates\CalendarEventDetail", objModel)


		End Function

		<HttpPost>
		Function EditCalendarEvent(objModel As CalendarEventDetailViewModel) As ActionResult

			Dim objReport = objReportRepository.RetrieveCalendarReport(objModel.CalendarReportID)
			objModel.AvailableTables = objReportRepository.GetChildTables(objReport.BaseTableID, True)

			ModelState.Clear()
			Return PartialView("EditorTemplates\CalendarEventDetail", objModel)
		End Function

		<HttpPost>
		Sub PostCalendarEvent(objModel As CalendarEventDetailViewModel)

			Dim objReport = objReportRepository.RetrieveCalendarReport(objModel.CalendarReportID)
			Dim original = objReport.Events.Where(Function(m) m.EventKey = objModel.EventKey).FirstOrDefault

			If Not original Is Nothing Then
				objReport.Events.Remove(original)
			End If

			objReport.Events.Add(objModel)

		End Sub

		<HttpPost, ValidateInput(False)>
	 Function ChangeEventBaseTable(objModel As CalendarEventDetailViewModel) As ActionResult

			objModel.ChangeBaseTable()

			ModelState.Clear()
			Return PartialView("EditorTemplates\CalendarEventDetail", objModel)

		End Function

		<HttpPost>
		Sub RemoveCalendarEvent(objModel As CalendarEventDetailViewModel)

			Dim objReport = objReportRepository.RetrieveCalendarReport(objModel.CalendarReportID)
			Dim original = objReport.Events.Where(Function(m) m.EventKey = objModel.EventKey).FirstOrDefault

			If Not original Is Nothing Then
				objReport.Events.Remove(original)
			End If

		End Sub

		<HttpGet>
		Function GetAllTablesInReport(reportID As Integer) As JsonResult

			Dim objItems = objReportRepository.GetAllTablesInReport(reportID)
			Return Json(objItems, JsonRequestBehavior.AllowGet)

		End Function

		<HttpPost, ValidateInput(False)>
	 Function ChangeBaseTable(objModel As CustomReportModel) As ActionResult

			objReportRepository.SetBaseTable(objModel)
			ModelState.Clear()
			Return View("UTIL_DEF_CUSTOMREPORT", objModel)

		End Function

	End Class

End Namespace