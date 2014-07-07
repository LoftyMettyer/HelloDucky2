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

Namespace Controllers

	Public Class ReportsController
		Inherits Controller

		Dim objReportRepository As New Repository.ReportRepository

		<HttpGet>
		Function util_def_customreport() As ActionResult

			Dim objModel As CustomReportModel
			Dim iReportID As Integer = CInt(Session("utilid"))
			Dim sAction = Session("action").ToString

			Select Case Session("action").ToString
				Case "new"
					objModel = objReportRepository.NewCustomReport()

				Case "copy"
					objModel = objReportRepository.LoadCustomReport(iReportID, True, sAction)

				Case Else
					objModel = objReportRepository.LoadCustomReport(iReportID, False, sAction)

			End Select

			Return View(objModel)

		End Function

		<HttpPost, ValidateInput(False)>
	 Function util_def_customreport(objModel As CustomReportModel) As ActionResult

			If ModelState.IsValid Then
				objReportRepository.SaveReportDefinition(objModel)
				Session("reaction") = "CUSTOMREPORTS"
				Return RedirectToAction("confirmok", "home")
			Else
				objModel.BaseTables = objReportRepository.GetTables()
				Return View(objModel)
			End If

		End Function

		<HttpGet>
		Function util_def_mailmerge() As ActionResult

			Dim iReportID As Integer = CInt(Session("utilid"))
			Dim sAction = Session("action").ToString

			Dim objModel As New MailMergeModel

			Select Case Session("action").ToString
				Case "new"
					objModel = objReportRepository.NewMailMerge()

				Case "copy"
					objModel = objReportRepository.LoadMailMerge(iReportID, True, sAction)

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

			Select Case Session("action").ToString
				Case "new"
					objModel = objReportRepository.NewCrossTab()

				Case "copy"
					objModel = objReportRepository.LoadCrossTab(iReportID, True, sAction)

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
				Return View(objModel)
			End If

		End Function

		<HttpGet>
		Function util_def_calendarreport() As ActionResult

			Dim objModel As CalendarReportModel
			Dim iReportID As Integer = CInt(Session("utilid"))
			Dim sAction = Session("action").ToString

			Select Case Session("action").ToString
				Case "new"
					objModel = objReportRepository.NewCalendarReport()

				Case "copy"
					objModel = objReportRepository.LoadCalendarReport(iReportID, True, sAction)

				Case Else
					objModel = objReportRepository.LoadCalendarReport(iReportID, False, sAction)

			End Select

			Return View(objModel)

		End Function

		<HttpPost, ValidateInput(False)>
		Function util_def_calendarreport(objModel As CalendarReportModel) As ActionResult

			If ModelState.IsValid Then
				objReportRepository.SaveReportDefinition(objModel)
				Session("reaction") = "CALENDARREPORTS"
				Return RedirectToAction("confirmok", "home")
			Else
				objModel.BaseTables = objReportRepository.GetTables()
				Return View(objModel)
			End If

		End Function

		<HttpGet>
		Function GetAvailableColumns(baseTableID As Integer) As JsonResult

			'Dim blah = objReportRepository.getModel(1)

			Dim objColumns = objReportRepository.GetColumnsForTable(baseTableID)
			Dim results = New With {.total = 1, .page = 1, .records = 1, .rows = objColumns}
			Return Json(results, JsonRequestBehavior.AllowGet)

		End Function

	End Class

End Namespace