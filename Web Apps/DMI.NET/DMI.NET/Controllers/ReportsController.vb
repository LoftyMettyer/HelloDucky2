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


		Function Util_Def_CustomReport() As ActionResult

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

		Function Util_Def_MailMerge() As ActionResult

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

		Function Util_Def_CrossTab() As ActionResult

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

		Function Util_Def_CalendarReport() As ActionResult

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

		<ValidateInput(False)>
	 Function util_def_customreports_submit(objModel As CustomReportModel) As ActionResult

			objReportRepository.SaveReportDefinition(objModel)

			Session("reaction") = "CUSTOMREPORTS"
			Return RedirectToAction("confirmok", "home")

		End Function

		<ValidateInput(False)>
	 Function util_def_mailmerge_submit(objModel As MailMergeModel) As ActionResult

			objReportRepository.SaveReportDefinition(objModel)

			Session("reaction") = "MAILMERGE"
			Return RedirectToAction("confirmok", "home")

		End Function

		Function util_def_crosstabs_submit(objModel As CrossTabModel) As ActionResult

			objReportRepository.SaveReportDefinition(objModel)

			Session("reaction") = "CROSSTABS"
			Return RedirectToAction("confirmok", "home")

		End Function

		'<HttpPost()>
		Function util_def_calendarreports_submit(objModel As CalendarReportModel) As ActionResult

			objReportRepository.SaveReportDefinition(objModel)

			Session("reaction") = "CALENDARREPORTS"
			Return RedirectToAction("confirmok", "home")

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