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
Imports DMI.NET.Models.ObjectRequests
Imports HR.Intranet.Server
Imports System.Data.SqlClient
Imports System.IO

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
      Function util_def_9boxgrid() As ActionResult

         Dim iReportID As Integer = CInt(Session("utilid"))
         Dim iAction = ActionToUtilityAction(Session("action").ToString)
         Dim objModel = objReportRepository.LoadNineBoxGrid(iReportID, iAction)

         Return View(objModel)

      End Function

      <HttpGet>
      Function util_def_talentreport() As ActionResult

         Dim iReportID As Integer = CInt(Session("utilid"))
         Dim iAction = ActionToUtilityAction(Session("action").ToString)

         Dim objModel = objReportRepository.LoadTalentReport(iReportID, iAction)

         Return View(objModel)

      End Function

      <HttpGet>
      Function util_def_calendarreport() As ActionResult

         Dim iReportID As Integer = CInt(Session("utilid"))
         Dim iAction = ActionToUtilityAction(Session("action").ToString)
         Dim objModel = objReportRepository.LoadCalendarReport(iReportID, iAction)

         Return View(objModel)

      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
      Function util_def_customreport(objModel As CustomReportModel) As ActionResult

         Dim objSaveWarning As SaveWarningModel
         Dim deserializer = New JavaScriptSerializer()
         Dim hiddenColumnsCount As Integer

         objModel.Dependencies = objReportRepository.RetrieveDependencies(objModel.ID, UtilityType.utlCustomReport)

         If objModel.ColumnsAsString IsNot Nothing Then
            If objModel.ColumnsAsString.Length > 0 Then
               objModel.Columns = deserializer.Deserialize(Of List(Of ReportColumnItem))(objModel.ColumnsAsString)
            End If

            ' Check the column heading has value.
            For Each columnItem As ReportColumnItem In objModel.Columns
               If String.IsNullOrEmpty(columnItem.Heading.Trim()) And columnItem.IsHidden = False Then
                  ModelState.AddModelError("IsColumnHeaderEmpty", "The '" & columnItem.Name & "' column has a blank heading.")
                  Exit For
               End If

               ' Count the hidden columns to validate if all columns are hidden or not
               If (columnItem.IsHidden) Then
                  hiddenColumnsCount += 1
               End If
            Next

            ' Check the column headings are unique.
            Dim breakNestedLoop As Boolean
            For Each columnItem As ReportColumnItem In objModel.Columns
               For Each columnItemHeaderToCheck As ReportColumnItem In objModel.Columns

                  If columnItem.ID <> columnItemHeaderToCheck.ID AndAlso UCase(columnItem.Heading.Trim()) = UCase(columnItemHeaderToCheck.Heading.Trim()) AndAlso columnItemHeaderToCheck.IsHidden = False Then
                     ModelState.AddModelError("IsColumnHeaderUnique", "One or more columns / calculations in your report have a heading of '" & HttpUtility.UrlDecode(columnItemHeaderToCheck.Heading) & "'. " & "Column headings must be unique.")
                     breakNestedLoop = True
                     Exit For
                  End If
               Next
               If breakNestedLoop Then
                  Exit For
               End If
            Next

            If objModel.IsSummary AndAlso objModel.Columns.Where(Function(m) m.IsAverage OrElse m.IsCount OrElse m.IsTotal).LongCount() = 0 Then
               ModelState.AddModelError("IsSummaryOK", "There are no columns defined as aggregates for this summary report.")
            End If

            ' Validate Value On Change and Suppress Repeated Values checkboxes i.e. not checked if column is Hidden.
            If objModel.SortOrdersString IsNot Nothing Then
               If objModel.SortOrdersString.Length > 0 Then
                  objModel.SortOrders = deserializer.Deserialize(Of List(Of SortOrderViewModel))(objModel.SortOrdersString)
               End If

               For Each columnItem As ReportColumnItem In objModel.Columns
                  If columnItem.IsHidden = True Then
                     For Each sortorderitem In objModel.SortOrders
                        If sortorderitem.SuppressRepeated = True And columnItem.Name = sortorderitem.Name Then
                           ModelState.AddModelError("IsHidddenParmCorrect", "The column '" & columnItem.Name & "' has 'Suppress Repeated Values' ticked on the Sort Order tab. <br/><br/>Hidden columns can not have 'Suppress Repeated Values' or 'Value On Change' ticked.")
                           breakNestedLoop = True
                           Exit For
                        End If

                        If sortorderitem.ValueOnChange = True And columnItem.Name = sortorderitem.Name Then
                           ModelState.AddModelError("IsHidddenParmCorrect", "The column '" & columnItem.Name & "' has 'Value On Change' ticked on the Sort Order tab.   <br/><br/>Hidden columns can not have 'Suppress Repeated Values' or 'Value On Change' ticked.")
                           breakNestedLoop = True
                           Exit For
                        End If
                     Next
                  End If
                  If breakNestedLoop Then
                     Exit For
                  End If
               Next
            End If

            If objModel.IgnoreZerosForAggregates AndAlso objModel.Columns.Where(
                Function(m) (m.DataType = ColumnDataType.sqlInteger OrElse m.DataType = ColumnDataType.sqlNumeric) AndAlso (m.IsAverage OrElse m.IsCount OrElse m.IsTotal)
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

         ' If above all validation passed then check if all selected columns are hidden, if yes then save of defination is not allowed
         If ModelState.IsValid Then
            If (hiddenColumnsCount = objModel.Columns.Count) Then
               ModelState.AddModelError("AreAllColumnsHidden", "This definition cannot be saved as all columns / calculations selected are defined as hidden.")
            End If
         End If

         If objModel.ValidityStatus = ReportValidationStatus.ServerCheckComplete Then
            objReportRepository.SaveReportDefinition(objModel)

            Session("utilid") = objModel.ID
            Return RedirectToAction("Defsel", "Home")


         Else
            If ModelState.IsValid Then
               objSaveWarning = objReportRepository.ServerValidate(objModel)
            Else
               objSaveWarning = ModelState.ToWebMessage
            End If

            Return Json(objSaveWarning, JsonRequestBehavior.AllowGet)

         End If

      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
      Function util_def_mailmerge(objModel As MailMergeModel) As ActionResult

         Dim objSaveWarning As SaveWarningModel
         Dim deserializer = New JavaScriptSerializer()

         objModel.Dependencies = objReportRepository.RetrieveDependencies(objModel.ID, UtilityType.utlMailMerge)
         objModel.UploadTemplate = CType(objReportRepository.RetrieveParent(objModel.ID, UtilityType.utlMailMerge), MailMergeModel).UploadTemplate

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

         If objModel.Columns.Count > 250 Then
            objSaveWarning = New SaveWarningModel With {
               .ReportType = objModel.ReportType,
               .ID = objModel.ID,
               .ErrorCode = ReportValidationStatus.InvalidOnClient,
               .ErrorMessage = "A maximum of 250 columns are allowed for your mail merge."}
            Return Json(objSaveWarning, JsonRequestBehavior.AllowGet)
         End If

         If objModel.ValidityStatus = ReportValidationStatus.ServerCheckComplete Then

            objReportRepository.SaveReportDefinition(objModel)
            Session("utilid") = objModel.ID
            Return RedirectToAction("Defsel", "Home")

         Else

            If ModelState.IsValid Then
               objSaveWarning = objReportRepository.ServerValidate(objModel)
            Else
               objSaveWarning = ModelState.ToWebMessage
            End If

            Return Json(objSaveWarning, JsonRequestBehavior.AllowGet)

         End If

      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
      Function util_def_talentreport(objModel As TalentReportModel) As ActionResult

         Dim objSaveWarning As SaveWarningModel
         Dim deserializer = New JavaScriptSerializer()

         objModel.Dependencies = objReportRepository.RetrieveDependencies(objModel.ID, UtilityType.TalentReport)

         If objModel.ColumnsAsString IsNot Nothing Then
            If objModel.ColumnsAsString.Length > 0 Then
               objModel.Columns = deserializer.Deserialize(Of List(Of ReportColumnItem))(objModel.ColumnsAsString)
            End If

            '------------------------------------------------------------------

            ' Check the column heading has value.
            For Each columnItem As ReportColumnItem In objModel.Columns
               If String.IsNullOrEmpty(columnItem.Heading.Trim()) And columnItem.IsHidden = False Then
                  ModelState.AddModelError("IsColumnHeaderEmpty", "The '" & columnItem.Name & "' column has a blank heading.")
                  Exit For
               End If
            Next

            ' Check the column headings are unique.
            Dim breakNestedLoop As Boolean
            For Each columnItem As ReportColumnItem In objModel.Columns
               For Each columnItemHeaderToCheck As ReportColumnItem In objModel.Columns

                  If columnItem.ID <> columnItemHeaderToCheck.ID AndAlso UCase(columnItem.Heading.Trim()) = UCase(columnItemHeaderToCheck.Heading.Trim()) AndAlso columnItemHeaderToCheck.IsHidden = False Then
                     ModelState.AddModelError("IsColumnHeaderUnique", "One or more columns in your report have a heading of '" & HttpUtility.UrlDecode(columnItemHeaderToCheck.Heading) & "'. " & "Column headings must be unique.")
                     breakNestedLoop = True
                     Exit For
                  End If
               Next
               If breakNestedLoop Then
                  Exit For
               End If
            Next

            '------------------------------------------------------------------

         End If

         If objModel.SortOrdersString IsNot Nothing Then
            If objModel.SortOrdersString.Length > 0 Then
               objModel.SortOrders = deserializer.Deserialize(Of List(Of SortOrderViewModel))(objModel.SortOrdersString)
            End If
         End If

         If (objModel.BaseMinimumRatingColumnID > 0 AndAlso objModel.MatchChildRatingColumnID = 0) Then
            ModelState.AddModelError("IsRatingsOK", "Actual rating should be selected if minimum rating is selected.")
         End If

         If objModel.BaseMinimumRatingColumnID = 0 AndAlso objModel.MatchChildRatingColumnID > 0 Then
            ModelState.AddModelError("IsRatingsEmpty", "Actual rating should be none if minimum rating is not selected.")
         End If

         If objModel.BaseMinimumRatingColumnID = 0 AndAlso objModel.BasePreferredRatingColumnID > 0 Then
            ModelState.AddModelError("IsPreferredRatingsEmpty", "Preferred rating should be none if minimum rating is not selected.")
         End If

         'Check if role and person table match column has same datatype
         If (objModel.BaseChildColumnDataType > 0 AndAlso objModel.MatchChildColumnDataType > 0 AndAlso objModel.BaseChildColumnDataType <> objModel.MatchChildColumnDataType) Then
            ModelState.AddModelError("IsColumnSelectionOK", "Role match column and Person match column datatype should be same.")
         End If

         If objModel.BaseChildTableID > 0 AndAlso objModel.MatchChildTableID > 0 Then

            Dim matchReport = New MatchReportRun
            matchReport.SessionInfo = CType(Session("SessionContext"), SessionInfo)
            Dim blnChildOf1, blnChildOf2, isTableHasSameParent As Boolean
            Dim isBaseAndMatchChildTableAreSame = (objModel.MatchChildTableID = objModel.BaseChildTableID)

            'Gets the table name
            Dim matchChildTableName = matchReport.SessionInfo.Tables.GetById(objModel.MatchChildTableID).Name
            Dim baseChildTableName = matchReport.SessionInfo.Tables.GetById(objModel.BaseChildTableID).Name
            Dim baseTableName = matchReport.SessionInfo.Tables.GetById(objModel.BaseTableID).Name
            Dim matchTableName = matchReport.SessionInfo.Tables.GetById(objModel.MatchTableID).Name

            blnChildOf1 = matchReport.IsAChildOf((objModel.BaseChildTableID), objModel.BaseTableID)
            blnChildOf2 = matchReport.IsAChildOf((objModel.BaseChildTableID), objModel.MatchTableID)

            ' Check if base child table is child of both base and match table OR Not.
            If blnChildOf1 AndAlso blnChildOf2 Then
               isTableHasSameParent = True
               ModelState.AddModelError("IsBaseTableSelectionOK", "Cannot use the '" & baseChildTableName & "' table as it is a child table of both the '" & baseTableName & "' and the '" & matchTableName & "' tables.")
            End If

            blnChildOf1 = matchReport.IsAChildOf((objModel.MatchChildTableID), objModel.BaseTableID)
            blnChildOf2 = matchReport.IsAChildOf((objModel.MatchChildTableID), objModel.MatchTableID)

            ' Check if match child table is child of both base and match table OR Not.
            If blnChildOf1 AndAlso blnChildOf2 Then
               isTableHasSameParent = True
               ModelState.AddModelError("IsMatchTableSelectionOK", "Cannot use the '" & matchChildTableName & "' table as it is a child table of both the '" & baseTableName & "' and the '" & matchTableName & "' tables.")
            End If

            ' If both base & match child table selected are same then show only one warning message instead of showing same message twice.
            If isBaseAndMatchChildTableAreSame AndAlso isTableHasSameParent Then
               ModelState.Remove("IsBaseTableSelectionOK")
            End If

         End If

         If objModel.ValidityStatus = ReportValidationStatus.ServerCheckComplete Then

            objReportRepository.SaveReportDefinition(objModel)
            Session("utilid") = objModel.ID
            Return RedirectToAction("Defsel", "Home")

         Else

            If ModelState.IsValid Then
               objSaveWarning = objReportRepository.ServerValidate(objModel)
            Else
               objSaveWarning = ModelState.ToWebMessage
            End If

            Return Json(objSaveWarning, JsonRequestBehavior.AllowGet)

         End If

      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
      Function util_def_crosstab(objModel As CrossTabModel) As ActionResult

         Dim objSaveWarning As SaveWarningModel
         objModel.Dependencies = objReportRepository.RetrieveDependencies(objModel.ID, UtilityType.utlCrossTab)

         If objModel.ValidityStatus = ReportValidationStatus.ServerCheckComplete Then
            objReportRepository.SaveReportDefinition(objModel)
            Session("utilid") = objModel.ID
            Return RedirectToAction("DefSel", "Home")

         Else

            If ModelState.IsValid Then
               objSaveWarning = objReportRepository.ServerValidate(objModel)
            Else
               objSaveWarning = ModelState.ToWebMessage
            End If

            Return Json(objSaveWarning, JsonRequestBehavior.AllowGet)

         End If

      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
      Function util_def_9boxgrid(objModel As NineBoxGridModel) As ActionResult

         Dim objSaveWarning As SaveWarningModel
         objModel.Dependencies = objReportRepository.RetrieveDependencies(objModel.ID, UtilityType.utlNineBoxGrid)

         If objModel.ValidityStatus = ReportValidationStatus.ServerCheckComplete Then
            objReportRepository.SaveReportDefinition(objModel)
            Session("utilid") = objModel.ID
            Return RedirectToAction("DefSel", "Home")


         Else

            If ModelState.IsValid Then
               objSaveWarning = objReportRepository.ServerValidate(objModel)
            Else
               objSaveWarning = ModelState.ToWebMessage
            End If

            Return Json(objSaveWarning, JsonRequestBehavior.AllowGet)

         End If

      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
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
            Session("utilid") = objModel.ID
            Return RedirectToAction("DefSel", "Home")

         Else

            If ModelState.IsValid Then
               objSaveWarning = objReportRepository.ServerValidate(objModel)
               DoesSortColumnsMatchToReflectGroupByDescription(objModel, objSaveWarning)
            Else
               objSaveWarning = ModelState.ToWebMessage
            End If

            Return Json(objSaveWarning, JsonRequestBehavior.AllowGet)
         End If

      End Function

      <HttpGet>
      Function GetAvailableColumnsForTable(TableID As Integer) As JsonResult
         Dim objResults = objReportRepository.GetColumnsForTable(TableID)

         objResults.RemoveAll(Function(m) m.IsExpression OrElse (m.DataType = ColumnDataType.sqlOle Or m.DataType = ColumnDataType.sqlVarBinary))
         Return Json(objResults, JsonRequestBehavior.AllowGet)

      End Function

      <HttpGet>
      Function GetAvailableCharacterLookupsForTable(TableID As Integer) As JsonResult

         Dim objResults = objReportRepository.GetAvailableCharacterLookupsForTable(TableID)
         Return Json(objResults, JsonRequestBehavior.AllowGet)

      End Function

      <HttpGet>
      Function GetAvailableItemsForTable(TableID As Integer, reportID As Integer, reportType As UtilityType, selectionType As String) As JsonResult

         Dim objReport = objReportRepository.RetrieveParent(reportID, reportType)
         Dim objAvailable As List(Of ReportColumnItem)

         If selectionType = "C" Then
            objAvailable = objReportRepository.GetColumnsForTable(TableID)
            objAvailable.RemoveAll(Function(m) m.IsExpression OrElse (m.DataType = ColumnDataType.sqlOle Or m.DataType = ColumnDataType.sqlVarBinary))
         Else
            objAvailable = objReportRepository.GetCalculationsForTable(TableID)
            objAvailable.RemoveAll(Function(m) Not m.IsExpression)
         End If

         Dim results = New With {.total = 1, .page = 1, .records = 0, .rows = objAvailable}
         Return Json(results, JsonRequestBehavior.AllowGet)

      End Function

      <HttpGet>
      Function GetBaseTables(reportType As UtilityType) As JsonResult

         Dim objTables = objReportRepository.GetTables(reportType)
         Return Json(objTables, JsonRequestBehavior.AllowGet)

      End Function

      <HttpGet>
      Function GetChildTables(parentTableId As Integer) As JsonResult

         Dim objTables = objReportRepository.GetChildTables(parentTableId, False)
         Return Json(objTables, JsonRequestBehavior.AllowGet)

      End Function


      <HttpPost>
      <ValidateAntiForgeryToken>
      Function AddChildTable(ReportID As Integer) As ActionResult

         Dim objModel As New ChildTableViewModel With {.ReportID = ReportID, .ReportType = UtilityType.utlCustomReport}
         Dim objReport = CType(objReportRepository.RetrieveParent(objModel), CustomReportModel)

         objModel.AvailableTables = objReportRepository.GetChildTables(objReport.BaseTableID, False)

         For Each objTable In objReport.ChildTables
            objModel.AvailableTables.RemoveAll(Function(m) m.id = objTable.TableID)
         Next

         If objReport.ChildTables.Any() Then
            objModel.ID = objReport.ChildTables.Max(Function(m) m.ID) + 1
         Else
            objModel.ID = 1
         End If

         objModel.IsAdd = True

         Return PartialView("EditorTemplates\ReportChildTable", objModel)


      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
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
      <ValidateAntiForgeryToken>
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
      <ValidateAntiForgeryToken>
      Function AddCalendarEvent(ReportID As Integer) As ActionResult

         Dim objReport = objReportRepository.RetrieveCalendarReport(ReportID)

         Dim objModel As New CalendarEventDetailViewModel

         objModel.ID = 0
         objModel.TableID = objReport.BaseTableID
         objModel.ReportID = ReportID

         If objReport.Events.Any() Then
            objModel.EventKey = objReport.Events.Max(Function(m) m.EventKey) + 1
         Else
            objModel.EventKey = 0
         End If

         objModel.AvailableTables = objReportRepository.GetTablesWithEvents(objReport.BaseTableID)

         ModelState.Clear()
         Return PartialView("EditorTemplates\CalendarEventDetail", objModel)


      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
      Function EditCalendarEvent(objModel As CalendarEventDetailViewModel) As ActionResult

         Dim objReport = objReportRepository.RetrieveCalendarReport(objModel.ReportID)
         objModel.AvailableTables = objReportRepository.GetTablesWithEvents(objReport.BaseTableID)

         ModelState.Clear()
         Return PartialView("EditorTemplates\CalendarEventDetail", objModel)
      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
      Sub PostCalendarEvent(objModel As CalendarEventDetailViewModel)

         Dim objReport = objReportRepository.RetrieveCalendarReport(objModel.ReportID)
         Dim original = objReport.Events.Where(Function(m) m.EventKey = objModel.EventKey).FirstOrDefault

         If original IsNot Nothing Then
            objReport.Events.Remove(original)
         End If

         objReport.Events.Add(objModel)

      End Sub

      <HttpPost>
      <ValidateAntiForgeryToken>
      Function ChangeEventBaseTable(objModel As CalendarEventDetailViewModel) As ActionResult

         Dim objReport = objReportRepository.RetrieveCalendarReport(objModel.ReportID)

         objModel.ChangeBaseTable()
         objModel.AvailableTables = objReportRepository.GetTablesWithEvents(objReport.BaseTableID)

         ModelState.Clear()
         Return PartialView("EditorTemplates\CalendarEventDetail", objModel)

      End Function

      <HttpPost>
      Function ChangeEventLookupTable(objModel As CalendarEventDetailViewModel) As ActionResult 'No ValidateAntiForgeryToken necessary for this method: it's never invoked!

         Dim objReport = objReportRepository.RetrieveCalendarReport(objModel.ReportID)
         objModel.AvailableTables = objReportRepository.GetChildTables(objReport.BaseTableID, True)
         Return PartialView("EditorTemplates\CalendarEventDetail", objModel)

      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
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
      <ValidateAntiForgeryToken>
      Function ChangeBaseTable(ReportID As Integer, ReportType As UtilityType, BaseTableID As Integer) As JsonResult

         Dim iChildTablesAvailable As Integer

         Dim objReport = objReportRepository.RetrieveParent(New ReportColumnItem With {.ReportID = ReportID, .ReportType = ReportType})

         objReport.BaseTableID = BaseTableID
         objReport.SetBaseTable(BaseTableID)

         If ReportType = UtilityType.utlCustomReport Then
            iChildTablesAvailable = CType(objReport, CustomReportModel).ChildTablesAvailable
         End If

         Dim result = New With {.childTablesAvailable = iChildTablesAvailable, .sortOrdersAvailable = objReport.SortOrdersAvailable}
         Return Json(result, JsonRequestBehavior.AllowGet)

      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
      Function AddSortOrder(ReportID As Integer, ReportType As UtilityType) As ActionResult

         Dim objModel As New SortOrderViewModel

         objModel.ReportID = ReportID
         objModel.ReportType = ReportType
         objModel.IsNew = True

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
      <ValidateAntiForgeryToken>
      Function EditSortOrder(objModel As SortOrderViewModel) As ActionResult

         Dim objReport = objReportRepository.RetrieveParent(objModel)
         objModel.AvailableColumns = objReport.GetAvailableSortColumns(objModel)

         ModelState.Clear()
         Return PartialView("EditorTemplates\SortOrder", objModel)
      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
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
      <ValidateAntiForgeryToken>
      Sub RemoveSortOrder(objModel As SortOrderViewModel)

         Dim objReport As IReport
         objReport = objReportRepository.RetrieveParent(objModel)

         Dim original = objReport.SortOrders.Where(Function(m) m.ID = objModel.ID).FirstOrDefault

         If original IsNot Nothing Then
            objReport.SortOrders.Remove(original)
         End If

      End Sub

      <HttpPost>
      <ValidateAntiForgeryToken>
      Sub AddAllReportColumns(objModel As ReportColumnCollection)

         Dim objReport As ReportBaseModel
         objReport = CType(objReportRepository.RetrieveParent(objModel), ReportBaseModel)
         Dim objAllObjects As List(Of ReportColumnItem)

         If objModel.SelectionType = "C" Then
            objAllObjects = objReportRepository.GetColumnsForTable(objModel.ColumnsTableID)
         Else
            objAllObjects = objReportRepository.GetCalculationsForTable(objModel.ColumnsTableID)
         End If

         For Each ObjectID In objModel.Columns
            Dim objColumn = objAllObjects.First(Function(m) m.ID = ObjectID)

            'Concatenate table name and column name, if the column is not the calculated column
            If objColumn.IsExpression = False Then
               objColumn.Name = objModel.TableName + "." + objColumn.Name
            End If

            If objReport.ReportType = UtilityType.TalentReport Then
               objColumn.TableID = objModel.ColumnsTableID
            End If

            objReport.Columns.Add(objColumn)
         Next

      End Sub

      <HttpPost>
      <ValidateAntiForgeryToken>
      Sub AddReportColumn(objModel As ReportColumnItem)

         Dim objReport As ReportBaseModel
         objReport = CType(objReportRepository.RetrieveParent(objModel), ReportBaseModel)

         objReport.Columns.Add(objModel)

      End Sub

      <HttpPost>
      <ValidateAntiForgeryToken>
      Sub RemoveAllChildTables(objModel As ReportColumnItem)

         Dim objReport As CustomReportModel
         objReport = CType(objReportRepository.RetrieveParent(objModel), CustomReportModel)

         For Each objChildTable In objReport.ChildTables

            'Remove sort columns
            For Each iColumnID In objReport.Columns.Where(Function(m) m.TableID = objChildTable.TableID)
               objReport.SortOrders.RemoveAll(Function(m) m.ColumnID = iColumnID.ID)
            Next

            objReport.Columns.RemoveAll(Function(m) m.TableID = objChildTable.TableID)
         Next

         objReport.ChildTables.Clear()

      End Sub

      <HttpPost>
      <ValidateAntiForgeryToken>
      Sub RemoveChildTable(objModel As ReportColumnItem)

         Dim objReport As CustomReportModel
         objReport = CType(objReportRepository.RetrieveParent(objModel), CustomReportModel)

         objReport.ChildTables.RemoveAll(Function(m) m.ID = objModel.ID)

         'Remove sort columns
         For Each iColumnID In objReport.Columns.Where(Function(m) m.TableID = objModel.TableID)
            objReport.SortOrders.RemoveAll(Function(m) m.ColumnID = iColumnID.ID)
         Next

         objReport.Columns.RemoveAll(Function(m) m.TableID = objModel.TableID)

      End Sub

      <HttpPost>
      <ValidateAntiForgeryToken>
      Sub RemoveReportColumn(objModel As ReportColumnCollection)

         Dim objReport As ReportBaseModel
         objReport = CType(objReportRepository.RetrieveParent(objModel), ReportBaseModel)

         For Each iColumnID In objModel.Columns
            objReport.Columns.RemoveAll(Function(m) m.ID = iColumnID)
            objReport.SortOrders.RemoveAll(Function(m) m.ColumnID = iColumnID)
         Next

      End Sub

      <HttpPost>
      <ValidateAntiForgeryToken>
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

      ''' <summary>
      ''' Validates that the description columns and sort order columns does match when Group by Description is ticked.
      ''' </summary>
      ''' <param name="objModel">The Model</param>
      ''' <param name="objSaveWarning">The save warning object</param>
      Private Sub DoesSortColumnsMatchToReflectGroupByDescription(objModel As CalendarReportModel, objSaveWarning As SaveWarningModel)

         ' Validate only if group by description is checked and calculation description is not selected
         If (objModel.GroupByDescription = True AndAlso objModel.Description3ID = 0) Then
            Dim descriptionColumnsCount As Integer = 0

            If objModel.Description1ID > 0 Then
               'check if description column1 with id exist into sort order, if yes increment the count
               If objModel.SortOrders.Exists(Function(f) f.ColumnID = objModel.Description1ID) Then
                  descriptionColumnsCount += 1
               End If
            End If

            If objModel.Description2ID > 0 Then
               'check if description column2 with id exist into sort order, if yes increment the count
               If objModel.SortOrders.Exists(Function(f) f.ColumnID = objModel.Description2ID) Then
                  descriptionColumnsCount += 1
               End If
            End If

            ' Validates sort order columns count does match with the selected descriptions
            If objModel.SortOrders.Count() <> descriptionColumnsCount Then
               objSaveWarning.ErrorCode = ReportValidationStatus.Overwrite
               objSaveWarning.ErrorMessage = "The sort order does not reflect the selected Group By Description columns.<BR/><BR/> Are you sure you wish to continue ?"
            End If

         End If
      End Sub

      Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
         target = value
         Return value
      End Function

      <HttpPost()>
      <ValidateAntiForgeryToken>
      Function util_def_mailmerge_submittemplate(TemplateFile As HttpPostedFileBase, MailMergeId As Integer) As ActionResult
         Try

            Dim objReport As MailMergeModel
            objReport = CType(objReportRepository.RetrieveParent(MailMergeId, UtilityType.utlMailMerge), MailMergeModel)

            If Not TemplateFile Is Nothing Then

               Dim acceptedTypes As New List(Of String)(New String() {
                     "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                     "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
                     "application/msword"})

               If acceptedTypes.Contains(TemplateFile.ContentType) Then

                  ' Read input stream from request
                  objReport.UploadTemplate = New Byte(CInt(TemplateFile.InputStream.Length - 1)) {}
                  Dim offset As Integer = 0
                  Dim cnt As Integer = 0
                  While (InlineAssignHelper(cnt, TemplateFile.InputStream.Read(objReport.UploadTemplate, offset, 10))) > 0
                     offset += cnt
                  End While

                  objReport.UploadTemplateName = Path.GetFileName(TemplateFile.FileName)

               Else
                  Return New HttpStatusCodeResult(400, "Please select a Microsoft Word document or template file")

               End If

            End If

         Catch ex As Exception
            Session("ErrorTitle") = "File upload"
            Session("ErrorText") = "You could not upload the template file because of the following error:<p>" & FormatError(ex.Message)
         End Try

      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
      Public Function util_def_mailmerge_downloadtemplate(MailMergeId As Integer) As FilePathResult

         Dim objReport As MailMergeModel
         objReport = CType(objReportRepository.RetrieveParent(MailMergeId, UtilityType.utlMailMerge), MailMergeModel)

         Dim downloadTokenValue As String = Request("download_token_value_id")

         Try
            '  Dim template = CType(objRow.Item(0), Byte())
            '     Dim fileName = objRow.Item("TemplateName").ToString

            ' Download the file
            Response.ContentType = "application/octet-stream"
            Response.Clear()
            Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client
            Response.AddHeader("Content-Disposition", String.Format("attachment;filename=""{0}""", objReport.UploadTemplateName))
            Response.OutputStream.Write(objReport.UploadTemplate, 0, objReport.UploadTemplate.Length)
            Response.End()
            Response.Flush()

         Catch ex As Exception

         End Try

      End Function



      <HttpPost>
      <ValidateAntiForgeryToken>
      Sub RemoveSelectedTableColumns(objModel As ReportColumnCollection)

         Dim objReport As ReportBaseModel
         objReport = CType(objReportRepository.RetrieveParent(objModel), ReportBaseModel)

         Dim objAllObjects As List(Of ReportColumnItem)
         objAllObjects = objReportRepository.GetColumnsForTable(objModel.ColumnsTableID)

         For Each iColumnID In objReport.Columns
            Dim objColumn = objAllObjects.Where(Function(m) m.ID = iColumnID.ID).FirstOrDefault
            If (Not IsNothing(objColumn)) Then
               Dim sortColumn = objReport.SortOrders.Where(Function(m) m.ColumnID = objColumn.ID).FirstOrDefault
               objReport.SortOrders.Remove(sortColumn)
            End If
         Next
         objReport.Columns.RemoveAll(Function(m) m.TableID = objModel.ColumnsTableID)

      End Sub

      <HttpPost>
      <ValidateAntiForgeryToken>
      Function changePersonTable(ReportID As Integer, ReportType As UtilityType, MatchTableID As Integer, BaseTableID As Integer) As JsonResult

         Dim objReport As TalentReportModel

         Dim iChildTablesAvailable As Integer

         objReport = CType(objReportRepository.RetrieveParent(ReportID, ReportType), TalentReportModel)
         objReport.MatchTableID = MatchTableID
         objReport.BaseTableID = BaseTableID
         objReport.SetBaseTable(BaseTableID)

         Dim result = New With {.childTablesAvailable = iChildTablesAvailable, .sortOrdersAvailable = objReport.SortOrdersAvailable}
         Return Json(result, JsonRequestBehavior.AllowGet)

      End Function

      <HttpGet>
      Function util_def_organisationreport() As ActionResult

         Dim iReportID As Integer = CInt(Session("utilid"))
         Dim iAction = ActionToUtilityAction(Session("action").ToString)
         Dim objModel = objReportRepository.LoadOrganisationReport(iReportID, iAction)

         Return View(objModel)

      End Function


      <HttpPost>
      <ValidateAntiForgeryToken>
      Function util_def_organisationreport(objModel As OrganisationReportModel) As ActionResult

         Dim objSaveWarning As SaveWarningModel
         Dim deserializer = New JavaScriptSerializer()

         objModel.Dependencies = objReportRepository.RetrieveDependencies(objModel.ID, UtilityType.OrgReporting)

         If objModel.ColumnsAsString IsNot Nothing Then
            If objModel.ColumnsAsString.Length > 0 Then
               objModel.Columns = deserializer.Deserialize(Of List(Of ReportColumnItem))(objModel.ColumnsAsString)
            End If

            'Only 50 columns selections are allowed
            If objModel.Columns.Count > 50 Then
               ModelState.AddModelError("IsSelectedColumnsCountValid", "A maximum of 50 columns are allowed for your organisation report.")
            End If

            'Check if base view is selected or not
            If objModel.BaseViewID = 0 Then
               ModelState.AddModelError("IsBaseViewSelected", "A base view must be selected.")
            End If

            For Each columnItem As ReportColumnItem In objModel.Columns
                    ' Check the column font size has value in range.
                    If (columnItem.FontSize < 6 OrElse columnItem.FontSize > 30 OrElse columnItem.FontSize = Nothing) Then
                        ModelState.AddModelError("IsFontSizeEmpty", "The '" & columnItem.Name & "' column does not have valid font size.")
                    End If

                    ' Check the column height has value in range.
                    If ((columnItem.DataType = -3) AndAlso (columnItem.Height < 3 OrElse columnItem.Height > 6 OrElse columnItem.Height = Nothing)) Then
                        ModelState.AddModelError("IsHeightEmpty", "The '" & columnItem.Name & "' column does not have valid Height (Rows).")
                    End If

                    If ((columnItem.DataType <> -3) AndAlso (columnItem.Height < 1 OrElse columnItem.Height > 6 OrElse columnItem.Height = Nothing)) Then
                        ModelState.AddModelError("IsHeightEmpty", "The '" & columnItem.Name & "' column does not have valid Height (Rows).")
                    End If
                Next

            End If

            If objModel.FilterColumnsAsString IsNot Nothing Then
            If objModel.FilterColumnsAsString.Length > 0 Then
               objModel.FiltersFieldList = deserializer.Deserialize(Of List(Of OrganisationReportFilterItem))(objModel.FilterColumnsAsString)
            End If
         End If

         If objModel.ValidityStatus = ReportValidationStatus.ServerCheckComplete Then

            objReportRepository.SaveReportDefinition(objModel)
            Session("utilid") = objModel.ID
            Return RedirectToAction("Defsel", "Home")

         Else

            If ModelState.IsValid Then
               objSaveWarning = objReportRepository.ServerValidate(objModel)
            Else
               objSaveWarning = ModelState.ToWebMessage
            End If

            Return Json(objSaveWarning, JsonRequestBehavior.AllowGet)

         End If

      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
      Sub AddReportFilters(objOrgModel As OrganisationReportFilterItem)

         Dim objReport As OrganisationReportModel
         objReport = CType(objReportRepository.RetrieveParent(objOrgModel), OrganisationReportModel)

         objReport.FiltersFieldList.Add(objOrgModel)

      End Sub

      <HttpGet>
      Function GetFilterColumns(ViewID As Integer) As JsonResult
         Dim objResults = objReportRepository.GetFilterColumns(ViewID)
         objResults.RemoveAll(Function(m) (m.FieldDataType = ColumnDataType.sqlOle Or m.FieldDataType = ColumnDataType.sqlVarBinary))
         Return Json(objResults, JsonRequestBehavior.AllowGet)

      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
      Sub ChangeBaseView(ReportID As Integer, ReportType As UtilityType, BaseViewId As Integer)

         Dim objReport = objReportRepository.RetrieveOrganisationReport(ReportID)
         'Remove previous selected base view
         objReport.AllAvailableViewList.RemoveAll(Function(x) x.id = objReport.BaseViewID)

         Dim availableViewList As New List(Of ReportTableItem)
         availableViewList.Add(objReport.BaseViewList.FirstOrDefault(Function(x) x.id = BaseViewId))
         availableViewList.AddRange(objReport.AllAvailableViewList)
         objReport.AllAvailableViewList = availableViewList

         'Set new baseview id
         objReport.BaseViewID = BaseViewId
      End Sub

      <HttpGet>
      Function GetAllAvailableViews(ReportID As Integer) As JsonResult
         Dim objReport = objReportRepository.RetrieveOrganisationReport(ReportID)
         Return Json(objReport.AllAvailableViewList, JsonRequestBehavior.AllowGet)
      End Function

      <HttpGet>
      Function GetAvailableItemsForView(ReportID As Integer, viewOrTableId As Integer, Optional IsTable As Boolean = False) As JsonResult

         Dim objReport = objReportRepository.RetrieveOrganisationReport(ReportID)
         Dim objResults As New List(Of ReportColumnItem)

         ' Based on selected view or table fetch all permitted columns associated to it.
         If (IsTable) Then
            objResults = objReportRepository.GetColumnsForTable(viewOrTableId)
         Else
            Dim getViewId As Integer = If(viewOrTableId > 0, viewOrTableId, (objReport.BaseViewList.First()).id)
            objResults = objReportRepository.GetViewFilterColumns(getViewId)
         End If

         For Each item As ReportColumnItem In objReport.Columns
            If item IsNot Nothing Then
               If IsTable Then
                  objResults.RemoveAll(Function(m) m.ID = item.ID)
               Else
                  objResults.Remove(item)
               End If

            End If
         Next

         Dim results = New With {.total = 1, .page = 1, .records = 0, .rows = objResults}
         Return Json(results, JsonRequestBehavior.AllowGet)

      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
      Sub AddOrganisationReportColumn(objModel As ReportColumnItem)

         Dim objReport As OrganisationReportModel
         objReport = CType(objReportRepository.RetrieveParent(objModel), OrganisationReportModel)
         objReport.Columns.Add(objModel)

      End Sub

      'Add all available organisation report columns to selected columns
      <HttpPost>
      <ValidateAntiForgeryToken>
      Sub AddAllOrganisationReportColumn(objModel As ReportColumnCollection, viewId As Integer, Optional IsTable As Boolean = False)

         Dim objReport As OrganisationReportModel
         objReport = CType(objReportRepository.RetrieveParent(objModel), OrganisationReportModel)
         Dim objAllObjects As List(Of ReportColumnItem)

         If (IsTable) Then
            objAllObjects = objReportRepository.GetColumnsForTable(viewId)
         Else
            Dim getViewId As Integer = If(viewId > 0, viewId, (objReport.BaseViewList.First()).id)
            objAllObjects = objReportRepository.GetViewFilterColumns(getViewId)
         End If

         For Each ObjectID In objModel.Columns

            Dim objColumn = objAllObjects.First(Function(m) m.ID = ObjectID)

            If objColumn IsNot Nothing Then
               objReport.Columns.Add(objColumn)
            End If
         Next

      End Sub

      'Remove all selected organisation report column from selected columns
      <HttpPost>
      <ValidateAntiForgeryToken>
      Sub RemoveOrganisationReportColumn(objModel As ReportColumnCollection)

         Dim objReport As ReportBaseModel
         objReport = CType(objReportRepository.RetrieveParent(objModel), ReportBaseModel)

         For Each iColumnID In objModel.Columns
            objReport.Columns.RemoveAll(Function(m) m.ID = iColumnID)
         Next

      End Sub

      'Remove all available organisation report columns from selected columns
      <HttpPost>
      <ValidateAntiForgeryToken>
      Sub RemoveAllOrganisationReportColumns(objModel As ReportColumnItem)

         Dim objReport As ReportBaseModel
         objReport = CType(objReportRepository.RetrieveParent(objModel), ReportBaseModel)
         objReport.Columns.Clear()

      End Sub

      <HttpPost>
      <ValidateAntiForgeryToken>
      Function ValidateSelectedColumn(ReportID As Integer, GridData As String) As JsonResult

         Dim isValidSelection = True

         Try
            If GridData IsNot Nothing Then
               If GridData.Length > 0 Then
                  Dim deserializer = New JavaScriptSerializer()
                  Dim columns = deserializer.Deserialize(Of List(Of ReportColumnItem))(GridData)

                  ' Check the column font size has value in range.
                  For Each columnItem As ReportColumnItem In columns
                            If (columnItem.FontSize < 6 OrElse columnItem.FontSize > 30 OrElse columnItem.FontSize = Nothing) Then
                                isValidSelection = False
                                Response.StatusCode = System.Net.HttpStatusCode.BadRequest
                                Dim data = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = "Error", .ErrorMessage = "The '" & columnItem.Name & "' column does not have valid font size."}
                                Return Json(data)
                                Exit For
                            End If

                            If ((columnItem.DataType = -3) AndAlso (columnItem.Height < 3 OrElse columnItem.Height > 6 OrElse columnItem.Height = Nothing)) Then
                                isValidSelection = False
                                Response.StatusCode = System.Net.HttpStatusCode.BadRequest
                                Dim data = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = "Error", .ErrorMessage = "The '" & columnItem.Name & "' column does not have valid Height (Rows)."}
                                Return Json(data)
                                Exit For
                            End If

                            If ((columnItem.DataType <> -3) AndAlso (columnItem.Height < 1 OrElse columnItem.Height > 6 OrElse columnItem.Height = Nothing)) Then
                                isValidSelection = False
                                Response.StatusCode = System.Net.HttpStatusCode.BadRequest
                                Dim data = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = "Error", .ErrorMessage = "The '" & columnItem.Name & "' column does not have valid Height (Rows)."}
                                Return Json(data)
                                Exit For
                            End If
                        Next

               End If
            End If

         Catch ex As Exception
            Response.StatusCode = System.Net.HttpStatusCode.BadRequest
            Dim data = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = "Error", .ErrorMessage = ex.Message}
            Return Json(data)
         End Try

         Return Json(isValidSelection)
      End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
      Function ShowPreviewPopup(ReportID As Integer, GridData As String) As ActionResult

         Try

            Dim objReport = objReportRepository.RetrieveOrganisationReport(ReportID)
            Dim objModel As New OrganisationReportModel

            If GridData IsNot Nothing Then
               If GridData.Length > 0 Then
                  Dim deserializer = New JavaScriptSerializer()
                  Dim columns = deserializer.Deserialize(Of List(Of ReportColumnItem))(GridData)
                  objModel.PreviewColumnList = objReport.ProcessColumnsForPreview(columns)
                  objModel.BaseViewID = objReport.BaseViewID
                  objModel.PostBasedTableId = objReport.PostBasedTableId
               End If
            End If
            Return PartialView("_PreviewOrganisation", objModel)
         Catch ex As Exception

         End Try
      End Function

   End Class

End Namespace