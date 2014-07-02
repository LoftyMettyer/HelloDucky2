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

		Private _reports As New Collection(Of ReportBaseModel)

		Private objSessionInfo As SessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
		Private objDataAccess As clsDataAccess = CType(HttpContext.Current.Session("DatabaseAccess"), clsDataAccess)

		Private _Username As String = HttpContext.Current.Session("username").ToString
		'		Private _Action As String = HttpContext.Current.Session("action").ToString


		'  Public Property BaseTable As Table
		Public Function LoadCustomReport(ID As Integer, bIsCopy As Boolean, Action As String) As CustomReportModel

			Dim objModel As New CustomReportModel

			Dim prmErrMsg = New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmName = New SqlParameter("psReportName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmOwner = New SqlParameter("psReportOwner", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmDescription = New SqlParameter("psReportDesc", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmBaseTableID = New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmAllRecords = New SqlParameter("pfAllRecords", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmPicklistID = New SqlParameter("piPicklistID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmPicklistName = New SqlParameter("psPicklistName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmPicklistHidden = New SqlParameter("pfPicklistHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmFilterID = New SqlParameter("piFilterID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmFilterName = New SqlParameter("psFilterName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmFilterHidden = New SqlParameter("pfFilterHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmParent1TableID = New SqlParameter("piParent1TableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmParent1TableName = New SqlParameter("psParent1Name", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmParent1FilterID = New SqlParameter("piParent1FilterID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmParent1FilterName = New SqlParameter("psParent1FilterName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmParent1FilterHidden = New SqlParameter("pfParent1FilterHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmParent2TableID = New SqlParameter("piParent2TableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmParent2TableName = New SqlParameter("psParent2Name", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmParent2FilterID = New SqlParameter("piParent2FilterID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmParent2FilterName = New SqlParameter("psParent2FilterName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmParent2FilterHidden = New SqlParameter("pfParent2FilterHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}

			Dim prmSummary = New SqlParameter("pfSummary", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmPrintFilterHeader = New SqlParameter("pfPrintFilterHeader", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputPreview = New SqlParameter("pfOutputPreview", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputFormat = New SqlParameter("piOutputFormat", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmOutputScreen = New SqlParameter("pfOutputScreen", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputPrinter = New SqlParameter("pfOutputPrinter", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputPrinterName = New SqlParameter("psOutputPrinterName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmOutputSave = New SqlParameter("pfOutputSave", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputSaveExisting = New SqlParameter("piOutputSaveExisting", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmOutputEmail = New SqlParameter("pfOutputEmail", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputEmailAddr = New SqlParameter("piOutputEmailAddr", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmOutputEmailName = New SqlParameter("psOutputEmailName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmOutputEmailSubject = New SqlParameter("psOutputEmailSubject", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmOutputEmailAttachAs = New SqlParameter("psOutputEmailAttachAs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmOutputFilename = New SqlParameter("psOutputFilename", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmTimestamp = New SqlParameter("piTimestamp", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmParent1AllRecords = New SqlParameter("pfParent1AllRecords", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmParent1PicklistID = New SqlParameter("piParent1PicklistID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmParent1PicklistName = New SqlParameter("psParent1PicklistName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmParent1PicklistHidden = New SqlParameter("pfParent1PicklistHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmParent2AllRecords = New SqlParameter("pfParent2AllRecords", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmParent2PicklistID = New SqlParameter("piParent2PicklistID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmParent2PicklistName = New SqlParameter("psParent2PicklistName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmParent2PicklistHidden = New SqlParameter("pfParent2PicklistHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmInfo = New SqlParameter("psInfoMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmIgnoreZeros = New SqlParameter("pfIgnoreZeros", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}

			Try

				'' TODO -- tidy these up.
				Dim lngAction = HttpContext.Current.Session("action")
				Dim sUserName = HttpContext.Current.Session("username")

				'' TODO - tidy up this proc to return a dataset instead of millions of bloody parameters!
				Dim rstDefinition As DataSet = objDataAccess.GetDataSet("sp_ASRIntGetReportDefinition" _
					, New SqlParameter("piReportID", SqlDbType.Int) With {.Value = CInt(ID)} _
					, New SqlParameter("psCurrentUser", SqlDbType.VarChar, 255) With {.Value = sUserName} _
					, New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = lngAction} _
					, prmErrMsg, prmName, prmOwner, prmDescription, prmBaseTableID, prmAllRecords, prmPicklistID, prmPicklistName, prmPicklistHidden _
					, prmFilterID, prmFilterName, prmFilterHidden _
					, prmParent1TableID, prmParent1TableName, prmParent1FilterID, prmParent1FilterName, prmParent1FilterHidden _
					, prmParent2TableID, prmParent2TableName, prmParent2FilterID, prmParent2FilterName, prmParent2FilterHidden _
					, prmSummary, prmPrintFilterHeader, prmOutputPreview, prmOutputFormat, prmOutputScreen, prmOutputPrinter, prmOutputPrinterName _
					, prmOutputSave, prmOutputSaveExisting _
					, prmOutputEmail, prmOutputEmailAddr, prmOutputEmailName, prmOutputEmailSubject, prmOutputEmailAttachAs _
					, prmOutputFilename, prmTimestamp, prmParent1AllRecords, prmParent1PicklistID, prmParent1PicklistName, prmParent1PicklistHidden _
					, prmParent2AllRecords, prmParent2PicklistID, prmParent2PicklistName, prmParent2PicklistHidden _
					, prmInfo, prmIgnoreZeros)

				'	 Use Dapper!
				'	

				'PopulateDefintion(objModel, rstDefinition.Tables(1))
				objModel.BaseTableID = CInt(prmBaseTableID.Value)
				objModel.BaseTables = GetTables()

				objModel.Name = prmName.Value.ToString
				objModel.Description = prmDescription.Value.ToString
				objModel.Owner = prmOwner.ToString
				objModel.FilterID = CInt(prmFilterID.Value)
				objModel.PicklistID = CInt(prmPicklistID.Value)

				Select Case CBool(prmAllRecords.Value)
					Case True
						objModel.SelectionType = RecordSelectionType.AllRecords
					Case Else
						If objModel.FilterID > 0 Then
							objModel.SelectionType = RecordSelectionType.Filter
						Else
							objModel.SelectionType = RecordSelectionType.Picklist
						End If

				End Select

				objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCustomReport, ID, bIsCopy)

				objModel.Parent1.ID = CInt(prmParent1TableID.Value)
				objModel.Parent1.Name = prmParent1TableName.Value.ToString
				objModel.Parent1.FilterID = CInt(prmParent1FilterID.Value)
				objModel.Parent1.PicklistID = CInt(prmParent1PicklistID.Value)

				objModel.Parent2.ID = CInt(prmParent2TableID.Value)
				objModel.Parent2.Name = prmParent2TableName.Value.ToString
				objModel.Parent2.FilterID = CInt(prmParent2FilterID.Value)
				objModel.Parent2.PicklistID = CInt(prmParent2PicklistID.Value)

				' '' TODO - Load columns tab (needs dynamic based on table selection)
				'objModel.Columns.Available = GetColumnsForTable(objModel.BaseTableID)
				objModel.Columns.BaseTableID = objModel.BaseTableID

				'' todo replace with dapper
				objModel.Columns.Selected = New Collection(Of ReportColumnItem)
				Dim rstChildColumns = objDataAccess.GetDataTable("SELECT * FROM ASRSysCustomReportsDetails WHERE CustomReportID = " & ID, CommandType.Text)
				For Each objRow As System.Data.DataRow In rstChildColumns.Rows
					Dim objItem As New ReportColumnItem() With {
						.CustomReportId = ID,
						.Heading = objRow("Heading").ToString,
						.id = CInt(objRow("ColExprID")),
						.Name = "TODO-lookup columnid-" & objRow("ColExprID").ToString,
						.Sequence = CInt(objRow("Sequence")),
						.Size = CInt(objRow("size")),
						.Decimals = CInt(objRow("dp")),
						.IsAverage = CBool(objRow("avge")),
						.IsCount = CBool(objRow("cnt")),
						.IsTotal = CBool(objRow("tot")),
						.IsHidden = CBool(objRow("hidden")),
						.IsGroupWithNext = CBool(objRow("groupwithnextcolumn"))}
					objModel.Columns.Selected.Add(objItem)

					If CInt(objRow("SortOrderSequence")) > 0 Then
						Dim objSortItem As New ReportSortItem() With {
							.ColumnID = CInt(objRow("colexprid")),
							.BreakOnChange = CBool(objRow("BOC")),
							.ValueOnChange = CBool(objRow("VOC")),
							.PageOnChange = CBool(objRow("POC")),
							.Sequence = CInt(objRow("SortOrderSequence")),
							.Order = objRow("sortorder").ToString,
							.SuppressRepeated = CBool(objRow("repetition"))}
						objModel.SortOrderColumns.Add(objSortItem)
					End If

				Next

				objModel.Output.Format = CType(prmOutputFormat.Value, OutputFormats)
				objModel.Output.SendAsEmail = CBool(prmOutputEmail.Value)
				objModel.Output.EmailAttachAs = prmOutputEmailAttachAs.Value.ToString

				Dim rstChildTables = objDataAccess.GetDataTable("sp_ASRIntGetReportChilds", CommandType.StoredProcedure _
					, New SqlParameter("piReportID", SqlDbType.Int) With {.Value = ID})

				' TODO - replace with dapper
				For Each objRow As DataRow In rstChildTables.Rows
					objModel.ChildTables.Add(New ReportChildTables() With {
									.TableName = objRow("table").ToString,
									.FilterName = objRow("filter").ToString,
									.OrderName = objRow("order").ToString,
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

			Catch ex As Exception
				Throw

			End Try

			Return objModel

		End Function

		' TODO
		Public Function LoadMailMerge(ID As Integer, bIsCopy As Boolean, Action As String) As MailMergeModel

			Dim objModel As New MailMergeModel
			Dim objItem As ReportColumnItem
			Dim objSort As ReportSortItem

			Dim prmErrMsg = New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmName = New SqlParameter("psReportName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmOwner = New SqlParameter("psReportOwner", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmDescription = New SqlParameter("psReportDesc", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmBaseTableID = New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmSelection = New SqlParameter("piSelection", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmPicklistID = New SqlParameter("piPicklistID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmPicklistName = New SqlParameter("psPicklistName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmPicklistHidden = New SqlParameter("pfPicklistHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmFilterID = New SqlParameter("piFilterID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmFilterName = New SqlParameter("psFilterName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmFilterHidden = New SqlParameter("pfFilterHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputFormat = New SqlParameter("piOutputFormat", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmOutputSave = New SqlParameter("pfOutputSave", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputFileName = New SqlParameter("psOutputFileName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmEmailAddrID = New SqlParameter("piEmailAddrID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmEmailSubject = New SqlParameter("psEmailSubject", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmTemplateFileName = New SqlParameter("psTemplateFileName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmOutputScreen = New SqlParameter("pfOutputScreen", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmEmailAsAttachment = New SqlParameter("pfEmailAsAttachment", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmEmailAttachmentName = New SqlParameter("psEmailAttachmentName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmSuppressBlanks = New SqlParameter("pfSuppressBlanks", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmPauseBeforeMerge = New SqlParameter("pfPauseBeforeMerge", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputPrinter = New SqlParameter("pfOutputPrinter", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputPrinterName = New SqlParameter("psOutputPrinterName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmDocumentMapID = New SqlParameter("piDocumentMapID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmManualDocManHeader = New SqlParameter("pfManualDocManHeader", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmTimestamp = New SqlParameter("piTimestamp", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmWarningMsg = New SqlParameter("psWarningMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

			Dim rstDefinition = objDataAccess.GetDataSet("spASRIntGetMailMergeDefinition" _
				, New SqlParameter("@piReportID", SqlDbType.Int) With {.Value = ID} _
				, New SqlParameter("@psCurrentUser", SqlDbType.VarChar, 255) With {.Value = _Username} _
				, New SqlParameter("@psAction", SqlDbType.VarChar, 255) With {.Value = Action} _
				, prmErrMsg, prmName, prmOwner, prmDescription, prmBaseTableID _
				, prmSelection, prmPicklistID, prmPicklistName, prmPicklistHidden _
				, prmFilterID, prmFilterName, prmFilterHidden _
				, prmOutputFormat, prmOutputSave, prmOutputFileName, prmEmailAddrID, prmEmailSubject _
				, prmTemplateFileName, prmOutputScreen, prmEmailAsAttachment, prmEmailAttachmentName, prmSuppressBlanks _
				, prmPauseBeforeMerge, prmOutputPrinter, prmOutputPrinterName _
				, prmDocumentMapID, prmManualDocManHeader, prmTimestamp, prmWarningMsg)

			' Columns
			For Each objRow As DataRow In rstDefinition.Tables(0).Rows
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
			For Each objRow As DataRow In rstDefinition.Tables(1).Rows
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
			For Each objRow As DataRow In rstDefinition.Tables(2).Rows
				objSort = New ReportSortItem
				objSort.TableID = CInt(objRow("tableid"))
				objSort.ColumnID = CInt(objRow("columnId"))
				objSort.ColumnName = objRow("columnname").ToString
				objSort.Order = objRow("sortorder").ToString
				objSort.Sequence = CInt(objRow("sequence"))
				objModel.SortOrderColumns.Add(objSort)
			Next

			Dim sSQL = String.Format("SELECT * FROM ASRSysMailMergeName WHERE MailMergeID = {0}", ID)
			Dim dtDefinition As DataTable = objDataAccess.GetDataTable(sSQL, CommandType.Text)

			objModel.FilterName = prmFilterName.Value.ToString
			objModel.PicklistName = prmPicklistName.Value.ToString

			PopulateDefintion(objModel, dtDefinition)
			objModel.GroupAccess = GetUtilityAccess(UtilityType.utlMailMerge, ID, bIsCopy)

			'objModel.Columns.Available = GetColumnsForTable(objModel.BaseTableID)
			objModel.Columns.BaseTableID = objModel.BaseTableID

			If dtDefinition.Rows.Count = 1 Then

				Dim row As DataRow = dtDefinition.Rows(0)

				objModel.TemplateName = row("TemplateFileName").ToString()
				objModel.OutputFormat = CType(row("OutputFormat"), MailMergeOutputTypes)
				objModel.SendToPrinter = CBool(row("OutputPrinter"))
				objModel.PrinterName = row("OutputPrinterName").ToString()
				objModel.SaveTofile = CBool(row("OutputSave"))
				objModel.Filename = row("OutputFileName").ToString
				objModel.SendAsEmail = (objModel.OutputFormat = 1)
				objModel.EmailGroupID = CInt(row("EmailAddrID"))
				objModel.Subject = row("EmailSubject").ToString()
				objModel.SendAsAttachment = CBool(row("EmailAsAttachment"))
				objModel.AttachAs = row("EmailAttachmentName").ToString()

				objModel.SuppressBlankLines = CBool(prmSuppressBlanks.Value)
				objModel.PauseBeforeMerge = CBool(prmPauseBeforeMerge.Value)

			End If

			If bIsCopy Then
				objModel.ID = 0
			Else
				objModel.ID = ID
			End If

			_reports.Add(objModel)

			Return objModel

		End Function

		Public Function NewCrossTab() As CrossTabModel

			Dim objModel As New CrossTabModel

			objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCrossTab, 0, False)
			objModel.BaseTables = GetTables()

			Return objModel

		End Function

		Public Function NewCalendarReport() As CalendarReportModel

			Dim objModel As New CalendarReportModel

			objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCalendarReport, 0, False)
			objModel.BaseTables = GetTables()

			Return objModel

		End Function

		Public Function NewCustomReport() As CustomReportModel

			Dim objModel As New CustomReportModel

			objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCustomReport, 0, False)
			objModel.BaseTables = GetTables()

			Return objModel

		End Function

		Public Function NewMailMerge() As MailMergeModel

			Dim objModel As New MailMergeModel

			objModel.GroupAccess = GetUtilityAccess(UtilityType.utlMailMerge, 0, False)
			objModel.BaseTables = GetTables()

			Return objModel

		End Function


		' TODO
		Public Function LoadCrossTab(ID As Integer, bIsCopy As Boolean, Action As String) As CrossTabModel

			Dim objModel As New CrossTabModel

			Try

				Dim prmErrMsg = New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmName = New SqlParameter("psReportName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmOwner = New SqlParameter("psReportOwner", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmDescription = New SqlParameter("psReportDesc", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmBaseTableID = New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmAllRecords = New SqlParameter("pfAllRecords", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmPicklistID = New SqlParameter("piPicklistID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmPicklistName = New SqlParameter("psPicklistName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmPicklistHidden = New SqlParameter("pfPicklistHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmFilterID = New SqlParameter("piFilterID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmFilterName = New SqlParameter("psFilterName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmFilterHidden = New SqlParameter("pfFilterHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmPrintFilter = New SqlParameter("pfPrintFilterHeader", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmHColID = New SqlParameter("HColID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmHStart = New SqlParameter("HStart", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
				Dim prmHStop = New SqlParameter("HStop", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
				Dim prmHStep = New SqlParameter("HStep", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
				Dim prmVColID = New SqlParameter("VColID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmVStart = New SqlParameter("VStart", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
				Dim prmVStop = New SqlParameter("VStop", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
				Dim prmVStep = New SqlParameter("VStep", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
				Dim prmPColID = New SqlParameter("PColID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmPStart = New SqlParameter("PStart", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
				Dim prmPStop = New SqlParameter("PStop", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
				Dim prmPStep = New SqlParameter("PStep", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
				Dim prmIType = New SqlParameter("IType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmIColID = New SqlParameter("IColID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmPercentage = New SqlParameter("Percentage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmPerPage = New SqlParameter("PerPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmSuppress = New SqlParameter("Suppress", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmThousand = New SqlParameter("Thousand", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmOutputPreview = New SqlParameter("pfOutputPreview", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmOutputFormat = New SqlParameter("piOutputFormat", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmOutputScreen = New SqlParameter("pfOutputScreen", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmOutputPrinter = New SqlParameter("pfOutputPrinter", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmOutputPrinterName = New SqlParameter("psOutputPrinterName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmOutputSave = New SqlParameter("pfOutputSave", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmOutputSaveExisting = New SqlParameter("piOutputSaveExisting", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmOutputEmail = New SqlParameter("pfOutputEmail", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmOutputEmailAddr = New SqlParameter("piOutputEmailAddr", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmOutputEmailAddrName = New SqlParameter("psOutputEmailName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmOutputEmailSubject = New SqlParameter("psOutputEmailSubject", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmOutputEmailAttachAs = New SqlParameter("psOutputEmailAttachAs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmOutputFilename = New SqlParameter("psOutputFilename", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmTimestamp = New SqlParameter("piTimestamp", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

				objDataAccess.GetFromSP("sp_ASRIntGetCrossTabDefinition", _
						New SqlParameter("piReportID", SqlDbType.Int) With {.Value = ID}, _
						New SqlParameter("psCurrentUser", SqlDbType.VarChar, 255) With {.Value = _Username}, _
						New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = Action}, _
						prmErrMsg, prmName, prmOwner, prmDescription, prmBaseTableID, _
						prmAllRecords, prmPicklistID, prmPicklistName, prmPicklistHidden, prmFilterID, prmFilterName, prmFilterHidden, _
						prmPrintFilter, prmHColID, prmHStart, prmHStop, prmHStep, prmVColID, prmVStart, prmVStop, prmVStep, prmPColID, _
						prmPStart, prmPStop, prmPStep, prmIType, prmIColID, prmPercentage, prmPerPage, prmSuppress, prmThousand, _
						prmOutputPreview, prmOutputFormat, prmOutputScreen, prmOutputPrinter, prmOutputPrinterName, prmOutputSave, prmOutputSaveExisting, _
						prmOutputEmail, prmOutputEmailAddr, prmOutputEmailAddrName, prmOutputEmailSubject, prmOutputEmailAttachAs, prmOutputFilename, prmTimestamp)

				objModel.FilterName = prmFilterName.Value.ToString
				objModel.PicklistName = prmPicklistName.Value.ToString

				' Definition tab
				' REPLACE WITH dapper!!!!
				Dim sSQL = String.Format("SELECT * FROM ASRSysCrossTab WHERE CrossTabID = {0}", ID)
				Dim dtDefinition As DataTable = objDataAccess.GetDataTable(sSQL, CommandType.Text)

				PopulateDefintion(objModel, dtDefinition)
				objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCrossTab, ID, bIsCopy)

				' Columns tab
				' '' TODO - Load columns tab (needs dynamic based on table selection)
				objModel.AvailableColumns = GetColumnsForTable(objModel.BaseTableID)

				If dtDefinition.Rows.Count = 1 Then

					Dim row As DataRow = dtDefinition.Rows(0)

					objModel.HorizontalID = CInt(row("HorizontalColID"))
					objModel.HorizontalStart = CInt(row("HorizontalStart"))
					objModel.HorizontalStop = CInt(row("HorizontalStop"))
					objModel.HorizontalIncrement = CInt(row("HorizontalStep"))

					objModel.VerticalID = CInt(row("VerticalColID"))
					objModel.VerticalStart = CInt(row("VerticalStart"))
					objModel.VerticalStop = CInt(row("VerticalStop"))
					objModel.VerticalIncrement = CInt(row("VerticalStep"))

					objModel.PageBreakID = CInt(row("PageBreakColID"))
					objModel.PageBreakStart = CInt(row("PageBreakStart"))
					objModel.PageBreakStop = CInt(row("PageBreakStop"))
					objModel.PageBreakIncrement = CInt(row("PageBreakStep"))

					objModel.IntersectionID = CInt(row("IntersectionColID"))
					objModel.IntersectionType = CType(row("IntersectionType"), IntersectionType)
					objModel.PercentageOfType = CBool(row("Percentage"))
					objModel.PercentageOfPage = CBool(row("PercentageofPage"))
					objModel.SuppressZeros = CBool(row("SuppressZeros"))
					objModel.UseThousandSeparators = CBool(row("ThousandSeparators"))

				End If

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

			objModel.Name = "calrep TODO"

			Dim con = objDataAccess.Connection

			'objModel.Events = con.Query(Of CalendarEventDetail)("SELECT * FROM ASRSysCalendarReportEvents")
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

			objDataAccess.ExecuteSP("sp_ASRIntSaveMailMerge" _
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
				, New SqlParameter("@psEmailSubject", SqlDbType.VarChar, -1) With {.Value = objModel.Subject} _
				, New SqlParameter("@psTemplateFileName", SqlDbType.VarChar, -1) With {.Value = objModel.TemplateName} _
				, New SqlParameter("@pfOutputScreen", SqlDbType.Bit) With {.Value = objModel.DisplayOutputOnScreen} _
				, New SqlParameter("@psUserName", SqlDbType.VarChar, 255) With {.Value = _Username} _
				, New SqlParameter("@pfEmailAsAttachment", SqlDbType.Bit) With {.Value = objModel.SendAsAttachment} _
				, New SqlParameter("@psEmailAttachmentName", SqlDbType.VarChar, -1) With {.Value = objModel.AttachAs} _
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


			_reports.Remove(objModel)

			Return True

		End Function

		Public Function SaveReportDefinition(objModel As CrossTabModel) As Boolean

			Try

				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = objModel.ID}

				' TODO old access stuff - needs updating
				Dim psAccess As String = ""	' Request.Form("txtSend_access")}
				Dim psJobsToHide As String = ""	' Request.Form("txtSend_jobsToHide")
				Dim psJobsToHideGroups As String = ""	' Request.Form("txtSend_jobsToHideGroups")}

				psAccess = UtilityAccessAsString(objModel.GroupAccess)

				objDataAccess.ExecuteSP("sp_ASRIntSaveCrossTab", _
						New SqlParameter("psName", SqlDbType.VarChar, 255) With {.Value = objModel.Name}, _
						New SqlParameter("psDescription", SqlDbType.VarChar, -1) With {.Value = objModel.Description}, _
						New SqlParameter("piTableID", SqlDbType.Int) With {.Value = objModel.BaseTableID}, _
						New SqlParameter("piSelection", SqlDbType.Int) With {.Value = objModel.SelectionType}, _
						New SqlParameter("piPicklistID", SqlDbType.Int) With {.Value = objModel.PicklistID}, _
						New SqlParameter("piFilterID", SqlDbType.Int) With {.Value = objModel.FilterID}, _
						New SqlParameter("pfPrintFilter", SqlDbType.Bit) With {.Value = objModel.DisplayTitleInReportHeader}, _
						New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = _Username}, _
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
						New SqlParameter("pfOutputEmail", SqlDbType.Bit) With {.Value = objModel.Output.SendAsEmail}, _
						New SqlParameter("piOutputEmailAddr", SqlDbType.Int) With {.Value = objModel.Output.EmailGroupID}, _
						New SqlParameter("psOutputEmailSubject", SqlDbType.VarChar, -1) With {.Value = objModel.Output.EmailSubject}, _
						New SqlParameter("psOutputEmailAttachAs", SqlDbType.VarChar, -1) With {.Value = objModel.Output.EmailAttachAs}, _
						New SqlParameter("psOutputFilename", SqlDbType.VarChar, -1) With {.Value = objModel.Output.Filename}, _
						New SqlParameter("psAccess", SqlDbType.VarChar, -1) With {.Value = psAccess}, _
						New SqlParameter("psJobsToHide", SqlDbType.VarChar, -1) With {.Value = psJobsToHide}, _
						New SqlParameter("psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = psJobsToHideGroups}, _
						prmID)

			Catch
				Throw

			End Try

			Return True
		End Function

		Public Function SaveReportDefinition(objModel As CustomReportModel) As Boolean

			' save it

			' HACK ALERT! Do this properly!!!
			Try

				Dim sSQL As String = String.Format("UPDATE ASRSysCustomReportsName SET Name = '{1}', OutputEmailAttachAs = '{2}' WHERE ID = {0}" _
																			, objModel.ID, objModel.Name, objModel.Output.EmailAttachAs)
				objDataAccess.ExecuteSql(sSQL)

				sSQL = String.Format("DELETE ASRSysCustomReportsChildDetails WHERE CustomReportID = {0}", objModel.ID)
				objDataAccess.ExecuteSql(sSQL)

				For Each objChild In objModel.ChildTables
					sSQL = String.Format("INSERT ASRSysCustomReportsChildDetails (CustomReportID, ChildTable, ChildFilter, ChildMaxRecords, ChildOrder) VALUES ({0}, {1}, {2}, {3}, {4})", _
															objModel.ID, objChild.TableID, objChild.FilterID, objChild.Records, objChild.OrderID)
					objDataAccess.ExecuteSql(sSQL)
				Next


				sSQL = String.Format("DELETE ASRSysCustomReportAccess WHERE ID = {0}", objModel.ID)
				objDataAccess.ExecuteSql(sSQL)

				For Each objChild In objModel.GroupAccess
					sSQL = String.Format("INSERT ASRSysCustomReportAccess (id, groupname, access) VALUES ({0}, '{1}', '{2}')", _
															objModel.ID, objChild.Name, objChild.Access)
					objDataAccess.ExecuteSql(sSQL)
				Next




			Catch ex As Exception

			End Try

			Return True

		End Function

		Public Function SaveReportDefinition(objModel As CalendarReportModel) As Boolean

			Return True
		End Function

		Private Function GetUtilityAccess(utilType As UtilityType, ID As Integer, IsCopy As Boolean) As Collection(Of GroupAccess)

			Dim objAccess As New Collection(Of GroupAccess)

			Try

				'Dim con = objDataAccess.Connection

				'objAccess = con.Query(Of GroupAccess)("spASRIntGetUtilityAccessRecords")
				'objModel.GroupAccess = GetUtilityAccess(UtilityType.utlCalendarReport, ID, bIsCopy)


				Dim rstAccessInfo As DataTable = objDataAccess.GetDataTable("spASRIntGetUtilityAccessRecords", CommandType.StoredProcedure _
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
		Private Function ReportColumnsAsString(objColumns As Collection(Of ReportColumnItem), objSortColumns As Collection(Of ReportSortItem)) As String

			Dim sColumns As String = ""
			Dim sOrderString As String

			Dim iCount As Integer = 1
			For Each objItem In objColumns

				' this could be improvbe with some linq or whatever! No panic because the whole function could be tidied up
				sOrderString = "||0||"
				For Each objSortItem In objSortColumns
					If objSortItem.ColumnID = objItem.id Then
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

		Private Function GetTables() As Collection(Of SelectListItem)

			Dim objSessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)
			Dim objItems As New Collection(Of SelectListItem)

			For Each objTable In objSessionInfo.Tables
				Dim objItem As New SelectListItem() With {.Value = CStr(objTable.ID), .Text = objTable.Name}
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

					' Lets get dapper in here!
					If data.Columns.Contains("TableID") Then
						outputModel.BaseTableID = CInt(row("TableID"))
					Else
						outputModel.BaseTableID = CInt(row("BaseTable"))
					End If

					outputModel.Name = row("name").ToString
					outputModel.Description = row("description").ToString
					outputModel.Owner = row("username").ToString
					outputModel.FilterID = CInt(row("FilterID"))
					outputModel.PicklistID = CInt(row("PicklistID"))

					'TODO - tidy up with proc pulling back values - probably using dapper
					' Crappy hack as custom reports and calendar reports use a boolean in allrecords, mail merge and cross tabs store the selection type. Rubbish!
					If data.Columns.Contains("Selection") Then
						outputModel.SelectionType = CType(row("Selection"), RecordSelectionType)
					Else
						Dim bAllRecs As Boolean = CBool(row("AllRecords"))
						If bAllRecs Then
							outputModel.SelectionType = RecordSelectionType.AllRecords
						Else
							If outputModel.FilterID > 0 Then
								outputModel.SelectionType = RecordSelectionType.Filter
							Else
								outputModel.SelectionType = RecordSelectionType.Picklist
							End If

						End If

					End If

					If data.Columns.Contains("PrintFilterHeader") Then
						outputModel.DisplayTitleInReportHeader = CBool(row("PrintFilterHeader"))
					End If

				End If

				' Load the base tables
				outputModel.BaseTables = GetTables()

			Catch ex As Exception
				Throw

			End Try

		End Sub


		' can be done with dapper?
		Private Sub PopulateOutput(outputModel As ReportOutputModel, data As DataTable)

			Try

				If data.Rows.Count = 1 Then

					Dim row As DataRow = data.Rows(0)

					outputModel.IsPreview = CBool(row("OutputPreview"))
					outputModel.Format = CType(row("OutputFormat"), OutputFormats)
					outputModel.ToScreen = CBool(row("OutputScreen"))
					outputModel.ToPrinter = CBool(row("OutputPrinter"))
					outputModel.PrinterName = row("OutputPrinterName").ToString()
					outputModel.SaveToFile = CBool(row("OutputSave"))
					outputModel.Filename = row("OutputFileName").ToString
					outputModel.SaveExisting = CType(row("OutputSaveExisting"), ExistingFile)
					outputModel.SendAsEmail = CBool(row("OutputEmail"))
					outputModel.EmailGroupID = CInt(row("OutputEmailAddr"))
					'objModel.Output.EmailAddress = GetEmailGroupName(CInt(objRow("OutputEmailAddr")))
					outputModel.EmailSubject = row("OutputEmailSubject").ToString()
					outputModel.EmailAttachAs = row("OutputEmailAttachAs").ToString()

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