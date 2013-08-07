<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/bundles/utilities_calendarreport_run")%>" type="text/javascript"></script>

<object
	classid="CLSID:8E2F1EF1-3812-4678-A084-16384DE3EA6D"
	codebase="cabs/COAInt_CalRepKey.cab#version=1,0,0,2"
	id="ctlKey"
	name="ctlKey"
	width="0"
	height="0"
	style="VISIBILITY: hidden; width: 0px; height: 0px">
</object>

<%
	Dim icount As Integer
	Dim fok As Boolean
	Dim objCalendar As HR.Intranet.Server.CalendarReport
	Dim fNotCancelled As Boolean
	Dim fBadUtilDef As Boolean
	Dim fNoRecords As Boolean
	Dim blnShowCalendar As Boolean
	'Dim CalRep_UtilID As Integer
	Dim aPrompts
		
	fBadUtilDef = (Session("utiltype") = "") Or _
		 (Session("utilname") = "") Or _
		 (Session("utilid") = "") Or _
		 (Session("action") = "")
	
	fok = Not fBadUtilDef
	fNotCancelled = True
	
	'objCalendar = Nothing
	Session("objCalendar" & Session("UtilID")) = Nothing
	Session("objCalendar" & Session("UtilID")) = ""
	
	If fok Then
		' Create the reference to the DLL (Report Class)
		objCalendar = New HR.Intranet.Server.CalendarReport
				
		' Pass required info to the DLL
		objCalendar.Username = Session("username")
		CallByName(objCalendar, "Connection", CallType.Let, Session("databaseConnection"))
		objCalendar.CalendarReportID = Session("utilid")
		objCalendar.ClientDateFormat = Session("LocaleDateFormat")
		objCalendar.LocalDecimalSeparator = Session("LocaleDecimalSeparator")
		objCalendar.SingleRecordID = Session("singleRecordID")
		
		aPrompts = Session("Prompts_" & Session("utiltype") & "_" & Session("UtilID"))
		If fok Then
			fok = objCalendar.SetPromptedValues(aPrompts)
			fNotCancelled = Response.IsClientConnected
			If fok Then fok = fNotCancelled
		End If

		If fok Then
			fok = objCalendar.GetCalendarReportDefinition
			fNotCancelled = Response.IsClientConnected
			If fok Then fok = fNotCancelled
		End If
		
		If fok Then
			fok = objCalendar.GetEventsCollection
			fNotCancelled = Response.IsClientConnected
			If fok Then fok = fNotCancelled
		End If

		If fok Then
			fok = objCalendar.GetOrderArray
			fNotCancelled = Response.IsClientConnected
			If fok Then fok = fNotCancelled
		End If

		If fok Then
			fok = objCalendar.GenerateSQL
			fNotCancelled = Response.IsClientConnected
			If fok Then fok = fNotCancelled
		End If

		If fok Then
			fok = objCalendar.ExecuteSql
			fNotCancelled = Response.IsClientConnected
			If fok Then fok = fNotCancelled
		End If

		If fok Then
			fok = objCalendar.Initialise_WP_Region
			fNotCancelled = Response.IsClientConnected
			If fok Then fok = fNotCancelled
		End If
		
		objCalendar.SetLastRun()

		fNoRecords = objCalendar.NoRecords
		
		If fok Then
			If Response.IsClientConnected Then
				objCalendar.Cancelled = False
			Else
				objCalendar.Cancelled = True
			End If
		Else
			If Not fNoRecords Then
				If fNotCancelled Then
					objCalendar.FailedMessage = objCalendar.ErrorString
					objCalendar.Failed = True
				Else
					objCalendar.Cancelled = True
				End If
			End If
		End If
		
		blnShowCalendar = (objCalendar.OutputPreview Or (objCalendar.OutputFormat = 0 And objCalendar.OutputScreen))
		
		Session("objCalendar" & Session("UtilID")) = objCalendar

				
	End If

	If fok Then
%>
<input type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">
<input type='hidden' id="txtOK" name="txtOK" value="True">
<%
	Dim objUser As New HR.Intranet.Server.clsSettings
	Dim cmdEmailAddr As Object
		
	Dim arrayDefinition
	Dim arrayColumnsDefinition
	Dim arrayDataDefinition
	Dim arrayStyles
	Dim arrayMerges
	Dim INPUT_VALUE
	Dim prmEmailGroupID As Object
	Dim rstEmailAddr As Object
	Dim sErrorDescription As String
	Dim iLoop As Integer
	Dim sEmailAddresses As String
		
	Session("CalRepUtilID") = Request.Form("utilid")
		
	If blnShowCalendar Then
		Response.Write("<input type='hidden' id=txtPreview name=txtPreview value=1>" & vbCrLf)
	Else
		Response.Write("<input type='hidden' id=txtPreview name=txtPreview value=0>" & vbCrLf)
	End If
		
	If blnShowCalendar Then
%>

<div id="calendarframeset">

	<div id="dataframe" data-framesource="util_run_calendarreport_data" style="display: block;">
		<%Html.RenderPartial("~/views/home/util_run_calendarreport_data.ascx")%>
	</div>

	<div id="navframeset">
		<div id="calendarworkframe" data-framesource="util_run_calendarreport_nav" style="display: block; height: 83px">
			<%Html.RenderPartial("~/views/home/util_run_calendarreport_nav.ascx")%>
		</div>

		<div id="workframefiller" data-framesource="util_run_calendarreport_nav" style="display: block ;font-size: xx-small">
			<%Html.RenderPartial("~/views/home/util_run_calendarreport_navfiller.ascx")%>
		</div>
	</div>

	<div id="calendarframe_calendar" data-framesource="util_run_calendarreport_calendar" style="display: block; overflow: auto; font-size: xx-small; height: 200px">
	 	<%Html.RenderPartial("~/views/home/util_run_calendarreport_calendar.ascx")%>
	</div>
	<div style="height: 10px"></div>
</div>

<div id="optionsframeset" style="height: 130px">
	<div id="calendarframe_key" data-framesource="util_run_calendarreport_key" style="display: block; width: 70%; float: left">
		<%Html.RenderPartial("~/views/home/util_run_calendarreport_key.ascx")%>
	</div>

	<div style="width: 5%; float: left"></div>

	<div id="calendarframe_options" data-framesource="util_run_calendarreport_options" style="display: block; width: 25%; float: left">
		<%Html.RenderPartial("~/views/home/util_run_calendarreport_options.ascx")%>
	</div>

	<table valign="top" align="right" width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
		<tr>
			<td>&nbsp;</td>
			<td width="80">
				<input type="button" id="cmdOutput" name="cmdOutput" value="Output" style="WIDTH: 100%" class="btn"
					onclick="ExportCalendarDataPrompt();"/>
			</td>
			<td width="10"></td>
			<td width="80">
				<input type="button" id="cmdClose" name="cmdClose" value="Close" style="WIDTH: 100%" class="btn"
					onclick="closeclick();"/>
			</td>
			<td width="5">&nbsp;</td>
		</tr>
	</table>
	<input type="hidden" id="txtCalRep_UtilID" name="txtCalRep_UtilID" value='<%=Request("CalRepUtilID")%>'>

</div>

<div id="outputoptions" data-framesource="util_run_outputoptions" style="display: none;">
	<%	Html.RenderPartial("~/Views/Home/util_run_outputoptions.ascx")%>
</div>

<%
Else
	'*****************************************
	'DO THE OUTPUT WITHOUT RUNNING TO PREVIEW
	'*****************************************
	If fok Then
		fok = objCalendar.OutputGridDefinition
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCalendar.OutputGridColumns
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCalendar.OutputReport(True)
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then

		arrayDefinition = objCalendar.OutputArray_Definition
		arrayColumnsDefinition = objCalendar.OutputArray_Columns
		arrayDataDefinition = objCalendar.OutputArray_Data
	End If
%>


<form id="frmOutput" name="frmOutput">
	<table align="center" class="outline" cellpadding="5" cellspacing="0">
		<tr>
			<td>
				<table align="center" class="invisible" cellpadding="0" cellspacing="0">
					<tr>
						<td colspan="3" height="10"></td>
					</tr>
					<tr>
						<td width="20"></td>
						<td align="center" id="tdDisplay">Outputting Calendar Report.&nbsp;Please Wait...
						</td>
						<td width="20"></td>
					</tr>
					<tr>
						<td colspan="3" height="20"></td>
					</tr>
					<tr>
						<td width="20"></td>
						<td align="center">
							<input id="Cancel" style="WIDTH: 80px" type="button" width="80" value="Cancel" name="Cancel" class="btn"
								onclick="closeclick();" />
						</td>
						<td width="20"></td>
					</tr>
					<tr>
						<td colspan="5" height="10"></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
		id="grdCalendarOutput"
		name="grdCalendarOutput"
		codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
		style="HEIGHT: 0px; VISIBILITY: visible; WIDTH: 0px; display: block"
		width="0">
		<%
			For icount = 1 To UBound(arrayDefinition)
				Response.Write(arrayDefinition(icount))
			Next

			For icount = 1 To UBound(arrayColumnsDefinition)
				Response.Write(arrayColumnsDefinition(icount))
			Next
				
			For icount = 1 To UBound(arrayDataDefinition)
				Response.Write(arrayDataDefinition(icount))
			Next

		%>
	</object>

	<%
				
		If fok Then
			arrayStyles = objCalendar.OutputArray_Styles
			arrayMerges = objCalendar.OutputArray_Merges
		End If
				
		'************************* START OF HIDDEN GRID ******************************
		Response.Write("<OBJECT classid=""clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1""    codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6"" id=ssHiddenGrid name=ssHiddenGrid style=""visibility: visible; display: block; HEIGHT: 0px; WIDTH: 0px"" WIDTH=0 HEIGHT=0>" & vbCrLf)
		Response.Write("	<PARAM NAME=""ScrollBars"" VALUE=""4"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""_Version"" VALUE=""196617"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Cols"" VALUE=""1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""GroupHeaders"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""ColumnHeaders"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""GroupHeadLines"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""HeadLines"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Col.Count"" VALUE=""1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""DividerType"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""BevelColorFrame"" VALUE=""-2147483642"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""BevelColorHighlight"" VALUE=""-2147483628"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""BevelColorShadow"" VALUE=""-2147483632"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""BevelColorFace"" VALUE=""-2147483633"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
		Response.Write("	<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""SelectTypeRow"" VALUE=""2"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""SelectByCell"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""RowNavigation"" VALUE=""1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""MaxSelectedRows"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
		Response.Write("	<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
		Response.Write("	<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""BackColorEven"" VALUE=""16777215"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""BackColorOdd"" VALUE=""16777215"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""RowHeight"" VALUE=""238"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
		Response.Write("	<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns.Count"" VALUE=""1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).Width"" VALUE=""1000"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).Visible"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).Columns.Count"" VALUE=""1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).Caption"" VALUE=""PageBreak"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).Name"" VALUE=""PageBreak"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).Alignment"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).CaptionAlignment"" VALUE=""2"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).Bound"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).AllowSizing"" VALUE=""1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).DataField"" VALUE=""Column 0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).DataType"" VALUE=""8"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).Level"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).NumberFormat"" VALUE="""">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).Case"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).FieldLen"" VALUE=""4096"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).VertScrollBar"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).Locked"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).Style"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).ButtonsAlways"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).RowCount"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).ColCount"" VALUE=""1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).HasForeColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).HasBackColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).HeadForeColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).HeadBackColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).ForeColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).BackColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).HeadStyleSet"" VALUE="""">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).StyleSet"" VALUE="""">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).Nullable"" VALUE=""1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).Mask"" VALUE="""">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).PromptInclude"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).ClipMode"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Columns(0).PromptChar"" VALUE=""95"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""BatchUpdate"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""_ExtentX"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""_ExtentY"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
		Response.Write("	<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""BackColor"" VALUE=""16777215"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)

		Response.Write("</OBJECT>" & vbCrLf)

		'***************************** END OF HIDDEN GRID **************************************				
		Response.Write("<INPUT type='hidden' id=txtCalendarPageCount name=txtCalendarPageCount value=" & UBound(arrayMerges) & ">" & vbCrLf)
		
	%>
</form>

<%

	If fok Then
						
		Dim iPage As Integer
		Dim iStyle As Integer
		Dim iMerge As Integer
		Dim arrayPageStyles
		Dim arrayPageMerges
			
		For iPage = 0 To UBound(arrayMerges)
			arrayPageMerges = arrayMerges(iPage)
			Response.Write("<form id=frmCalendarMerge_" & iPage & " name=frmCalendarMerge_" & iPage & " style=""visibility:hidden;display:none"">" & vbCrLf)
			For iMerge = 0 To UBound(arrayPageMerges)
				INPUT_VALUE = arrayPageMerges(iMerge)
				Response.Write("	<INPUT type=hidden name=Merge_" & iPage & "_" & iMerge & " ID=Merge_" & iPage & "_" & iMerge & " VALUE=""" & INPUT_VALUE & """>" & vbCrLf)
			Next
			Response.Write("</form>" & vbCrLf)
		Next

		For iPage = 0 To UBound(arrayStyles)
			arrayPageStyles = arrayStyles(iPage)
			Response.Write("<form id=frmCalendarStyle_" & iPage & " name=frmCalendarStyle_" & iPage & " style=""visibility:hidden;display:none"">" & vbCrLf)
			For iStyle = 0 To UBound(arrayPageStyles)
				INPUT_VALUE = arrayPageStyles(iStyle)
				Response.Write("	<INPUT type=hidden name=Style_" & iPage & "_" & iStyle & " ID=Style_" & iPage & "_" & iStyle & " VALUE=""" & INPUT_VALUE & """>" & vbCrLf)
			Next
			Response.Write("</form>" & vbCrLf)
		Next
	End If
			
	If fok Then
		objCalendar.OutputArray_Clear()
	End If

	'Write the function that Outputs the report to the Output Classes in the Client DLL.
	Response.Write("<script type=""text/javascript"">" & vbCrLf)

	Response.Write("function outputCalendarReport() " & vbCrLf)
	Response.Write("	{" & vbCrLf & vbCrLf)
				
	Response.Write("	var lngPageColumnCount = 3;" & vbCrLf)
	Response.Write("  var lngActualRow = new Number(0);" & vbCrLf)
	Response.Write("  var blnSettingsDone = false;" & vbCrLf)
	Response.Write("	var sColHeading = new String(''); " & vbCrLf)
	Response.Write("	var iColDataType = new Number(12); " & vbCrLf)
	Response.Write("	var iColDecimals = new Number(0); " & vbCrLf)
	Response.Write("  var blnNewPage = false;" & vbCrLf)
	Response.Write("  var lngPageCount = new Number(0);" & vbCrLf)

	Response.Write("  var strType = new String('');" & vbCrLf)
	Response.Write("  var lngStartCol = new Number(0);" & vbCrLf)
	Response.Write("  var lngStartRow = new Number(0);" & vbCrLf)
	Response.Write("  var lngEndCol = new Number(0);" & vbCrLf)
	Response.Write("  var lngEndRow = new Number(0);" & vbCrLf)
	Response.Write("  var lngBackCol = new Number(0);" & vbCrLf)
	Response.Write("  var lngForeCol = new Number(0);" & vbCrLf)
	Response.Write("  var blnBold = false;" & vbCrLf)
	Response.Write("  var blnUnderline = false;" & vbCrLf)
	Response.Write("  var blnGridlines = false;" & vbCrLf)
	
	CallByName(objUser, "Connection", CallType.Let, Session("databaseConnection"))
			
	Response.Write("  window.parent.ASRIntranetOutput.UserName = """ & CleanStringForJavaScript(Session("Username")) & """;" & vbCrLf)
	Response.Write("  window.parent.ASRIntranetOutput.SaveAsValues = """ & CleanStringForJavaScript(Session("OfficeSaveAsValues")) & """;" & vbCrLf)

	Response.Write("  frmMenuFrame = window.parent.parent.opener.window.parent.frames(""menuframe"");" & vbCrLf)
		
	Response.Write("	window.parent.ASRIntranetOutput.SettingOptions(")
	Response.Write("""" & CleanStringForJavaScript(objUser.GetUserSetting("Output", "WordTemplate", "")) & """, ")
	Response.Write("""" & CleanStringForJavaScript(objUser.GetUserSetting("Output", "ExcelTemplate", "")) & """, ")

	If (objUser.GetUserSetting("Output", "ExcelGridlines", "0") = "1") Then
		Response.Write("true, ")
	Else
		Response.Write("false, ")
	End If

	If (objUser.GetUserSetting("Output", "ExcelHeaders", "0") = "1") Then
		Response.Write("true, ")
	Else
		Response.Write("false, ")
	End If

	If (objUser.GetUserSetting("Output", "AutoFitCols", "1") = "1") Then
		Response.Write("true, ")
	Else
		Response.Write("false, ")
	End If

	If (objUser.GetUserSetting("Output", "Landscape", "1") = "1") Then
		Response.Write("true, " & vbCrLf)
	Else
		Response.Write("false, " & vbCrLf)
	End If
				 
	Response.Write("frmMenuFrame.document.all.item(""txtSysPerm_EMAILGROUPS_VIEW"").value);" & vbCrLf)

	Response.Write("  window.parent.ASRIntranetOutput.SettingLocations(")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleCol", "3")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleRow", "2")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataCol", "2")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataRow", "4")) & ");" & vbCrLf)

	Response.Write("  window.parent.ASRIntranetOutput.SettingTitle(")
	If (objUser.GetUserSetting("Output", "TitleGridLines", "0") = "1") Then
		Response.Write("true, ")
	Else
		Response.Write("false, ")
	End If

	If (objUser.GetUserSetting("Output", "TitleBold", "1") = "1") Then
		Response.Write("true, ")
	Else
		Response.Write("false, ")
	End If

	If (objUser.GetUserSetting("Output", "TitleUnderline", "0") = "1") Then
		Response.Write("true, ")
	Else
		Response.Write("false, ")
	End If

	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleForecolour", "6697779")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215"))) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "TitleForecolour", "6697779"))) & ");" & vbCrLf)

	Response.Write("window.parent.ASRIntranetOutput.SettingHeading(")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingGridLines", "1")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingBold", "1")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingUnderline", "0")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingBackcolour", "16248553")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingForecolour", "6697779")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "HeadingBackcolour", "16248553"))) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "HeadingForecolour", "6697779"))) & ");" & vbCrLf)

	Response.Write("window.parent.ASRIntranetOutput.SettingData(")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataGridLines", "1")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataBold", "0")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataUnderline", "0")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataBackcolour", "15988214")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataForecolour", "6697779")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "DataBackcolour", "15988214"))) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "DataForecolour", "6697779"))) & ");" & vbCrLf)
			
	Dim lngFormat As Long
	Dim blnScreen As Boolean
	Dim blnPrinter As Boolean
	Dim strPrinterName As String
	Dim blnSave As Boolean
	Dim lngSaveExisting As Long
	Dim blnEmail As Boolean
	Dim lngEmailGroupID As Long
	Dim strEmailSubject As String
	Dim strEmailAttachAs As String
	Dim strFileName As String
			
	lngFormat = CleanStringForJavaScript(objCalendar.OutputFormat)
	blnScreen = CleanStringForJavaScript(LCase(objCalendar.OutputScreen))
	blnPrinter = CleanStringForJavaScript(LCase(objCalendar.OutputPrinter))
	strPrinterName = CleanStringForJavaScript(objCalendar.OutputPrinterName)
	blnSave = CleanStringForJavaScript(LCase(objCalendar.OutputSave))
	lngSaveExisting = CleanStringForJavaScript(objCalendar.OutputSaveExisting)
	blnEmail = CleanStringForJavaScript(LCase(objCalendar.OutputEmail))
	lngEmailGroupID = CLng(objCalendar.OutputEmailID)
	strEmailSubject = CleanStringForJavaScript(objCalendar.OutputEmailSubject)
	strEmailAttachAs = CleanStringForJavaScript(objCalendar.OutputEmailAttachAs)
	strFileName = CleanStringForJavaScript(objCalendar.OutputFilename)

	If (blnEmail) And (lngEmailGroupID > 0) Then
			
		cmdEmailAddr = CreateObject("ADODB.Command")
		cmdEmailAddr.CommandText = "spASRIntGetEmailGroupAddresses"
		cmdEmailAddr.CommandType = 4 ' Stored procedure
		cmdEmailAddr.ActiveConnection = Session("databaseConnection")

		prmEmailGroupID = cmdEmailAddr.CreateParameter("EmailGroupID", 3, 1) ' 3=integer, 1=input
		cmdEmailAddr.Parameters.Append(prmEmailGroupID)
		prmEmailGroupID.value = CleanNumeric(lngEmailGroupID)

		Err.Clear()
		rstEmailAddr = cmdEmailAddr.Execute

		If (Err.Number <> 0) Then
			sErrorDescription = "Error getting the email addresses for group." & vbCrLf & FormatError(Err.Description)
		End If

		If Len(sErrorDescription) = 0 Then
			iLoop = 1
			Do While Not rstEmailAddr.EOF
				If iLoop > 1 Then
					sEmailAddresses = sEmailAddresses & ";"
				End If
				sEmailAddresses = sEmailAddresses & rstEmailAddr.Fields("Fixed").Value
				rstEmailAddr.MoveNext()
				iLoop = iLoop + 1
			Loop
					
			' Release the ADO recordset object.
			rstEmailAddr.close()
		End If
						
		rstEmailAddr = Nothing
		cmdEmailAddr = Nothing
	End If
			
	Response.Write("fok = window.parent.ASRIntranetOutput.SetOptions(false, " & _
																					lngFormat & "," & blnScreen & ", " & _
																					blnPrinter & ",""" & strPrinterName & """, " & _
																					blnSave & "," & lngSaveExisting & ", " & _
																					blnEmail & ", """ & CleanStringForJavaScript(sEmailAddresses) & """, """ & _
																					strEmailSubject & """,""" & strEmailAttachAs & """,""" & strFileName & """);" & vbCrLf)
			
	Response.Write("if (fok == true) {" & vbCrLf)
	If (objCalendar.OutputFormat = 0) And (objCalendar.OutputPrinter) Then
		Response.Write("	window.parent.ASRIntranetOutput.SetPrinter();" & vbCrLf)
		Response.Write("  dataOnlyPrint();" & vbCrLf)
		Response.Write("	window.parent.ASRIntranetOutput.ResetDefaultPrinter();" & vbCrLf)
	Else
		Response.Write("if (window.parent.ASRIntranetOutput.GetFile() == true) " & vbCrLf)
		Response.Write("	{" & vbCrLf)
		Response.Write("	window.parent.ASRIntranetOutput.InitialiseStyles();" & vbCrLf)
		Response.Write("	window.parent.ASRIntranetOutput.ResetStyles();" & vbCrLf)
		Response.Write("	window.parent.ASRIntranetOutput.ResetColumns();" & vbCrLf)
		Response.Write("	window.parent.ASRIntranetOutput.ResetMerges();" & vbCrLf)

		Response.Write("	window.parent.ASRIntranetOutput.HeaderRows = 1;" & vbCrLf)
		Response.Write("	window.parent.ASRIntranetOutput.HeaderCols = 0;" & vbCrLf)
		Response.Write("	window.parent.ASRIntranetOutput.SizeColumnsIndependently = true;" & vbCrLf)
	
		Response.Write("	window.parent.ASRIntranetOutput.ArrayDim((lngPageColumnCount-1), 0);" & vbCrLf & vbCrLf)
		Response.Write("  frmOutput.grdCalendarOutput.focus();")

		Response.Write("  frmOutput.grdCalendarOutput.MoveFirst();" & vbCrLf)
		Response.Write("  for (var lngRow=0; lngRow<frmOutput.grdCalendarOutput.Rows; lngRow++)" & vbCrLf)
		Response.Write("		{" & vbCrLf)
		Response.Write("		bm = frmOutput.grdCalendarOutput.AddItemBookmark(lngRow);" & vbCrLf)
	
	
		Response.Write("		if (lngRow == (frmOutput.grdCalendarOutput.Rows - 1))" & vbCrLf)
		Response.Write("			{" & vbCrLf)
		Response.Write("			sBreakValue = frmOutput.grdCalendarOutput.Columns(1).CellText(bm);" & vbCrLf)
		Response.Write("			if ((sBreakValue == 'Key') && (" & lngFormat & " != 4)) " & vbCrLf)
		Response.Write("				{ " & vbCrLf)
		Response.Write("				window.parent.ASRIntranetOutput.AddPage(replace(frmOutput.grdCalendarOutput.Caption,'&&','&') ,sBreakValue);" & vbCrLf)
		Response.Write("				} " & vbCrLf)
		Response.Write("			else " & vbCrLf)
		Response.Write("				{ " & vbCrLf)
		Response.Write("				window.parent.ASRIntranetOutput.AddPage(replace(frmOutput.grdCalendarOutput.Caption,'&&','&') + ' - ' + sBreakValue,sBreakValue);" & vbCrLf)
		Response.Write("				} " & vbCrLf)

		Response.Write("			var frmMerge = document.forms('frmCalendarMerge_'+lngPageCount);" & vbCrLf)
	
		Response.Write("			var dataCollection = frmMerge.elements;" & vbCrLf)
		Response.Write("			if (dataCollection!=null) " & vbCrLf)
		Response.Write("				{" & vbCrLf)
		Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
		Response.Write("					{" & vbCrLf)
		Response.Write("					strMergeString = dataCollection.item(i).value;" & vbCrLf)
		Response.Write("					if (strMergeString != '')" & vbCrLf)
		Response.Write("						{" & vbCrLf)
		Response.Write("						lngStartCol = Number(mergeArgument(strMergeString,'STARTCOL'));" & vbCrLf)
		Response.Write("						lngStartRow = Number(mergeArgument(strMergeString,'STARTROW'));" & vbCrLf)
		Response.Write("						lngEndCol = Number(mergeArgument(strMergeString,'ENDCOL'));" & vbCrLf)
		Response.Write("						lngEndRow = Number(mergeArgument(strMergeString,'ENDROW'));" & vbCrLf)
		Response.Write("						window.parent.ASRIntranetOutput.AddMerge(lngStartCol,lngStartRow,lngEndCol,lngEndRow);" & vbCrLf)
		Response.Write("						}" & vbCrLf)
		Response.Write("					}" & vbCrLf)
		Response.Write("				}" & vbCrLf)

		Response.Write("			var frmStyle = document.forms('frmCalendarStyle_'+lngPageCount);" & vbCrLf)
		Response.Write("			var dataCollection = frmStyle.elements;" & vbCrLf)
		Response.Write("			if (dataCollection!=null) " & vbCrLf)
		Response.Write("				{" & vbCrLf)
		Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
		Response.Write("					{" & vbCrLf)
		Response.Write("					strStyleString = dataCollection.item(i).value;" & vbCrLf)
		Response.Write("					if (strStyleString != '')" & vbCrLf)
		Response.Write("						{" & vbCrLf)
		Response.Write("						strType = styleArgument(strStyleString,'TYPE');" & vbCrLf)
		Response.Write("						lngStartCol = Number(styleArgument(strStyleString,'STARTCOL'));" & vbCrLf)
		Response.Write("						lngStartRow = Number(styleArgument(strStyleString,'STARTROW'));" & vbCrLf)
		Response.Write("						lngEndCol = Number(styleArgument(strStyleString,'ENDCOL'));" & vbCrLf)
		Response.Write("						lngEndRow = Number(styleArgument(strStyleString,'ENDROW'));" & vbCrLf)
		Response.Write("						lngBackCol = Number(styleArgument(strStyleString,'BACKCOLOR'));" & vbCrLf)
		Response.Write("						lngForeCol = Number(styleArgument(strStyleString,'FORECOLOR'));" & vbCrLf)
		Response.Write("						blnBold = styleArgument(strStyleString,'BOLD');" & vbCrLf)
		Response.Write("						blnUnderline = styleArgument(strStyleString,'UNDERLINE');" & vbCrLf)
		Response.Write("						blnGridlines = styleArgument(strStyleString,'GRIDLINES');" & vbCrLf)
		Response.Write("						window.parent.ASRIntranetOutput.AddStyle(strType,lngStartCol,lngStartRow,lngEndCol,lngEndRow,lngBackCol,lngForeCol,blnBold,blnUnderline,blnGridlines);" & vbCrLf)
		Response.Write("						}" & vbCrLf)
		Response.Write("					}" & vbCrLf)
		Response.Write("				}" & vbCrLf)
	
		Response.Write("			for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)" & vbCrLf)
		Response.Write("				{" & vbCrLf)
		Response.Write("				window.parent.ASRIntranetOutput.AddColumn(sColHeading, iColDataType, iColDecimals, false);" & vbCrLf)
		Response.Write("				}" & vbCrLf)
		Response.Write("			window.parent.ASRIntranetOutput.DataArray();" & vbCrLf)
		Response.Write("			blnBreakCheck = true;" & vbCrLf)
		Response.Write("			sBreakValue = '';" & vbCrLf)
		Response.Write("			lngActualRow = 0;" & vbCrLf)
		Response.Write("			}" & vbCrLf)
	
		Response.Write("    else if ((frmOutput.grdCalendarOutput.Columns(0).CellText(bm) == '*')" & vbCrLf)
		Response.Write("					&& (!blnBreakCheck))" & vbCrLf)
		Response.Write("			{" & vbCrLf)
		Response.Write("			sBreakValue = frmOutput.grdCalendarOutput.Columns(1).CellText(bm);" & vbCrLf)
		Response.Write("			window.parent.ASRIntranetOutput.AddPage(replace(frmOutput.grdCalendarOutput.Caption,'&&','&') + ' - ' + sBreakValue,sBreakValue);" & vbCrLf)
	
		Response.Write("			var frmMerge = document.forms('frmCalendarMerge_'+lngPageCount);" & vbCrLf)
		Response.Write("			var dataCollection = frmMerge.elements;" & vbCrLf)
		Response.Write("			if (dataCollection!=null) " & vbCrLf)
		Response.Write("				{" & vbCrLf)
		Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
		Response.Write("					{" & vbCrLf)
		Response.Write("					strMergeString = dataCollection.item(i).value;" & vbCrLf)
		Response.Write("					if (strMergeString != '')" & vbCrLf)
		Response.Write("						{" & vbCrLf)
		Response.Write("						lngStartCol = Number(mergeArgument(strMergeString,'STARTCOL'));" & vbCrLf)
		Response.Write("						lngStartRow = Number(mergeArgument(strMergeString,'STARTROW'));" & vbCrLf)
		Response.Write("						lngEndCol = Number(mergeArgument(strMergeString,'ENDCOL'));" & vbCrLf)
		Response.Write("						lngEndRow = Number(mergeArgument(strMergeString,'ENDROW'));" & vbCrLf)
		Response.Write("						window.parent.ASRIntranetOutput.AddMerge(lngStartCol,lngStartRow,lngEndCol,lngEndRow);" & vbCrLf)
		Response.Write("						}" & vbCrLf)
		Response.Write("					}" & vbCrLf)
		Response.Write("				}" & vbCrLf)

		Response.Write("			var frmStyle = document.forms('frmCalendarStyle_'+lngPageCount);" & vbCrLf)
		Response.Write("			var dataCollection = frmStyle.elements;" & vbCrLf)
		Response.Write("			if (dataCollection!=null) " & vbCrLf)
		Response.Write("				{" & vbCrLf)
		Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
		Response.Write("					{" & vbCrLf)
		Response.Write("					strStyleString = dataCollection.item(i).value;" & vbCrLf)
		Response.Write("					if (strStyleString != '')" & vbCrLf)
		Response.Write("						{" & vbCrLf)
		Response.Write("						strType = styleArgument(strStyleString,'TYPE');" & vbCrLf)
		Response.Write("						lngStartCol = Number(styleArgument(strStyleString,'STARTCOL'));" & vbCrLf)
		Response.Write("						lngStartRow = Number(styleArgument(strStyleString,'STARTROW'));" & vbCrLf)
		Response.Write("						lngEndCol = Number(styleArgument(strStyleString,'ENDCOL'));" & vbCrLf)
		Response.Write("						lngEndRow = Number(styleArgument(strStyleString,'ENDROW'));" & vbCrLf)
		Response.Write("						lngBackCol = Number(styleArgument(strStyleString,'BACKCOLOR'));" & vbCrLf)
		Response.Write("						lngForeCol = Number(styleArgument(strStyleString,'FORECOLOR'));" & vbCrLf)
		Response.Write("						blnBold = styleArgument(strStyleString,'BOLD');" & vbCrLf)
		Response.Write("						blnUnderline = styleArgument(strStyleString,'UNDERLINE');" & vbCrLf)
		Response.Write("						blnGridlines = styleArgument(strStyleString,'GRIDLINES');" & vbCrLf)
		Response.Write("						window.parent.ASRIntranetOutput.AddStyle(strType,lngStartCol,lngStartRow,lngEndCol,lngEndRow,lngBackCol,lngForeCol,blnBold,blnUnderline,blnGridlines);" & vbCrLf)
		Response.Write("						}" & vbCrLf)
		Response.Write("					}" & vbCrLf)
		Response.Write("				}" & vbCrLf)
	
		Response.Write("			for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)" & vbCrLf)
		Response.Write("				{" & vbCrLf)
		Response.Write("				window.parent.ASRIntranetOutput.AddColumn(sColHeading, iColDataType, iColDecimals, false);" & vbCrLf)
		Response.Write("				}" & vbCrLf)
		Response.Write("      window.parent.ASRIntranetOutput.DataArray();" & vbCrLf)
		Response.Write("			lngPageColumnCount = frmOutput.grdCalendarOutput.Columns.Count;" & vbCrLf)
		Response.Write("			if (!blnSettingsDone)" & vbCrLf)
		Response.Write("				{" & vbCrLf)
		Response.Write("				window.parent.ASRIntranetOutput.HeaderRows = 2;" & vbCrLf)
		Response.Write("				window.parent.ASRIntranetOutput.HeaderCols = 1;" & vbCrLf)
		Response.Write("				window.parent.ASRIntranetOutput.SizeColumnsIndependently = true;" & vbCrLf)
		Response.Write("				blnSettingsDone = true;" & vbCrLf)
		Response.Write("				}" & vbCrLf)
		Response.Write("			window.parent.ASRIntranetOutput.InitialiseStyles();" & vbCrLf)
		Response.Write("			window.parent.ASRIntranetOutput.ResetStyles();" & vbCrLf)
		Response.Write("			window.parent.ASRIntranetOutput.ResetColumns();" & vbCrLf)
		Response.Write("			window.parent.ASRIntranetOutput.ResetMerges();" & vbCrLf)
		Response.Write("			lngPageCount++;" & vbCrLf)
		Response.Write("			window.parent.ASRIntranetOutput.ArrayDim((lngPageColumnCount-1), 0);" & vbCrLf)
		Response.Write("			blnBreakCheck = true;" & vbCrLf)
		Response.Write("			sBreakValue = '';" & vbCrLf)
		Response.Write("			lngActualRow = 0;" & vbCrLf)
		Response.Write("			blnNewPage = true;" & vbCrLf)
		Response.Write("			}" & vbCrLf & vbCrLf)

		Response.Write("		else if (frmOutput.grdCalendarOutput.Columns(0).CellText(bm) != '*')" & vbCrLf)
		Response.Write("			{" & vbCrLf)
		Response.Write("			blnBreakCheck = false;" & vbCrLf)
		Response.Write("			blnNewPage = false;" & vbCrLf)
		Response.Write("			if (lngActualRow > 0)" & vbCrLf)
		Response.Write("				{" & vbCrLf)
		Response.Write("				window.parent.ASRIntranetOutput.ArrayReDim();" & vbCrLf)
		Response.Write("				}" & vbCrLf)
		Response.Write("			for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)" & vbCrLf)
		Response.Write("				{" & vbCrLf)
		Response.Write("				window.parent.ASRIntranetOutput.ArrayAddTo(lngCol, (lngActualRow), frmOutput.grdCalendarOutput.Columns(lngCol).CellText(bm));" & vbCrLf)
		Response.Write("				}" & vbCrLf)
		Response.Write("			}" & vbCrLf)
	
	
		Response.Write("		if (!blnNewPage) " & vbCrLf)
		Response.Write("			{" & vbCrLf)
		Response.Write("			lngActualRow = lngActualRow + 1; " & vbCrLf)
		Response.Write("			}" & vbCrLf)
		Response.Write("		}" & vbCrLf)
		Response.Write("    window.parent.ASRIntranetOutput.Complete();" & vbCrLf)
		Response.Write("    window.parent.parent.ShowDataFrame();" & vbCrLf)
		Response.Write("	}" & vbCrLf)
	End If

	Response.Write("}" & vbCrLf)

	If Not objCalendar.OutputPreview Then
		Response.Write("  window.parent.frmError.txtEventLogID.value = """ & CleanStringForJavaScript(objCalendar.EventLogID) & """;" & vbCrLf)
		Response.Write("  if (frmOriginalDefinition.txtCancelPrint.value == 1) {" & vbCrLf)
		Response.Write("    window.parent.parent.raiseError('',false,true);" & vbCrLf)
		Response.Write("  }" & vbCrLf)
		Response.Write("  else if (window.parent.ASRIntranetOutput.ErrorMessage != '') {" & vbCrLf)
		Response.Write("    window.parent.raiseError(window.parent.ASRIntranetOutput.ErrorMessage,false,false);" & vbCrLf)
		Response.Write("  }" & vbCrLf)
		Response.Write("  else {" & vbCrLf)
		Response.Write("    window.parent.raiseError('',true,false);" & vbCrLf)
		Response.Write("  }" & vbCrLf)
	Else
		Response.Write("  sUtilTypeDesc = window.parent.parent.parent.frames(""top"").frmPopup.txtUtilTypeDesc.value;" & vbCrLf)
		Response.Write("  if (window.parent.ASRIntranetOutput.ErrorMessage != """") {" & vbCrLf)
		Response.Write("    OpenHR.messageBox(sUtilTypeDesc+"" output failed.\n\n"" + window.parent.ASRIntranetOutput.ErrorMessage,48,""Calendar Report"");" & vbCrLf)
		Response.Write("  }" & vbCrLf)
		Response.Write("  else {" & vbCrLf)
		Response.Write("    OpenHR.messageBox(sUtilTypeDesc+"" output complete."",64,""Calendar Report"");" & vbCrLf)
		Response.Write("  }" & vbCrLf)
	End If
					
	Response.Write("	}" & vbCrLf)
	Response.Write("</script>" & vbCrLf & vbCrLf)
End If
Else
If fBadUtilDef Then
%>

<input type='hidden' id="txtOK" name="txtOK" value="False">
<table align="center" class="outline" cellpadding="5" cellspacing="0">
	<tr>
		<td>
			<table class="invisible" cellspacing="0" cellpadding="0">
				<tr>
					<td colspan="3" height="10"></td>
				</tr>
				<tr>
					<td colspan="3" align="center">
						<h3>Error</h3>
					</td>
				</tr>
				<tr>
					<td width="20" height="10"></td>
					<td>
						<h4>Not all session variables found</h4>
					</td>
					<td width="20"></td>
				</tr>
				<tr>
					<td width="20" height="10"></td>
					<td>Type = <%Session("utiltype").ToString()%>
					</td>
					<td width="20"></td>
				</tr>
				<tr>
					<td width="20" height="10"></td>
					<td>Utility Name = <%Session("utilname").ToString()%>
					</td>
					<td width="20"></td>
				</tr>
				<tr>
					<td width="20" height="10"></td>
					<td>Utility ID = <%Session("utilid").ToString()%>
					</td>
					<td width="20"></td>
				</tr>
				<tr>
					<td width="20" height="10"></td>
					<td>Action = <%Session("action").ToString()%>
					</td>
					<td width="20"></td>
				</tr>
				<tr>
					<td colspan="3" height="10">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="3" height="10" align="center">
						<input type="button" value="Close" name="cmdClose" style="WIDTH: 80px" width="80" id="cmdClose" class="btn"
							onclick="closeclick();" />
					</td>
				</tr>
				<tr>
					<td colspan="3" height="10"></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<input type="hidden" id="txtSuccessFlag" name="txtSuccessFlag" value="1">

<%
Else
%>

<input type='hidden' id="txtOK" name="txtOK" value="False">
<form id="frmPopup" name="frmPopup">
	<table align="center" class="outline" cellpadding="5" cellspacing="0">
		<tr>
			<td>
				<table class="invisible" cellspacing="0" cellpadding="0">
					<tr>
						<td colspan="3" height="10"></td>
					</tr>
					<%
						Dim sCloseFunction As String
		
						Response.Write("			  <tr> " & vbCrLf)
						Response.Write("			    <td width=20 height=10></td> " & vbCrLf)
						Response.Write("			    <td align=center> " & vbCrLf)

						If fNoRecords Then
							Response.Write("						<H4>Calendar Report '" & Session("utilname") & "' Completed successfully.</H4>" & vbCrLf)
							sCloseFunction = "closeclick();"
						Else
							Response.Write("						<H4>Calendar Report '" & Session("utilname") & "' Failed." & vbCrLf)
							sCloseFunction = "closeclick();"
						End If
						Response.Write("			    </td>" & vbCrLf)
						Response.Write("			    <td width=20></td> " & vbCrLf)
						Response.Write("			  </tr>" & vbCrLf)
					%>
					<tr>
						<td width="20" height="10"></td>
						<td align="center" nowrap>
							<%=objCalendar.ErrorString%>
						</td>
						<td width="20"></td>
					</tr>
					<tr>
						<td colspan="3" height="10">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="3" height="10" align="center">
							<input type="button" value="Close" name="cmdClose" style="WIDTH: 80px" width="80" id="cmdClose" class="btn"
								onclick="closeclick();" />
						</td>
					</tr>
					<tr>
						<td colspan="3" height="10"></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</form>
<input type="hidden" id="Hidden3" name="txtSuccessFlag" value="1">
<input type='hidden' id="txtPreview" name="txtPreview" value="0">
<%
End If
End If

Response.Write("<input type=hidden id=txtTitle name=txtTitle value=""" & Replace(objCalendar.CalendarReportName, """", "&quot;") & """>" & vbCrLf)
objCalendar = Nothing
%>

<form id="frmOriginalDefinition" style="visibility: hidden; display: none">
	<%
		Dim sErrMsg As String = ""
		Response.Write("	<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(Session("utilname").ToString(), """", "&quot;") & """>" & vbCrLf)
		Response.Write("	<INPUT type='hidden' id=txtDefn_ErrMsg name=txtDefn_ErrMsg value=""" & sErrMsg & """>" & vbCrLf)
	%>
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%Session("username").ToString()%>">
	<input type="hidden" id="txtDateFormat" name="txtDateFormat" value="<%Session("LocaleDateFormat").ToString()%>">
	<input type="hidden" id="txtCancelPrint" name="txtCancelPrint">
	<input type="hidden" id="txtOptionsDone" name="txtOptionsDone">
	<input type="hidden" id="txtOptionsPortrait" name="txtOptionsPortrait">
	<input type="hidden" id="txtOptionsMarginLeft" name="txtOptionsMarginLeft">
	<input type="hidden" id="txtOptionsMarginRight" name="txtOptionsMarginRight">
	<input type="hidden" id="txtOptionsMarginTop" name="txtOptionsMarginTop">
	<input type="hidden" id="txtOptionsMarginBottom" name="txtOptionsMarginBottom">
	<input type="hidden" id="txtOptionsCopies" name="txtOptionsCopies">
	<input type="hidden" id="txtCalRep_UtilID" name="txtCalRep_UtilID" value="<%Session("UtilID").ToString()%>">
</form>


<script type="text/javascript">

	$("#reportframe").show();

	util_run_calendarreport_main_window_onload();
	$(".popup").dialog('option', 'title', $("#txtTitle").val());
	$("#top").hide();
	$("#calendarframeset").show();

</script>
