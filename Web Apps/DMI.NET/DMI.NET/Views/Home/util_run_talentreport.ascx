<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="DMI.NET.Models.Responses" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_customreports")%>" type="text/javascript"></script>
<script src="<%:Url.LatestContent("~/Scripts/jquery/jquery.sparkline.min.js") %>"></script>

<% 

    Dim matchReport = New MatchReportRun
    matchReport.SessionInfo = CType(Session("SessionContext"), SessionInfo)

    Dim prompts = Session("Prompts_" & Session("utiltype") & "_" & Session("utilid"))

    matchReport.SetPromptedValues(prompts)
    matchReport.MatchReportID = Cint(Session("utilid"))
    dim result = matchReport.RunMatchReport()

    dim combinedString = string.Join( ",", result.ToArray() )


    '   dim tabbedData = string.Join(vbTab, result.Data.ToArray() )

    '   			Dim message As New PostResponse With {
    '		.Message = combinedString '"Match Report not yet implemented!"
    '}



    Dim fok As Boolean = True
    Dim fNotCancelled As Boolean = True

    ' Bind the data to the grid if atleast one non-hidden column available
    If (matchReport.ReportDataTable IsNot Nothing) Then

        gridReportData.DataSource = matchReport.ReportDataTable

        gridReportData.DataBind()

    Else
        Response.Write("No output generated. Check your data.")
    End If

    If fok Then
        Response.Write("<div>")
        Response.Write("			<table name=tblGrid id=tblGrid height=100% width=100% class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
        Response.Write("				<tr>" & vbCrLf)
        Response.Write("					<td ALIGN=center colspan=12 NAME='tdOutputMSG' ID='tdOutputMSG'>" & vbCrLf)
%>

<form id="formReportData" runat="server">
	<asp:GridView ID="gridReportData" runat="server"
		AllowPaging="False"
		GridLines="None"
		CssClass="visibletablecolumn"
		ClientIDMode="Static">
	</asp:GridView>
</form>

<%		
	Response.Write("					</td>" & vbCrLf)
	Response.Write("				</tr>" & vbCrLf)
	Response.Write("				<tr>" & vbCrLf)
	Response.Write("					<td colspan=12 height=10></td>" & vbCrLf)
	Response.Write("				</tr>" & vbCrLf)

	Response.Write("				<tr height=25>" & vbCrLf)
	Response.Write("					<td width=20></td>" & vbCrLf)
	Response.Write("					<td colspan=8>" & vbCrLf)
	Response.Write("            <div>")
	Response.Write("						<table WIDTH=""100%"" class=""invisible"" CELLSPACING=0 CELLPADDING=0>" & vbCrLf)
	Response.Write("							<tr>" & vbCrLf)
	Response.Write("								<td>" & vbCrLf)
	Response.Write("								</td>" & vbCrLf)
	Response.Write("								<td>&nbsp;</td>" & vbCrLf)
	Response.Write("								<td width=20>" & vbCrLf)
	Response.Write("								</td>" & vbCrLf)
	Response.Write("							</tr>" & vbCrLf)
	Response.Write("						</table>" & vbCrLf)
	Response.Write("</div>")
	Response.Write("					</td>" & vbCrLf)
	Response.Write("					<td width=10></td>" & vbCrLf)
	Response.Write("					<td width=80> " & vbCrLf)
	Response.Write("					</td>" & vbCrLf)
	Response.Write("					<td width=20></td>" & vbCrLf)
	Response.Write("				</tr>" & vbCrLf)
	Response.Write("				<tr>" & vbCrLf)
	Response.Write("					<td colspan=12 height=10></td>" & vbCrLf)
	Response.Write("				</tr>" & vbCrLf)
	Response.Write("			</table>" & vbCrLf)
	Response.Write("      </div>")

End If
%>

<input type='hidden' id="txtNoRecs" name="txtNoRecs" value="<%:matchReport.NoRecords%>">
<input type='hidden' id="txtDefn_Name" name="txtDefn_Name" value="<%:matchReport.ReportCaption%>">
<input type='hidden' id="txtDefn_ErrMsg" name="txtDefn_ErrMsg" value="<%:matchReport.ErrorString%>">


<form action="util_run_talentreport_downloadoutput" method="post" id="frmExportData" name="frmExportData" target="submit-iframe">
	<input type="hidden" id="txtPreview" name="txtPreview" value="<%=matchReport.OutputPreview%>">
	<input type="hidden" id="txtFormat" name="txtFormat" value="<%=matchReport.OutputFormat%>">
	<input type="hidden" id="txtScreen" name="txtScreen" value="<%=matchReport.OutputScreen%>">
	<input type="hidden" id="txtEmailGroupAddr" name="txtEmailGroupAddr" value="">
	<input type="hidden" id="txtFileName" name="txtFileName" value="<%=matchReport.OutputFilename%>">
	<input type="hidden" id="txtUtilType" name="txtUtilType" value="<%=session("utilType")%>">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value="<%=Session("utilID")%>">
	<input type="hidden" id="download_token_value_id" name="download_token_value_id"/>
	<%=Html.AntiForgeryToken()%>
</form>

<script type="text/javascript">
	//Shrink to fit, or set to 100px per column?
	var ShrinkToFit = false;
	var gridWidth;
	var gridHeight;
	// first get the size from the window
	// if that didn't work, get it from the body
	var size = {
		MakeWidth: $('#divUtilRunForm').width(),
		MakeHeight: $('#reportworkframe').height()
	};
	//Get count of visible columns
	if (menu_isSSIMode()) {
		try {
			gridWidth = $('#reportworkframe').width();
			gridHeight = $('#reportworkframe').height() - 100;
		} catch (e) {
			gridWidth = 'auto';
			gridHeight = 'auto';
		}
		ShrinkToFit = true;
	} else {
		//DMI options.
		
		var iVisibleCount = Number("<%:matchReport.DisplayColumns.Count%>");
		if ((iVisibleCount *100) < size.MakeWidth) ShrinkToFit = true;
		gridWidth = (size.MakeWidth);
		gridHeight = (size.MakeHeight);
	}

		var newFormat = OpenHR.getLocaleDateString();
		var srcFormat = newFormat;	
	var grid = $("#gridReportData"),
			getColumnIndexByName = function (columnName) {
			var cm = grid.jqGrid('getGridParam', 'colModel');
			for (var i = 0, l = cm.length; i < l; i++) {
				if (cm[i].name === columnName) {
					return i; // return the index
				}
			}
			return -1;
		};

		tableToGrid("#gridReportData", {
			shrinkToFit: ShrinkToFit,
			width: gridWidth,
			height: gridHeight,
			ignoreCase: true,
			colNames: [
				<%Dim iColCount As Integer = 0
		For Each objItem In matchReport.DisplayColumns
		Dim sColumnName = HttpUtility.HtmlEncode(matchReport.ReportDataTable.Columns(iColCount).ColumnName)
			Response.Write(String.Format("{0}'{1}'", IIf(iColCount > 0, ", ", ""), sColumnName))
			iColCount += 1
		Next%>
			],
			colModel: [
		<%
	iColCount = 0
	
	For Each objItem In matchReport.DisplayColumns
		Dim sColumnName = HttpUtility.HtmlEncode(matchReport.ReportDataTable.Columns(iColCount).ColumnName.Replace(" ", "_").Replace("""", "_"))
	
		
		Dim iColumnWidth As Integer = 100
		
		If objItem.IsNumeric Then
			Response.Write(String.Format("{0}{{name:'", IIf(iColCount > 0, ", ", "")) & sColumnName & "',align:'right', width: '" & iColumnWidth.ToString() & "'}")
		ElseIf objItem.DataType = ColumnDataType.sqlDate Then
		  Response.Write(String.Format("{0}{{name:'", IIf(iColCount > 0, ", ", "")) & sColumnName & "', edittype: 'date', align: 'left',  formatter: 'date', formatoptions: { srcformat: srcFormat, newformat: newFormat, disabled: true, width: '" & iColumnWidth.ToString() & "' }}")
		ElseIf sColumnName = "talentchart" then
			Response.Write(String.Format("{0}{{name:'", IIf(iColCount > 0, ", ", "")) & sColumnName & "', width: '160'}")
		Else
			Response.Write(String.Format("{0}{{name:'", IIf(iColCount > 0, ", ", "")) & sColumnName & "', width: '" & iColumnWidth.ToString() & "'}")
		End If
		
		iColCount += 1
	Next
			%>
			],
			cmTemplate: { sortable: false },
			rowNum: 200000,
			loadComplete: function () {				
			  $('#gridReportData').hideCol("rowType");			 
				stylejqGrid();
				$('#gridReportData').setGridWidth($('#main').width());
				setTimeout(function () {
					var index = getColumnIndexByName('talentchart');

					$('#gridReportData').find('tr.jqgrow td:nth-child(' + (index + 1) + ')').each(function (index, value) {
						var ar;
						try {
							ar = $.parseJSON($(this).text());
							if (ar && ar.length > 0) {
								var talentchartCellObject = this;
								$(talentchartCellObject).html("<table width='100%'></table>");

								$.each(ar, function (index, obj) {
									var target = obj.PrefScore,
											performance = obj.ActualScore,
											range1 = 100,
											range2 = target,
											range3 = obj.MinScore;
									var graphData = [target, performance, range1, range2, range3];

									var cell1css = "'width:80px;white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-size: small;'";
									var chartTitleText = "Minimum Score: " + obj.MinScore +
																				"\nPreferred Score: " + obj.PrefScore +
																				"\nActual Score: " + obj.ActualScore;

									$(talentchartCellObject).find("table").append("<tr ><td style='width:80px;padding:2px;border: 0;'><div style=" + cell1css + ">" + obj.Competency + "</div></td>" + 
										"<td style='width:150px;border:0;' title='" + chartTitleText + "' class='graph_" + index + "'></td></tr>");

									//Create the graph and add it to the 2nd cell.
									$(talentchartCellObject).find(".graph_" + index).sparkline(graphData, { type: 'bullet', targetColor: 'red', width: '150px' });
								});

							}
						} catch (e) { }
					});
				}, 100);
			}
		});
	
		$('#gview_gridReportData td').css('white-space', 'pre-line');

		function stylejqGrid() {
			//jqGrid style overrides
			$('#gview_gridReportData tr.jqgrow td').css('vertical-align', 'top'); //float text to top, in case of multi-line cells
			$('#gview_gridReportData tr.footrow td').css('vertical-align', 'top'); //float text to top, in case of multi-line footers
			$('#gview_gridReportData .s-ico span').css('display', 'none'); //hide the sort order icons - they don't tie in to the dataview model.
		}
		
		if (menu_isSSIMode()) $('#gbox_gridReportData').css('margin', '0 auto'); //center the report in self-service screen.
</script>

