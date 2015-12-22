<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl(Of DMI.NET.Models.PopoutChartModel)" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server.Interfaces" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%="" %>

<%
	Dim fMultiAxis As Boolean = (Model.chartTableID2 > 0 Or Model.chartTableID3 > 0)
	Dim objChart As IChart
																								
	If fMultiAxis Then
		objChart = New HR.Intranet.Server.clsMultiAxisChart
	Else
		objChart = New HR.Intranet.Server.clsChart
	End If

	objChart.SessionInfo = CType(Session("SessionContext"), SessionInfo)

	Dim mrstChartData As DataTable
			
	If fMultiAxis Then
		mrstChartData = objChart.GetChartData(Model.chartTableID, Model.chartColumnID, Model.chartFilterID, Model.aggregateType, Model.elementType, Model.chartTableID2, Model.chartColumnID2, Model.chartTableID3, Model.chartColumnID3, Model.sortOrderID, Model.sortDirection, Model.colourID)
	Else
		mrstChartData = objChart.GetChartData(Model.chartTableID, Model.chartColumnID, Model.chartFilterID, Model.aggregateType, Model.elementType, 0, 0, 0, 0, Model.sortOrderID, Model.sortDirection, Model.colourID)
	End If
			
	If Not mrstChartData Is Nothing Then
		If mrstChartData.Rows.Count > 500 Then mrstChartData = Nothing ' limit to 500 rows as get row buffer limit exceeded error.
	End If

	Dim chartAggregateType As ChartAggregateType = Model.aggregateType
%>

<div id="popout_Wrapper" style="position: absolute; top: 0; bottom: 0; left: 0; right: 0;">
	<div id="popout_chartDataDiv" style="width: 90%; height: 90%; display: none"></div>
	<div id="popout_chartImageDiv" style="width: 90%; height: 90%;" >
		<img id="popout_chartImage" src="" alt="Loading chart..." />
	</div>
	<div id="popout_Toolbar" style="position: relative; bottom: 0;">
		<table class="aligncenter" style="border: 1px solid; background-color: #ccc">
			<tr style='font-family: Verdana; font-size: x-small'>
				<td>
					<%--<input type="button" value="Redraw chart" onclick="loadChart();" /></td>--%>  
				<td>Chart Type:
				<select id='popout_selChartType' onchange='loadChart()'>
					<option value="0" <%=IIf(Model.chartType = 0, "selected", "")%>>3D Bar</option>
					<option value="1" <%=IIf(Model.chartType = 1, "selected", "")%>>2D Bar</option>
					<option value="2" <%=IIf(Model.chartType = 2, "selected", "")%>>3D Line</option>
					<option value="3" <%=IIf(Model.chartType = 3, "selected", "")%>>2D Line</option>
					<option value="4" <%=IIf(Model.chartType = 4, "selected", "")%>>3D Area</option>
					<option value="5" <%=IIf(Model.chartType = 5, "selected", "")%>>2D Area</option>
					<option value="6" <%=IIf(Model.chartType = 6, "selected", "")%>>3D Step</option>
					<option value="7" <%=IIf(Model.chartType = 7, "selected", "")%>>2D Step</option>
					<option value="14" <%=IIf(Model.chartType = 14, "selected", "")%>>2D Pie</option>
					<option value="16" <%=IIf(Model.chartType = 16, "selected", "")%>>2D XY</option>
				</select></td>
				<td>Show Legend:<input id='chkshowLegend' type='checkbox' onchange='loadChart()' <%=IIf(Model.showLegend = True, "checked", "")%> /></td>
				<td>Stack Series:<input id='chkstackSeries' type='checkbox' onchange='loadChart()' <%=IIf(Model.stackSeries = True, "checked", "")%> /></td>
				<td>Show Gridlines:<input id='chkShowGrid' type='checkbox' onchange='loadChart()' <%=IIf(Model.showGrid = True, "checked", "")%> /></td>
				<td>Show Values As:<select id='lstValueType' onchange='loadChart()'>
					<option value="Values" <%=IIf(Model.showPercentages = True, "selected", "")%>>Values</option>
					<option value="Percentages" <%=IIf(Model.showPercentages = True, "selected", "")%>>Percentages</option>
				</select></td>
				<%If Model.multiAxis Then%>
				<td>Rotate X:
				<select id='Inclination' onchange='loadChart()'>
					<option value='-90'>-90</option>
					<option value='-70'>-70</option>
					<option value='-50'>-50</option>
					<option value='-30'>-30</option>
					<option value='-10'>-10</option>
					<option value='0'>0</option>
					<option value='10' selected='selected'>10</option>
					<option value='30'>30</option>
					<option value='50'>50</option>
					<option value='70'>70</option>
					<option value='90'>90</option>
				</select></td>
				<td>Rotate Y:
				<select id='Rotation' onchange='loadChart()'>
					<option value='-110'>-110</option>
					<option value='-90'>-90</option>
					<option value='-70'>-70</option>
					<option value='-50'>-50</option>
					<option value='-30'>-30</option>
					<option value='-10'>-10</option>
					<option value='10' selected='selected'>10</option>
					<option value='30'>30</option>
					<option value='50'>50</option>
					<option value='70'>70</option>
					<option value='90'>90</option>
					<option value='110'>110</option>
				</select></td>
				<td>
					<input value='Reset rotation' id='btnResetRotation' type='button' onclick='resetOrientation()' /></td>
				<%End If%>
				<td>Show Datagrid:<input id='chkshowDatagrid' type='checkbox' onchange='toggleGrid()' /></td>
				<td>
					<input value='Print' id='btnPrint' type='button' onclick='OpenHR.printDiv("popout_Wrapper")' /></td>
			</tr>
		</table>
	</div>

	<div class="widgetplaceholder datagrid" id="chartDatagrid" style="display: none;">
		<table id="DataTable_popout" class="cellspace0 cellpadding5 rulesAll" style="width: 100%; vertical-align: top; border: 3px solid lightgray">
			<%If mrstChartData.Rows.Count > 0 AndAlso (TryCast(mrstChartData.Rows(0)(0), String) <> "No Access" And TryCast(mrstChartData.Rows(0)(0), String) <> "No Data") Then%>
			<thead>
				<tr>
					<th style="font-weight: normal; text-align: left; cursor: default">
						<%=Left(NullSafeString(Model.chartColumnName), 50)%>
					</th>
					<%If fMultiAxis Then%>
					<th style="font-weight: normal; text-align: left; cursor: default">
						<%=Trim(Left(NullSafeString(Model.chartColumnName2), 50))%>
					</th>
					<th style="font-weight: normal; text-align: right; cursor: default">
						<%Else%>
					<th style="font-weight: normal; text-align: right; cursor: default">
						<%End If%>
						<%Response.Write(chartAggregateType.ToString())%>
					</th>
				</tr>
			</thead>
			<tbody>
				<%
					If mrstChartData.Rows.Count > 0 Then
						For Each objRow As DataRow In mrstChartData.Rows
				%>
				<tr>
					<td class="bordered" style="width: 150px; text-align: left; white-space: nowrap">
						<%If fMultiAxis Then%>
						<%=Trim(Left(NullSafeString(objRow(1)), 50))%>
						<%Else%>
						<%=Trim(Left(NullSafeString(objRow(0)), 50))%>
						<%End If%>
					</td>
					<%If fMultiAxis Then%>
					<td class="bordered" style="text-align: left; white-space: nowrap">
						<div style="width: 150px; white-space: nowrap">
							<%=Trim(Left(NullSafeString(objRow(3)), 50))%>
						</div>
					</td>
					<%End If%>
					<td class="bordered" style="text-align: right; vertical-align: top; padding-bottom: 0; white-space: nowrap; overflow: hidden">
						<%If fMultiAxis Then%>
						<%=Trim(Left(NullSafeString(objRow(4)), 50))%>
						<%Else%>
						<%=Trim(Left(NullSafeString(objRow(1)), 50))%>

						<%End If%>
					</td>
				</tr>
				<%    
											
				Next
				%>
			</tbody>
			<%
			End If
		Else
			%>
			<tr>
				<td class="bordered" style="text-align: center;" rowspan="3">No matching records found</td>
			</tr>
			<script type="text/javascript">
				// No data on this chart, adjust UI accordingly
			<%If Session("CurrentLayout").ToString() <> Layout.tiles.ToString() Then%>
				$("#WidgetPlaceHolder_popout").css('height', "40px"); //Reduce the size of the parent div ('widgetplaceholder')
				$("#WidgetPlaceHolder_popout").children(0).css('border', 'none'); //Remove the border of the table
				<%End If%>
			</script>
			<%End If%>
		</table>
		<script type="text/javascript">
			//Attach table sorter to the table
			$("#DataTable_popout").tablesorter();
		</script>
	</div>
</div>

<script type="text/javascript">
	function loadChart() {
		var windowHeight = $('#popout_chartImageDiv').height() - 80;
		var windowWidth = $('#popout_chartImageDiv').width();
		var chartType = document.getElementById("popout_selChartType").value;
		var chartShowLegend = (document.getElementById("chkshowLegend").checked == true);
		var chartStackSeries = (document.getElementById("chkstackSeries").checked == true);
		var chartShowGridlines = (document.getElementById("chkShowGrid").checked == true);
		var chartShowPercentages = (document.getElementById("lstValueType").value == "Percentages");
		var psUrl;
		var rotateX = 10;
		var rotateY = 10;

		if ('<%=model.multiAxis%>' == 'True') {
			rotateX = document.getElementById("Inclination").value;
			rotateY = document.getElementById("Rotation").value;
			psUrl = "GetMultiAxisChart?";
		} else {
			psUrl = "GetChart?";
		}
		psUrl +=
			"height=" + windowHeight + "&width=" + windowWidth + "&ShowLegend=" + chartShowLegend + "&DottedGrid=" + chartShowGridlines + "&ShowValues=true&Stack=" + chartStackSeries + "&ShowPercent=" + chartShowPercentages +
				"&ChartType=" + chartType +
				"&TableID=<%=Model.chartTableID%>" +
				"&ColumnID=<%=Model.chartColumnID%>" +
				"&FilterID=<%=Model.chartFilterID%>" +
				"&AggregateType=<%=Model.aggregateType%>" +
				"&ElementType=<%=Model.elementType%>";

		if ('<%=model.multiAxis%>' == 'True') {
			psUrl += "&RotateX=" + rotateX +
				"&RotateY=" + rotateY +
				"&TableID_2=<%=Model.chartTableID2%>" +
				"&ColumnID_2=<%=Model.chartColumnID2%>" +
				"&TableID_3=<%=Model.chartTableID3%>" +
				"&ColumnID_3=<%=Model.chartColumnID3%>";
		}

		psUrl += "&SortOrderID=<%=Model.sortOrderID%>" +
			"&SortDirection=<%=Model.sortDirection%>" +
			"&ColourID=<%=Model.colourID%>" +
			"&ShowLabels=True" +
			"&Title=<%=Model.chartText%>";
		document.getElementById("popout_chartImage").src = psUrl;

	}

	function resetOrientation() {
		document.getElementById("Inclination").value = 10;
		document.getElementById("Rotation").value = 10;
		loadChart();
	}

	setTimeout("loadChart()", 500);



	//function loadData() {

	//	var ajaxCall = "GetChartDataAsHTML?TableID=1&ColumnID=107&FilterID=12513&AggregateType=0&ElementType=2&TableID_2=0&ColumnID_2=0&TableID_3=0&ColumnID_3=0&SortOrderID=0&SortDirection=0&ColourID=0&Title=Headcount%20by%20Department&MultiAxisChart=False";

	//	ajaxCall = "GetChartDataAsHTML?" +
	//		"Table"

	//	$.ajax({
	//		url: ajaxCall, type: "GET", dataType: "text", async: true, success: function (data) {
	//			try {
	//				$("#popout_chartDataDiv").html("");
	//				$("#popout_chartDataDiv").html(data);
	//			} catch (e) { }
	//		}
	//	});
	//	$(document).ready(function () {
	//		$("#chartData").setGridWidth($("#popout_chartDataDiv").width());
	//		$("#chartData").setGridHeight($("#popout_chartDataDiv").height());
	//		$(window).on("debouncedresize", function () {
	//			$("#chartData").setGridWidth($("#popout_chartDataDiv").width());
	//			$("#chartData").setGridHeight($("#popout_chartDataDiv").height());
	//			loadChart();
	//		});
	//	});
	//};

	//function ScriptjQueryDebouncedResizeCallback() {
	//	loadData();
	//} function ScriptjQueryCallback() {
	//	var head = document.getElementsByTagName("head")[0];
	//	var ScriptjQueryUI = document.createElement("script");

	//	ScriptjQueryUI.type = "text/javascript";
	//	ScriptjQueryUI.src = pathToResources + "/Scripts/jquery/jquery-ui-1.9.2.custom.js";
	//	ScriptjQueryUI.onreadystatechange = ScriptjQueryUICallback;
	//	ScriptjQueryUI.onload = ScriptjQueryUICallback;
	//	head.appendChild(ScriptjQueryUI);
	//};

	//function ScriptjQueryUICallback() {
	//	var head = document.getElementsByTagName("head")[0];
	//	var ScriptjqGrid = document.createElement("script");
	//	ScriptjqGrid.type = "text/javascript";
	//	ScriptjqGrid.src = pathToResources + "/Scripts/jquery/jquery.jqGrid.src.js";
	//	ScriptjqGrid.onreadystatechange = ScriptjqGridCallback;
	//	ScriptjqGrid.onload = ScriptjqGridCallback;
	//	head.appendChild(ScriptjqGrid);
	//}

	//function ScriptjqGridCallback() {
	//	var head = document.getElementsByTagName("head")[0];
	//	var jQueryDebouncedResize = document.createElement("script");
	//	jQueryDebouncedResize.type = "text/javascript";
	//	jQueryDebouncedResize.src = pathToResources + "/Scripts/jquery/jquery-debouncedresize.js";
	//	jQueryDebouncedResize.onreadystatechange = ScriptjQueryDebouncedResizeCallback;
	//	jQueryDebouncedResize.onload = ScriptjQueryDebouncedResizeCallback;
	//	head.appendChild(jQueryDebouncedResize);
	//}

	//var pathToResources = window.location.pathname.substring(0, window.location.pathname.substring(1).indexOf("/") + 1);
	//if (pathToResources == "/Home") {
	//	pathToResources = "";
	//};
	//var head = document.getElementsByTagName("head")[0];
	//var ScriptjQuery = document.createElement("script");
	//ScriptjQuery.type = "text/javascript";

	//ScriptjQuery.src = pathToResources + "/Scripts/jquery/jquery-1.8.3.js";
	//ScriptjQuery.onreadystatechange = ScriptjQueryCallback;
	//ScriptjQuery.onload = ScriptjQueryCallback;
	//head.appendChild(ScriptjQuery);
	//var DMIthemeLink = document.createElement("link");
	//DMIthemeLink.rel = "stylesheet";
	//DMIthemeLink.type = "text/css";
	//DMIthemeLink.href = pathToResources + "/Content/themes/redmond-segoe/jquery-ui.min.css";
	//head.appendChild(DMIthemeLink);

	function toggleGrid() {
		$('#popout_chartImageDiv').toggle();
		$('#chartDatagrid').toggle();

		//enable/disable toolbar buttons
		var toolbarDisabled = false;
		if ($('#chartDatagrid').is(":visible")) toolbarDisabled = true;

		$('#popout_Toolbar input, #popout_Toolbar select').prop('disabled', toolbarDisabled);
		$('#chkshowDatagrid').prop('disabled', false);
		$('#btnPrint').prop('disabled', false);
	}

</script>



