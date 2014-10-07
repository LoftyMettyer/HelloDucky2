<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET.Code" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<script type="text/javascript">

	$("#top").hide();

	$(window).bind('resize', function () {
		$("#ssOutputGrid").setGridWidth($('#main').width(), true);
		//$("#ssOutputGrid").setGridHeight("230px", true);
	}).trigger('resize');
</script>


<%
	Dim objCrossTab As CrossTab
		
	objCrossTab = CType(Session("objCrossTab" & Session("UtilID")), CrossTab)
	If (objCrossTab.ErrorString = "") Then

		Response.Write("<script type=""text/javascript"">" & vbCrLf)
		Response.Write("   function util_run_crosstabs_window_onload() {" & vbCrLf)

		Response.Write("	crosstab_loadAddRecords();" & vbCrLf)
		Response.Write("    frmError.txtEventLogID.value = """ & CleanStringForJavaScript(objCrossTab.EventLogID) & """;" & vbCrLf)
		Response.Write("  }" & vbCrLf)
		Response.Write("</script>" & vbCrLf)

		objCrossTab.EventLogChangeHeaderStatus(EventLog_Status.elsSuccessful)

	Else
		
		If objCrossTab.NoRecords Then
			objCrossTab.EventLogChangeHeaderStatus(EventLog_Status.elsSuccessful)
		Else
			objCrossTab.EventLogChangeHeaderStatus(EventLog_Status.elsFailed)
		End If
		
		If objCrossTab.ErrorString <> "" Then
			objCrossTab.FailedMessage = objCrossTab.ErrorString
		End If

	End If
%>

<div>
	<table id="ssOutputGrid" style="display: none;">
		<tbody>
			<tr>
				<th>_</th>
			</tr>
		</tbody>
	</table>

	<table id="tblNineBox" style="display: none;">
		<tr>
			<td class="yaxismajor" rowspan="3">
				<p class="rot270">Potential</p>
			</td>
			<td class="yaxisminor">
				<p class="rot270">High</p>
			</td>
			<td id="nineBoxR1C1">1</td>
			<td id="nineBoxR1C2">2</td>
			<td id="nineBoxR1C3">3</td>
		</tr>
		<tr>
			<td class="yaxisminor">
				<p class="rot270">Medium</p>
			</td>
			<td id="nineBoxR2C1">4</td>
			<td id="nineBoxR2C2">5</td>
			<td id="nineBoxR2C3">6</td>
		</tr>
		<tr>
			<td class="yaxisminor">
				<p class="rot270">Low</p>
			</td>
			<td id="nineBoxR3C1">7</td>
			<td id="nineBoxR3C2">8</td>
			<td id="nineBoxR3C3">9</td>
		</tr>
		<tr>
			<td colspan="2" rowspan="2" class="xaxis"></td>
			<td class="xaxisminor">Low</td>
			<td class="xaxisminor">Medium</td>
			<td class="xaxisminor">High</td>
		</tr>
		<tr>
			<td colspan="3" class="xaxisminor">Performance</td>
		</tr>
	</table>


</div>


<%--Session("utiltype = 1 		Cross Tab
Session("utiltype = 2		Custom Report
Session("utiltype = 9 		Mail Merge
Session("utiltype = 15 		Absence Breakdown
Session("utiltype = 16 		Bradford Factor
Session("utiltype = 17 		Calendar Report--%>
<br />

<div id="divCrossTabOptions">
	<%--Not a Absence Breakdown so show all components--%>
	<div id="CrossTabsIntersectionControls" style="float: left; width: 45%">
		<div id="Div1" style="font-weight: bold;">Intersection</div>
		<div style="width: 80px; float: left">
			<label>Column :</label>
		</div>
		<div>
			<input id="txtIntersectionColumn" name="txtIntersectionColumn" class="text textdisabled" style="WIDTH: 200px;" disabled="disabled">
		</div>
		<div style="width: 80px; float: left">
			<label>Type :</label>
		</div>
		<div>
			<select id="cboIntersectionType" name="cboIntersectionType" class="combo" style="WIDTH: 205px" onchange="UpdateGrid()"></select>
		</div>
		<div id="CrossTabCheckBoxes" style="float: left; margin-left: 80px; padding-top: 10px">
		<input type="checkbox" id="chkSuppressZeros" name="chkSuppressZeros" value="checkbox"
			onclick="UpdateGrid()" />
		<label
			for="chkSuppressZeros"
			class="checkbox"
			tabindex="0">
			Suppress Zeros<br>
		</label>

		<input type="checkbox" id="chkPercentType" name="chkPercentType" value="checkbox"
			onclick="chkPercentType_Click()" />
		<label
			for="chkPercentType"
			class="checkbox"
			tabindex="0">
			<%
				If objCrossTab.CrossTabType <> CrossTabType.cttAbsenceBreakdown Then
					Response.Write(" Percentage of Type")
				End If
			%>
		</label>
		<br>
		<input type="checkbox" id="chkPercentPage" name="chkPercentPage" value="checkbox"
			onclick="UpdateGrid();" />
		<label
			for="chkPercentPage"
			class="checkbox"
			tabindex="0">
			<%
				If objCrossTab.CrossTabType <> CrossTabType.cttAbsenceBreakdown Then
					Response.Write(" Percentage of Page")
				End If
			%>
		</label>
		<br>
		<input type="checkbox" id="chkUse1000" name="chkUse1000" value="checkbox"
			onclick="UpdateGrid()" />
		<label
			for="chkUse1000"
			class="checkbox"
			tabindex="0">
			<%
				If objCrossTab.CrossTabType <> CrossTabType.cttAbsenceBreakdown Then
					Response.Write(" Use 1000 Separators (,)")
				End If
			%>
		</label>
	</div>
	</div>
		
	<div id="PageControls" style="float: left">
		

		<div id="CrossTabPage" name="CrossTabPage">
			<div id="txtWordVer" style="font-weight: bold;">Page</div>
			<div style="width: 80px; float: left">
				<label>Column :</label>
			</div>
			<div>
				<input id="txtPageColumn" name="txtPageColumn" style="WIDTH: 200px;" class="text textdisabled" disabled="disabled">
			</div>

			<div style="width: 80px; float: left">
				<label>Value :</label>
			</div>
			<div>
				<select id="cboPage" name="cboPage" class="combo" style="WIDTH: 205px" onchange="UpdateGrid()"></select>
			</div>
		</div>
	</div>

</div>

<form id="frmOriginalDefinition">
	<%
		Response.Write("	<input type='hidden' id='txtDefn_Name' name='txtDefn_Name' value='" & objCrossTab.CrossTabName.ToString() & "'>" & vbCrLf)
		Response.Write("	<input type='hidden' id=txtDefn_ErrMsg name=txtDefn_ErrMsg value=""" & objCrossTab.ErrorString & """>" & vbCrLf)
	%>
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
	<input type="hidden" id="txtDateFormat" name="txtDateFormat" value="<%=session("LocaleDateFormat")%>">

	<input type="hidden" id="txtCurrentPrintPage" name="txtCurrentPrintPage">
	<input type="hidden" id="txtCancelPrint" name="txtCancelPrint">
	<input type="hidden" id="txtOptionsDone" name="txtOptionsDone">
	<input type="hidden" id="txtOptionsPortrait" name="txtOptionsPortrait">
	<input type="hidden" id="txtOptionsMarginLeft" name="txtOptionsMarginLeft">
	<input type="hidden" id="txtOptionsMarginRight" name="txtOptionsMarginRight">
	<input type="hidden" id="txtOptionsMarginTop" name="txtOptionsMarginTop">
	<input type="hidden" id="txtOptionsMarginBottom" name="txtOptionsMarginBottom">
	<input type="hidden" id="txtOptionsCopies" name="txtOptionsCopies">
</form>

<select style="visibility: hidden; display: none" id="cboDummy" name="cboDummy">
</select>
