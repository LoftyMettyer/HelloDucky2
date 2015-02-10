<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%-- For other devs: Do not remove below line. --%>
<%="" %>
<%-- For other devs: Do not remove above line. --%>
<script type="text/javascript">

	$("#top").hide();

	$(window).bind("resize", function () {
		$("#ssOutputGrid").setGridWidth($("#main").width(), true);		
	}).trigger("resize");
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
	
	<%If objCrossTab.CrossTabType = CrossTabType.ctt9GridBox Then%>
	<table id="tblNineBox" style="display: none;">
		<tr>
			<td class="yaxismajor" rowspan="3">
				<p class="rot270"><%:objCrossTab.YAxisLabel%></p>
			</td>
			<td class="yaxisminor">
				<p class="rot270"><%:objCrossTab.YAxisSubLabel1%></p>
			</td>
			<td id="nineBoxR1C1" class="nineBoxGridCell">
				<p><%:objCrossTab.Description1%></p>
				<p></p>
			</td>
			<td id="nineBoxR1C2" class="nineBoxGridCell">
				<p><%:objCrossTab.Description2%></p>
				<p></p>
			</td>
			<td id="nineBoxR1C3" class="nineBoxGridCell">
				<p><%:objCrossTab.Description3%></p>
				<p></p>
			</td>
		</tr>
		<tr>
			<td class="yaxisminor">
				<p class="rot270"><%:objCrossTab.YAxisSubLabel2%></p>
			</td>
			<td id="nineBoxR2C1" class="nineBoxGridCell">
				<p><%:objCrossTab.Description4%></p>
				<p></p>
			</td>
			<td id="nineBoxR2C2" class="nineBoxGridCell">
				<p><%:objCrossTab.Description5%></p>
				<p></p>
			</td>
			<td id="nineBoxR2C3" class="nineBoxGridCell">
				<p><%:objCrossTab.Description6%></p>
				<p></p>
			</td>
		</tr>
		<tr>
			<td class="yaxisminor">
				<p class="rot270"><%:objCrossTab.YAxisSubLabel3%></p>
			</td>
			<td id="nineBoxR3C1" class="nineBoxGridCell">
				<p><%:objCrossTab.Description7%></p>
				<p></p>
			</td>
			<td id="nineBoxR3C2" class="nineBoxGridCell">
				<p><%:objCrossTab.Description8%></p>
				<p></p>
			</td>
			<td id="nineBoxR3C3" class="nineBoxGridCell">
				<p><%:objCrossTab.Description9%></p>
				<p></p>
			</td>
		</tr>
		<tr>
			<td colspan="2" rowspan="2" class="xaxis"></td>
			<td class="xaxisminor"><%:objCrossTab.XAxisSubLabel1%></td>
			<td class="xaxisminor"><%:objCrossTab.XAxisSubLabel2%></td>
			<td class="xaxisminor"><%:objCrossTab.XAxisSubLabel3%></td>
		</tr>
		<tr>
			<td colspan="3" class="xaxisminor xaxismajor"><%:objCrossTab.XAxisLabel%></td>
		</tr>
	</table>
	<%End If%>

</div>

<%--Session("utiltype = 1 		Cross Tab
Session("utiltype = 2		Custom Report
Session("utiltype = 9 		Mail Merge
Session("utiltype = 15 		Absence Breakdown
Session("utiltype = 16 		Bradford Factor
Session("utiltype = 17 		Calendar Report--%>
<br />
<fieldset class="CTT<%=objCrossTab.CrossTabType%>" style="border: 0;">
	<div id="divCrossTabOptions">
		<%--Not a Absence Breakdown so show all components--%>
		<div id="CrossTabsIntersectionControls" style="float: left; width: 33%;">
			<div id="Div1" style="font-weight: bold;margin-bottom: 8px;">Intersection</div>
			<div style="float: left; font-weight: normal; margin-left: 5px;">
				<label name="txtIntersectionColumn" id="txtIntersectionColumn">Type :</label>
			</div>
			<div>
				<select id="cboIntersectionType" name="cboIntersectionType" class="combo" style="width: 205px" onchange="UpdateGrid()"></select>
			</div>
		</div>


		<div id="PageControls" style="float: left; width: 33%;">
			<div id="CrossTabPage" name="CrossTabPage">
				<div id="txtWordVer" style="font-weight: bold;margin-bottom: 8px;">Page</div>
				<div style="float: left; font-weight: normal; margin-left: 5px;">
					<label id="txtPageColumn" name="txtPageColumn">Page :</label>
				</div>
				<div>
					<select id="cboPage" name="cboPage" class="combo" style="width: 205px" onchange="UpdateGrid()"></select>
				</div>
			</div>
		</div>

		<div id="CrossTabCheckBoxes" style="float: left;width: 33%;">
			<div style="font-weight: bold;
				<%If objCrossTab.CrossTabType = CrossTabType.cttAbsenceBreakdown Then Response.Write(";margin-bottom: 8px;")%>">Options</div>
			<div style="font-weight: normal">
				<input type="checkbox" id="chkSuppressZeros" name="chkSuppressZeros" value="checkbox"
					onclick="UpdateGrid()" />
				<label class="checkbox" for="chkSuppressZeros" tabindex="0">
					Suppress Zeros
				</label>
				<br />
				<input type="checkbox" id="chkUse1000" name="chkUse1000" value="checkbox"
					onclick="UpdateGrid()" />
				<label class="checkbox" for="chkUse1000" tabindex="0">
					<%
						If objCrossTab.CrossTabType <> CrossTabType.cttAbsenceBreakdown Then
							Response.Write(" Use 1000 Separators")
						End If
					%>
				</label>
			</div>
			<%	Dim noshowpercentages As String
				noshowpercentages = ""
				If objCrossTab.CrossTabType = CrossTabType.ctt9GridBox Then
					noshowpercentages = "display: none"
				End If%>

			<div style="font-weight: normal;<%=noshowpercentages%>">			
					<input id="chkPercentType" name="chkPercentType" onclick="chkPercentType_Click()" type="checkbox" value="checkbox" />
				<label class="checkbox" for="chkPercentType" tabindex="0">
					<%
						If objCrossTab.CrossTabType <> CrossTabType.cttAbsenceBreakdown Then
							Response.Write(" Percentage of Type")
						End If
					%>
				</label>
				<br />
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
			</div>			
		</div>
	</div>
</fieldset>
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

<script type="text/javascript">
	$('#nineBoxR1C1').css('background-color', '<%="#" & objCrossTab.ColorDesc1%>');
	$('#nineBoxR1C2').css('background-color', '<%="#" & objCrossTab.ColorDesc2%>');
	$('#nineBoxR1C3').css('background-color', '<%="#" & objCrossTab.ColorDesc3%>');
	$('#nineBoxR2C1').css('background-color', '<%="#" & objCrossTab.ColorDesc4%>');
	$('#nineBoxR2C2').css('background-color', '<%="#" & objCrossTab.ColorDesc5%>');
	$('#nineBoxR2C3').css('background-color', '<%="#" & objCrossTab.ColorDesc6%>');
	$('#nineBoxR3C1').css('background-color', '<%="#" & objCrossTab.ColorDesc7%>');
	$('#nineBoxR3C2').css('background-color', '<%="#" & objCrossTab.ColorDesc8%>');
	$('#nineBoxR3C3').css('background-color', '<%="#" & objCrossTab.ColorDesc9%>');

</script>
