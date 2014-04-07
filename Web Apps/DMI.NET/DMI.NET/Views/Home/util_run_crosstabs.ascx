<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
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
		Response.Write("<form Name=frmPopup ID=frmPopup>" & vbCrLf)
		Response.Write("<table align=center class=""outline="" cellPadding=5 cellSpacing=0>" & vbCrLf)
		Response.Write("	<TR>" & vbCrLf)
		Response.Write("		<TD>" & vbCrLf)
		Response.Write("			<table class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
		Response.Write("			  <tr>" & vbCrLf)
		Response.Write("			    <td colspan=3 height=10></td>" & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			  <tr> " & vbCrLf)
		Response.Write("			    <td width=20 height=10></td> " & vbCrLf)
		Response.Write("			    <td align=center> " & vbCrLf)

		If objCrossTab.NoRecords Then
			objCrossTab.EventLogChangeHeaderStatus(EventLog_Status.elsSuccessful)
		Else
			objCrossTab.EventLogChangeHeaderStatus(EventLog_Status.elsFailed)
		End If

		Response.Write("			    </td>" & vbCrLf)
		Response.Write("			    <td width=20></td> " & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			  <tr> " & vbCrLf)
		Response.Write("			    <td width=20 height=10></td> " & vbCrLf)
		Response.Write("			    <td align=center nowrap>" & objCrossTab.ErrorString & vbCrLf)
		Response.Write("			    </td>" & vbCrLf)
		Response.Write("			    <td width=20></td> " & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			  <tr>" & vbCrLf)
		Response.Write("			    <td colspan=3 height=10>&nbsp;</td>" & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			  <tr> " & vbCrLf)
		Response.Write("			    <td colspan=3 height=10 align=center> " & vbCrLf)
		Response.Write("						<input type=button id=cmdClose name=cmdClose value=Close style=""WIDTH: 80px"" width=80px class=""btn""" & vbCrLf)
		Response.Write("                      onclick=""closeclick();"""" />" & vbCrLf)
		Response.Write("			    </td>" & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			  <tr> " & vbCrLf)
		Response.Write("			    <td colspan=3 height=10></td>" & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			</table>" & vbCrLf)
		Response.Write("		</td>" & vbCrLf)
		Response.Write("	</tr>" & vbCrLf)
		Response.Write("</table>" & vbCrLf)
		Response.Write("</form>" & vbCrLf)
		
		If objCrossTab.ErrorString <> "" Then
			objCrossTab.FailedMessage = objCrossTab.ErrorString
		End If

		Response.End()
	End If
%>

<div>
	<table id="ssOutputGrid">
		<tbody>
			<tr>
				<th>_</th>
			</tr>
		</tbody>
	</table>
</div>


<%--Session("utiltype = 1 		Cross Tab
Session("utiltype = 2		Custom Report
Session("utiltype = 9 		Mail Merge
Session("utiltype = 15 		Absence Breakdown
Session("utiltype = 16 		Bradford Factor
Session("utiltype = 17 		Calendar Report--%>
<br />

<div>
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
	</div>
	<div id="PageControls" style="float: left">
		<div id="txtWordVer" style="font-weight: bold;">Page</div>

		<div id="CrossTabPage" name="CrossTabPage">
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

<%--<%If CLng(Session("utiltype")) <> 15 Then%>

<%Else%>
<%--It is an Absence Breakdown so show only Intersection Type combo and suppress Zero components--%>
<%-- %><input type="HIDDEN" id="HIDDEN1" name="txtIntersectionColumn" style="BACKGROUND-COLOR: threedface; width: 200px">
<%End If%>--%>

<form id="frmOriginalDefinition">
	<%
		Response.Write("	<input type='hidden' id='txtDefn_Name' name='txtDefn_Name' value='" & objCrossTab.CrossTabName.ToString() & "'>" & vbCrLf)
	%>
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
	<input type="hidden" id="txtDateFormat" name="txtDateFormat" value="<%=session("LocaleDateFormat")%>">
	<input type="hidden" id="txtDatabase" name="txtDatabase" value="<%=session("database")%>">

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
