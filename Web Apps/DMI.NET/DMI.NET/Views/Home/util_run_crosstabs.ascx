<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<script type="text/javascript">
	$("#top").hide();
	$(".popup").dialog('option', 'title', $("#txtDefn_Name").val());
	//$(".popup").height("470px");
	
	$(window).bind('resize', function () {
		$("#ssOutputGrid").setGridWidth($('#main').width(), true);
		$("#ssOutputGrid").setGridHeight("230px", true);
	}).trigger('resize');
</script>


<%
	Dim objCrossTab As HR.Intranet.Server.CrossTab
		
	objCrossTab = CType(Session("objCrossTab" & Session("UtilID")), CrossTab)
	If (objCrossTab.ErrorString = "") Then

		Response.Write("<script type=""text/javascript"">" & vbCrLf)
		Response.Write("   function util_run_crosstabs_window_onload() {" & vbCrLf)

		Response.Write("	crosstab_loadAddRecords();" & vbCrLf)
		Response.Write("    frmError.txtEventLogID.value = """ & CleanStringForJavaScript(objCrossTab.EventLogID) & """;" & vbCrLf)
		Response.Write("  }" & vbCrLf)
		Response.Write("</script>" & vbCrLf)

		objCrossTab.EventLogChangeHeaderStatus(3)	 'Successful

	Else
		Response.Write("<FORM Name=frmPopup ID=frmPopup>" & vbCrLf)
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
			If objCrossTab.CrossTabType = 3 Then
				Response.Write("						<H4>Absence Breakdown Completed successfully.</H4>" & vbCrLf)
			Else
				Response.Write("						<H4>Cross Tab '" & Session("utilname").ToString.Trim & "' Completed successfully.</H4>" & vbCrLf)
			End If
			objCrossTab.EventLogChangeHeaderStatus(3)		 'Successful
		Else
			If objCrossTab.CrossTabType = 3 Then
				Response.Write("						<H4>Absence Breakdown Failed." & vbCrLf)
			Else
				Response.Write("						<H4>Cross Tab '" & Session("utilname").ToString.Trim & "' Failed." & vbCrLf)
			End If
			objCrossTab.EventLogChangeHeaderStatus(2)		 'Failed
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
		Response.Write("</FORM>" & vbCrLf)
		
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

<div>
	<%	If CLng(Session("utiltype")) <> 15 Then%>
		<div style="float: left; font-weight: bold; padding-top: 10px; padding-bottom: 5px; width: 59%">
			<Label>Intersection</label>
		</div>
		<div id="txtWordVer" style=" font-weight: bold; padding-top: 10px; padding-bottom: 5px">Page</div>
	<%Else%>
		<div style="float: left; font-weight: bold; padding-top: 10px;padding-bottom: 5px; width: 100%">Intersection</div>
	<%End If%>
</div>

<div style="width: 35%;position: static; float: left; white-space: nowrap; text-align: right">
	<%	If CLng(Session("utiltype")) = 15 Then%>
	<input type="HIDDEN" id="txtIntersectionColumn" name="txtIntersectionColumn" style="BACKGROUND-COLOR: threedface; width: 200px">
	<%	Else%>
	<label>Column :</label>
	<input id="txtIntersectionColumn" name="txtIntersectionColumn" class="text textdisabled" style="WIDTH: 200px;" disabled="disabled">
	<%	End If%>
	<div>
		<label>Type :</label>
		<select id="cboIntersectionType" name="cboIntersectionType" class="combo" style="WIDTH: 205px" onchange="UpdateGrid()"></select>
	</div>
</div>

<div style="width: 20%; padding-left: 5px;position :static; white-space: nowrap; float: left">
		<input type="checkbox" id="Checkbox3" name="chkSuppressZeros" value="checkbox"
		onclick="UpdateGrid()" />
	<label
		for="chkSuppressZeros"
		class="checkbox"
		tabindex="0">
		Suppress Zeros<br>
	</label>

	<input type="checkbox" id="Checkbox1" name="chkPercentType" value="checkbox"
		onclick="chkPercentType_Click()" />
	<label
		for="chkPercentType"
		class="checkbox"
		tabindex="0">
		<%
			If objCrossTab.CrossTabType <> 3 Then
				Response.Write(" Percentage of Type")
			End If
		%>
	</label>
	<br>
	<input type="checkbox" id="Checkbox2" name="chkPercentPage" value="checkbox"
		onclick="UpdateGrid();" />
	<label
		for="chkPercentPage"
		class="checkbox"
		tabindex="0">
		<%
			If objCrossTab.CrossTabType <> 3 Then
				Response.Write(" Percentage of Page")
			End If
		%>
	</label>
	<br>

	<input type="checkbox" id="Checkbox4" name="chkUse1000" value="checkbox"
		onclick="UpdateGrid()" />
	<label
		for="chkUse1000"
		class="checkbox"
		tabindex="0">
		<%
			If objCrossTab.CrossTabType <> 3 Then
				Response.Write(" Use 1000 Separators (,)")
			End If
		%>
	</label>
</div>

<div style="width: 35%;float: left;white-space: nowrap;padding-left: 30px;text-align: right;" id="CrossTabPage" name="CrossTabPage">
	<label style="padding-left: 5px;text-align: left">Column :</label>
	<input id="txtPageColumn" name="txtPageColumn" style="width: 200px" class="text textdisabled" disabled="disabled">
	<br/>
	<label style="padding-left: 5px;text-align: left">Value :</label>
	<select id="cboPage" name="cboPage" style="WIDTH: 205px" class="combo" onchange="UpdateGrid()"></select>
</div>

<div style="clear: left">
	<input type="button" id="cmdOutput" name="cmdOutput" value="Output" style="WIDTH: 80px" onclick="ViewExportOptions();" />
	<input type="button" id="cmdClose" name="cmdClose" value="Close" style="WIDTH: 80px" class="btn" onclick="try { closeclick(); } catch (e) { }" />
</div>


<form id="frmOriginalDefinition">
	<input type="hidden" id="txtDefn_Name" name="txtDefn_Name" value="<%=session("utilname")%>">
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

<form target="Output" action="util_run_outputoptions" method="post" id="frmExportData" name="frmExportData">
	<input type="hidden" id="txtPreview" name="txtPreview" value="">
	<input type="hidden" id="txtFormat" name="txtFormat" value="0">
	<input type="hidden" id="txtScreen" name="txtScreen" value="">
	<input type="hidden" id="txtPrinter" name="txtPrinter" value="">
	<input type="hidden" id="txtPrinterName" name="txtPrinterName" value="">
	<input type="hidden" id="txtSave" name="txtSave" value="">
	<input type="hidden" id="txtSaveExisting" name="txtSaveExisting" value="">
	<input type="hidden" id="txtEmail" name="txtEmail" value="">
	<input type="hidden" id="txtEmailAddr" name="txtEmailAddr" value="">
	<input type="hidden" id="txtEmailAddrName" name="txtEmailAddrName" value="">
	<input type="hidden" id="txtEmailSubject" name="txtEmailSubject" value="">
	<input type="hidden" id="txtEmailAttachAs" name="txtEmailAttachAs" value="">
	<input type="hidden" id="txtEmailGroupAddr" name="txtEmailGroupAddr" value="">
	<input type="hidden" id="txtFileName" name="txtFileName" value="">
	<input type="hidden" id="txtUtilType" name="txtUtilType" value="<%=session("utilType")%>">
</form>

<%--<select style="visibility: hidden; display: none" id="cboDummy" name="cboDummy">
</select>--%>
