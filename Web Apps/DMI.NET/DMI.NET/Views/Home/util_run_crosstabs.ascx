<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<script type="text/javascript">
	$("#top").hide();
	$(".popup").dialog('option', 'title', $("#txtDefn_Name").val());

	$(window).bind('resize', function () {
		$("#ssOutputGrid").setGridWidth($('#main').width(), true);
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

<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
	<tr>
		<td>
			<table height="100%" width="100%" class="invisible" cellspacing="0" cellpadding="0">
				<tr>
					<td colspan="50">
						<table class='outline' style='width: 100%; height:auto' id="ssOutputGrid">
							<tbody>
								<tr class='header' style="text-align: left;" >
									<th>_</th>
								</tr>
							</tbody>
						</table>
					</td>
				</tr>
				<tr height="5">
					<td colspan="50">
						<table width="100%" class="outline" cellspacing="0" cellpadding="0">
							<tr height="5">
								<td></td>
							</tr>

							<tr>
								<td>&nbsp;&nbsp;<u>Intersection</u>
									<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
									<td>
										<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
											<%	If CLng(Session("utiltype")) = 15 Then%>
											<input type="HIDDEN" id="txtIntersectionColumn" name="txtIntersectionColumn" style="BACKGROUND-COLOR: threedface; WIDTH: 100%"
												readonly>
										</table>
									</td>
								<% Else%>
								<td width="20"></td>
								<td width="100">Column :</td>
								<td width="5"></td>
								<td width="300">
									<input id="txtIntersectionColumn" name="txtIntersectionColumn" class="text textdisabled" style="WIDTH: 100%"
										disabled="disabled"></td>
								<tr height="5">
									<td></td>
								</tr>
								<% End If%>
								<td width="20"></td>
								<td width="100" valign="top">Type :</td>
								<td width="5"></td>
								<td width="300" valign="top">
									<select id="cboIntersectionType" name="cboIntersectionType" class="combo" style="WIDTH: 100%" onchange="UpdateGrid()"></select>
								</td>
								<td width="20"></td>

						</table>

						<tr height="5">
							<td></td>
						</tr>

			</table>
		</td>
		<td width="20%" valign="top" nowrap>

			<input type="checkbox" id="chkPercentType" name="chkPercentType" value="checkbox"
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

			<input type="checkbox" id="chkPercentPage" name="chkPercentPage" value="checkbox"
				onclick="UpdateGrid();"/>
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

			<input type="checkbox" id="chkSuppressZeros" name="chkSuppressZeros" value="checkbox"
				onclick="UpdateGrid()" />
			<label
				for="chkSuppressZeros"
				class="checkbox"
				tabindex="0">
				Suppress Zeros<br>
			</label>

			<input type="checkbox" id="chkUse1000" name="chkUse1000" value="checkbox"
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

			<br>
		</td>
		<td id="CrossTabPage" name="CrossTabPage">&nbsp;&nbsp;<u>Page</u>
			<table width="100%" outline="invisible" cellspacing="0" cellpadding="0">
				<tr>
					<td width="20"></td>
					<td width="100">Column :</td>
					<td width="5"></td>
					<td width="300">
						<input id="txtPageColumn" name="txtPageColumn" style="WIDTH: 100%" class="text textdisabled" disabled="disabled"></td>
					<tr height="5">
						<td></td>
					</tr>
					<td width="20"></td>
					<td width="100">Value :</td>
					<td width="5"></td>
					<td width="300">
						<select id="cboPage" name="cboPage" style="WIDTH: 100%" class="combo" onchange="UpdateGrid()">
						</select>
					</td>
					<td width="20"></td>
				</tr>
			</table>
		</td>
	</tr>

	<tr height="1">
		<td width="40%"></td>
		<td width="150"></td>
		<td></td>
	</tr>
</table>

<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
	<tr height="5">
		<td></td>
	</tr>
	<tr height="5">
		<td colspan="3">
			<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
				<td align="RIGHT">
					<input type="button" id="cmdOutput" name="cmdOutput" value="Output" style="WIDTH: 80px"
						onclick="ViewExportOptions();" />
				</td>
				<td width="15"></td>
				<td width="5" align="RIGHT">
					<input type="button" id="cmdClose" name="cmdClose" value="Close" style="WIDTH: 80px" class="btn"
						onclick="try { closeclick(); } catch (e) { }" />
				</td>
			</table>
		</td>
	</tr>
	<tr height="5">
		<td></td>
	</tr>
</table>


</table>
</table>
</table>


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

<select style="visibility: hidden; display: none" id="cboDummy" name="cboDummy">
</select>
