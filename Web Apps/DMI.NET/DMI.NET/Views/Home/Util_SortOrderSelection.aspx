<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<script src="<%: Url.Content("~/Scripts/FormScripts/Util_Def_CustomReports.js") %>" type="text/javascript"></script>

<!DOCTYPE html>
<html>
<head>
	<title>OpenHR Intranet</title>
	<script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
	<script id="officebarscript" src="<%: Url.Content("~/Scripts/officebar/jquery.officebar.js") %>" type="text/javascript"></script>
	<link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/fonts/SSI80v194934/style.css")%>" rel="stylesheet" />
</head>

	<script type="text/javascript">
		function util_SortOrderSelection_onload() {
			$("input[type=submit], input[type=button], button")
				.button();
			$("input").addClass("ui-widget ui-widget-content ui-corner-all");
			$("input").removeClass("text");

			$("select").addClass("ui-widget ui-corner-tl ui-corner-bl");
			$("select").removeClass("text");

			var iResizeBy, iNewWidth, iNewHeight, iNewLeft, iNewTop;
			var frmPopup = document.getElementById("frmPopup");

			// Resize the grid to show all prompted values.
			iResizeBy = frmPopup.offsetParent.scrollWidth - frmPopup.offsetParent.clientWidth;
			if (frmPopup.offsetParent.offsetWidth + iResizeBy > screen.width) {
				window.dialogWidth = new String(screen.width) + "px";
			} else {
				iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length - 2));
				iNewWidth = iNewWidth + iResizeBy;
				window.dialogWidth = new String(iNewWidth) + "px";
			}

			<% If (Session("utiltype") = 2) Then%>
			checkColumnOptions(false);
			<%End If%>

			iResizeBy = frmPopup.offsetParent.scrollHeight - frmPopup.offsetParent.clientHeight;
			if (frmPopup.offsetParent.offsetHeight + iResizeBy > screen.height) {
				window.dialogHeight = new String(screen.height) + "px";
			} else {
				iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length - 2));
				iNewHeight = iNewHeight + iResizeBy;
				window.dialogHeight = new String(iNewHeight) + "px";
			}

			iNewLeft = (screen.width - frmPopup.offsetParent.offsetWidth) / 2;
			iNewTop = (screen.height - frmPopup.offsetParent.offsetHeight) / 2;

			window.dialogLeft = new String(iNewLeft) + "px";
			window.dialogTop = new String(iNewTop) + "px";
		};

		function selectedColumnParameter(psDefnString, psParameter) {
			var iCharIndex;
			var sDefn;

			sDefn = new String(psDefnString);

			iCharIndex = sDefn.indexOf("	");
			if (iCharIndex >= 0) {
				if (psParameter == "TYPE") return sDefn.substr(0, iCharIndex);
				sDefn = sDefn.substr(iCharIndex + 1);
				iCharIndex = sDefn.indexOf("	");
				if (iCharIndex >= 0) {
					if (psParameter == "TABLEID") return sDefn.substr(0, iCharIndex);
					sDefn = sDefn.substr(iCharIndex + 1);
					iCharIndex = sDefn.indexOf("	");
					if (iCharIndex >= 0) {
						if (psParameter == "COLUMNID") return sDefn.substr(0, iCharIndex);
						sDefn = sDefn.substr(iCharIndex + 1);
						iCharIndex = sDefn.indexOf("	");
						if (iCharIndex >= 0) {
							if (psParameter == "DISPLAY") return sDefn.substr(0, iCharIndex);
							sDefn = sDefn.substr(iCharIndex + 1);
							iCharIndex = sDefn.indexOf("	");
							if (iCharIndex >= 0) {
								if (psParameter == "SIZE") return sDefn.substr(0, iCharIndex);
								sDefn = sDefn.substr(iCharIndex + 1);
								iCharIndex = sDefn.indexOf("	");
								if (iCharIndex >= 0) {
									if (psParameter == "DECIMALS") return sDefn.substr(0, iCharIndex);
									sDefn = sDefn.substr(iCharIndex + 1);
									iCharIndex = sDefn.indexOf("	");
									if (iCharIndex >= 0) {
										if (psParameter == "HIDDEN") return sDefn.substr(0, iCharIndex);
										sDefn = sDefn.substr(iCharIndex + 1);
										iCharIndex = sDefn.indexOf("	");
										if (iCharIndex >= 0) {
											if (psParameter == "NUMERIC") return sDefn.substr(0, iCharIndex);
											sDefn = sDefn.substr(iCharIndex + 1);
											iCharIndex = sDefn.indexOf("	");
											if (iCharIndex >= 0) {
												if (psParameter == "HEADING") return sDefn.substr(0, iCharIndex);
												sDefn = sDefn.substr(iCharIndex + 1);
												iCharIndex = sDefn.indexOf("	");
												if (iCharIndex >= 0) {
													if (psParameter == "AVERAGE") return sDefn.substr(0, iCharIndex);
													sDefn = sDefn.substr(iCharIndex + 1);
													iCharIndex = sDefn.indexOf("	");
													if (iCharIndex >= 0) {
														if (psParameter == "COUNT") return sDefn.substr(0, iCharIndex);
														sDefn = sDefn.substr(iCharIndex + 1);

														if (psParameter == "TOTAL") return sDefn;
													}
												}
											}
										}
									}
								}
							}
						}
					}
				}
			}

			return "";
		}

		function getTableIDFromSelectedColumns(piColumnID) {

			var frmDef = window.dialogArguments.OpenHR.getForm("workframe", "frmDefinition");
			var frmUseful = window.dialogArguments.OpenHR.getForm("workframe", "frmUseful");
			var frmOrig = window.dialogArguments.OpenHR.getForm("workframe", "frmOriginalDefinition");


			if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
				frmDef.ssOleDBGridSelectedColumns.Redraw = false;
				frmDef.ssOleDBGridSelectedColumns.MoveFirst();
				for (var i = 0; i < frmDef.ssOleDBGridSelectedColumns.Rows; i++) {
					if (frmDef.ssOleDBGridSelectedColumns.Columns(0).Text == 'C') {
						if (frmDef.ssOleDBGridSelectedColumns.Columns(2).Text == piColumnID) {
							frmDef.ssOleDBGridSelectedColumns.Redraw = true;
							return frmDef.ssOleDBGridSelectedColumns.Columns(1).Text;
						}
					}
					frmDef.ssOleDBGridSelectedColumns.MoveNext();
				}
				frmDef.ssOleDBGridSelectedColumns.Redraw = true;
			}
			else {
				var dataCollection = frmOrig.elements;
				var sControlName;
				var tmpColID;
				var tmpTabID;

				if (dataCollection != null) {
					for (i = 0; i < dataCollection.length; i++) {
						sControlName = dataCollection.item(i).name;
						sControlName = sControlName.substr(0, 20);
						if (sControlName == "txtReportDefnColumn_") {
							tmpColID = selectedColumnParameter(dataCollection.item(i).value, 'COLUMNID');
							if (tmpColID == piColumnID) {
								tmpTabID = selectedColumnParameter(dataCollection.item(i).value, 'TABLEID');
								return tmpTabID;
							}
						}
					}
				}
			}
			return '';
		}

		function checkColumnOptions(pbFromCheckBox) {

			var frmPopup = document.getElementById("frmPopup");
			var parWin = window.dialogArguments;
			var sKey = new String('C' + frmPopup.cboColumn.options[frmPopup.cboColumn.selectedIndex].value);

			if (parWin.setGirdCol(sKey)) {
				var bBreak = parWin.getCurrentColProp('Break');
				var bPage = parWin.getCurrentColProp('Page');
				var bHidden = parWin.getCurrentColProp('Hidden');
				var bRepetition = parWin.getCurrentColProp('Repetition');

				if ((!pbFromCheckBox && bBreak == true) || (pbFromCheckBox && frmPopup.chkBOC.checked == true)) {
					checkbox_disable(frmPopup.chkBOC, false);
					frmPopup.chkPOC.checked = false;
					checkbox_disable(frmPopup.chkPOC, true);
				}
				else if ((!pbFromCheckBox && bPage == true) || (pbFromCheckBox && frmPopup.chkPOC.checked == true)) {
					frmPopup.chkBOC.checked = false;
					checkbox_disable(frmPopup.chkBOC, true);
					checkbox_disable(frmPopup.chkPOC, false);
				}
				else {
					checkbox_disable(frmPopup.chkBOC, false);
					checkbox_disable(frmPopup.chkPOC, false);
				}

				if (bHidden) {
					frmPopup.chkVOC.checked = false;
					checkbox_disable(frmPopup.chkVOC, true);
				}
				else {
					checkbox_disable(frmPopup.chkVOC, false);
				}

				if (bHidden || bRepetition) {
					frmPopup.chkSRV.checked = false;
					checkbox_disable(frmPopup.chkSRV, true);
				}
				else {
					checkbox_disable(frmPopup.chkSRV, false);
				}
			}
		}
	</script>

	<%
		Response.Write("<script type=""text/javascript"">" & vbCrLf)
		Response.Write("function setForm()" & vbCrLf)
		Response.Write("	{" & vbCrLf)
		Response.Write("	var frmPopup = document.getElementById('frmPopup');" & vbCrLf)
		Response.Write("	sAdd = frmPopup.cboColumn.value + '	' + frmPopup.cboColumn.options(frmPopup.cboColumn.options.selectedIndex).text;" & vbCrLf)
		
		Response.Write("	if (frmPopup.optAscending.checked == true) " & vbCrLf)
		Response.Write("		{" & vbCrLf)
		Response.Write("		sAdd = sAdd + '	' + 'Asc';" & vbCrLf)
		Response.Write("		}" & vbCrLf)
		Response.Write("	else " & vbCrLf)
		Response.Write("		{" & vbCrLf)
		Response.Write("		sAdd = sAdd + '	' + 'Desc';" & vbCrLf)
		Response.Write("		}" & vbCrLf)

		If (Session("utiltype") = 2) Then

			Response.Write("		if (frmPopup.chkBOC.checked == true) " & vbCrLf)
			Response.Write("			{" & vbCrLf)
			Response.Write("			sAdd = sAdd + '	' + '-1';" & vbCrLf)
			Response.Write("			}" & vbCrLf)
			Response.Write("		else " & vbCrLf)
			Response.Write("			{" & vbCrLf)
			Response.Write("			sAdd = sAdd + '	' + '0';" & vbCrLf)
			Response.Write("			}" & vbCrLf)

			Response.Write("		if (frmPopup.chkPOC.checked == true) " & vbCrLf)
			Response.Write("			{" & vbCrLf)
			Response.Write("			sAdd = sAdd + '	' + '-1';" & vbCrLf)
			Response.Write("			}" & vbCrLf)
			Response.Write("		else " & vbCrLf)
			Response.Write("			{" & vbCrLf)
			Response.Write("			sAdd = sAdd + '	' + '0';" & vbCrLf)
			Response.Write("			}" & vbCrLf)
			
			Response.Write("		if (frmPopup.chkVOC.checked == true) " & vbCrLf)
			Response.Write("			{" & vbCrLf)
			Response.Write("			sAdd = sAdd + '	' + '-1';" & vbCrLf)
			Response.Write("			}" & vbCrLf)
			Response.Write("		else " & vbCrLf)
			Response.Write("			{" & vbCrLf)
			Response.Write("			sAdd = sAdd + '	' + '0';" & vbCrLf)
			Response.Write("			}" & vbCrLf)

			Response.Write("		if (frmPopup.chkSRV.checked == true) " & vbCrLf)
			Response.Write("			{" & vbCrLf)
			Response.Write("			sAdd = sAdd + '	' + '-1';" & vbCrLf)
			Response.Write("			}" & vbCrLf)
			Response.Write("		else " & vbCrLf)
			Response.Write("			{" & vbCrLf)
			Response.Write("			sAdd = sAdd + '	' + '0';" & vbCrLf)
			Response.Write("			}" & vbCrLf)
		
			Response.Write("		sAdd = sAdd + '	' + getTableIDFromSelectedColumns(frmPopup.cboColumn.value);" & vbCrLf)

		End If
		
		Response.Write("	if(frmPopup.txtEditing.value == 'true') " & vbCrLf)
		Response.Write("		{" & vbCrLf)
		Response.Write("		window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').columns(0).text = frmPopup.cboColumn.value;" & vbCrLf)
		Response.Write("		window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').columns(1).text = frmPopup.cboColumn.options(frmPopup.cboColumn.options.selectedIndex).text;" & vbCrLf)

		Response.Write("		if (frmPopup.optAscending.checked == true) " & vbCrLf)
		Response.Write("			{" & vbCrLf)
		Response.Write("			window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').columns(2).text = 'Asc';" & vbCrLf)
		Response.Write("			}" & vbCrLf)
		Response.Write("		else " & vbCrLf)
		Response.Write("			{" & vbCrLf)
		Response.Write("			window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').columns(2).text = 'Desc';" & vbCrLf)
		Response.Write("			}" & vbCrLf)

		If (Session("utiltype") = 2) Then
			Response.Write("		var sKey = new String('C'+frmPopup.cboColumn.options[frmPopup.cboColumn.selectedIndex].value);" & vbCrLf)
			Response.Write("		if (window.dialogArguments.setGirdCol(sKey))" & vbCrLf)
			Response.Write("			{" & vbCrLf)
			Response.Write("			window.dialogArguments.updateCurrentColProp('break', frmPopup.chkBOC.checked);" & vbCrLf)
			Response.Write("			window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').columns(3).value = frmPopup.chkBOC.checked;" & vbCrLf)
			Response.Write("			window.dialogArguments.updateCurrentColProp('page', frmPopup.chkPOC.checked);" & vbCrLf)
			Response.Write("			window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').columns(4).value = frmPopup.chkPOC.checked;" & vbCrLf)
			Response.Write("			window.dialogArguments.updateCurrentColProp('value', frmPopup.chkVOC.checked);" & vbCrLf)
			Response.Write("			window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').columns(5).value = frmPopup.chkVOC.checked;" & vbCrLf)
			Response.Write("			window.dialogArguments.updateCurrentColProp('hide', frmPopup.chkSRV.checked);" & vbCrLf)
			Response.Write("			window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').columns(6).value = frmPopup.chkSRV.checked;" & vbCrLf)
			Response.Write("			}" & vbCrLf)
		End If

		Response.Write("		}" & vbCrLf)
		Response.Write("	else " & vbCrLf)
		Response.Write("		{" & vbCrLf)
	
		If (Session("utiltype") = 2) Then
			Response.Write("		var sKey = new String('C'+frmPopup.cboColumn.options[frmPopup.cboColumn.selectedIndex].value);" & vbCrLf)
			Response.Write("		if (window.dialogArguments.setGirdCol(sKey))" & vbCrLf)
			Response.Write("			{" & vbCrLf)
			Response.Write("			window.dialogArguments.updateCurrentColProp('break', frmPopup.chkBOC.checked);" & vbCrLf)
			Response.Write("			window.dialogArguments.updateCurrentColProp('page', frmPopup.chkPOC.checked);" & vbCrLf)
			Response.Write("			window.dialogArguments.updateCurrentColProp('value', frmPopup.chkVOC.checked);" & vbCrLf)
			Response.Write("			window.dialogArguments.updateCurrentColProp('hide', frmPopup.chkSRV.checked);" & vbCrLf)
			Response.Write("			}" & vbCrLf)
		End If
	
		Response.Write("		window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').additem(sAdd);" & vbCrLf)
		Response.Write("		window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').movelast();" & vbCrLf)
		Response.Write("		window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').selbookmarks.add(window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').bookmark);" & vbCrLf)
		Response.Write("		}" & vbCrLf)
		Response.Write("	self.close();" & vbCrLf)
		Response.Write("	return false;" & vbCrLf)

		Response.Write("	}" & vbCrLf)
		Response.Write("</script>" & vbCrLf)
	%>

<body>
	<div>
		<form id="frmPopup" name="frmPopup" onsubmit="return setForm();">
			<table align="center" class="outline">
				<tr>
					<td>
						<table class="invisible" width="100%">
							<tr height="10">
								<td height="10" colspan="5" align="center">
									<%
										' Get the order records.
										
										Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

										Dim rstSortOrder = objDataAccess.GetFromSP("spASRIntGetSortOrderColumns" _
											, New SqlParameter("psIncludedColumns", SqlDbType.VarChar, -1) With {.Value = Request("txtSortInclude")} _
											, New SqlParameter("psExcludedColumns", SqlDbType.VarChar, -1) With {.Value = Request("txtSortExclude")})
							
										If rstSortOrder.Rows.Count = 0 Then
									%>
									<h3>Warning</h3>
								</td>
							</tr>
							<tr>
								<td width="20" height="10"></td>
								<td colspan="3" style="text-align: center">There are no non-calculated columns to add!</td>
								<td width="20"></td>
							</tr>
							<tr>
								<td colspan="5" height="10">&nbsp;</td>
							</tr>
							<tr>
								<td colspan="5" height="10" align="center">
									<input type="button" class="btn" value="Close" name="cmdClose" style="WIDTH: 80px" width="80" id="cmdClose"
										onclick="self.close();" />
								</td>
							</tr>
							<tr>
								<td colspan="5" height="10"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>

		<%
			Response.End()
		End If
		%>

		<h3>Select Column</h3>
		<table>
			<tr>
				<td width="20">&nbsp;</td>
				<td nowrap>Column :</td>
				<td width="20">&nbsp;</td>
				<td>
					<%
						If Request("txtSortEditing") = "false" Then
					%>
					<input type="hidden" disabled id="txtEditing" name="txtEditing" value="false" />
					<%
					Else
					%>
					<input type="hidden" disabled id="txtEditing" name="txtEditing" value="true" />
					<%
					End If

					If (Session("utiltype") = 2) Then
					%>
					<select id="cboColumn" name="cboColumn" class="combo" onchange="checkColumnOptions();">
						<%
						Else
						%>
						<select id="cboColumn" name="cboColumn" class="combo">
							<%
							End If

							For Each objRow As DataRow In rstSortOrder.Rows
								If Not InStr(Request("txtSortExclude"), objRow("columnID").ToString()) Then
									Response.Write("<option value=" & Chr(34) & objRow("columnID") & Chr(34))
									If Request("txtSortEditing") = "true" Then
										If (CInt(objRow("columnID")) = CInt(Request("txtSortColumnID"))) Then
											Response.Write(" selected")
										End If
									End If
									Response.Write(">" & objRow("columnName").ToString() & "</option>" & vbCrLf)
								End If
							Next
							%>
						</select>
					</select>
				</td>
				<td width="20">&nbsp;</td>
			</tr>
			<tr height="10">
				<td colspan="4"></td>
			</tr>
			<tr>
				<td width="20">&nbsp;</td>
				<td nowrap>Order :</td>
				<td width="20">&nbsp;</td>
				<td>
					<input type="radio" checked id="optAscending" name="optOrder" value="radiobutton"
						<%
						If (Request("txtSortEditing") = "true") Then
							If (Request("txtSortOrder") = "Asc") Then
								Response.Write(" checked ")
							End If
						End If
%> />
					<label tabindex="-1"
						for="optAscending"
						class="radio">
						Ascending</label>
					<input type="radio" id="optDescending" name="optOrder" value="radiobutton"
						<%
						If (Request("txtSortEditing") = "true") Then
							If (Request("txtSortOrder") = "Desc") Then
								Response.Write(" checked ")
							End If
						End If
%> />
					<label tabindex="-1"
						for="optDescending"
						class="radio">
						Descending</label>
				</td>
				<td width="20">&nbsp;</td>
			</tr>
			<tr height="10">
				<td colspan="4"></td>
			</tr>
			<%
				If Session("utiltype") = 2 Then
			%>
			<tr>
				<td width="20">&nbsp;</td>
				<td nowrap>Break on Change :</td>
				<td width="20">&nbsp;</td>
				<td>
					<input type="checkbox" id="chkBOC" name="chkBOC" value="checkbox"
						<%
						If Request("txtSortBOC") = "-1" And (Request("txtSortEditing") = "true") Then
							Response.Write(" checked " & vbCrLf)
						End If%>
						onclick="checkColumnOptions(true);" />
				</td>
				<td width="20">&nbsp;</td>
			</tr>
			<tr height="10">
				<td colspan="4"></td>
			</tr>
			<tr>
				<td width="20">&nbsp;</td>
				<td nowrap>Page on Change :</td>
				<td width="20">&nbsp;</td>
				<td>
					<input type="checkbox" id="chkPOC" name="chkPOC" value="checkbox"
						<%
						If Request("txtSortPOC") = "-1" And (Request("txtSortEditing") = "true") Then
							Response.Write(" checked " & vbCrLf)
						End If
%>
						onclick="checkColumnOptions(true);" />
				</td>
				<td width="20">&nbsp;</td>
			</tr>
			<tr height="10">
				<td colspan="4"></td>
			</tr>
			<tr>
				<td width="20">&nbsp;</td>
				<td nowrap>Value on Change :</td>
				<td width="20">&nbsp;</td>
				<td>
					<input type="checkbox" id="chkVOC" name="chkVOC" value="checkbox"
						<%
						If Request("txtSortVOC") = "-1" And (Request("txtSortEditing") = "true") Then
							Response.Write(" checked " & vbCrLf)
						End If
%>>
				</td>
				<td width="20">&nbsp;</td>
			</tr>
			<tr height="10">
				<td colspan="4"></td>
			</tr>
			<tr>
				<td width="20">&nbsp;</td>
				<td nowrap>Suppress Repeated Values :</td>
				<td width="20">&nbsp;</td>
				<td>
					<input type="checkbox" id="chkSRV" name="chkSRV" value="checkbox"
						<%
						If Request("txtSortSRV") = "-1" And (Request("txtSortEditing") = "true") Then
							Response.Write(" checked " & vbCrLf)
						End If
%>
						onclick='checkColumnOptions(true);'>
				</td>
				<td width="20">&nbsp;</td>
			</tr>
			<%
			End If
			%>
			<tr height="20">
				<td colspan="4"></td>
			</tr>
			<tr>
				<td colspan="4">
					<table width="100%" class="invisible">
						<tr>
							<td>&nbsp;</td>
							<td width="10">
								<input id="cmdOK" type="button"
									class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br"
									value="OK" name="cmdOK" style="width: 80px" onclick="setForm();" />
							</td>
							<td width="10">&nbsp;</td>
							<td width="10">
								<input id="cmdCancel" type="button"
									class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br"
									value="Cancel" name="cmdCancel" style="width: 80px" onclick="self.close();" />
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</div>
</body>
</html>

<script type="text/javascript">
	util_SortOrderSelection_onload();
</script>