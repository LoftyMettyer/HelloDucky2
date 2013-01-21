<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css"/>
<script src="<%: Url.Content("~/Scripts/jquery-1.8.2.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/openhr.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>
<title>OpenHR Intranet</title>

<script type="text/javascript">
<!--
	window.onload = function() {

		var iResizeBy, iNewWidth, iNewHeight, iNewLeft, iNewTop;
		var frmPopup = document.getElementById("frmPopup");
		
		// Resize the grid to show all prompted values.
		iResizeBy = frmPopup.offsetParent.scrollWidth	- frmPopup.offsetParent.clientWidth;
		if (frmPopup.offsetParent.offsetWidth + iResizeBy > screen.width) {
			window.dialogWidth = new String(screen.width) + "px";
		}
		else {
			iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length-2));
			iNewWidth = iNewWidth + iResizeBy;
			window.dialogWidth = new String(iNewWidth) + "px";
		}

		<% If (session("utiltype") = 2) then %>
		checkColumnOptions(false);
		<%End If%>
	
		iResizeBy = frmPopup.offsetParent.scrollHeight	- frmPopup.offsetParent.clientHeight;
		if (frmPopup.offsetParent.offsetHeight + iResizeBy > screen.height) {
			window.dialogHeight = new String(screen.height) + "px";
		}
		else {
			iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length-2));
			iNewHeight = iNewHeight + iResizeBy;
			window.dialogHeight = new String(iNewHeight) + "px";
		}
		
		iNewLeft = (screen.width - frmPopup.offsetParent.offsetWidth) / 2;
		iNewTop = (screen.height - frmPopup.offsetParent.offsetHeight) / 2;
		
		window.dialogLeft = new String(iNewLeft) + "px";
		window.dialogTop = new String(iNewTop) + "px";
	}
-->	
</SCRIPT>

<script type="text/javascript">
<!--

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
		var frmDef = window.dialogArguments.parent.frames("workframe").document.forms("frmDefinition");
		var frmUseful = window.dialogArguments.parent.frames("workframe").document.forms("frmUseful");
		var frmOrig = window.dialogArguments.parent.frames("workframe").document.forms("frmOriginalDefinition");

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
		var parWin = window.dialogArguments.parent.frames("workframe");
		var sKey = new String('C' + frmPopup.cboColumn.options[frmPopup.cboColumn.selectedIndex].value);

		if (window.dialogArguments.parent.frames("workframe").setGirdCol(sKey)) {
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
	
-->
</SCRIPT>

<%
	Response.Write("<script>" & vbCrLf)
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

	if (session("utiltype") = 2) then

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

	end if
		
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

	if (session("utiltype") = 2) then
		Response.Write("		var sKey = new String('C'+frmPopup.cboColumn.options[frmPopup.cboColumn.selectedIndex].value);" & vbCrLf)
		Response.Write("		if (window.dialogArguments.parent.frames('workframe').setGirdCol(sKey))" & vbCrLf)
		Response.Write("			{" & vbCrLf)
		Response.Write("			window.dialogArguments.parent.frames('workframe').updateCurrentColProp('break', frmPopup.chkBOC.checked);" & vbCrLf)
		Response.Write("			window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').columns(3).value = frmPopup.chkBOC.checked;" & vbCrLf)
		Response.Write("			window.dialogArguments.parent.frames('workframe').updateCurrentColProp('page', frmPopup.chkPOC.checked);" & vbCrLf)
		Response.Write("			window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').columns(4).value = frmPopup.chkPOC.checked;" & vbCrLf)
		Response.Write("			window.dialogArguments.parent.frames('workframe').updateCurrentColProp('value', frmPopup.chkVOC.checked);" & vbCrLf)
		Response.Write("			window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').columns(5).value = frmPopup.chkVOC.checked;" & vbCrLf)
		Response.Write("			window.dialogArguments.parent.frames('workframe').updateCurrentColProp('hide', frmPopup.chkSRV.checked);" & vbCrLf)
		Response.Write("			window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').columns(6).value = frmPopup.chkSRV.checked;" & vbCrLf)
		Response.Write("			}" & vbCrLf)
	end if

	Response.Write("		}" & vbCrLf)
	Response.Write("	else " & vbCrLf)
	Response.Write("		{" & vbCrLf)
	
	if (session("utiltype") = 2) then
		Response.Write("		var sKey = new String('C'+frmPopup.cboColumn.options[frmPopup.cboColumn.selectedIndex].value);" & vbCrLf)
		Response.Write("		if (window.dialogArguments.parent.frames('workframe').setGirdCol(sKey))" & vbCrLf)
		Response.Write("			{" & vbCrLf)
		Response.Write("			window.dialogArguments.parent.frames('workframe').updateCurrentColProp('break', frmPopup.chkBOC.checked);" & vbCrLf)
		Response.Write("			window.dialogArguments.parent.frames('workframe').updateCurrentColProp('page', frmPopup.chkPOC.checked);" & vbCrLf)
		Response.Write("			window.dialogArguments.parent.frames('workframe').updateCurrentColProp('value', frmPopup.chkVOC.checked);" & vbCrLf)
		Response.Write("			window.dialogArguments.parent.frames('workframe').updateCurrentColProp('hide', frmPopup.chkSRV.checked);" & vbCrLf)
		Response.Write("			}" & vbCrLf)
	end if
	
	Response.Write("		window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').additem(sAdd);" & vbCrLf)
	Response.Write("		window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').movelast();" & vbCrLf)
	Response.Write("		window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').selbookmarks.add(window.dialogArguments.document.getElementById('ssOleDBGridSortOrder').bookmark);" & vbCrLf)
	Response.Write("		}" & vbCrLf)
	Response.Write("	self.close();" & vbCrLf)
	Response.Write("	return false;" & vbCrLf)

	Response.Write("	}" & vbCrLf)
	Response.Write("</script>" & vbCrLf)

%>
	
<form id=frmPopup name=frmPopup onsubmit="return setForm();">
	
<table align=center class="outline" cellpadding=5 cellspacing=0 width="95%">
	<tr>
		<td>
			<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
				<tr height=10> 
					<td height=10 colspan=5 align=center> 
<%
	' Get the order records.
	Dim cmdSortOrder = CreateObject("ADODB.Command")
	cmdSortOrder.CommandType = 4
	cmdSortOrder.ActiveConnection = Session("databaseConnection")
	cmdSortOrder.CommandText = "spASRIntGetSortOrderColumns"

	Dim prmIncluded = cmdSortOrder.CreateParameter("included", 200, 1, 8000) ' 200 = varchar, 1 = input, 8000=size
	cmdSortOrder.Parameters.Append(prmIncluded)
	prmIncluded.value = Request("txtSortInclude")

	Dim prmExcluded = cmdSortOrder.CreateParameter("excluded", 200, 1, 8000) ' 200 = varchar, 1 = input, 8000=size
	cmdSortOrder.Parameters.Append(prmExcluded)
	prmExcluded.value = Request("txtSortExclude")

	Dim rstSortOrder = cmdSortOrder.Execute
	
	if rstSortOrder.eof then
%>	
					    <H3>Error</H3>
                    </td>
				</tr> 
			    <tr> 
			        <td width=20 height=10></td> 
				    <td colspan=3>The are no non-calculated columns to add!</td>
				    <td width=20></td> 
			    </tr>
			    <tr>
			        <td colspan=5 height=10>&nbsp;</td>
			    </tr>
			    <tr> 
			        <td colspan=5 height=10 align=center> 
				        <input TYPE=button class="btn" VALUE=Close NAME=cmdClose style="WIDTH: 80px" width=80 id=cmdClose
				            OnClick=self.close();
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
				    </td>
			    </tr>
			        <tr> 
			            <td colspan=5 height=10></td>
		            </tr>
			    </table>
	    </td>
    </tr>
</table>
</form>
</body>
<%
		Response.End 
	end if
%>	
		
					    <H3>Select Column</H3>
				    </td>
			    </tr>
			    <tr>
			        <td width=20>&nbsp;</td>
				    <td nowrap>Column :</td>
				    <td width=20>&nbsp;</td>
				    <td>
<%
	if Request("txtSortEditing") = "false" then
%>	
					    <input type=hidden disabled id=txtEditing name=txtEditing value="false" />
<%
	else
%>
					    <input type=hidden disabled id=txtEditing name=txtEditing value="true" />
<%
	end if

	if (session("utiltype") = 2) then
%>
					    <select id=cboColumn name=cboColumn style="WIDTH: 100%" class="combo" onchange="checkColumnOptions();">
<%
	else
%>
					    <select id=cboColumn name=cboColumn style="WIDTH: 100%" class="combo">
<%
	end if

	do until rstSortOrder.eof
		if not instr(Request("txtSortExclude"), rstSortOrder.fields("columnID").value) then
		Response.Write("<option value=" & Chr(34) & rstSortOrder.fields("columnID").value & Chr(34))
			if Request("txtSortEditing") = "true" then
				if (rstSortOrder.fields("columnID").value = clng(Request("txtSortColumnID"))) then
				Response.Write(" selected")
				end if
			end if
		Response.Write(">" & rstSortOrder.fields("columnName").value & "</option>" & vbCrLf)
		end if
		rstSortOrder.movenext
	loop
%>
    					</select>
				    </td>
				    <td width=20>&nbsp;</td>
			    </tr>
			    <tr height=10> 
				    <td colspan="4"></td>
			    </tr> 
			    <tr>
					<td width=20>&nbsp;</td>
					<td nowrap>Order :</td>
					<td width=20>&nbsp;</td>
					<td>
						<input type="radio" checked id=optAscending name="optOrder" value="radiobutton"
<%
	if (Request("txtSortEditing") = "true") then
		if (Request("txtSortOrder") = "Asc") then 
			Response.Write( " checked ")
		end if
	end if
%>
                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                            onfocus="try{radio_onFocus(this);}catch(e){}"
                            onblur="try{radio_onBlur(this);}catch(e){}"/>
                        <label tabindex="-1"
                            for="optAscending"
                            class="radio"
                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Ascending</label>
						<input type="radio" id="optDescending" name="optOrder" value="radiobutton"
<%
	if (Request("txtSortEditing") = "true") then
		if (Request("txtSortOrder") = "Desc") then 
			Response.Write (" checked ")
		end if
	end if
%>
                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                            onfocus="try{radio_onFocus(this);}catch(e){}"
                            onblur="try{radio_onBlur(this);}catch(e){}"/>
                        <label tabindex="-1"
                            for="optDescending"
                            class="radio"
                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">Descending</label>
    				</td>
					<td width=20>&nbsp;</td>
				</tr>
				<tr height=10> 
					<td colspan="4"></td>
				</tr> 
<%
	if session("utiltype") = 2 then
%>	
		        <tr>
					<td width=20>&nbsp;</td>
					<td nowrap>Break on Change :</td>
					<td width=20>&nbsp;</td>
					<td>
						<input type="checkbox" id=chkBOC name=chkBOC value="checkbox" 
<%
		if Request("txtSortBOC") = "-1" and (Request("txtSortEditing") = "true") then
			Response.Write (" checked " & vbcrlf)
		end if
%>
                            onclick="checkColumnOptions(true);" />
					</td>
					<td width=20>&nbsp;</td>
				</tr>
				<tr height=10> 
					<td colspan=4></td>
				</tr> 
				<tr>
					<td width=20>&nbsp;</td>
					<td nowrap>Page on Change :</td>
					<td width=20>&nbsp;</td>
					<td>
						<input type="checkbox" id=chkPOC name=chkPOC value="checkbox"
<%
		if Request("txtSortPOC") = "-1" and (Request("txtSortEditing") = "true") then
			Response.Write (" checked " & vbcrlf)
		end if
%>
                            onclick="checkColumnOptions(true);" />
					</td>
					<td width=20>&nbsp;</td>
				</tr>
				<tr height=10> 
					<td colspan=4></td>
				</tr> 
				<tr>
					<td width=20>&nbsp;</td>
					<td nowrap>Value on Change :</td>
					<td width=20>&nbsp;</td>
					<td>
						<input type="checkbox" id=chkVOC name=chkVOC value="checkbox"
<%
		if Request("txtSortVOC") = "-1" and (Request("txtSortEditing") = "true") then
			Response.Write (" checked " & vbcrlf)
		end if
%>
						>
					</td>
					<td width=20>&nbsp;</td>
				</tr>
				<tr height=10> 
					<td colspan=4></td>
				</tr> 
				<tr>
					<td width=20>&nbsp;</td>
					<td nowrap>Suppress Repeated Values :</td>
					<td width=20>&nbsp;</td>
					<td>
						<input type="checkbox" id=chkSRV name=chkSRV value="checkbox"
<%
		if Request("txtSortSRV") = "-1" and (Request("txtSortEditing") = "true") then
			Response.Write( " checked " & vbcrlf)
		end if
%>
						    onclick='checkColumnOptions(true);'>
					</td>
					<td width=20>&nbsp;</td>
				</tr>
<%
	end if
%>
				<tr height=20> 
					<td colspan="4"></td>
				</tr> 
				<tr> 
					<td colspan="4"> 
						<table WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD>&nbsp;</TD>
								<TD width=10>
									<input id=cmdOK type=button class="btn" value=OK name=cmdOK style="WIDTH: 80px" width="80" onclick="setForm()"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD width=10>&nbsp;</TD>
								<TD width=10>
									<input id=cmdCancel type=button class="btn" value=Cancel name=cmdCancel style="WIDTH: 80px" width="80" onclick="self.close();" 
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
							</tr>
						</table>
    			    </td>
				</tr>
				<tr height=10> 
					<td colspan="4"></td>
				</tr> 
			</table>
        </td>
	</tr>
</table>
</form>
</body>
</html>
