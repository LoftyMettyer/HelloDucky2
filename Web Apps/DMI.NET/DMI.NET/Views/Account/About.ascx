<%
	Dim iNumRows As Integer
%>
<script type="text/javascript">
	/* Return to the default page. */
	function cancelClick() {
		$("#About").dialog("close");
		return false;
	}
</script>

<form method="post" id="frmAboutForm" name="frmAboutForm">
	<br>
	<table style="text-align: center; border-spacing: 5px; border-collapse: collapse;" class="outline">
		<tr>
			<td>
				<table style="text-align: center; border-spacing: 5px; border-collapse: collapse;" class="invisible">
					<tr>
						<td colspan="6" height="10"></td>
					</tr>
					<tr>
						<td width="40"></td>
						<td colspan="4">
							<h3 align="center"></h3>
						</td>
						<td width="40"></td>
					</tr>
					<%If Len(Session("Server")) = 0 Then
							iNumRows = 12
						Else
							iNumRows = 16
						End If
					%>
					<tr>
						<td width="40" rowspan="<%=iNumRows %>"></td>
						<td width="20" rowspan="<%=iNumRows %>"></td>
						<td style="vertical-align: top; text-align: left; white-space: nowrap; padding-right: 10px;">OpenHR :&nbsp;
						</td>
						<td style="vertical-align: top; text-align: left; white-space: nowrap;">Version
                                <%=session("Version")%>
						</td>
						<td width="40" rowspan="<%=iNumRows %>"></td>
					</tr>
					<%If Len(Session("Server")) > 0 Then%>
					<tr>
						<td style="vertical-align: top; text-align: left; white-space: nowrap; padding-right: 10px;">Server :
						</td>
						<td style="vertical-align: top; text-align: left; white-space: nowrap;">
							<%=session("Server")%>
						</td>
					</tr>
					<tr>
						<td style="vertical-align: top; text-align: left; white-space: nowrap; padding-right: 10px;">Database :
						</td>
						<td style="vertical-align: top; text-align: left; white-space: nowrap;">
							<%=session("Database")%>
						</td>
					</tr>
					<tr>
						<td style="vertical-align: top; text-align: left; white-space: nowrap; padding-right: 10px;">Current user :
						</td>
						<td style="vertical-align: top; text-align: left; white-space: nowrap;">
							<%=session("Username")%>
						</td>
					</tr>
					<tr>
						<td style="vertical-align: top; text-align: left; white-space: nowrap; padding-right: 10px;">User Group :
						</td>
						<td style="vertical-align: top; text-align: left; white-space: nowrap;">
							<%=session("Usergroup")%>
						</td>
					</tr>
					<%End If%>
					<tr>
						<td colspan="2" style="vertical-align: top; text-align: left; white-space: nowrap;">
							<br />
							Copyright © Advanced Business Software and Solutions Ltd 2012
						</td>
					</tr>
					<tr>
						<td colspan="2" style="vertical-align: top; text-align: left; white-space: nowrap;">
							<a target="Advanced Website" href="http://www.advancedcomputersoftware.com/abs" class="hypertext"
								onmouseover="try{hypertextARef_onMouseOver(this);}catch(e){}"
								onmouseout="try{hypertextARef_onMouseOut(this);}catch(e){}"
								onfocus="try{hypertextARef_onFocus(this);}catch(e){}"
								onblur="try{hypertextARef_onBlur(this);}catch(e){}">http://www.advancedcomputersoftware.com/abs
							</a>
						</td>
					</tr>
					<tr>
						<td colspan="2" style="vertical-align: top; text-align: left; white-space: nowrap;">&nbsp;
						</td>
					</tr>
					<tr>
						<td colspan="2" style="vertical-align: top; text-align: left; white-space: nowrap;">Contacts for Support :
						</td>
					</tr>
					<tr>
						<td style="vertical-align: top; text-align: left; white-space: nowrap; padding-right: 10px;">Telephone :
						</td>
						<td style="vertical-align: top; text-align: left; white-space: nowrap;">
							<%If Session("SupportTelNo") = "" Then%>
                              08451 609 999
                            <%Else
                            		Response.Write(Session("SupportTelNo"))
                            	End If%>
						</td>
					</tr>
					<tr>
						<td style="vertical-align: top; text-align: left; white-space: nowrap; padding-right: 10px;">Email :
						</td>
						<td style="vertical-align: top; text-align: left; white-space: nowrap;">
							<%If Session("SupportEmail") = "" Then%>
							<a href="mailto://service.delivery@advancedcomputersoftware.com?subject=OpenHR Support Query - Data Manager Intranet" class="hypertext"
								onmouseover="try{hypertextARef_onMouseOver(this);}catch(e){}"
								onmouseout="try{hypertextARef_onMouseOut(this);}catch(e){}"
								onfocus="try{hypertextARef_onFocus(this);}catch(e){}"
								onblur="try{hypertextARef_onBlur(this);}catch(e){}">service.delivery@advancedcomputersoftware.com</a>
							<%Else%>
							<a href="mailto://<%=session("SupportEmail") %>?subject=OpenHR Support Query - Data Manager Intranet" class="hypertext"
								onmouseover="try{hypertextARef_onMouseOver(this);}catch(e){}"
								onmouseout="try{hypertextARef_onMouseOut(this);}catch(e){}"
								onfocus="try{hypertextARef_onFocus(this);}catch(e){}"
								onblur="try{hypertextARef_onBlur(this);}catch(e){}">
								<%=session("SupportEmail") %></a>
							<%End If%>
						</td>
					</tr>
					<tr>
						<td style="vertical-align: top; text-align: left; white-space: nowrap; padding-right: 10px;">Web site :
						</td>
						<td style="vertical-align: top; text-align: left; white-space: nowrap;">
							<%If Session("SupportWebpage") = "" Then%>
							<a target="AdvancedSupportWebsite" href="http://webfirst.advancedcomputersoftware.com" class="hypertext"
								onmouseover="try{hypertextARef_onMouseOver(this);}catch(e){}"
								onmouseout="try{hypertextARef_onMouseOut(this);}catch(e){}"
								onfocus="try{hypertextARef_onFocus(this);}catch(e){}"
								onblur="try{hypertextARef_onBlur(this);}catch(e){}">http://webfirst.advancedcomputersoftware.com</a>
							<%Else%>
							<a target="AdvancedSupportWebsite" href="<%=session("SupportWebpage") %>" class="hypertext"
								onmouseover="try{hypertextARef_onMouseOver(this);}catch(e){}"
								onmouseout="try{hypertextARef_onMouseOut(this);}catch(e){}"
								onfocus="try{hypertextARef_onFocus(this);}catch(e){}"
								onblur="try{hypertextARef_onBlur(this);}catch(e){}">
								<%=session("SupportWebpage") %></a>
							<%End If%>
						</td>
					</tr>
					<tr>
						<td colspan="2" style="vertical-align: top; text-align: left; white-space: nowrap;">&nbsp;
						</td>
					</tr>
					<tr>
						<td colspan="2" style="vertical-align: top; text-align: left; white-space: nowrap;">
							<a target="AdvancedConnectWebsite" href="http://www.advancedconnect.co.uk/" class="hypertext"
								onmouseover="try{hypertextARef_onMouseOver(this);}catch(e){}"
								onmouseout="try{hypertextARef_onMouseOut(this);}catch(e){}"
								onfocus="try{hypertextARef_onFocus(this);}catch(e){}"
								onblur="try{hypertextARef_onBlur(this);}catch(e){}">Visit Advanced Connect for the latest OpenHR news and events</a>
						</td>
					</tr>
					<tr>
						<td colspan="6" style="vertical-align: top; text-align: left; white-space: nowrap;">&nbsp;
						</td>
					</tr>
					<tr>
						<td colspan="6" style="text-align: center">
							<input id="btnCancel" name="btnCancel" type="button" class="btn" value="OK" style="width: 75px" width="75"
								onclick="cancelClick()"
								onmouseover="try{button_onMouseOver(this);}catch(e){}"
								onmouseout="try{button_onMouseOut(this);}catch(e){}"
								onfocus="try{button_onFocus(this);}catch(e){}"
								onblur="try{button_onBlur(this);}catch(e){}" />
						</td>
					</tr>
					<tr>
						<td colspan="7" height="10"></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</form>

