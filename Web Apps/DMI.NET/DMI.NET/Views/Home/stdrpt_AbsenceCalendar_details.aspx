<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

		<form id="frmDetails" name="frmDetails">
				<table class="outline" cellpadding="0" cellspacing="7" width="100%">
						<tr>
								<td>Start&nbsp;Date :</td>
								<td><%Response.Write(Request("txtStartDate"))%>&nbsp;
		<%Response.Write(Request("txtStartSession"))%>
								</td>
						</tr>
						<tr>
								<td>End Date :</td>
								<td><%Response.Write(Request("txtEndDate"))%>&nbsp;
		<%Response.Write(Request("txtEndSession"))%>
								</td>
						</tr>
						<tr>
								<td>Duration :</td>
								<td><%Response.Write(Request("txtDuration"))%></td>
						</tr>
						<tr>
								<td>Type :</td>
								<td><%Response.Write(Request("txtType"))%></td>
						</tr>
						<tr>
								<td>Type Code :</td>
								<td><%Response.Write(Request("txtTypeCode"))%></td>
						</tr>
						<tr>
								<td>Calendar Code :</td>
								<td><%Response.Write(Request("txtCalCode"))%></td>
						</tr>
						<tr>
								<td>Reason :</td>
								<td><%Response.Write(Request("txtReason"))%></td>
						</tr>

						<% If Request("txtDisableRegions") = "False" Then%>
						<tr>
								<td>Region :</td>
								<td><%Response.Write(Request("txtRegion"))%></td>
						</tr>
						<% End If%>

						<% If Request("txtDisableWPs") = "False" Then%>
						<tr>
								<td>Working Pattern :</td>
								<td>
										<%
											Dim objAbsenceCalendar As New HR.Intranet.Server.AbsenceCalendar
											Response.Write(objAbsenceCalendar.HTML_WorkingPattern(Request("txtWorkingPattern")))
											objAbsenceCalendar = Nothing
										%>
								</td>
						</tr>
						<% End If%>
	
	</table>
</form>

<div style="text-align: center; padding-top: 20px">
	<input id="cmdOK" type="button" value="OK" class="btn" name="cmdOK" onclick="closeAbsenceDetails();" />
</div>

<script type="text/javascript">

	function closeAbsenceDetails() {
		$("#DisplayAbsenceCalendarEventDetail").dialog("close");
	}

	$("#DisplayAbsenceCalendarEventDetail").dialog('option', 'title', 'Absence Details');

</script>