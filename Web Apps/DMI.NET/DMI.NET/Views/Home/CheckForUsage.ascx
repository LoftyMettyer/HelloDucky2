

<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>


<%
	
	Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
	
	Dim rstUsage = objDatabase.GetUtilityUsage(CInt(CleanNumeric(Session("utiltype"))), CInt(CleanNumeric(Session("utilid"))))
		
	' The utility isnt used in any batch jobs, so we can delete it
	If rstUsage.Rows.Count = 0 Then
		Response.Redirect("util_delete")
	End If
%>

<div <%=session("BodyTag")%>>
<table align=center class="outline" cellPadding=5 cellSpacing=0>
	<TR>
		<TD>
			<table class="invisible" cellspacing="0" cellpadding=0>
					<tr> 
							<td colspan=3 height=10></td>
					</tr>
				
					<tr> 
							<td colspan=3 align=center> 
							<H3>Usage Check</H3>
							</td>
					</tr>
				
					<tr height=10> 
					<td width=20>&nbsp;</td>
							<td>
									Could not <%=session("action")%> '<%=session("utilname")%>' because it is used in the following:<BR><BR>
<%
	

	For Each objRow As DataRow In rstUsage.Rows

		Dim sDescription As String = objRow("description").ToString()
		sDescription = Replace(sDescription, "<", "&lt;")
		sDescription = Replace(sDescription, ">", "&gt;")

		Response.Write(sDescription & "<BR>" & vbCrLf)

	Next
%>
							</td>
					<td width=20>&nbsp;</td>
					</tr>

					<tr>
					<td height=10 colspan=3></td>
				</tr>
				 
					<tr> 
							<td colspan=3 height=10 align=center> 
									<input id="cmdOK" name="cmdOK" type="button" value="OK" style="WIDTH: 75px" class="btn" onclick="OpenHR.submitForm(frmUsage);" />
							</td>
					</tr>
					<tr>
					<td height=10 colspan=3></td>
				</tr>

			</table>
		</TD>
	</TR>
</table>
	
<form name="frmDelete" method="post" action="util_delete" id="frmDelete">
</form>

<form name="frmUsage" method="post" action="defsel" id="frmUsage">
</form>

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto">
	<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>
</div>
