<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%
	' call the sp to retrieve the items that the current utility
	' is used in.
	Dim cmdUsage = CreateObject("ADODB.Command")
	cmdUsage.CommandText = "sp_ASRIntDefUsage"
	cmdUsage.CommandType = 4
	cmdUsage.ActiveConnection = Session("databaseConnection")

	Dim prmType = cmdUsage.CreateParameter("type", 3, 1)
	cmdUsage.Parameters.Append(prmType)
  prmType.value = cleanNumeric(session("utiltype"))
				
	Dim prmId = cmdUsage.CreateParameter("id", 3, 1)
	cmdUsage.Parameters.Append(prmId)
	prmId.value = cleanNumeric(Session("utilid"))

	Err.Clear()
	Dim rstUsage = cmdUsage.Execute
	
	' The util isnt used in any batch jobs, so we can delete it
	if rstUsage.eof then
		Response.Redirect("util_delete")
	end if
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
	Do While Not rstUsage.EOF
		Dim sDescription As String = rstUsage.Fields("description").Value
		sDescription = Replace(sDescription, "<", "&lt;")
		sDescription = Replace(sDescription, ">", "&gt;")

		Response.Write(sDescription & "<BR>" & vbCrLf)

		rstUsage.MoveNext()
	Loop
%>
			        </td>
					<td width=20>&nbsp;</td>
			    </tr>

			    <tr>
					<td height=10 colspan=3></td>
				</tr>
				 
			    <tr> 
			        <td colspan=3 height=10 align=center> 
			            <input id="cmdOK" name="cmdOK" type="button" value="OK" style="WIDTH: 75px" class="btn" onclick="OpenHR.submitForm(frmUsage);"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
			        </td>
			    </tr>
			    <tr>
					<td height=10 colspan=3></td>
				</tr>
				<%session("utilid") = Request.Form("utilid")%>
			</table>
		</TD>
	</TR>
</table>

<form name="frmUsage" method="post" action="defsel" id="frmUsage">
</form>

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto">
	<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>
</div>
