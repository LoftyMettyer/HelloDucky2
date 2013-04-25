<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<% 
	Response.Expires = -1 
%>

<!DOCTYPE html>
<html>
<head>
    <script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>

    <title>OpenHR Intranet</title>

    <script type="text/javascript">
        function defselproperties_window_onload() {
            // Resize the popup.
            var frmPopup = document.getElementById("frmPopup");
            var iResizeBy = frmPopup.offsetParent.scrollHeight - frmPopup.offsetParent.clientHeight;
            if (frmPopup.offsetParent.offsetHeight + iResizeBy > screen.height) {
                window.parent.dialogHeight = new String(screen.height) + "px";
            }
            else {
                var iNewHeight = new Number(window.parent.dialogHeight.substr(0, window.parent.dialogHeight.length - 2));
                iNewHeight = iNewHeight + iResizeBy;
                window.parent.dialogHeight = new String(iNewHeight) + "px";
            }
        }

    </script>

</head>

<body onload="onBlur=self.focus()" leftmargin="20" topmargin="20" bottommargin="20" rightmargin="5">
    <form name="frmPopup" id="frmPopup">
        <table align=center class="outline" cellPadding=5 cellSpacing=0 height=100%> 
    <tr>
	    <td>
		    <table align=center class="invisible" cellPadding=0 cellSpacing=0 height=100%> 
			    <tr height=10>
			        <td colSpan=5 height=10></td>
			    </tr>
			    <tr height=10>
			        <td colSpan=5>
			            <H3 align=center>Definition Properties</H3>
			        </td>
			    </tr>
<%
	Dim cmdDefPropRecords = CreateObject("ADODB.Command")
cmdDefPropRecords.CommandText = "sp_ASRIntDefProperties"
cmdDefPropRecords.CommandType = 4 ' Stored Procedure

	cmdDefPropRecords.ActiveConnection = Session("databaseConnection")

	Dim prmType = cmdDefPropRecords.CreateParameter("type", 3, 1)
	cmdDefPropRecords.Parameters.Append(prmType)
prmType.value = cleanNumeric(clng(Request("utiltype")))

	Dim prmID = cmdDefPropRecords.CreateParameter("id", 3, 1)
	cmdDefPropRecords.Parameters.Append(prmID)
prmID.value = cleanNumeric(clng(Request("prop_id")))

	Err.Clear()
	Dim rsDefProp = cmdDefPropRecords.Execute()

	Dim sCreated As String, sSaved As String, sRun As String
	
	If rsDefProp.BOF And rsDefProp.EOF Then
		sCreated = "<Unknown>"
		sSaved = "<Unknown>"
		sRun = "<Unknown>"
	Else
		sCreated = rsDefProp.Fields("CreatedDate").Value & "  by " & rsDefProp.Fields("Createdby").Value
		If sCreated = "  by " Then sCreated = "<Unknown>"
		sSaved = rsDefProp.Fields("SavedDate").Value & "  by " & rsDefProp.Fields("Savedby").Value
		If sSaved = "  by " Then sSaved = "<Unknown>"
		sRun = rsDefProp.Fields("RunDate").Value & "  by " & rsDefProp.Fields("Runby").Value
		If sRun = "  by " Then sRun = "<Unknown>"
	End If
				   
	rsDefProp = Nothing
	cmdDefPropRecords = Nothing
%>
			    <tr height=10> 
				    <td width=20></td>
				    <td nowrap>
				        Name :
				    </td>
					<td width=20></td>
				    <td>
				        <input name="textfield" style="WIDTH: 300px;" value ="<%Response.Write( replace(Request("prop_name"),chr(34),"&quot;")) %>" class="text textdisabled" disabled="disabled">
				    </td>
					<td width=20></td>
				</tr>
				<tr height=10> 
					<td width=20></td>
				    <td nowrap>Created :</td>
					<td width=20></td>
				    <td>
				        <input name="textfield2" style="WIDTH: 300px;" value="<% Response.Write( sCreated) %>" class="text textdisabled" disabled="disabled">
				    </td>
					<td width=20></td>
				</tr>
				<tr height=10>
					<td width=20></td>
				    <td nowrap>Last Save :</td>
					<td width=20></td>
				    <td>
				        <input name="textfield3" style="WIDTH: 300px" class="text textdisabled" disabled="disabled" value="<% Response.Write(sSaved)%>">
				    </td>
					<td width=20></td>
				</tr>
<% if Request("utiltype") <> 10 and _
	Request("utiltype") <> 11 and _
	Request("utiltype") <> 12 then %>
				<tr height=10>
					<td width=20></td>
				    <td nowrap>Last Run :</td>
					<td width=20></td>
				    <td>
				        <input name="textfield4" style="WIDTH: 300px" class="text textdisabled" disabled="disabled" value="<% Response.Write(sRun)%>">
				    </td>
					<td width=20></td>
				</tr>
<% end if%>			
				<tr>
					<td width=20 rowspan=4></td>
				    <td nowrap valign=top>Current Usage :</td>
					<td width=20 rowspan=4></td>
				    <td rowspan=4>
						<SELECT size=2 id=select1 name=select1 style="WIDTH: 300px; height:100%" class="combo">
<%
	cmdDefPropRecords = CreateObject("ADODB.Command")
cmdDefPropRecords.CommandText = "sp_ASRIntDefUsage"
cmdDefPropRecords.CommandType = 4 ' Stored Procedure

	cmdDefPropRecords.ActiveConnection = Session("databaseConnection")

	prmType = cmdDefPropRecords.CreateParameter("type", 3, 1)
	cmdDefPropRecords.Parameters.Append(prmType)
prmType.value = cleanNumeric(clng(Request("utiltype")))

	prmID = cmdDefPropRecords.CreateParameter("id", 3, 1)
	cmdDefPropRecords.Parameters.Append(prmID)
prmID.value = cleanNumeric(clng(Request("prop_id")))

	Err.Clear()
	rsDefProp = cmdDefPropRecords.Execute

if rsDefProp.BOF and rsDefProp.EOF then
		Response.Write("<OPTION>&lt;None&gt;</OPTION>")
else
	do while not rsDefProp.EOF
			Dim sDescription As String = CStr(rsDefProp.Fields("description").Value)
		sDescription = replace(sDescription, "<", "&lt;")
		sDescription = replace(sDescription, ">", "&gt;")
			Response.Write("<OPTION>" & sDescription & "</OPTION>")
		rsDefProp.MoveNext
	loop
end if
				   
	rsDefProp = Nothing
	cmdDefPropRecords = Nothing
%>
						</SELECT>
				    </td>
					<td width=20 rowspan=4></td>
				</tr>
				<tr>
				    <td>&nbsp;</td>
				</tr>
				<tr>
				    <td>&nbsp;</td>
				</tr>
				<tr>
				    <td>&nbsp;</td>
				</tr>

				<tr height=10> 
				    <td colspan=4 align=right> 
			            <input type="button" name="cmdClose" value="Close" style="WIDTH: 80px" width="80"
			                onClick="self.close()" 
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
                    </td>
        			<td width=20></td>
				</tr>
			    <tr height=10>
			        <td colSpan=5 height=5></td>
			    </tr>
		    </table>
		</td>
	</tr>
</table>

    </form>
<script type="text/javascript">
    window.onload = defselproperties_window_onload;
</script>
</body>
</html>
