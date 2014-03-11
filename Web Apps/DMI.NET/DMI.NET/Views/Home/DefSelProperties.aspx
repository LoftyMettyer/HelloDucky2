﻿<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<% 
	Response.Expires = -1 
%>

<!DOCTYPE html>
<html>
<head>
	<script src="<%: Url.LatestContent("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script id="officebarscript" src="<%: Url.LatestContent("~/Scripts/officebar/jquery.officebar.js")%>" type="text/javascript"></script>
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-admin-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/fonts/SSI80v194934/style.css")%>" rel="stylesheet" />

	<title>OpenHR Intranet</title>

	<script type="text/javascript">
		function defselproperties_window_onload() {

			$("input[type=submit], input[type=button], button")
				.button();
			$("input").addClass("ui-widget ui-widget-content ui-corner-all");
			$("input").removeClass("text");

		}

	</script>

</head>

<body leftmargin="20" topmargin="20" bottommargin="20" rightmargin="5">
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
	
	Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

	Dim rsDefProp = objDataAccess.GetFromSP("sp_ASRIntDefProperties" _
		, New SqlParameter("intType", SqlDbType.Int) With {.Value = CInt(Request("utiltype"))} _
		, New SqlParameter("intID", SqlDbType.Int) With {.Value = CInt(Request("prop_id"))})
								
	Dim sCreated As String, sSaved As String, sRun As String
	
	If rsDefProp.Rows.Count = 0 Then
		sCreated = "<Unknown>"
		sSaved = "<Unknown>"
		sRun = "<Unknown>"
	Else
		sCreated = rsDefProp.Rows(0)("CreatedDate").ToString() & "  by " & rsDefProp.Rows(0)("Createdby").ToString
		If sCreated = "  by " Then sCreated = "<Unknown>"
		sSaved = rsDefProp.Rows(0)("SavedDate").ToString() & "  by " & rsDefProp.Rows(0)("Savedby").ToString()
		If sSaved = "  by " Then sSaved = "<Unknown>"
		sRun = rsDefProp.Rows(0)("RunDate").ToString() & "  by " & rsDefProp.Rows(0)("Runby").ToString()
		If sRun = "  by " Then sRun = "<Unknown>"
	End If
					 
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
						<select size=2 id=select1 name=select1 style="WIDTH: 300px; height:100%" class="combo">
<%
	
	Try
		Dim rsDefUsage = objDatabase.GetUtilityUsage(CInt(Request("utiltype")), CInt(Request("prop_id")))

		If rsDefUsage.Rows.Count = 0 Then
			Response.Write("<option>&lt;None&gt;</option>")
		Else
			For Each objRow As DataRow In rsDefUsage.Rows
				Dim sDescription As String = objRow("description").ToString()
				sDescription = Replace(sDescription, "<", "&lt;")
				sDescription = Replace(sDescription, ">", "&gt;")
				Response.Write("<option>" & sDescription & "</option>")
			Next
		End If

	Catch ex As Exception
		Throw
		
	End Try
					 
%>
						</select>
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
											onClick="self.close()" />
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
		defselproperties_window_onload();
</script>
</body>
</html>
