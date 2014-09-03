<%@ import Namespace="System.Web.Configuration" %>
<%@ Import Namespace="DMI.NET.Code" %>
<%
	Dim iNumRows As Integer
%>
<html>
<head>
	<script type="text/javascript">
		/* Return to the default page. */
		function about_cancelClick() {
			$("#About").dialog("close");
			return false;
		}

		<%_txtLocalServerValue = ApplicationSettings.LoginPage_Server%>
		<%_txtLocalDatabaseValue = ApplicationSettings.LoginPage_Database%>
	</script>

	<script runat="server">
		Private _txtLocalServerValue As String
		Private _txtLocalDatabaseValue As String
</script>
</head>
<body>

<form method="post" id="frmAboutForm" name="frmAboutForm">
	<div class="pageTitleDiv" style="margin-bottom: 15px">
		<span class="pageTitle" id="PopupReportDefinition_PageTitle">About OpenHR</span>
	</div>

	<table style="text-align: center; border-spacing: 5px; border-collapse: collapse;" class="outline">
		<tr>
			<td>
				<table style="text-align: center; border-spacing: 5px; border-collapse: collapse;" class="invisible">
					<tr>
						<td width="40"></td>
						<td colspan="4">
							<h3 align="center"></h3>
						</td>
						<td width="40"></td>
					</tr>
					<%If Len(ApplicationSettings.LoginPage_Server) = 0 Then
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
						<td style="vertical-align: top; text-align: left; white-space: nowrap;">Version <%=session("Version")%></td>
						<td width="40" rowspan="<%=iNumRows %>"></td>
					</tr>

					<%If Len(ApplicationSettings.LoginPage_Server) > 0 Then%>
						<tr>
							<td style="vertical-align: top; text-align: left; white-space: nowrap; padding-right: 10px;">Server : 
							</td>
							<td style="vertical-align: top; text-align: left; white-space: nowrap;">
								<%=ApplicationSettings.LoginPage_Server%>
							</td>
						</tr>
						<tr>
							<td style="vertical-align: top; text-align: left; white-space: nowrap; padding-right: 10px;">Database : 
							</td>
							<td style="vertical-align: top; text-align: left; white-space: nowrap;">
								<%=ApplicationSettings.LoginPage_Database%>
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
					<%Else%>
						<%--Get Server and DB from web config--%>
						<tr>
							<td style="vertical-align: top; text-align: left; white-space: nowrap; padding-right: 10px;">Server : 
							</td>
							<td style="vertical-align: top; text-align: left; white-space: nowrap;">
								<%=_txtLocalServerValue%>
							</td>
						</tr>
						<tr>
							<td style="vertical-align: top; text-align: left; white-space: nowrap; padding-right: 10px;">Database : 
							</td>
							<td style="vertical-align: top; text-align: left; white-space: nowrap;">
								<%=_txtLocalDatabaseValue%>
							</td>
						</tr>
					<%End If%>
						<tr>
							<td style="vertical-align: top; text-align: left; white-space: nowrap; padding-right: 10px;">User Locale :
							</td>
							<td style="vertical-align: top; text-align: left; white-space: nowrap;">
								<span id="spnAbout_LocaleCultureName"></span>
							</td>
						</tr>
					<tr>
						<td colspan="2" style="vertical-align: top; text-align: left; white-space: nowrap;">
							<br />
							Copyright © Advanced Business Software and Solutions Ltd 2014
						</td>
					</tr>
					<tr>
						<td colspan="2" style="vertical-align: top; text-align: left; white-space: nowrap;">
							<a target="Advanced Website" href="http://www.advancedcomputersoftware.com/abs" class="hypertext">
								http://www.advancedcomputersoftware.com/abs
							</a>
						</td>
					</tr>
					<tr>
						<td colspan="2" style="vertical-align: top; text-align: left; white-space: nowrap;">&nbsp;
						</td>
					</tr>
					<tr>
						<td colspan="2" style="vertical-align: top; text-align: left; white-space: nowrap;">Contacts for Customer Services : 
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
							<a href="mailto://service.delivery@advancedcomputersoftware.com?subject=OpenHR Support Query - Intranet Login" class="hypertext">
								service.delivery@advancedcomputersoftware.com</a>
							<%Else%>
							<a href="mailto://<%=session("SupportEmail") %>?subject=OpenHR Support Query - Web Login" class="hypertext">
								<%=session("SupportEmail") %></a>
							<%End If%>
						</td>
					</tr>
					<tr>
						<td style="vertical-align: top; text-align: left; white-space: nowrap; padding-right: 10px;">Web site :
						</td>
						<td style="vertical-align: top; text-align: left; white-space: nowrap;">
							<%If Session("SupportWebpage") = "" Then%>
							<a target="AdvancedSupportWebsite" href="http://webfirst.advancedcomputersoftware.com" class="hypertext">
								http://webfirst.advancedcomputersoftware.com</a>
							<%Else%>
							<a target="AdvancedSupportWebsite" href="<%=session("SupportWebpage") %>" class="hypertext">
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
							<a target="AdvancedConnectWebsite" href="http://www.advancedconnect.co.uk/" class="hypertext">
								Visit Advanced Connect for the latest OpenHR news and events</a>
						</td>
					</tr>
					<tr>
						<td colspan="6" style="vertical-align: top; text-align: left; white-space: nowrap;">&nbsp;
						</td>
					</tr>
					
					<tr>
						<td colspan="7">
							Current Users
							<table id="currentLoggedInUsers">	
							</table>
						</td>
					</tr>

					<tr>
						<td colspan="6" style="text-align: center">
							<input id="btnCancel" name="btnCancel" type="button" class="btn" value="OK" style="width: 75px" width="75"
								onclick="about_cancelClick()"
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
	
</form></body>
</html>

<script type="text/javascript">
	$("#spnAbout_LocaleCultureName")[0].innerHTML = window.UserLocale;



	$(document).ready(function() {

		var licence = $.connection['LicenceHub'];

		licence['client'].currentUserList = function(userList) {

			$("#currentLoggedInUsers").jqGrid('GridUnload');

			$("#currentLoggedInUsers").jqGrid({
				datatype: 'jsonstring',
				datastr: userList,
				mtype: 'GET',
				jsonReader: {
					root: "rows", //array containing actual data
					page: "page", //current page
					total: "total", //total pages for the query
					records: "records", //total number of records
					repeatitems: false,
					id: "UserName" //index of the column with the PK in it
				},
				colNames: ['User Name', 'Device', 'Area'],
				colModel: [
					{ name: 'UserName', index: 'UserName' },
					{ name: 'Device', index: 'Device' },
					{ name: 'WebAreaName', index: 'WebAreaName' }
				],
				viewrecords: true,
				width: 450,
				height: 90,
				sortname: 'User',
				sortorder: "desc",
				rowNum: 10000,
			});
		}
	});



</script>


