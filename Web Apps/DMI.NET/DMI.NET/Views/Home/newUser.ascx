<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>

<script type="text/javascript">
	function newUser_window_onload() {
		//window.parent.document.all.item("workframeset").cols = "*, 0";
		$("#workframe").attr("data-framesource", "NEWUSER");
		
		// Get menu to refresh the menu.
		//window.parent.frames("menuframe").refreshMenu();
		menu_refreshMenu();

		//Set focus on the dropdown list of users if it exists.
		var ctlNewUsers = frmNewUserForm.selNewUser;
		if (ctlNewUsers != null) {
			ctlNewUsers.focus();
		}
		
		$("#optionframe").hide();
		$("#workframe").show();
	}
</script>

<script type="text/javascript">
	/* Submit the new user login. */
	function SubmitNewUserDetails() {
		//frmNewUserForm.submit();
		OpenHR.submitForm(frmNewUserForm);
	}
	/* Return to the default page. */
	function cancelClick() {
		window.location.href = "main";  // "default.asp";
	}
	/* Go to the default page. */
	function okClick() {
		window.location.href = "main";  // "default.asp";
	}
</script>

<div <%=session("BodyTag")%>>
	<form action="newUser_Submit" method="post" id="frmNewUserForm" name="frmNewUserForm">

<%
			
	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
			
	Try
		Dim rstLogins = objDataAccess.GetDataTable("spASRIntGetAvailableLoginsFromAssembly", CommandType.StoredProcedure)

		If (rstLogins.Rows.Count = 0) Then
			' No available logins.
		%>
		<br>
		<table align="center" class="outline" cellpadding="5" cellspacing="0">
			<tr>
				<td>
					<table align="center" class="invisible" cellpadding="0" cellspacing="0">
						<tr>
							<td colspan="3" height="10"></td>
						</tr>
						<tr>
							<td colspan="3" align="center">
								<h3>New User</h3>
							</td>
						</tr>
						<tr>
							<td width="20"></td>
							<td>No available user logins.</td>
							<td width="20"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<%
		Else
			' Display the available logins.
		%>
		<br>
		<table align="center" class="outline" cellpadding="5" cellspacing="0">
			<tr>
				<td>
					<table align="center" class="invisible" cellpadding="0" cellspacing="0">
						<tr>
							<td colspan="5" height="10"></td>
						</tr>
						<tr>
							<td align="center" colspan="5">
								<h3>New User</h3>
							</td>
						</tr>
						<tr>
							<td width="20"></td>
							<td align="right" nowrap>User Login :</td>
							<td width="20"></td>
							<td align="left">
								<select id="selNewUser" class="combo" name="selNewUser" style="WIDTH: 200px;">
									<%
										For Each objRow As DataRow In rstLogins.Rows
										%>
											<option value="<%=Replace(objRow("name").ToString(), """", "&quot;")%>"><%=objRow("name").ToString()%></option>
										<%
										Next
									%>
								</select>
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td colspan="5" height="20"></td>
						</tr>

						<tr>
							<td colspan="5">
								<table class="invisible" cellspacing="0" cellpadding="0" align="center">
									<td align="center">
										<input id="submitNewUserDetails" name="submitNewUserDetails" type="button" class="btn" value="OK" style="WIDTH: 75px" width="75"
											onclick="SubmitNewUserDetails()" />
									</td>
								</table>
							</td>
						</tr>
						<tr>
							<td colspan="5" height="10"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<%

		End If

	Catch ex As Exception
			%>
		<br>
		<table align="center" class="outline" cellpadding="5" cellspacing="0">
			<tr>
				<td>
					<table align="center" class="invisible" cellpadding="0" cellspacing="0">
						<tr>
							<td colspan="3" height="10"></td>
						</tr>
						<tr>
							<td colspan="3" align="center">
								<h3>New User</h3>
							</td>
						</tr>
						<tr>
							<td width="20"></td>
							<td>Unable to get the list of available logins.
												<br>
								<%=FormatError(ex.Message)%>
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td colspan="3" height="20"></td>
						</tr>
						<tr>
							<td colspan="3" align="center">
								<input type="button" value="OK" name="GoBack" class="btn" style="HEIGHT: 24px; WIDTH: 75px" width="75" id="cmdGoBack"
									onclick="okClick()" />
							</td>
						</tr>
						<tr>
							<td colspan="3" height="10"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<%

	End Try

		%>

		<%=Html.AntiForgeryToken()%>
	</form>

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
			<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
			<%=Html.AntiForgeryToken()%>
	</form>

</div>

<script type="text/javascript">newUser_window_onload();</script>
