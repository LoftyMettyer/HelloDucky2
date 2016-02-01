<%@ Page Title="Title" Language="VB" Inherits="System.Web.Mvc.ViewPage" MasterPageFile="~/Views/Shared/Site.Master" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
	<script type="text/javascript">
		/* Validate the password change, and change the user's password
		on the SQL database if everything is okay. */
		function SubmitPasswordDetails() {
			var sCurrentPassword;
			var sNewPassword1;
			var sNewPassword2;
			var fChangeOK;
			var frmPasswordChangeForm = document.getElementById('frmPasswordChangeForm');

			fChangeOK = true;
			sCurrentPassword = frmPasswordChangeForm.txtCurrentPassword.value.toLowerCase();
			sNewPassword1 = frmPasswordChangeForm.txtPassword1.value.toLowerCase();
			sNewPassword2 = frmPasswordChangeForm.txtPassword2.value.toLowerCase();

			/* Check that the two new passwords are the same. */
			if (sNewPassword1 != sNewPassword2) {
				alert("The confirmation password is not correct.");
				fChangeOK = false;
				frmPasswordChangeForm.txtPassword2.value = "";
				frmPasswordChangeForm.txtPassword2.focus();
			}

			/* Check that the new password is different to the old one. */
			if (sNewPassword1 == sCurrentPassword) {
				alert("The new password cannot be the same as the old one.");
				fChangeOK = false;
				frmPasswordChangeForm.txtPassword1.value = "";
				frmPasswordChangeForm.txtPassword2.value = "";
				frmPasswordChangeForm.txtPassword1.focus();
			}

			/* If everything is okay, submit the password change. */
			if (fChangeOK) {

				//Don't use OPenHR.submitForm, as we don't want results to go to workframe.
				//OpenHR.submitForm(frmPasswordChangeForm);
				frmPasswordChangeForm.submit();
			}

		}

	</script>

	<div class="divLogin">
		<div class="ui-dialog-titlebar ui-widget-header loginTitleBar">
			<img alt="about OpenHR" title="About OpenHR Web" src="<%= Url.Content("~/Content/images/help32.png")%>" />
		</div>

		<div <%=session("BodyTag")%> class="centered" style="top: 190px; position: absolute; left: 37%;">

			<form action="forcedPasswordChange_Submit" method="post" id="frmPasswordChangeForm" name="frmPasswordChangeForm">
				<table class="outline" align="center" cellpadding="5" cellspacing="0">
					<tr>
						<td>
							<table align="center" class="invisible" cellpadding="0" cellspacing="0" width="100%" height="100%">
								<tr>
									<td>
										<table align="center" class="invisible" cellpadding="0" cellspacing="0">
											<tr>
												<td colspan="5" height="40">
													<h3 align="center">You must change your password</h3>
												</td>
											</tr>
											<tr>
												<td width="20"></td>
												<td align="left" nowrap>Current Password :</td>
												<td width="20"></td>
												<td align="left">
													<input id="txtCurrentPassword" name="txtCurrentPassword" type="password" class="text" style="width: 200px">
												</td>
												<td width="20"></td>
											</tr>
											<tr>
												<td width="20"></td>
												<td align="left" nowrap>New Password :</td>
												<td width="20"></td>
												<td align="left">
													<input id="txtPassword1" name="txtPassword1" type="password" class="text" style="width: 200px">
												</td>
												<td width="20"></td>
											</tr>
											<tr>
												<td width="20"></td>
												<td align="left" nowrap>Confirm New Password :</td>
												<td width="20"></td>
												<td align="left">
													<input id="txtPassword2" name="txtPassword2" type="password" class="text" style="width: 200px">
												</td>
												<td width="20"></td>
											</tr>
											<tr>
												<td colspan="5" height="20"></td>
											</tr>
                      

                      <tr style="text-align: left">
                        <td width="20"></td>
                        <td colspan="3" style="text-align:center">
                          <img src="<%=Url.Action("GetCaptcha","Generic")%>" />
                        </td>
                      </tr>

                      <tr style="text-align: left">
                        <td width="20"></td>
                        <td colspan="3" width="40" align="center">Please type the above characters to ensure that you a person, not an automated program submitting this request.</td>
                      </tr>
                                            
                     <tr>
                        <td>
                          <br />
                        </td>
                     </tr>                      
                      
                     <tr>
							          <td width="20"></td>
							          <td align="left" nowrap>Verify :</td>
							          <td width="20"></td>
							          <td align="left">
								          <input id="txtVerify" name="txtVerify" class="text" autocomplete="off" style="WIDTH: 200px; margin-top: 1px; margin-bottom: 1px">
							          </td>
							          <td width="20"></td>
						          </tr>
                                               
                     <tr>
                        <td>
                          <br />
                        </td>
                     </tr>           

											<tr>
												<td colspan="5" align="center">
													<input id="submitPasswordDetails" name="submitPasswordDetails" type="button" class="btn" value="OK" style="width: 100px" width="100" onclick="SubmitPasswordDetails()" />
												</td>
											</tr>
											<tr>
												<td colspan="5" height="10">

													<% 
														If Not Session("ErrorText") Is Nothing Then
															If Len(Session("ErrorText").ToString()) > 0 Then
																Response.Write(Session("ErrorText").ToString())
															End If
														End If
													%>


												</td>


											</tr>
										</table>
									</td>
								</tr>
							</table>
						</td>
					</tr>


				</table>

				<%=Html.AntiForgeryToken()%>

			</form>


		</div>
	</div>

	<script type="text/javascript">
		//Set up button click events
		$('.loginTitleBar img').click(function () {
			OpenHR.showAboutPopup();
		});
	</script>
</asp:Content>
