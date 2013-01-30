<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

    <script type="text/javascript">
        function PasswordChange_window_onload() {
            
            $("#workframe").attr("data-framesource", "PASSWORDCHANGE");

            var fOK = true;

            var sErrMsg = document.getElementById("txtErrorDescription").value;

            if (sErrMsg.length > 0) {
                fOK = false;
                //window.parent.frames("menuframe").OpenHR.messageBox(sErrMsg);
                OpenHR.messageBox(sErrMsg);

                window.parent.location.replace("login");
            }

            if (fOK == true) {
                // Get menu to refresh the menu.
                //window.parent.frames("menuframe").refreshMenu();
                menu_refreshMenu();

                if (frmPasswordChangeForm.txtUserSessionCount.value < 2) {
                    frmPasswordChangeForm.txtCurrentPassword.focus();
                }
            }
        }

        /* Validate the password change, and change the user's password on the SQL database if everything is okay. */
        function SubmitPasswordDetails() {
            var sCurrentPassword;
            var sNewPassword1;
            var sNewPassword2;
            var fChangeOK;
            var sErrorMessage;

            var frmPasswordChangeForm = OpenHR.getForm("workframe", "frmPasswordChangeForm");

            fChangeOK = true;
            sCurrentPassword = frmPasswordChangeForm.txtCurrentPassword.value.toLowerCase();
            sNewPassword1 = frmPasswordChangeForm.txtPassword1.value.toLowerCase();
            sNewPassword2 = frmPasswordChangeForm.txtPassword2.value.toLowerCase();

            /* Check that the two new passwords are the same. */
            if (sNewPassword1 != sNewPassword2) {
                //window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("The confirmation password is not correct.");
                OpenHR.messageBox("The confirmation password is not correct.");
                fChangeOK = false;
                frmPasswordChangeForm.txtPassword2.value = "";
                frmPasswordChangeForm.txtPassword2.focus();
            }

            /* Check that the new password is different to the old one. */
            if (sNewPassword1 == sCurrentPassword) {
                OpenHR.messageBox("The new password cannot be the same as the old one.");
                fChangeOK = false;
                frmPasswordChangeForm.txtPassword1.value = "";
                frmPasswordChangeForm.txtPassword2.value = "";
                frmPasswordChangeForm.txtPassword1.focus();
            }

            /* Check if the new password is long enough. */
            if ((fChangeOK) && (txtMinPasswordLength.value > 0) && (txtMinPasswordLength.value > sNewPassword1.length)) {
                sErrorMessage = "The password must be at least ";
                sErrorMessage = sErrorMessage.concat(txtMinPasswordLength.value);
                sErrorMessage = sErrorMessage.concat(" characters long.");
                OpenHR.messageBox(sErrorMessage);
                fChangeOK = false;
                frmPasswordChangeForm.txtPassword1.value = "";
                frmPasswordChangeForm.txtPassword2.value = "";
                frmPasswordChangeForm.txtPassword1.focus();
            }

            /* If everything is okay, submit the password change. */
            if (fChangeOK) {                
                OpenHR.submitForm(frmPasswordChangeForm);

            }
        }

        /* Return to the default page. */
        function cancelClick() {
            window.location = "main";
        }
    </script>

    <!--Client script to handle the screen events.-->

    <div <%=session("BodyTag")%>>
        <form action="passwordChange_Submit" method="post" id="frmPasswordChangeForm" name="frmPasswordChangeForm">
            <br>
            <table align="center" class="outline" cellpadding="5" cellspacing="0">
                <tr>
                    <td>
                        <table align="center" class="invisible" cellpadding="0" cellspacing="0">
                            <tr>
                                <td colspan="5" height="10"></td>
                            </tr>
                            <tr>
                                <td colspan="5">
                                    <h3 align="center">Change Password</h3>
                                </td>
                            </tr>
                            <%
                                ' Force password change only if there are no other users logged in with the same name.
                                Dim cmdCheckUserSessions = CreateObject("ADODB.Command")
                                cmdCheckUserSessions.CommandText = "spASRGetCurrentUsersCountOnServer"
                                cmdCheckUserSessions.CommandType = 4 ' Stored procedure.
                                cmdCheckUserSessions.ActiveConnection = Session("databaseConnection")

                                Dim prmCount = cmdCheckUserSessions.CreateParameter("count", 3, 2) ' 3=integer, 2=output
                                cmdCheckUserSessions.Parameters.Append(prmCount)

                                Dim prmUserName = cmdCheckUserSessions.CreateParameter("userName", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
                                cmdCheckUserSessions.Parameters.Append(prmUserName)
                                prmUserName.value = Session("Username")

                                Err.Clear()
                                cmdCheckUserSessions.Execute()
			
                                Dim iUserSessionCount = CLng(cmdCheckUserSessions.Parameters("count").Value)
                                cmdCheckUserSessions = Nothing

                                Response.Write("<INPUT type='hidden' id=txtUserSessionCount name=txtUserSessionCount value=" & iUserSessionCount & ">")
			
                                If iUserSessionCount < 2 Then
                            %>
                            <tr>
                                <td width="20"></td>
                                <td align="left" nowrap>Current Password :</td>
                                <td width="20"></td>
                                <td align="left">
                                    <input id="txtCurrentPassword" name="txtCurrentPassword" type="password" class="text" style="WIDTH: 200px">
                                </td>
                                <td width="20"></td>
                            </tr>
                            <tr>
                                <td width="20"></td>
                                <td align="left" nowrap>New Password :</td>
                                <td width="20"></td>
                                <td align="left">
                                    <input id="txtPassword1" name="txtPassword1" type="password" class="text" style="WIDTH: 200px">
                                </td>
                                <td width="20"></td>
                            </tr>
                            <tr>
                                <td width="20"></td>
                                <td align="left" nowrap>Confirm New Password :</td>
                                <td width="20"></td>
                                <td align="left">
                                    <input id="txtPassword2" name="txtPassword2" type="password" class="text" style="WIDTH: 200px">
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
                                            <input id="submitPasswordDetails" name="submitPasswordDetails" type="button" class="btn" value="OK" style="WIDTH: 75px" width="75"
                                                onclick="SubmitPasswordDetails()"
                                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                onfocus="try{button_onFocus(this);}catch(e){}"
                                                onblur="try{button_onBlur(this);}catch(e){}" />
                                        </td>
                                        <td width="20"></td>
                                        <td align="center">
                                            <input id="btnCancel" name="btnCancel" type="button" class="btn" value="Cancel" style="WIDTH: 75px" width="75"
                                                onclick="cancelClick()"
                                                onmouseover="try{button_onMouseOver(this);}catch(e){}"
                                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                onfocus="try{button_onFocus(this);}catch(e){}"
                                                onblur="try{button_onBlur(this);}catch(e){}" />
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
            <p></p>
            <%
            Else
            %>
            <tr>
                <td width="20"></td>
                <td colspan="3">
                    <%
                        Dim sErrorText = "You cannot change your password.<p>The account is currently being used by "
                        If iUserSessionCount > 2 Then
                            sErrorText = sErrorText & iUserSessionCount & " users"
                        Else
                            sErrorText = sErrorText & "another user"
                        End If
                        sErrorText = sErrorText & " in the system."
                        Response.Write(sErrorText)
                    %>
                </td>
                <td width="20"></td>
            </tr>

            <tr>
                <td colspan="5" height="20"></td>
            </tr>

            <tr>
                <td colspan="5" height="10" align="center">
                    <input type="button" value="Cancel" name="btnCancel" class="btn" style="WIDTH: 80px" width="80" id="Button1"
                        onclick="cancelClick()"
                        onmouseover="try{button_onMouseOver(this);}catch(e){}"
                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                        onfocus="try{button_onFocus(this);}catch(e){}"
                        onblur="try{button_onBlur(this);}catch(e){}" />
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
        %>
        </form>

        <form action="passwordChange_Submit" method="post" id="frmGoto" name="frmGoto">
            <input type="hidden" id="txtAction" name="txtAction">
            <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
        </form>
    </div>
    <%
        On Error Resume Next
	
        Dim sErrorDescription = ""
        
        ' Get the minimum password length.
        Dim cmdPwdLength = CreateObject("ADODB.Command")
        cmdPwdLength.CommandText = "sp_ASRIntGetMinimumPasswordLength"
        cmdPwdLength.CommandType = 4 ' Stored Procedure
        cmdPwdLength.ActiveConnection = Session("databaseConnection")
		
        Dim prmMinPasswordLength = cmdPwdLength.CreateParameter("MinPasswordLength", 3, 2) ' 3 = integer, 2 = output
        cmdPwdLength.Parameters.Append(prmMinPasswordLength)

        Err.Clear()
        cmdPwdLength.Execute()
        If (Err.Number <> 0) Then
            sErrorDescription = "The minimum password length could not be determined." & vbCrLf & FormatError(Err.Description)
        End If

        Response.Write("<INPUT type='hidden' id=txtMinPasswordLength name=txtMinPasswordLength value=" & cmdPwdLength.Parameters("MinPasswordLength").Value & ">" & vbCrLf)

        ' Release the ADO command object.
        cmdPwdLength = Nothing

        Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
    %>

    <script type="type/javascript"> PasswordChange_window_onload(); </script>

