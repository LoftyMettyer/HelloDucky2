<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

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
      On Error Resume Next

      ' Display a list of available logins if there are any.
      Dim cmdLogins = Server.CreateObject("ADODB.Command")
      cmdLogins.CommandText = "spASRIntGetAvailableLoginsFromAssembly"
      cmdLogins.CommandType = 4 ' Stored Procedure
      cmdLogins.ActiveConnection = Session("databaseConnection")

      Err.Clear()
      Dim rstLogins = cmdLogins.Execute
  
      If (Err.Number <> 0) Then
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
                <%=formatError(err.description)%>
              </td>
              <td width="20"></td>
            </tr>
            <tr>
              <td colspan="3" height="20"></td>
            </tr>
            <tr>
              <td colspan="3" align="center">
                <input type="button" value="OK" name="GoBack" class="btn" style="HEIGHT: 24px; WIDTH: 75px" width="75" id="cmdGoBack"
                  onclick="okClick()"
                  onmouseover="try{button_onMouseOver(this);}catch(e){}"
                  onmouseout="try{button_onMouseOut(this);}catch(e){}"
                  onfocus="try{button_onFocus(this);}catch(e){}"
                  onblur="try{button_onBlur(this);}catch(e){}" />
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
    Else
      If (rstLogins.BOF And rstLogins.EOF) Then
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
            <tr>
              <td colspan="3" height="20"></td>
            </tr>
            <tr>
              <td colspan="3" align="center">
                <input type="button" class="btn" value="OK" name="GoBack" style="HEIGHT: 24px; WIDTH: 75px" width="75" id="Button1"
                  onclick="okClick()"
                  onmouseover="try{button_onMouseOver(this);}catch(e){}"
                  onmouseout="try{button_onMouseOut(this);}catch(e){}"
                  onfocus="try{button_onFocus(this);}catch(e){}"
                  onblur="try{button_onBlur(this);}catch(e){}" />
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
                    Do While Not rstLogins.EOF
                  %>
                  <option value="<%=replace(rstLogins.Fields("name").Value, """", "&quot;")%>"><%=rstLogins.Fields("name").Value%></option>
                  <%
                    rstLogins.MoveNext()
                  Loop
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
                      onclick="SubmitNewUserDetails()"
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
    <%
    End If

    ' Release the ADO recordset and command objects.
    rstLogins.close()
  End If

  rstLogins = Nothing
  cmdLogins = Nothing
    %>
  </form>

  <script type="text/javascript">newUser_window_onload();</script>


    <form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
        <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
    </form>


</div>
