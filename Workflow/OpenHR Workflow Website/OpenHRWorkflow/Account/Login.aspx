<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="Login.aspx.vb" Inherits="Login" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">
    <script type="text/javascript">
    // <!CDATA[
        window.onload = function() {
            document.getElementById('ctl00_mainCPH_txtUserName').focus();
        };

        function submitCheck() {

            var header = 'Login Failed';

            if (document.getElementById('ctl00_mainCPH_txtUserName').value.length === 0) {
                showDialog(header, 'Username is required.');
                return false;
            }
            if (document.getElementById('ctl00_mainCPH_txtPassword').value.length === 0) {
                showDialog(header, 'Password is required.');
                return false;
            }
            return true;
        }
    // ]]>
    </script>
</asp:Content>

<asp:Content ID="main" ContentPlaceHolderID="mainCPH" Runat="Server">
    
    <label id="lblWelcome" runat="server">Welcome</label>

    <table class="controlgrid">
        <tr>
            <td><label id="lblUserName" runat="server">Username</label></td>
            <td ><input id="txtUserName" runat="server" type="text" /></td>
        </tr>
        <tr>
            <td><label id="lblPassword" runat="server">Password</label></td>
            <td><input id="txtPassword" runat="server" type="password"/></td>
        </tr>
        <tr>
            <td><label  id="lblRememberPwd" runat="server">Remember me</label></td>
            <td><input id="chkRememberPwd" runat="server" type="checkbox" /></td>
        </tr>
    </table>
                             
    <div class="copyright">Copyright © Advanced Business Software and Solutions Ltd 2012</div>

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
    <ol class="footer-buttons col3">
        <li>
            <asp:LinkButton ID="btnLoginButton" runat="server"  OnClientClick="return submitCheck();">
                <asp:Image runat="server" ID="btnLogin"/>
                <asp:Label runat="server" ID="btnLogin_Label"/>
            </asp:LinkButton>
        </li>
        <li>
            <asp:HyperLink runat="server" NavigateUrl="~/Account/ForgottenLogin.aspx">
                <asp:Image runat="server" ID="btnForgotPwd"  />
                <asp:Label runat="server" ID="btnForgotPwd_label" />           
            </asp:HyperLink>
        </li>
        <li>
            <asp:HyperLink runat="server" NavigateUrl="~/Account/Registration.aspx">
                <asp:Image runat="server" ID="btnRegister"/>
                <asp:Label runat="server" ID="btnRegister_label"/>
            </asp:HyperLink>
        </li>
    </ol>

</asp:Content>