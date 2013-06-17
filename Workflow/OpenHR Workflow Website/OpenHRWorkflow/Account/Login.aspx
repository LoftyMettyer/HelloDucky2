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
    
    <table style="height: 100%; width: 100%">
        <tr style="height: 40px">
            <td style="width: 33%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnLogin"  runat="server" OnClientClick="return submitCheck();"/></td>
            <td style="width: 33%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnForgotPwd" runat="server"/></td>
            <td style="width: 33%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnRegister" runat="server" /></td>
        </tr>
        <tr style="height: 17px">
            <td style="width: 33%; text-align: center; overflow: hidden"><label runat="server" id="btnLogin_label"></label></td>
            <td style="width: 33%; text-align: center; overflow: hidden"><label runat="server" id="btnForgotPwd_label"></label></td>
            <td style="width: 33%; text-align: center; overflow: hidden"><label runat="server" id="btnRegister_label"></label></td>
        </tr>
    </table>

</asp:Content>