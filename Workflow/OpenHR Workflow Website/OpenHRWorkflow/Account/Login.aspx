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
    
    <asp:Label runat="server" ID="lblWelcome" Text="Welcome"/>

    <table class="controlgrid">
        <tr>
            <td><asp:Label runat="server" ID="lblUserName" Text="Username"/></td>
            <td><asp:TextBox runat="server" ID="txtUserName"/></td>
        </tr>
        <tr>
            <td><asp:Label runat="server" ID="lblPassword" Text="Password"/></td>
            <td><asp:TextBox runat="server" ID="txtPassword" TextMode="Password" /></td>            
        </tr>
        <tr>
            <td><asp:Label runat="server" ID="lblRememberPwd" Text="Remember me"/></td>
            <td><asp:CheckBox runat="server" ID="chkRememberPwd" /></td>            
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