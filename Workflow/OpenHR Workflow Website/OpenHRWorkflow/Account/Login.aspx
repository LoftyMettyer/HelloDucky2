<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="Login.aspx.vb" Inherits="Login" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">
    <script type="text/javascript">
    // <!CDATA[
        function submitCheck() {
            return true;
            var header = 'Login Failed';

            if (document.getElementById('<%= txtUserName.ClientID %>').value.trim().length === 0) {
                showDialog(header, 'Username is required.');
                return false;
            }
            if (document.getElementById('<%= txtPassword.ClientID %>').value.trim().length === 0) {
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
    
    <asp:ValidationSummary runat="server" ShowMessageBox="False" />

    <table class="controlgrid">
        <tr>
            <td><asp:Label runat="server" ID="lblUserName" Text="Username" AssociatedControlID="txtUserName" /></td>
            <td>
                <asp:TextBox runat="server" ID="txtUserName"/>
                <%--    TODO finish or remove --%>
                <asp:RequiredFieldValidator runat="server" ControlToValidate="txtUserName" ErrorMessage="Username is required"></asp:RequiredFieldValidator>
            </td>
        </tr>
        <tr>
            <td><asp:Label runat="server" ID="lblPassword" Text="Password"/></td>
            <td><asp:TextBox runat="server" ID="txtPassword" TextMode="Password" /></td>            
        </tr>
        <tr>
            <td><asp:Label runat="server" ID="lblRememberPwd" Text="Remember me" AssociatedControlID="chkRememberPwd" /></td>
            <td><asp:CheckBox runat="server" ID="chkRememberPwd" /></td>            
        </tr>
    </table>
                             
    <div class="copyright">Copyright © Advanced Business Software and Solutions Ltd 2012</div>

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
    <ul class="footer-buttons col3">
        <li>
            <asp:LinkButton runat="server" ID="btnLogin" OnClientClick="return submitCheck();">
                <asp:Image runat="server"/>
                <asp:Label runat="server"/>
            </asp:LinkButton>
        </li>
        <li>
            <asp:HyperLink runat="server" ID="btnForgotPwd" NavigateUrl="~/Account/ForgottenLogin.aspx">
                <asp:Image runat="server"/>
                <asp:Label runat="server"/>
            </asp:HyperLink>
        </li>
        <li>
            <asp:HyperLink runat="server" ID="btnRegister" NavigateUrl="~/Account/Registration.aspx">
                <asp:Image runat="server"/>
                <asp:Label runat="server"/>
            </asp:HyperLink>
        </li>
    </ul>

</asp:Content>