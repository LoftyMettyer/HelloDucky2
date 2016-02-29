<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" Inherits="OpenHRWorkflow.Login" Codebehind="Login.aspx.vb" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">
    <script type="text/javascript">
    // <!CDATA[
        function submitCheck() {

            var header = 'Login Failed';

            if (document.getElementById('<%=txtUserName.ClientID%>').value.trim().length === 0) {
                showDialog(header, 'Username is required.');
                return false;
            }
            if (document.getElementById('<%=txtPassword.ClientID%>').value.trim().length === 0) {
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
    
    <asp:ValidationSummary runat="server" Font-Size="11px" />
  
    <table style="display:none">
        <tr>
            <asp:TextBox runat="server" style="display:none"/>
            <asp:TextBox runat="server" TextMode="Password" style="visibility:hidden" Height="0"/>
        </tr>        
    </table>

    <table class="controlgrid">


        <tr>
            <td><asp:Label runat="server" ID="lblUserName" Text="Username" AssociatedControlID="txtUserName"/></td>
            <td><asp:TextBox runat="server" ID="txtUserName" autocomplete="Off"/></td>
        </tr>
        <tr>
            <td><asp:Label runat="server" ID="lblPassword" Text="Password"/></td>
            <td><asp:TextBox runat="server" ID="txtPassword" autocomplete="Off" TextMode="Password" /></td>            
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
            <asp:Button runat="server" ID="btnLogin2" OnClientClick="return submitCheck();" style="display: none;"/>

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