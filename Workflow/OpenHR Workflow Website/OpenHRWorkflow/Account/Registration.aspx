<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="Registration.aspx.vb" Inherits="Registration" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">
    <script type="text/javascript">
        // <!CDATA[
        window.onload = function() {
            document.getElementById('ctl00_mainCPH_txtEmail').setAttribute('type', 'email');
        };

        function submitCheck() {

            var header = 'Registration Failed';

            if (document.getElementById('ctl00_mainCPH_txtEmail').value.length === 0) {
                showDialog(header, 'Email address is required.');
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
            <td><asp:Label runat="server" ID="lblEmail" Text="Email"/></td>
            <td><asp:TextBox runat="server" ID="txtEmail"/></td>
        </tr>
    </table>

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
    <ul class="footer-buttons col2">
        <li>
            <asp:LinkButton runat="server" ID="btnRegister" OnClientClick="return submitCheck();">
                <asp:Image runat="server"/>
                <asp:Label runat="server"/>
            </asp:LinkButton>
        </li>
        <li>
            <asp:HyperLink runat="server" ID="btnHome" NavigateUrl="~/Account/Login.aspx">
                <asp:Image runat="server"/>
                <asp:Label runat="server"/>
            </asp:HyperLink>
        </li>
    </ul>
</asp:Content>