<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="Registration.aspx.vb" Inherits="Registration" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">
    <script type="text/javascript">
        // <!CDATA[
        window.onload = function() {
            document.getElementById('<%= txtEmail.ClientID %>').setAttribute('type', 'email');
        };

        function submitCheck() {

            var header = 'Registration Failed';

            if (document.getElementById('<%= txtEmail.ClientID %>').value.trim().length === 0) {
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
            <asp:Button runat="server" ID="btnSubmit2" OnClientClick="return submitCheck();" style="display: none;"/>

            <asp:LinkButton runat="server" ID="btnSubmit" OnClientClick="return submitCheck();">
                <asp:Image runat="server"/>
                <asp:Label runat="server"/>
            </asp:LinkButton>
        </li>
        <li>
            <asp:HyperLink runat="server" ID="btnCancel" NavigateUrl="~/Account/Login.aspx">
                <asp:Image runat="server"/>
                <asp:Label runat="server"/>
            </asp:HyperLink>
        </li>
    </ul>
</asp:Content>