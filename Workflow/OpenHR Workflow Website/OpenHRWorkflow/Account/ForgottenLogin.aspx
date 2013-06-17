<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="ForgottenLogin.aspx.vb" Inherits="ForgottenLogin" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">   
    <script  type="text/javascript">
        // <!CDATA[
        window.onload = function() {
            document.getElementById('ctl00_mainCPH_txtEmail').focus();
        };

        function submitCheck() {

            var header = 'Request Failed';

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
    
    <label id="lblWelcome" runat="server">Welcome</label>

    <table class="controlgrid">
        <tr>
            <td ><label id="lblEmail" runat="server">Email</label></td>
            <td><input id="txtEmail" runat="server" /></td>
        </tr>
                
    </table>

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
    <ol class="footer-buttons col2">
        <li>
            <asp:LinkButton runat="server" ID="btnSubmitButton" OnClientClick="return submitCheck();">
                <asp:Image runat="server" ID="btnSubmit"/>
                <asp:Label runat="server" ID="btnSubmit_Label"/>
            </asp:LinkButton>
        </li>
        <li>
            <asp:HyperLink runat="server" NavigateUrl="~/Account/Login.aspx">
                <asp:Image runat="server" ID="btnCancel" />
                <asp:Label runat="server" ID="btnCancel_Label" />
            </asp:HyperLink>
        </li>
    </ol>

</asp:Content>