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
            <a href="javascript:void(0);" onclick="document.getElementById('ctl00_footerCPH_btnSubmitButton').click();">
                <asp:Image ID="btnSubmit" runat="server" /><asp:Label ID="btnSubmit_Label" runat="server"/>
            </a>
            <asp:ImageButton ID="btnSubmitButton" runat="server" OnClientClick="return submitCheck();"/>
        </li>
        <li><a href="Login.aspx"><asp:Image ID="btnCancel" runat="server" /><asp:Label ID="btnCancel_Label" runat="server"/></a></li>
    </ol>

</asp:Content>