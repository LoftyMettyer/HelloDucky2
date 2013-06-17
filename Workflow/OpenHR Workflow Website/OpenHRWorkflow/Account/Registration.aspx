<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="Registration.aspx.vb" Inherits="Registration" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">
    <script type="text/javascript">
        // <!CDATA[
        window.onload = function() {
            document.getElementById('ctl00_mainCPH_txtEmail').focus();
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
    
    <label  id="lblWelcome" runat="server">Welcome</label>
                              
    <table class="controlgrid">
        <tr>
            <td><label id="lblEmail" runat="server">Email</label></td>
            <td><input id="txtEmail" runat="server"/></td>
        </tr>
    </table>

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
     <ol class="footer-buttons col2">
        <li>
            <a href="javascript:void(0);" onclick="document.getElementById('ctl00_footerCPH_btnRegisterButton').click();">
                <asp:Image ID="btnRegister" runat="server" /><asp:Label ID="btnRegister_Label" runat="server"/>
            </a>
            <asp:ImageButton ID="btnRegisterButton" runat="server" OnClientClick="return submitCheck();"/>
        </li>
        <li><a href="Login.aspx"><asp:Image ID="btnHome" runat="server" /><asp:Label ID="btnHome_Label" runat="server"/></a></li>
    </ol>

</asp:Content>