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
    
    <table style="height: 100%; width: 100%">
        <tr style="height: 40px">
            <td style="width: 50%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnSubmit" runat="server" OnClientClick="return submitCheck();"/></td>                
            <td style="width: 50%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnCancel" runat="server" /></td>
        </tr>
        <tr style="height: 17px">
            <td style="width: 50%; text-align: center; overflow: hidden"><label runat="server" id="btnSubmit_label"></label></td>
            <td style="width: 50%; text-align: center; overflow: hidden"><label runat="server" id="btnCancel_label"></label></td>
        </tr>
    </table>

</asp:Content>