<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="ChangePassword.aspx.vb" Inherits="ChangePassword" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">
    <script type="text/javascript">
    // <!CDATA[
        window.onload = function() {
            document.getElementById('ctl00_mainCPH_txtCurrPassword').focus();
        };

        function submitCheck() {

            var header = 'Change Password Failed';

            if (document.getElementById('ctl00_mainCPH_txtCurrPassword').value.length === 0) {
                showDialog(header, 'Current Password is required.');
                return false;
            }
            if (document.getElementById('ctl00_mainCPH_txtNewPassword').value.length === 0) {
                showDialog(header, 'New Password is required.');
                return false;
            }
            if (document.getElementById('ctl00_mainCPH_txtConfPassword').value.length === 0) {
                showDialog(header, 'Confirm Password is required.');
                return false;
            }

            if (document.getElementById('ctl00_mainCPH_txtNewPassword').value != document.getElementById('ctl00_mainCPH_txtConfPassword').value) {
                showDialog(header, 'New Password and Confirm Password do not match.');
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
            <td><label id="lblCurrPassword" runat="server">Current Password</label></td>
            <td><input id="txtCurrPassword" runat="server" type="password"/></td>
        </tr>
        <tr>
            <td><label id="lblNewPassword" runat="server">New Password</label></td>
            <td><input id="txtNewPassword" runat="server" type="password"/></td>
        </tr>
        <tr>
            <td><label id="lblConfPassword" runat="server">Confirm New Password</label></td>
            <td><input id="txtConfPassword" runat="server" type="password"/></td>
        </tr>
    </table>

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
    <ol class="footer-buttons col2">
        <li>
            <asp:LinkButton runat="server" ID="btnSubmitButton"  OnClientClick="return submitCheck();">
                <asp:Image ID="btnSubmit" runat="server"/>
                <asp:Label ID="btnSubmit_label" runat="server"/>
            </asp:LinkButton>
        </li>
        <li>
            <asp:HyperLink runat="server" NavigateUrl="~/Home.aspx">
                <asp:Image runat="server" ID="btnCancel"/>
                <asp:Label runat="server" ID="btnCancel_Label"/>
            </asp:HyperLink>
        </li>
    </ol>

</asp:Content>