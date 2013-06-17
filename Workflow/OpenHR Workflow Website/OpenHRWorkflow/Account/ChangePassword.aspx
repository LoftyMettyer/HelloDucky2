﻿<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="ChangePassword.aspx.vb" Inherits="ChangePassword" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">
    <script type="text/javascript">
    // <!CDATA[
        function submitCheck() {

            var header = 'Change Password Failed';

            if (document.getElementById('ctl00_mainCPH_txtCurrPassword').value.trim().length === 0) {
                showDialog(header, 'Current Password is required.');
                return false;
            }
            if (document.getElementById('ctl00_mainCPH_txtNewPassword').value.trim().length === 0) {
                showDialog(header, 'New Password is required.');
                return false;
            }
            if (document.getElementById('ctl00_mainCPH_txtConfPassword').value.trim().length === 0) {
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
    
    <asp:Label runat="server" ID="lblWelcome" Text="Welcome"/>

    <table class="controlgrid">
        <tr>
            <td><asp:Label runat="server" ID="lblCurrPassword" Text="Current Password"/></td>
            <td><asp:TextBox runat="server" ID="txtCurrPassword" TextMode="Password" /></td>
        </tr>
        <tr>
            <td><asp:Label runat="server" ID="lblNewPassword" Text="New Password"/></td>
            <td><asp:TextBox runat="server" ID="txtNewPassword" TextMode="Password" /></td>
        </tr>
        <tr>
            <td><asp:Label runat="server" ID="lblConfPassword" Text="Confirm Password"/></td>
            <td><asp:TextBox runat="server" ID="txtConfPassword" TextMode="Password" /></td>
        </tr>
    </table>

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
    <ul class="footer-buttons col2">
        <li>
            <asp:LinkButton runat="server" ID="btnSubmit" OnClientClick="return submitCheck();">
                <asp:Image runat="server"/>
                <asp:Label runat="server"/>
            </asp:LinkButton>
        </li>
        <li>
            <asp:HyperLink runat="server" ID="btnCancel" NavigateUrl="~/Home.aspx">
                <asp:Image runat="server"/>
                <asp:Label runat="server"/>
            </asp:HyperLink>
        </li>
    </ul>
</asp:Content>