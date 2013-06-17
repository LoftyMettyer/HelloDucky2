<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="Home.aspx.vb" Inherits="Home" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">
</asp:Content>

<asp:Content ID="main" ContentPlaceHolderID="mainCPH" Runat="Server">
    
    <label id="lblWelcome" runat="server">Welcome</label>
    <label id="lblNothingTodo" runat="server">Nothing Todo</label>

    <div runat="server" id="pnlWFList" />

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
    <ol class="footer-buttons col3">
        <li>
            <a href="PendingSteps.aspx">
                <asp:Image ID="btnToDoList" runat="server" />
                <asp:Label ID="btnToDoList_Label" runat="server"/>
                <label ID="lblWFCount" style="position: absolute; background-color: Red; color: White; top: 6px; left: 50%; width: 13px; height: 13px; font-size: 11px;font-weight: bold; padding: 1px 2px 1px 2px; margin-left:5px; border-radius: 30px; box-shadow: 1px 1px 1px gray;" runat="server"/>
            </a>
        </li>
        <li><a href="Account/ChangePassword.aspx"><asp:Image ID="btnChangePwd" runat="server" /><asp:Label ID="btnChangePwd_Label" runat="server"/></a></li>
        <li>
            <a href="javascript:void(0);" onclick="document.getElementById('ctl00_footerCPH_btnLogoutButton').click();">
                <asp:Image ID="btnLogout" runat="server" /><asp:Label ID="btnLogout_Label" runat="server"/>
            </a>
            <asp:ImageButton ID="btnLogoutButton" runat="server"/>
        </li>
    </ol>
        
</asp:Content>