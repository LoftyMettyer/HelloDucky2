<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="Home.aspx.vb" Inherits="Home" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">
</asp:Content>

<asp:Content ID="main" ContentPlaceHolderID="mainCPH" Runat="Server">
    
    <asp:Label runat="server" ID="lblWelcome" Text="Welcome"/>
    <asp:Label runat="server" ID="lblNothingTodo" Text="Nothing Todo"/>

    <div runat="server" id="pnlWFList" />

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
    <ol class="footer-buttons col3">
        <li>
            <asp:HyperLink runat="server" NavigateUrl="~/PendingSteps.aspx">
                <asp:Image ID="btnToDoList" runat="server" />
                <asp:Label ID="btnToDoList_Label" runat="server"/>
                <label ID="lblWFCount" style="position: absolute; background-color: Red; color: White; top: 6px; left: 50%; width: 13px; height: 13px; font-size: 11px;font-weight: bold; padding: 1px 2px 1px 2px; margin-left:5px; border-radius: 30px; box-shadow: 1px 1px 1px gray;" runat="server"/>
            </asp:HyperLink>
        </li>
        <li>
            <asp:HyperLink runat="server" ID="btnChangePwdButton" NavigateUrl="~/Account/ChangePassword.aspx">
                <asp:Image runat="server" ID="btnChangePwd"/>
                <asp:Label runat="server" ID="btnChangePwd_Label"/>
            </asp:HyperLink>
        </li>
        <li>
            <asp:LinkButton runat="server" ID="btnLogoutButton">
                <asp:Image runat="server" ID="btnLogout"/>
                <asp:Label runat="server" ID="btnLogout_Label"/>
            </asp:LinkButton>
        </li>
    </ol>
        
</asp:Content>