<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="Home.aspx.vb" Inherits="Home" %>

<asp:Content ID="main" ContentPlaceHolderID="mainCPH" Runat="Server">
    
    <asp:Label runat="server" ID="lblWelcome" Text="Welcome"/>
    <asp:Label runat="server" ID="lblNothingTodo" Text="No items"/>
    
    <ul runat="server" id="workflowList" class="workflow-list"/>

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
    <ul class="footer-buttons col3">
        <li>
            <asp:HyperLink runat="server" ID="btnTodoList" NavigateUrl="~/PendingSteps.aspx">
                <asp:Image runat="server"/>
                <asp:Label runat="server"/>
                <asp:Label runat="server" ID="lblWFCount" CssClass="step-count" />
            </asp:HyperLink>
        </li>
        <li>
            <asp:HyperLink runat="server" ID="btnChangePwd" NavigateUrl="~/Account/ChangePassword.aspx">
                <asp:Image runat="server"/>
                <asp:Label runat="server"/>
            </asp:HyperLink>
        </li>
        <li>
            <asp:LinkButton runat="server" ID="btnLogout">
                <asp:Image runat="server"/>
                <asp:Label runat="server"/>
            </asp:LinkButton>
        </li>
    </ul>
        
</asp:Content>