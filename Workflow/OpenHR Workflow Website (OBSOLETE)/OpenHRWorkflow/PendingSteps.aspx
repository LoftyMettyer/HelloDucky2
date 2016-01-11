<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="PendingSteps.aspx.vb" Inherits="PendingSteps" %>

<asp:Content ID="main" ContentPlaceHolderID="mainCPH" Runat="Server">
    
    <asp:Label runat="server" ID="lblInstruction" Text="Welcome"/>
    <asp:Label runat="server" ID="lblNothingTodo" Text="No items"/>

    <ul runat="server" id="workflowList" class="workflow-list"/>

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
    <ul class="footer-buttons col2">
        <li>
            <asp:HyperLink runat="server" ID="btnRefresh" NavigateUrl="~/PendingSteps.aspx">
                <asp:Image runat="server"/>
                <asp:Label runat="server"/>
            </asp:HyperLink>
        </li>
        <li>
            <asp:HyperLink runat="server" ID="btnCancel" NavigateUrl="~/Home.aspx">
                <asp:Image runat="server"/>
                <asp:Label runat="server"/>
            </asp:HyperLink>
        </li>
    </ul>

</asp:Content>