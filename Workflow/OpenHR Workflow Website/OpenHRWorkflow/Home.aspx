<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="Home.aspx.vb" Inherits="Home" %>

<asp:Content ID="main" ContentPlaceHolderID="mainCPH" Runat="Server">
    
    <asp:Label runat="server" ID="lblWelcome" Text="Welcome"/>
    <asp:Label runat="server" ID="lblNothingTodo" Text="No items"/>

    <div runat="server" id="pnlWFList" />

<%-- TODO Prob should be using databinding, much simpler than building stuff in code.
    <asp:ListView ID="workflowList" runat="server">
        <LayoutTemplate>
            <ul>
                <li id="ItemPlaceHolder" runat="server"></li>
            </ul>
        </LayoutTemplate>
        <ItemTemplate>
            <li>
                <asp:HyperLink runat="server" NavigateUrl='<%# Eval("Url") %>'>
                    <asp:Image runat="server" ImageUrl='<%# Eval("Image") %>'/>
                    <asp:Label runat="server" Text='<%# Eval("Name") %>'/>
                </asp:HyperLink>
            </li>
        </ItemTemplate>
    </asp:ListView>--%>

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