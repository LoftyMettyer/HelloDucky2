<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="PendingSteps.aspx.vb" Inherits="PendingSteps" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">
</asp:Content>

<asp:Content ID="main" ContentPlaceHolderID="mainCPH" Runat="Server">
    
    <asp:Label runat="server" ID="lblInstruction" Text="Welcome"/>
    <asp:Label runat="server" ID="lblNothingTodo" Text="Nothing Todo"/>

    <div runat="server" id="pnlWFList" />   

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
    <ol class="footer-buttons col2">
        <li>
            <asp:HyperLink runat="server" NavigateUrl="~/PendingSteps.aspx">
                <asp:Image runat="server" ID="btnRefresh"/>
                <asp:Label runat="server" ID="btnRefresh_Label"/>
            </asp:HyperLink>
        </li>
        <li>
            <asp:HyperLink runat="server" NavigateUrl="~/Home.aspx">
                <asp:Image runat="server" ID="btnCancel"/>
                <asp:Label runat="server" ID="btnCancel_Label"/>
            </asp:HyperLink>
        </li>
    </ol>

</asp:Content>