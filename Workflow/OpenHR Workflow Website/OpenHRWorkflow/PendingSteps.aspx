<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="PendingSteps.aspx.vb" Inherits="PendingSteps" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">
</asp:Content>

<asp:Content ID="main" ContentPlaceHolderID="mainCPH" Runat="Server">
    
    <label id="lblInstruction" runat="server">Welcome</label>
    <label id="lblNothingTodo" runat="server">Nothing Todo</label>

    <div runat="server" id="pnlWFList" />   

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
    <ol class="footer-buttons col2">
        <li><a href="PendingSteps.aspx"><asp:Image ID="btnRefresh" runat="server" /><asp:Label ID="btnRefresh_Label" runat="server"/></a></li>
        <li><a href="Home.aspx"><asp:Image ID="btnCancel" runat="server" /><asp:Label ID="btnCancel_Label" runat="server"/></a></li>
    </ol>

</asp:Content>