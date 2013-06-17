<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="PendingSteps.aspx.vb" Inherits="PendingSteps" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">
</asp:Content>

<asp:Content ID="main" ContentPlaceHolderID="mainCPH" Runat="Server">
    
    <label id="lblInstruction" runat="server">Welcome</label>
    <label id="lblNothingTodo" runat="server">Nothing Todo</label>

    <div runat="server" id="pnlWFList" />   

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
    <table style="height: 100%; width: 100%">
        <tr style="height: 40px">
            <td style="width: 50%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnRefresh" runat="server" /></td>
            <td style="width: 50%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnCancel" runat="server" /></td>
        </tr>
        <tr style="height: 17px">
            <td style="width: 50%; text-align: center; overflow: hidden"><label runat="server" id="btnRefresh_label"></label></td>
            <td style="width: 50%; text-align: center; overflow: hidden"><label runat="server" id="btnCancel_label"></label></td>
        </tr>
    </table>

</asp:Content>