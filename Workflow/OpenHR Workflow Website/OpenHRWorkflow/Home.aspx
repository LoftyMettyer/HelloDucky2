<%@ Page Title="" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="Home.aspx.vb" Inherits="Home" %>

<asp:Content ID="head" ContentPlaceHolderID="headCPH" Runat="Server">
</asp:Content>

<asp:Content ID="main" ContentPlaceHolderID="mainCPH" Runat="Server">
    
    <label id="lblWelcome" runat="server">Welcome</label>
    <label id="lblNothingTodo" runat="server">Nothing Todo</label>

    <div runat="server" id="pnlWFList" />

</asp:Content>

<asp:Content ID="footer" ContentPlaceHolderID="footerCPH" Runat="Server">
    
    <table style="height: 100%; width: 100%">
        <tr style="height: 40px">
            <td style="width: 33%; text-align: center; overflow: hidden">
                <div style="position: relative; width: 40px; height: 100%; margin: auto">
                    <asp:ImageButton ID="btnToDoList" runat="server"/>
                    <div id="pnlWFCount" runat="server" style="position: absolute; top: 0px; right: -6px; padding: 1px 2px 1px 2px; background-color: Red; color: White; font-family: verdana; font-weight: bold; font-size: 0.75em; border-radius: 30px; box-shadow: 1px 1px 1px gray;">
                        <label id="lblWFCount" runat="server"></label>
                    </div>
                </div>
            </td>
            <td style="width: 33%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnChangePwd" runat="server" /></td>
            <td style="width: 33%; text-align: center; overflow: hidden"><asp:ImageButton ID="btnLogout" runat="server" /></td>
        </tr>
        <tr style="height: 17px">
            <td style="width: 33%; text-align: center; overflow: hidden"><label runat="server" id="btnToDoList_label"></label></td>
            <td style="width: 33%; text-align: center; overflow: hidden"><label runat="server" id="btnChangePwd_label"></label></td>
            <td style="width: 33%; text-align: center; overflow: hidden"><label runat="server" id="btnLogout_label"></label></td>
        </tr>
    </table>

</asp:Content>