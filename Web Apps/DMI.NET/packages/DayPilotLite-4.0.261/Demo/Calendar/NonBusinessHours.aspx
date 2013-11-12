<%@ Page Language="C#" MasterPageFile="~/Demo.master" AutoEventWireup="true" CodeFile="NonBusinessHours.aspx.cs" 
Inherits="NonBusinessHours" Title="Non-Business Hours | DayPilot Lite for ASP.NET WebForms Demo" Culture="en-US" %>
<%@ Register Assembly="DayPilot" Namespace="DayPilot.Web.Ui" TagPrefix="DayPilot" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div>
        <asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged">
            <asp:ListItem Value="hide">Hide non-business hours</asp:ListItem>
            <asp:ListItem Value="hideifpossible">Hide non-business hours (if possible)</asp:ListItem>
            <asp:ListItem Value="show">Show non-business hours</asp:ListItem>
        </asp:DropDownList><br />
        <br />
        <daypilot:daypilotcalendar id="DayPilotCalendar1" runat="server" DataStartField="Start"
            dataendfield="End" datatextfield="Name" datavaluefield="Id" nonbusinesshours="Hide" Days="2"
        CssOnly="true"
        CssClassPrefix="calendar_white"
            
            ></daypilot:daypilotcalendar>
    </div>
</asp:Content>