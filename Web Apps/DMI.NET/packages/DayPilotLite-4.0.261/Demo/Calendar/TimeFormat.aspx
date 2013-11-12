<%@ Page Language="C#" MasterPageFile="~/Demo.master" AutoEventWireup="true" CodeFile="TimeFormat.aspx.cs" 
Inherits="TimeFormat" Title="Time Format (12/24 Hours) | DayPilot Lite for ASP.NET WebForms Demo" Culture="en-US" %>

<%@ Register Assembly="DayPilot" Namespace="DayPilot.Web.Ui" TagPrefix="DayPilot" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged">
        <asp:ListItem Value="12">12-hours clock</asp:ListItem>
        <asp:ListItem Value="24" Selected="true">24-hours clock</asp:ListItem>
    </asp:DropDownList><br />
    <br />
    <daypilot:daypilotcalendar id="DayPilotCalendar1" runat="server" DataEndField="end" DataStartField="start" DataTextField="name" DataValueField="id" TimeFormat="Clock24Hours"
        CssOnly="true"
        CssClassPrefix="calendar_white"
    
    ></daypilot:daypilotcalendar>
</asp:Content>

