<%@ Page Language="C#" MasterPageFile="~/Demo.master" AutoEventWireup="true" CodeFile="Events.aspx.cs"
    Inherits="Events" Title="Event Handling | DayPilot Lite for ASP.NET WebForms Demo" Culture="en-US" %>

<%@ Register Assembly="DayPilot" Namespace="DayPilot.Web.Ui" TagPrefix="DayPilot" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged">
        <asp:ListItem Value="client">Client-side event handling</asp:ListItem>
        <asp:ListItem Value="server">Server-side event handling</asp:ListItem>
        <asp:ListItem Value="disabled">Disabled</asp:ListItem>
    </asp:DropDownList><br />
    <br />
    <asp:Label ID="Label1" runat="server" Text=""></asp:Label>
    <br />
    <br />
    <DayPilot:DayPilotCalendar ID="DayPilotCalendar1" runat="server" 
        DataSourceID="XmlDataSource1"
        DataTextField="name" 
        DataValueField="id" 
        StartDate="2007-01-01" 
        TimeFormat="Clock12Hours"
        DataStartField="Start" 
        DataEndField="End" 
        Days="7" 
        OnEventClick="DayPilotCalendar1_EventClick"
        OnTimeRangeSelected="DayPilotCalendar1_TimeRangeSelected"
        EventClickHandling="JavaScript"
        TimeRangeSelectedHandling="JavaScript"
        CssOnly="true"
        CssClassPrefix="calendar_white"

        ></DayPilot:DayPilotCalendar>
    &nbsp;<br />
    <asp:XmlDataSource ID="XmlDataSource1" runat="server" DataFile="~/App_Data/TestingData.xml">
    </asp:XmlDataSource>
</asp:Content>
