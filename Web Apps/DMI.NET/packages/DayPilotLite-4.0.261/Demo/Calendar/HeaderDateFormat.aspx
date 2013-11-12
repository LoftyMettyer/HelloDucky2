<%@ Page Language="C#" MasterPageFile="~/Demo.master" AutoEventWireup="true" CodeFile="HeaderDateFormat.aspx.cs" 
Inherits="HeaderDateFormat" Title="Header Date Format | DayPilot Lite for ASP.NET WebForms Demo" Culture="en-US" %>

<%@ Register Assembly="DayPilot" Namespace="DayPilot.Web.Ui" TagPrefix="DayPilot" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged">
        <asp:ListItem Value="d">Short date (d)</asp:ListItem>
        <asp:ListItem Value="D">Long date (D)</asp:ListItem>
        <asp:ListItem Value="m">Month day (m)</asp:ListItem>
        <asp:ListItem Value="dddd">Custom (dddd)</asp:ListItem>
        <asp:ListItem Value="ddd">Custom (ddd)</asp:ListItem>
        <asp:ListItem Value="dd">Custom (dd)</asp:ListItem>
        <asp:ListItem Value="%d">Custom (%d)</asp:ListItem>
        <asp:ListItem Value="M/dd">Custom (M/dd)</asp:ListItem>
        <asp:ListItem Value="M/d">Custom (M/d)</asp:ListItem>
        <asp:ListItem Value="d MMMM yyyy">Custom (d MMMM yyyy)</asp:ListItem>
        <asp:ListItem Value="yyyy-MM-dd">Custom (yyyy-MM-dd)</asp:ListItem>
    </asp:DropDownList><br />
    <br />
    <daypilot:daypilotcalendar id="DayPilotCalendar1" runat="server" DataEndField="end" DataStartField="start" DataTextField="name" DataValueField="id" Width="100%"
        CssOnly="true"
        CssClassPrefix="calendar_white"
    
    ></daypilot:daypilotcalendar>
</asp:Content>

