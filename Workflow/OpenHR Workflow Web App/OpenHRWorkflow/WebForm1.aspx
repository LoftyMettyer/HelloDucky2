<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WebForm1.aspx.vb" Inherits="OpenHRWorkflow.WebForm1" Trace="False" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
       <asp:GridView ID="GridView1" runat="server" AllowPaging="True" 
          AllowSorting="True" EnableViewState="False" DataSourceID="SqlDataSource1">
       </asp:GridView>
       <asp:ObjectDataSource ID="ObjectDataSource1" runat="server"></asp:ObjectDataSource>
       <asp:SqlDataSource ID="SqlDataSource1" runat="server" EnableCaching="True"></asp:SqlDataSource>
    </div>
    </form>
</body>
</html>
