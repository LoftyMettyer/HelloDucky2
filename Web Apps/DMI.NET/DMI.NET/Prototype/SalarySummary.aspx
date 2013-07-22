<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SalarySummary.aspx.vb" Inherits="DMI.NET.SalarySummary" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<style type="text/css">
	h3 { color: blue; }
	table, th, td { border: 1px solid black; }
	td { width: 100px;}
	tr.alt td { background-color: #EAF2D3; color: #000000; }
</style>
<body>
    <form id="form1" runat="server">
    <div>
	    <h3>Your salary summary</h3>
				<asp:Table runat="server" ID="SalarySummaryTable" CssClass=""></asp:Table>
    </div>
    </form>
</body>
</html>
