<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>

<html>
<head runat="server">
    
    <link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />
    <script src="<%: Url.Content("~/Scripts/jquery-1.8.2.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/openhr.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>

    <title>OpenHR Intranet</title>
        <%
            Session("selectionType") = Request.Form("selectionType")
            Session("selectionTableID") = Request.Form("txtTableID")
	
            Session("selectedIDs1") = Request.Form("selectedIDs1")
            Session("picklistSelectionDataLoading") = True
        %>

 <script type="text/javascript">

     function picklistSelectionMain_window_onload() {
         $("#picklistdialog").attr("data-framesource", "PICKLISTSELECTIONMAIN");     
     }

     function loadAddRecords()
     {
         var iCount;
         
         iCount = new Number(txtLoadCount.value);
         txtLoadCount.value = iCount + 1;

         if (iCount > 0) {	
             var dataForm = OpenHR.getForm("dataframe","frmGetData");

             dataForm.txtTableID.value = txtTableID.value;
             dataForm.txtViewID.value = txtViewID.value;
             dataForm.txtOrderID.value = txtOrderID.value;
             dataForm.txtFirstRecPos.value = 1;
             dataForm.txtCurrentRecCount.value = 0;
             dataForm.txtPageAction.value = "LOAD";

             refreshData();
         }
     }

 </script>

</head>

<body>
    <div data-framesource="picklistSelectionMain">

    	<input type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">
        <input type='hidden' id="txtTableID" name="txtTableID" value="0">
        <input type='hidden' id="txtViewID" name="txtViewID" value="0">
        <input type='hidden' id="txtOrderID" name="txtOrderID" value="0">
        <input type='hidden' id="txtSelectionType" name="txtSelectionType" value='<%=Request.Form("selectionType")%>'>
        <input type='hidden' id="txtSelectionTableID" name="txtSelectionTableID" value='<%=Request.Form("selectionTableID")%>'>

        <div id="mainframeset">
            <div id="workframe" data-framesource="picklistSelection" style="display: none"><%Html.RenderPartial("~/views/home/picklistSelection.ascx")%></div>
            <div id="dataframe" data-framesource="picklistSelectionData" style="display: none"><%Html.RenderPartial("~/views/home/picklistSelectionData.ascx")%></div>
        </div>

    </div>
</body>
</html>

<script type="text/javascript">
    picklistSelection_window_onload();
    picklistSelectionData_window_onload();
    picklistSelectionMain_window_onload();
</script>
