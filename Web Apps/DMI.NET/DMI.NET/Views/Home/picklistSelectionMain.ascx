﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%
    Session("selectionType") = Request.Form("selectionType")
    Session("selectionTableID") = Request.Form("txtTableID")
	
    Session("selectedIDs1") = Request.Form("selectedIDs1")
    Session("picklistSelectionDataLoading") = True
%>

<script type="text/javascript">

    function picklistSelectionMain_window_onload() {
        $("#picklistmainframeset").attr("data-framesource", "PICKLISTSELECTIONMAIN");
    }

    function loadAddRecords() {

        var iCount;
         
        iCount = new Number(txtLoadCount.value);
        txtLoadCount.value = iCount + 1;
         
        if (iCount > 0) {	
            var dataForm = OpenHR.getForm("dataframe", "frmPicklistGetData");

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


<div id="picklistSelection" data-framesource="picklistSelectionMain">

    <input type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">
    <input type='hidden' id="txtTableID" name="txtTableID" value="0">
    <input type='hidden' id="txtViewID" name="txtViewID" value="0">
    <input type='hidden' id="txtOrderID" name="txtOrderID" value="0">
    <input type='hidden' id="txtSelectionType" name="txtSelectionType" value='<%=Request.Form("selectionType")%>'>
    <input type='hidden' id="txtSelectionTableID" name="txtSelectionTableID" value='<%=Request.Form("selectionTableID")%>'>

    <div id="picklistmainframeset">
        <div id="picklistworkframe" data-framesource="picklistSelection"><%Html.RenderPartial("~/views/home/picklistSelection.ascx")%></div>
        <div id="picklistdataframe" data-framesource="picklistSelectionData" style="display: none"><%Html.RenderPartial("~/views/home/picklistSelectionData.ascx")%></div>
    </div>

</div>

<script type="text/javascript">
    picklistSelectionMain_window_onload();
    picklistSelection_window_onload();
    picklistSelectionData_window_onload();
    picklistSelection_addhandlers();

    //$("#workframeset").hide();
   // $("#workframe").hide();
    $("#picklistworkframe").show();
    $("#picklistSelection").show();
    $("#picklistmainframeset").show();
    $("#reportframe").show();

</script>
