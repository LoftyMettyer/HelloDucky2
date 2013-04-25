<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<%
    Dim fok As Boolean
    Dim objCrossTab As HR.Intranet.Server.CrossTab
    Dim fNotCancelled As Boolean
    Dim lngEventLogID As Long
    Dim aPrompts As Object
    
    Session("objCrossTab" & Session("UtilID")) = Nothing
    Session("CT_Mode") = ""
    Session("CT_PageNumber") = ""
    Session("CT_IntersectionType") = ""
    Session("CT_ShowPercentage") = ""
    Session("CT_PercentageOfPage") = ""
    Session("CT_SupressZeros") = ""
    Session("CT_Use1000") = ""

    If Session("utiltype") = "" Or _
       Session("utilname") = "" Or _
       Session("utilid") = "" Or _
       Session("action") = "" Then
	      
        Response.Write("Error : Not all session variables found...<HR>")
        Response.Write("Type = " & Session("utiltype") & "<BR>")
        Response.Write("UtilName = " & Session("utilname") & "<BR>")
        Response.Write("UtilID = " & Session("utilid") & "<BR>")
        Response.Write("Action = " & Session("action") & "<BR>")
        Response.End()
    End If

    ' Create the reference to the DLL (Report Class)
    objCrossTab = New HR.Intranet.Server.CrossTab
    Session("objCrossTab" & Session("UtilID")) = Nothing

    ' Pass required info to the DLL
    objCrossTab.Username = Session("username").ToString()
    CallByName(objCrossTab, "Connection", CallType.Let, Session("databaseConnection"))    
    objCrossTab.CrossTabID = Session("utilid")
    objCrossTab.ClientDateFormat = Session("LocaleDateFormat")
    objCrossTab.LocalDecimalSeparator = Session("LocaleDecimalSeparator")

    fok = True

    objCrossTab.CreateTablesCollection()

    aPrompts = Session("Prompts_" & Session("utiltype") & "_" & Session("utilid"))
    If fok Then
        fok = objCrossTab.SetPromptedValues(aPrompts)
        fNotCancelled = Response.IsClientConnected
        If fok Then fok = fNotCancelled
    End If

    If fok Then
        fok = objCrossTab.RetreiveDefinition
        fNotCancelled = Response.IsClientConnected
        If fok Then fok = fNotCancelled
    End If

    If fok Then
        lngEventLogID = objCrossTab.EventLogAddHeader
        fok = (lngEventLogID > 0)
        fNotCancelled = Response.IsClientConnected
        If fok Then fok = fNotCancelled
    End If

    If fok Then
        fok = objCrossTab.UDFFunctions(True)
        fNotCancelled = Response.IsClientConnected
        If fok Then fok = fNotCancelled
    End If

    If fok Then
        fok = objCrossTab.CreateTempTable
        fNotCancelled = Response.IsClientConnected
        If fok Then fok = fNotCancelled
    End If

    If fok Then
        fok = objCrossTab.UDFFunctions(False)
        fNotCancelled = Response.IsClientConnected
        If fok Then fok = fNotCancelled
    End If

    If fok Then
        fok = objCrossTab.GetHeadingsAndSearches
        fNotCancelled = Response.IsClientConnected
        If fok Then fok = fNotCancelled
    End If

    If fok Then
        fok = objCrossTab.BuildTypeArray
        fNotCancelled = Response.IsClientConnected
        If fok Then fok = fNotCancelled
    End If

    If fok Then
        fok = objCrossTab.BuildDataArrays
        fNotCancelled = Response.IsClientConnected
        If fok Then fok = fNotCancelled
    End If

    Session("objCrossTab" & Session("UtilID")) = objCrossTab

%>

<script type="text/javascript">
	function loadAddRecords() {
        
		var iCount;
		iCount = new Number(txtLoadCount.value);
		txtLoadCount.value = iCount + 1;
		if (iCount > 0) {
			var frmGetData = OpenHR.getForm("reportdataframe", "frmGetCrossTabData");
			<% Response.Write("frmGetData.txtUtilID.value = """ & Session("utilid") & """;" & vbCrLf)%>
			getData("LOAD", 0, 0, 0, 0, 0, 0);
		}
	}
</script>


<input type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">

<div id="reportworkframe" data-framesource="util_run_crosstabs" style="display: block;">
    <%Html.RenderPartial("~/views/home/util_run_crosstabs.ascx")%>
</div>

<div id="reportdataframe" data-framesource="util_run_crosstabsData" style="display: none;" accesskey="">
    <%Html.RenderPartial("~/views/home/util_run_crosstabsData.ascx")%>
</div>

<div id="reportbreakdownframe" data-framesource="util_run_crosstabsBreakdown" style="display: none;" accesskey="">   
    <%Html.RenderPartial("~/views/home/util_run_crosstabsBreakdown.ascx")%>
</div>



<form id="frmOutput" name="frmOutput">
    <input type="hidden" id="fok" name="fok" value="">
    <input type="hidden" id="cancelled" name="cancelled" value="">
    <input type="hidden" id="statusmessage" name="statusmessage" value="">
</form>

<script type="text/javascript">

	util_run_crosstabs_window_onload();

	$("#reportframe").show();

	$("#top").hide();
	$("#reportworkframe").show();

</script>
