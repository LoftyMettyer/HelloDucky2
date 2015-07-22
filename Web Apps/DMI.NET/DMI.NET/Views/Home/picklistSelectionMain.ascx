<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl(of DMI.NET.Models.ObjectRequests.PicklistSelectionModel)" %>

<%
	Session("selectionType") = Model.Type
	Session("selectionTableID") = Model.TableID
	Session("selectedIDs1") = Model.IDs1
	Session("picklistSelectionDataLoading") = True
%>

<script type="text/javascript">

	function loadAddRecords() {

		var iCount;
		iCount = new Number($('#txtLoadCount').val());
		$('#txtLoadCount').val(iCount + 1);

		if (iCount > 0) {
			var dataForm = OpenHR.getForm("picklistdataframe", "frmPicklistGetData");

			dataForm.txtTableID.value = frmUseful.txtTableID.value;
			dataForm.txtViewID.value = $('#selectView').val();
			dataForm.txtOrderID.value = $('#selectOrder').val();	// txtOrderID.value;
			dataForm.txtFirstRecPos.value = 1;
			dataForm.txtCurrentRecCount.value = 0;
			dataForm.txtPageAction.value = "LOAD";

			picklist_refreshData();
		}
	}

</script>
<div id="divPicklistSelectionMain">
	<input type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">
	<input type='hidden' id="txtTableID" name="txtTableID" value="0">
	<input type='hidden' id="txtViewID" name="txtViewID" value="0">
	<input type='hidden' id="txtOrderID" name="txtOrderID" value="0">
	<input type='hidden' id="txtSelectionType" name="txtSelectionType" value='<%:Model.Type%>'>
	<input type='hidden' id="txtSelectionTableID" name="txtSelectionTableID" value='<%:Model.TableID%>'>
</div>
<div id="picklistworkframe" data-framesource="picklistSelection" style="display: block"><%Html.RenderPartial("~/views/home/picklistSelection.ascx")%></div>
<div id="picklistdataframe" data-framesource="picklistSelectionData" style="display: none"><%Html.RenderPartial("~/views/home/picklistSelectionData.ascx")%></div>


<script type="text/javascript">
	$("#reportframe").show();

	picklistSelection_window_onload();

	$('.popup').bind('dialogclose', function () {
		closeclick();
		$("#optionframe").hide();
	});
</script>
