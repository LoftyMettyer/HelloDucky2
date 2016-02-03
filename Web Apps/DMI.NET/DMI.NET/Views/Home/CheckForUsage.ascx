<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl(Of DMI.NET.Models.ObjectRequests.DefselModel)" %>
<%@ Import Namespace="DMI.NET.Code.Extensions" %>

<div class="pageTitleDiv" style="margin-bottom: 15px">
	<span class="pageTitle" id="PopupReportDefinition_PageTitle">'<%:Model.utilName%>' is in use</span>
</div>

<fieldset id="definitionusagediv">
	<table id="definitionUsage"></table>
</fieldset>

<fieldset class="genericbuttonpopupalignment" id="defselPropertiesPopup">
	<input type="button" value="Close" onclick="closeThisPopup();" />
</fieldset>

<script type="text/javascript">

	$("#definitionUsage").jqGrid({
		datatype: "jsonstring",
		datastr: '<%:Model.Usage.ToJsonResult%>',
		mtype: 'GET',
		jsonReader: {
			root: "rows", //array containing actual data
			page: "page", //current page
			total: "total", //total pages for the query
			records: "records", //total number of records
			repeatitems: false,
			id: "Name" //index of the column with the PK in it
		},
		colNames: ['Name'],
		colModel: [
		{ name: 'Name', index: 'Name', align: "left" }],
		rowNum: 10000,
		width: 'auto',
		height: '300px',
		autowidth: true,
		rowTotal: 50,
		rowList: [10, 20, 30],
		shrinkToFit: true,
		pager: '#pcrud',
		sortname: 'Name',
		loadonce: true,
		autoencode: true,
		viewrecords: true,
		sortorder: "asc",
		cmTemplate: { sortable: false },
		loadComplete: function (data) {
			$('fieldset').css("border", "0");
		}
	});


	function closeThisPopup() {
		$("#divPopupReportDefinition").dialog("close");
		$("#divPopupReportDefinition").empty();
	}

	$(function() {

		<%If Model.Usage Is Nothing Then%>
			OpenHR.modalPrompt("<%:Model.Status%>", 0, "Confirm").then(function(answer) {
				if (answer === 1) {
						menu_LoadDefSel_Inside_Frame(<%:CInt(Model.utiltype)%>, 0, <%:Model.txtTableID%>, true);
				}
			});
		<%Else%>
			$('#divPopupReportDefinition').dialog("open");
			var dialogWidth = screen.width / 2;
			$('#divPopupReportDefinition').dialog("option", "width", dialogWidth);
			$("#definitionUsage").jqGrid('setGridWidth', 820);		
		<%End If%>

	});

</script>


