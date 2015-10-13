@Imports DMI.NET
@Imports DMI.NET.Code.Extensions
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of ViewModels.DefinitionPropertiesViewModel)

<div class="pageTitleDiv" style="margin-bottom: 15px">
	<span class="pageTitle" id="PopupReportDefinition_PageTitle">Properties</span>
</div>

<fieldset id="definitionPropertiesFields">
	@Html.LabelFor(Function(m) m.Name, New With {.class = "labelProperty"})
	@Html.TextBoxFor(Function(m) m.NameUrlDecoded, New With {.disabled = True, .class = "inputProperty"})
	<br />
	@Html.LabelFor(Function(m) m.CreatedDate, New With {.class = "labelProperty"})
	@Html.TextBoxFor(Function(m) m.CreatedDate, New With {.disabled = True, .class = "inputProperty"})
	<br />
	@Html.LabelFor(Function(m) m.LastSaveDate, New With {.class = "labelProperty"})
	@Html.TextBoxFor(Function(m) m.LastSaveDate, New With {.disabled = True, .class = "inputProperty"})
	<br />
	@Html.LabelFor(Function(m) m.LastRunDate, New With {.class = "labelProperty", .style = Model.LastRunHidden})
	@Html.TextBoxFor(Function(m) m.LastRunDate, New With {.disabled = True, .class = "inputProperty", .style = Model.LastRunHidden})
</fieldset>

<fieldset id="definitionusagediv">
	<legend class="fontsmalltitle">@Html.LabelFor(Function(m) m.Usage)</legend>
	<table id="definitionUsage"></table>
</fieldset>

<fieldset class="genericbuttonpopupalignment" id="defselPropertiesPopup">
	<input type="button" value="Close" onclick="closeThisPopup();" />
</fieldset>

<script type="text/javascript">
	$("#definitionUsage").jqGrid({
		datatype: "jsonstring",
		datastr: '@Model.Usage.ToJsonResult',
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
		cmTemplate: { sortable: false }
	});

	$('fieldset').css("border", "0");

	$("#definitionUsage").jqGrid('setGridWidth', 840);

	function closeThisPopup() {
		$("#divPopupReportDefinition").dialog("close");
		$("#divPopupReportDefinition").empty();
	}

</script>
