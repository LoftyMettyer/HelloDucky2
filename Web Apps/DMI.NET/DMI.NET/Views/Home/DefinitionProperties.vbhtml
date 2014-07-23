@Imports DMI.NET
@Imports DMI.NET.Code.Extensions
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of ViewModels.DefinitionPropertiesViewModel)

<style>
	.inputProperty {
		width: 500px;
	}

	.labelProperty {
		display: inline-block;
		width: 120px;
		text-align: left;
	}

	.buttonProperty {
		float: right;
	}

</style>

@Html.LabelFor(Function(m) m.Name, New With {.class = "labelProperty"})
@Html.TextBoxFor(Function(m) m.Name, New With {.disabled = True, .class = "inputProperty"})
<br/>
@Html.LabelFor(Function(m) m.CreatedDate, New With {.class = "labelProperty"})
@Html.TextBoxFor(Function(m) m.CreatedDate, New With {.disabled = True, .class = "inputProperty"})
<br />
@Html.LabelFor(Function(m) m.LastSaveDate, New With {.class = "labelProperty"})
@Html.TextBoxFor(Function(m) m.LastSaveDate, New With {.disabled = True, .class = "inputProperty"})
<br />
@Html.LabelFor(Function(m) m.LastRunDate, New With {.class = "labelProperty", .style = Model.LastRunHidden})
@Html.TextBoxFor(Function(m) m.LastRunDate, New With {.disabled = True, .class = "inputProperty", .style = Model.LastRunHidden})
<br />
<br />
@Html.LabelFor(Function(m) m.Usage)

<div class="stretchyfill">
	<table id="definitionUsage"></table>
</div>

<br/>

<input type="button" value="Close" onclick="closeThisPopup();" class="buttonProperty" />

<script type="text/javascript">

	function attachUsage() {
	
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
				{ name: 'Name', index: 'Name', width: 620 }],
			rowNum: 10,
			height: 300,
			autowidth: true,
			rowTotal: 50,
			rowList: [10, 20, 30],
			shrinkToFit: true,
			pager: '#pcrud',
			sortname: 'Name',
			loadonce: true,
			viewrecords: true,
			sortorder: "asc"
		});

	}

	function closeThisPopup() {
		$("#divPopupReportDefinition").dialog("close");
		$("#divPopupReportDefinition").empty();
	}

	$(function () {
		attachUsage();
	});


	</script>
