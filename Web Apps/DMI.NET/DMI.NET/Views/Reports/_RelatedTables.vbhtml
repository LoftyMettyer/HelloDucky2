@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Enums
@Imports DMI.NET.Classes
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of Models.CustomReportModel)


<div id="divReportParents">

	<fieldset @Model.Parent1.Visibility>
		<legend>Parent 1 :</legend>

		@Html.HiddenFor(Function(m) m.ChildTablesString, New With {.id = "txtCTAAS"})

		<input type="hidden" id="txtParent1ID" name="Parent1.ID" value="@Model.Parent1.ID" />
		Table:
		@Html.TextBoxFor(Function(m) m.Parent1.Name, New With {.readonly = "true"})
		<br />
		@Html.RadioButton("Parent1.Selectiontype", RecordSelectionType.AllRecords, Model.Parent1.SelectionType = RecordSelectionType.AllRecords, New With {.onclick = "changeRecordOption('Parent1','all')"})
		All Records
		<br />

		@Html.RadioButton("Parent1.SelectionType", RecordSelectionType.Picklist, Model.Parent1.SelectionType = RecordSelectionType.Picklist, New With {.onclick = "changeRecordOption('Parent1','picklist')"})
		Picklist
		<input type="hidden" id="txtParent1PicklistID" name="Parent1.PicklistID" value="@Model.Parent1.PicklistID" />
		@Html.TextBoxFor(Function(m) m.Parent1.PicklistName, New With {.id = "txtParent1Picklist", .readonly = "true"})
		@Html.EllipseButton("cmdParent1Picklist", "selectRecordOption('p1', 'picklist')", Model.Parent1.SelectionType = RecordSelectionType.Picklist)
		@Html.ValidationMessageFor(Function(m) m.Parent1.PicklistID)
		<br />

		@Html.RadioButton("Parent1.SelectionType", RecordSelectionType.Filter, Model.Parent1.SelectionType = RecordSelectionType.Filter, New With {.onclick = "changeRecordOption('Parent1','filter')"})
		Filter
		<input type="hidden" id="txtParent1FilterID" name="Parent1.FilterID" value="@Model.Parent1.FilterID" />
		@Html.TextBoxFor(Function(m) m.Parent1.FilterName, New With {.id = "txtParent1Filter", .readonly = "true"})
		@Html.EllipseButton("cmdParent1Filter", "selectRecordOption('p1', 'filter')", Model.Parent1.SelectionType = RecordSelectionType.Filter)
		@Html.ValidationMessageFor(Function(m) m.Parent1.FilterID)

	</fieldset>

	<fieldset @Model.Parent2.Visibility>
		<legend>Parent 2 :</legend>

		<input type="hidden" id="txtParent2ID" name="Parent2.ID" value="@Model.Parent2.ID" />
		@Html.TextBoxFor(Function(m) m.Parent2.Name, New With {.readonly = "true"})
		<br />
		@Html.RadioButton("Parent2.Selectiontype", RecordSelectionType.AllRecords, Model.Parent2.SelectionType = RecordSelectionType.AllRecords, New With {.onclick = "changeRecordOption('Parent2','all')"})
		All Records
		<br />
		@Html.RadioButton("Parent2.SelectionType", RecordSelectionType.Picklist, Model.Parent2.SelectionType = RecordSelectionType.Picklist, New With {.onclick = "changeRecordOption('Parent2','picklist')"})
		Picklist
		<input type="hidden" id="txtParent2PicklistID" name="Parent2.PicklistID" value="@Model.Parent2.PicklistID" />
		@Html.TextBoxFor(Function(m) m.Parent2.PicklistName, New With {.id = "txtParent2Picklist", .readonly = "true"})
		@Html.EllipseButton("cmdParent2Picklist", "selectRecordOption('p2', 'picklist')", Model.Parent2.SelectionType = RecordSelectionType.Picklist)
		@Html.ValidationMessageFor(Function(m) m.Parent2.PicklistID)
		<br />

		@Html.RadioButton("Parent2.SelectionType", RecordSelectionType.Filter, Model.Parent2.SelectionType = RecordSelectionType.Filter, New With {.onclick = "changeRecordOption('Parent2','filter')"})
		Filter
		<input type="hidden" id="txtParent2FilterID" name="Parent2.FilterID" value="@Model.Parent2.FilterID" />
		@Html.TextBoxFor(Function(m) m.Parent2.FilterName, New With {.id = "txtParent2Filter", .readonly = "true"})
		@Html.EllipseButton("cmdParent2Filter", "selectRecordOption('p2', 'filter')", Model.Parent2.SelectionType = RecordSelectionType.Filter)
		@Html.ValidationMessageFor(Function(m) m.Parent2.FilterID)

	</fieldset>

</div>

<br/>

<fieldset class="relatedtables">
	<legend>Child Tables :</legend>

	<div class="stretchyfill">
		<table id ="ChildTables"></table>
	</div>

	<div class="stretchyfixed">
		<input type="button" id="btnChildAdd" value="Add..." onclick="addChildTable();" />
		<br/>
		<input type="button" id="btnChildEdit" value="Edit..." onclick="editChildTable(0);" />
		<br />
		<input type="button" id="btnChildRemove" value="Remove" />
		<br />
		<input type="button" id="btnChildRemoveAll" value="Remove All" onclick="removeAllChildTables();" />				
	</div>

</fieldset>


@*@code
	For Each objEvent As ReportChildTables In Model.ChildTables
		Html.RenderPartial("EditorTemplates\ReportChildTable", objEvent)
	Next
End Code*@



<script type="text/javascript">


	function toggle_visibility(id) {

		var e = $("#Event_Detail_" + id)[0]
		if (e.style.display == 'block')
			e.style.display = 'none';
		else
			e.style.display = 'block';
	}


	// slightly hacked version of the orignal from util_def_customreports.js
	// passes in generated stirng of currently selected item (I'm sure this code can be cleverised.)
	function addChildTable() {

		var frmChild = $("#frmGetChildTable");
		OpenHR.submitForm(frmChild, "divPopupReportDefinition");

		$("#divPopupReportDefinition").dialog("open")

		@*var i;
		$.get('@Url.Action("AddChildTable", "Reports", New With {.ReportID = Model.ID})', function (data) {
			$('#tabs-2').append(data);
		});*@


		//   var chilTableID = $("")

		var sChildren = new String("");
		var sChildrenNames = new String("");


		var gridData = $("#ChildTables").getRowData();
		var postData = JSON.stringify(gridData);


		// swap in some json or such like to get the ids of child tables
		sChildren = "2	84	";
		sChildrenNames = "2	Absence	84	Absence_Requests	";


//		debugger;

		@*$.ajax({
			type: "GET",
			url: '@Url.Action("getChildTable", "Reports", Nothing)',
			data: postData,
			contentType: "application/json; charset=utf-8",
		//	dataType: "json",
			dataType: "html",
			async: true,
			cache: false,
			success: function (msg) {
				$("#messageBox").html(msg);
			},
			error: function (XMLHttpRequest, textStatus, errorThrown) {
				debugger;
				alert(textStatus);
			}
		});*@


		//var frmGetChild = $("#frmGetChildTable");
		////frmGetChild.values = blahs
		//OpenHR.submitForm(frmGetChild, "divGetChildTable")


	//	var datarow = { ID: 0, TableID: 2, FilterID: 22, OrderID: 84, TableName: 'Absence', FilterName: 'This Years Absence', OrderName: 'Start_Date', Records: 84 };
//		var su = jQuery("#ChildTables").jqGrid('addRowData', 99, datarow);



		////$("[id^=ChildTables] [id$=__TableID]").each {
		////}

		//$( "[id^=ChildTables] [id$=__TableID]" ).each(function() {
		//  debugger;
		//  sChildren += this.value + ',';
		//});

		//var dataCollection = frmTables.elements;
		//if (dataCollection != null) {
		//  sReqdControlName = new String("txtTableChildren_");
		//  sReqdControlName = sReqdControlName.concat(frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value);

		//  for (i = 0; i < dataCollection.length; i++) {
		//    sControlName = dataCollection.item(i).name;
		//    if (sControlName == sReqdControlName) {
		//      sChildren = dataCollection.item(i).value;
		//      frmCustomReportChilds.childrenString.value = sChildren;
		//      break;
		//    }
		//  }
		//}


		// NPG Remmed:
//    var sURL = "util_customreportchilds" +
//"?childTableID=" + "0" +
//  "&childTable=" + "" +
//    "&childFilterID=" + "0" +
//      "&childFilter=" + "" +
//        "&childOrderID=" + "0" +
//          "&childOrder=" + "" +
//            "&childRecords=" + "0" +
//              "&childrenString=" + escape(sChildren) +
//                "&childrenNames=" + escape(sChildrenNames) +
//                  "&selectedChildString=" + escape("''") +
//                    "&childAction=" + "NEW" +
//                      "&childMax=" + "5";
//    openDialog(sURL, 365, 275, "no", "no");




//    var itemIndex = $("#ChildTables tr").length - 1;
//    e.preventDefault();

//    var newItem = $("<tr><td><input name='ChildTables[" + itemIndex + "].Records' value='23'><td/></tr>");

//    //      var newItem = $("<tr><td><input id='ChildTables" + itemIndex + "__Id' type='hidden' value='' class='iHidden'  name='Interests[" + itemIndex + "].Id' /><input type='text' id='Interests_" + itemIndex + "__InterestText' name='Interests[" + itemIndex + "].InterestText'/></td><td><input type='checkbox' value='true'  id='Interests_" + itemIndex + "__IsExperienced' name='Interests[" + itemIndex + "].IsExperienced' /></tr>");
//    $("#ChildTables").append(newItem);

	}

	function editChildTable(rowID) {

	//	var frmChild = $("#frmGetChildTable");

		//debugger;

	//	debugger;
	//	var currentRow = $('#ChildTables').jqGrid('getGridParam', 'selarrrow');
		if (rowID == 0) {
			rowID = $('#ChildTables').jqGrid('getGridParam', 'selarrrow');
		}

	//	var frmChild = $("#frmGetChildTable");
		//var postData = JSON.stringify($("#ChildTables").getRowData(rowID));

		var gridData = $("#ChildTables").getRowData(rowID);

	//	var blah = $.toJSON($("#ChildTables").getRowData(rowID));

		

		//gridData = {
		//	ID: "3",
		//	ReportID: "3",
		//	TableID: "3",
		//	FilterID: "3",
		//	OrderID: "3",
		//	Records: "3",
		//	TableName: "hello",
		//	FilterName: "hello",
		//	OrderName: "hello"
		//};


		OpenHR.OpenDialog("Reports/EditChildTable", "divPopupReportDefinition", gridData); // postData);



	}

	function removeAllChildTables() {
		$("#ChildTables").empty();
	}

	function selectChildTable(rowID) {
		alert(rowID);
	}



	//Public Property ID As Integer Implements IJsonSerialize.ID

	//Public Property TableID As Integer
	//Public Property FilterID As Integer
	//Public Property OrderID As Integer
	//Public Property Records As Integer

	//' these are for display purposes (better way?)
	//Public Property TableName As String
	//Public Property FilterName As String
	//Public Property OrderName As String
		
	$(function () {

		jQuery("#ChildTables").jqGrid({
			datatype: "jsonstring",
			datastr: '@Model.ChildTables.ToJsonResult',
			mtype: 'GET',
			jsonReader: {
				root: "rows", //array containing actual data
				page: "page", //current page
				total: "total", //total pages for the query
				records: "records", //total number of records
				repeatitems: false,
				id: "TableID" //index of the column with the PK in it
			},
			colNames: ['ID', 'TableID', 'FilterID', 'OrderID', 'Table', 'Filter', 'Order', 'Records'],
			colModel: [
				{ name: 'ID', index: 'id', sorttype: 'int', hidden: true },
				{ name: 'TableID', index: 'TableID', width: 100, hidden: true },
				{ name: 'FilterID', index: 'FilterID', width: 100, hidden: true },
				{ name: 'OrderID', index: 'OrderID', width: 100, hidden: true },
				{ name: 'TableName', index: 'TableName', width: 100 },
				{ name: 'FilterName', index: 'FilterName', width: 100 },
				{ name: 'OrderName', index: 'OrderName', width: 100 },
			{ name: 'Records', index: 'Records', width: 100 }
			],
			rowNum: 10,
			autowidth: true,
			rowTotal: 50,
			rowList: [10, 20, 30],
			shrinkToFit: true,
			pager: '#pcrud',
			sortname: 'TableID',
			loadonce: true,
			viewrecords: true,
			sortorder: "desc",
			editurl: 'server.php', // this is dummy existing url
			ondblClickRow: function (rowID) {
				editChildTable(rowID);
			},
		});
		jQuery("#ChildTables").jqGrid('navGrid', '#pcrud', {});

	});

	
</script>