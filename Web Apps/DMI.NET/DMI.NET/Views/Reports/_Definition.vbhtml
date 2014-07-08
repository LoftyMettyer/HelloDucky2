@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Enums
@Inherits System.Web.Mvc.WebViewPage(Of Models.ReportBaseModel)

<fieldset>

	<div class="inner">
		<div class="left">

			@Html.LabelFor(Function(m) m.Name)
			@Html.TextBoxFor(Function(m) m.Name)
			@Html.ValidationMessageFor(Function(m) m.Name)

			<br />
			@Html.LabelFor(Function(m) m.Description)
			@Html.TextBox("description", Model.Description)
			@Html.ValidationMessageFor(Function(m) m.Description)

		</div>

		<div class="right">
			<div class="editor-field-greyed-out">
				Owner: @Html.TextBoxFor(Function(m) m.Owner, New With {.readonly = "true"})
			</div>
			<br />
			Access : @Html.Raw(Html.AccessGrid("GroupAccess", Model.GroupAccess, Nothing))
		</div>
	</div>

</fieldset>

<fieldset>
	<legend>Data :</legend>
	<br />

	<div class="inner">

		<div class="left">
			Base Table : @Html.DropDownList("BaseTableID", Model.BaseTables)
		</div>

		<div class="right">

			Records :
			<br />
			@Html.RadioButton("selectiontype", RecordSelectionType.AllRecords, Model.SelectionType = RecordSelectionType.AllRecords, New With {.onclick = "changeRecordOption('Base','all')"})
			All Records
			<br />

			@Html.RadioButton("selectiontype", RecordSelectionType.Picklist, Model.SelectionType = RecordSelectionType.Picklist, New With {.onclick = "changeRecordOption('Base','picklist')"})
			Picklist
			<input type="hidden" id="txtBasePicklistID" name="picklistID" value="@Model.PicklistID" />
			@Html.TextBoxFor(Function(m) m.PicklistName, New With {.id = "txtBasePicklist", .readonly = "true"})
			@Html.EllipseButton("cmdBasePicklist", "selectRecordOption('base', 'picklist')", Model.SelectionType = RecordSelectionType.Picklist)
			<br />

			@Html.RadioButton("selectiontype", RecordSelectionType.Filter, Model.SelectionType = RecordSelectionType.Filter, New With {.onclick = "changeRecordOption('Base','filter')"})
			Filter
			<input type="hidden" id="txtBaseFilterID" name="filterID" value="@Model.FilterID" />
			@Html.TextBoxFor(Function(m) m.FilterName, New With {.id = "txtBaseFilter", .readonly = "true"})
			@Html.EllipseButton("cmdBaseFilter", "selectRecordOption('base', 'filter')", Model.SelectionType = RecordSelectionType.Filter)
			<br />

			@Html.CheckBoxFor(Function(m) m.DisplayTitleInReportHeader)
			@Html.LabelFor(Function(m) m.DisplayTitleInReportHeader)
			<br />

			@Html.ValidationMessageFor(Function(m) m.PicklistID)
			@Html.ValidationMessageFor(Function(m) m.FilterID)

		</div>

		<input type="hidden" id="ctl_DefinitionChanged" name="HasChanged" value="False" />

		<input type="hidden" id="baseHidden" name="baseHidden">
		<input type="hidden" id="p1Hidden" name="p1Hidden">
		<input type="hidden" id="p2Hidden" name="p2Hidden">

	</div>

</fieldset>

  <form id="frmCustomReportChilds" name="frmCustomReportChilds" target="childselection" action="util_customreportchilds" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="childTableID" name="childTableID">
	<input type="hidden" id="childTable" name="childTable">
	<input type="hidden" id="childFilterID" name="childFilterID">
	<input type="hidden" id="childFilter" name="childFilter">
	<input type="hidden" id="childOrderID" name="childOrderID">
	<input type="hidden" id="childOrder" name="childOrder">
	<input type="hidden" id="childRecords" name="childRecords">
	<input type="hidden" id="childrenString" name="childrenString">
	<input type="hidden" id="childrenNames" name="childrenNames">
	<input type="hidden" id="selectedChildString" name="selectedChildString">
	<input type="hidden" id="childAction" name="childAction" value="NEW">
	<input type="hidden" id="childMax" name="childMax" value="5">
</form>


<script type="text/javascript">

	$(function () {

		// tighten up these input selectors
		$("#frmReportDefintion :input").on("change", function () { enableSaveButton(this); });


	});

	function warning() {
		return "You will lose your changes if you do not save before leaving this page.\n\nWhat do you want to do?";
	}

	function enableSaveButton() {

		if ($("#ctl_DefinitionChanged").val() == "False") {
			$("#ctl_DefinitionChanged").val("True");
			menu_toolbarEnableItem("mnutoolSaveRecord", true);
		}
		window.onbeforeunload = warning;
	}

	function changeRecordOption(psTable, psType) {

		if (psType == "all") {
			button_disable($("#cmd" + psTable + "Picklist")[0], true);
			button_disable($("#cmd" + psTable + "Filter")[0], true);
			$("#txt" + psTable + "Filter").val("");
			$("#txt" + psTable + "Picklist").val("");
			$("#txt" + psTable + "PicklistID").val(0);
			$("#txt" + psTable + "FilterID").val(0);
		}

		if (psType == "picklist") {
			button_disable($("#cmd" + psTable + "Picklist")[0], false)
			button_disable($("#cmd" + psTable + "Filter")[0], true)
			$("#txt" + psTable + "Filter").val("");
			$("#txt" + psTable + "FilterID").val(0);

			if ($("#txt" + psTable + "PicklistID").val() == 0) {
				$("#txt" + psTable + "Picklist").val("<None>");
			}

		}

		if (psType == "filter") {
			button_disable($("#cmd" + psTable + "Picklist")[0], true)
			button_disable($("#cmd" + psTable + "Filter")[0], false)
			$("#txt" + psTable + "Picklist").val("");
			$("#txt" + psTable + "PicklistID").val(0);
	
			if ($("#txt" + psTable + "FilterID").val() == 0) {
				$("#txt" + psTable + "Filter").val("<None>");
			}

		}

	}

  function selectRecordOption(psTable, psType) {

  	var sURL;										 
  	var frmRecordSelection = $("#frmRecordSelection")[0];
    var iCurrentID;

    if (psTable == 'base') {

    	var e = $("#BaseTableID")[0];
      var iTableID = e.options[e.selectedIndex].value;

      if (psType == 'picklist') {
      	iCurrentID = $("#txtBasePicklistID").val();
      }
      else {
        iCurrentID = $("#txtBaseFilterID").val();
      }
    }
    if (psTable == 'p1') {
      iTableID = $("#txtParent1ID").val();

      if (psType == 'picklist') {
        iCurrentID = $("#txtParent1PicklistID").val();
      }
      else {
        iCurrentID = $("#txtParent1FilterID").val();
      }
    }
    if (psTable == 'p2') {
      iTableID = $("#txtParent2ID").val();

      if (psType == 'picklist') {
        iCurrentID = $("#txtParent2PicklistID").val();
      }
      else {
        iCurrentID = $("#txtParent2FilterID").val();
      }
    }

    var strDefOwner = $("#Owner").val();
    var strCurrentUser = $("#Owner").val();
    var isOwner;

    strDefOwner = strDefOwner.toLowerCase();
    strCurrentUser = strCurrentUser.toLowerCase();

    if (strDefOwner == strCurrentUser) {
    	isOwner = '1';
    }
    else {
    	isOwner = '0';
    }

    sURL = "util_recordSelection" +
			"?recSelType=" + psType +
				"&recSelTableID=" + iTableID +
					"&recSelCurrentID=" + iCurrentID +
						"&recSelTable=" + psTable +
							"&recSelDefOwner=" + isOwner +
								"&recSelDefType=" + escape("Selection");
    openDialog(sURL, (screen.width) / 3 + 40, (screen.height) / 2, "no", "no");

  }



</script>

