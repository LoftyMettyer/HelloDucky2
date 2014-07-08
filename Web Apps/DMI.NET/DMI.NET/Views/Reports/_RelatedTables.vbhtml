@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Enums
@Inherits System.Web.Mvc.WebViewPage(Of Models.CustomReportModel)


<div id="divReportParents">

	<fieldset @Model.Parent1.Visibility>
		<legend>Parent 1 :</legend>

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

<fieldset>
	<legend>Child Tables :</legend>

  <div class="left">
    @Html.TableFor("ChildTables", Model.ChildTables, Nothing)
  </div>

  <div class="right">
		<input type="button" id="btnChildAdd" value="Add..." onclick="addChildTable();" />
    <br/>
    <input type="button" id="btnChildEdit" value="Edit..." />
    <br />
    <input type="button" id="btnChildRemove" value="Remove" />
    <br />
    <input type="button" id="btnChildRemoveAll" value="Remove All" onclick="removeAllChildTables();" />
  </div>

</fieldset>


<script type="text/javascript">

  // slightly hacked version of the orignal from util_def_customreports.js
  // passes in generated stirng of currently selected item (I'm sure this code can be cleverised.)
  function addChildTable() {

    //   var chilTableID = $("")

    var sChildren = new String("");
    var sChildrenNames = new String("");


    // swap in some json or such like to get the ids of child tables
    sChildren = "2	84	";
    sChildrenNames = "2	Absence	84	Absence_Requests	";

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

    var sURL = "util_customreportchilds" +
"?childTableID=" + "0" +
  "&childTable=" + "" +
    "&childFilterID=" + "0" +
      "&childFilter=" + "" +
        "&childOrderID=" + "0" +
          "&childOrder=" + "" +
            "&childRecords=" + "0" +
              "&childrenString=" + escape(sChildren) +
                "&childrenNames=" + escape(sChildrenNames) +
                  "&selectedChildString=" + escape("''") +
                    "&childAction=" + "NEW" +
                      "&childMax=" + "5";
    openDialog(sURL, 365, 275, "no", "no");




    var itemIndex = $("#ChildTables tr").length - 1;
    e.preventDefault();

    var newItem = $("<tr><td><input name='ChildTables[" + itemIndex + "].Records' value='23'><td/></tr>");

    //      var newItem = $("<tr><td><input id='ChildTables" + itemIndex + "__Id' type='hidden' value='' class='iHidden'  name='Interests[" + itemIndex + "].Id' /><input type='text' id='Interests_" + itemIndex + "__InterestText' name='Interests[" + itemIndex + "].InterestText'/></td><td><input type='checkbox' value='true'  id='Interests_" + itemIndex + "__IsExperienced' name='Interests[" + itemIndex + "].IsExperienced' /></tr>");
    $("#ChildTables").append(newItem);

  }

  function editChildTable(rowID) {
    alert(rowID);
  }

  function removeAllChildTables() {
    $("#ChildTables").empty();
  }

  function selectChildTable(rowID) {
    alert(rowID);
  }

    
  $(function () {

    tableToGrid("#ChildTables", {
      onSelectRow: function (rowID) {
        selectChildTable(rowID);
      },
      ondblClickRow: function (rowID) {
        editChildTable(rowID);
      },
      cmTemplate: { sortable: false },
      ignoreCase: true,
      pager: $('#pager-coldata'),
      rowList: [],        // disable page size dropdown
      pgbuttons: false,     // disable page control like next, back button
      pgtext: null,         // disable pager text like 'Page 0 of 10'
      viewrecords: false,    // disable current view record text like 'View 1-10 of 100'            
      rowNum: 1000
    });

    $("#btnChildAdd").click(function (e) {
      addChildTable();
    });

  });


</script>