@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.CustomReportModel)


@Code
  Layout = Nothing
End Code

<div id="divReportParents">

  <br/>

  <div style="float:left">
    Parent Table 1
    <input type="hidden" id="txtParent1ID" name="Parent1.ID" value="@Model.Parent1.ID" />
    <input type="text" disabled="disabled" id="txtParent1Name" value="@Model.Parent1.Name" />
    <label>
			@Html.RadioButton("Parent1.SelectionType", 0, Model.Parent1.SelectionType = Enums.RecordSelectionType.AllRecords)
			All Records
    </label>
    <br />
    <label>
			@Html.RadioButton("Parent1.SelectionType", 1, Model.Parent1.SelectionType = Enums.RecordSelectionType.Picklist)
			Picklist
      <input type="text" id="txtParent1PicklistID" name="Parent1.PicklistID" value="@Model.Parent1.PicklistID" />
      <input id="txtParent1Picklist" name="txtParent1Picklist" class="text textdisabled" disabled="disabled">
      <input id="cmdParent1Picklist" name="cmdParent1Picklist" type="button" value="..."
             onclick="selectRecordOption('p1', 'picklist')" />
    </label>
    <br />
    <label>
			@Html.RadioButton("Parent1.SelectionType", 2, Model.Parent1.SelectionType = Enums.RecordSelectionType.Filter)
			Filter
      <input type="text" id="txtParent1FilterID" name="Parent1.FilterID" value="@Model.Parent1.FilterID" />
      <input id="txtParent1Filter" name="txtParent1Filter" class="text textdisabled" disabled="disabled">
      <input id="cmdParent1Filter" name="cmdParent1Filter" type="button" value="..."
             onclick="selectRecordOption('p1', 'filter')" />
    </label>

  </div>

  <div style="float:right">
    Parent Table 2
    <input type="hidden" id="txtParent2ID" name="Parent2.ID" value="@Model.Parent2.ID" />
    <input type="text" disabled="disabled" id="txtParent2Name" value="@Model.Parent2.Name" />

    <label>
			@Html.RadioButton("Parent2.SelectionType", 0, Model.Parent2.SelectionType = Enums.RecordSelectionType.AllRecords)
      All Records
    </label>
    <br />
    <label>
			@Html.RadioButton("Parent2.SelectionType", 1, Model.Parent2.SelectionType = Enums.RecordSelectionType.Picklist)
			Picklist
      <input type="text" id="txtParent2PicklistID" name="Parent2.PicklistID" value="@Model.Parent2.PicklistID" />
      <input id="txtParent2Picklist" name="txtParent2Picklist" class="text textdisabled" disabled="disabled">
      <input id="cmdParent2Picklist" name="cmdParent2Picklist" type="button" value="..."
             onclick="selectRecordOption('p2', 'picklist')" />
    </label>
    <br />
    <label>
			@Html.RadioButton("Parent2.SelectionType", 0, Model.Parent2.SelectionType = Enums.RecordSelectionType.Filter)
			Filter
      <input type="text" id="txtParent2FilterID" name="Parent2.FilterID" value="@Model.Parent2.FilterID" />
      <input id="txtParent2Filter" name="txtParent2Filter" class="text textdisabled" disabled="disabled">
      <input id="cmdParent2Filter" name="cmdParent2Filter" type="button" value="..."
             onclick="selectRecordOption('p2', 'filter')" />
    </label>

  </div>

  <br />

</div>

<br/>

<br/>

Child Tables :
<div>
  <div class="left">
    @Html.TableFor("ChildTables", Model.ChildTables, Nothing)
  </div>

  <div class="right">
    <input type="button" id="btnChildAdd" value="Add Child Tabel" />
    <br/>
    <input type="button" id="btnChildEdit" value="Edit" />
    <br />
    <input type="button" id="btnChildRemove" value="Remove" />
    <br />
    <input type="button" id="btnChildRemoveAll" value="Remove All" onclick="removeAllChildTables();" />
  </div>

</div>


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
    alert("TODO remove child table");
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