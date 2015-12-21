@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.TalentReportModel)

@Code
	Layout = Nothing
End Code

<div>

  <form action="reports\util_def_mailmerge_downloadtemplate" style="display: none" method="post" id="frmDownloadTemplate" name="frmDownloadTemplate" target="submit-iframe">
    <input type="hidden" id="MailMergeId" name="MailMergeId" value="@Model.ID" />
    <input type="hidden" id="download_token_value_id" name="download_token_value_id" />
    @Html.AntiForgeryToken()
  </form>


  @Using (Html.BeginForm("util_def_talentreport", "Reports", FormMethod.Post, New With {.id = "frmReportDefintion", .name = "frmReportDefintion"}))

  @Html.HiddenFor(Function(m) m.ID)

  @<div id="tabs">
    <ul>
      <li><a href="#tabs-1">Definition</a></li>
      <li><a href="#report_definition_tab_columns">Columns</a></li>
      <li><a href="#report_definition_tab_order">Sort Order</a></li>
      <li><a href="#report_definition_tab_output">Output</a></li>
    </ul>

    <div id="tabs-1">
      @Code
      Html.RenderPartial("_Definition", Model)
      End Code

      <fieldset>
        Base Child : <select class="width70 floatright" name="BaseChildTableID" id="BaseChildTableID" onchange="refreshTalentReportBaseColumns(event.target);"></select>
      </fieldset>

      <fieldset>
        Base Column : <select class="width70 floatright" name="BaseChildColumnID" id="BaseChildColumnID"></select>
      </fieldset>

      <fieldset>
        Minimum Rating : <select class="width70 floatright" name="BaseMinimumRatingColumnID" id="BaseMinimumRatingColumnID"></select>
      </fieldset>

      <fieldset>
        Preferred Taing : <select class="width70 floatright" name="BasePreferredRatingColumnID" id="BasePreferredRatingColumnID"></select>
      </fieldset>

      <fieldset>
        <fieldset class="">
          Match Table : <select class="width70 floatright" name="MatchTableID" id="MatchTableID"></select>
        </fieldset>

        All/Picklist/Filter goes here
       <br/>

        <fieldset class="">
          Match Child : <select class="width70 floatright" name="MatchChildTableID" id="MatchChildTableID" onchange="refreshTalentReportMatchColumns(event.target);"></select>
        </fieldset>
        <fieldset class="">
          Match Column : <select class="width70 floatright" name="MatchChildColumnID" id="MatchChildColumnID"></select>
        </fieldset>
        <fieldset class="">
          Actual Rating : <select class="width70 floatright" name="MatchChildRatingColumnID" id="MatchChildRatingColumnID"></select>
        </fieldset>

        <br />
        Match Any / Match Against goes here

      </fieldset>


    </div>

    <div id="report_definition_tab_columns">
      @Code
        Html.RenderPartial("_ColumnSelection", Model)
      End Code
    </div>

    <div id="report_definition_tab_order">
      @Code
      Html.RenderPartial("_SortOrder", Model)
      End Code
    </div>

    <div id="report_definition_tab_output">
      @Code
      Html.RenderPartial("_Output", Model.Output)
      End Code
    </div>
  </div>
  @Html.AntiForgeryToken()
  End Using

</div>

<script type="text/javascript">

  function setTalentDefinitionDetails() {
    $('#MatchTableID').val("@Model.MatchTableID");
    refreshTalentReportChildTables();
  }

  function refreshTalentReportChildTables() {

    $.ajax({
      url: 'Reports/GetChildTables?parentTableId=' + $("#BaseTableID").val(),
      datatype: 'json',
      mtype: 'GET',
      cache: false,
      success: function (json) {

        var option = "";
        for (var i = 0; i < json.length; i++) {
          option += "<option value='" + json[i].id + "'>" + json[i].Name + "</option>";
        }
        $("select#BaseChildTableID").html(option);
        $('#BaseChildTableID').val("@Model.BaseChildTableID");
        refreshTalentReportBaseColumns();
      }
    });

    $.ajax({
      url: 'Reports/GetChildTables?parentTableId=' + $("#MatchTableID").val(),
      datatype: 'json',
      mtype: 'GET',
      cache: false,
      success: function (json) {

        var option = "";
        for (var i = 0; i < json.length; i++) {
          option += "<option value='" + json[i].ID + "'>" + json[i].Name + "</option>";
        }
        $("select#MatchChildTableID").html(option);
        $('#MatchChildTableID').val("@Model.MatchChildTableID");
        refreshTalentReportMatchColumns();

      }
    });

  }

  function refreshTalentReportBaseColumns() {

    $.ajax({
      url: 'Reports/GetAvailableColumnsForTable?TableID=' + $("#BaseChildTableID").val(),
      datatype: 'json',
      mtype: 'GET',
      cache: false,
      success: function (json) {

        var option = "";

        for (var i = 0; i < json.length; i++) {
          option += "<option value='" + json[i].ID + "' data-datatype='" + json[i].DataType + "' data-size='" + json[i].ColumnSize + "' data-decimals='" + json[i].Decimals + "'>" + json[i].Name + "</option>";
        }

        $("select#BaseChildColumnID").html(option);
        $("select#BaseMinimumRatingColumnID").html(option);
        $("select#BasePreferredRatingColumnID").html(option);

        $('#BaseChildColumnID').val("@Model.BaseChildColumnID");
        $('#BaseMinimumRatingColumnID').val("@Model.BaseMinimumRatingColumnID");
        $('#BasePreferredRatingColumnID').val("@Model.BasePreferredRatingColumnID");

      }
    });
  }

  function refreshTalentReportMatchColumns() {

    $.ajax({
      url: 'Reports/GetAvailableColumnsForTable?TableID=' + $("#MatchChildTableID").val(),
      datatype: 'json',
      mtype: 'GET',
      cache: false,
      success: function (json) {

        var option = "";

        for (var i = 0; i < json.length; i++) {
          option += "<option value='" + json[i].ID + "' data-datatype='" + json[i].DataType + "' data-size='" + json[i].ColumnSize + "' data-decimals='" + json[i].Decimals + "'>" + json[i].Name + "</option>";
        }

        $("select#MatchChildColumnID").html(option);
        $("select#MatchChildRatingColumnID").html(option);

        $('#MatchChildColumnID').val("@Model.MatchChildColumnID");
        $('#MatchChildRatingColumnID').val("@Model.MatchChildRatingColumnID");
      }
    });

  }


  $(function () {
    $("#tabs").tabs({
      activate: function (event, ui) {
        //Tab click event fired
        if (ui.newTab.text() == "Columns") {
          resizeColumnGrids();
        }
        if (ui.newTab.text() == "Sort Order") {
          //resize the Event Details grid to fit
          var workPageHeight = $('#workframeset').height();
          var gridTopPos = $('#divSortOrderDiv').position().top;
          var tabHeight = $('#tabs>.ui-tabs-nav').outerHeight();
          var marginHeight = 40;
          var gridHeight = workPageHeight - gridTopPos - tabHeight - marginHeight;
          $("#SortOrders").jqGrid('setGridHeight', gridHeight);

          var gridWidth = $('#divSortOrderDiv').width();
          $("#SortOrders").jqGrid('setGridWidth', gridWidth);
        }
      }
    });
    $('input[type=number]').numeric();
    $('#description, #Name').css('width', $('#BaseTableID').width());
  });
  function resizeColumnGrids() {
    var gridWidth = $('#columnsAvailable').width() - 10;
    $("#AvailableColumns").jqGrid('setGridWidth', gridWidth);
    $('#SelectedTableID').width(gridWidth);

    gridWidth = $('#columnsSelected').width() - 10;
    $("#SelectedColumns").jqGrid('setGridWidth', gridWidth);

    //var gridHeight = $('#columnsAvailable').parent().height() - 20;
    var gridHeight = screen.height / 3;
    $("#SelectedColumns").jqGrid('setGridHeight', gridHeight);
    $("#AvailableColumns").jqGrid('setGridHeight', gridHeight);

    //column aggregate widths
    $('.colAggregates').find('.tablecell').css('width', gridWidth / 3);
  }

  $("#workframe").attr("data-framesource", "UTIL_DEF_TALENTREPORT");
</script>
