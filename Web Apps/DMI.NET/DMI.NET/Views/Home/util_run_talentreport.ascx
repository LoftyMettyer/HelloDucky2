<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="DMI.NET.Repository" %>

<link href="<%: Url.LatestContent("~/Content/BulletGraph.css")%>" rel="stylesheet" type="text/css" />

<%
  dim repository As New ReportRepository
	Dim objReport = repository.LoadTalentReport(CInt(Session("utilid")), UtilityActionType.Edit)
%>

<div id="reportworkframe" data-framesource="util_run_talentreport" style="display: inline-block; width:100%; height: 100%;">
  <table id="gridReportData"></table> 
  
  <form action="util_run_talentreport_downloadoutput" method="post" id="frmExportData" name="frmExportData" target="submit-iframe">
	  <input type="hidden" id="txtPreview" name="txtPreview" value="<%=objReport.Output.IsPreview%>">
	  <input type="hidden" id="txtFormat" name="txtFormat" value="<%=objReport.Output.Format%>">
	  <input type="hidden" id="txtScreen" name="txtScreen" value="<%=objReport.Output.ToScreen%>">
	  <input type="hidden" id="txtPrinter" name="txtPrinter" value="<%=objReport.Output.ToPrinter%>">
	  <input type="hidden" id="txtPrinterName" name="txtPrinterName" value="<%=objReport.Output.PrinterName%>">
	  <input type="hidden" id="txtSave" name="txtSave" value="<%=objReport.Output.SaveToFile%>">
	  <input type="hidden" id="txtSaveExisting" name="txtSaveExisting" value="<%=objReport.Output.SaveExisting%>">
	  <input type="hidden" id="txtEmail" name="txtEmail" value="<%=objReport.Output.SendToEmail%>">
	  <input type="hidden" id="txtEmailAddr" name="txtEmailAddr" value="<%=objReport.Output.EmailGroupID%>">
	  <input type="hidden" id="txtEmailAddrName" name="txtEmailAddrName" value="<%=objReport.Output.EmailGroupName%>">
	  <input type="hidden" id="txtEmailSubject" name="txtEmailSubject" value="<%=objReport.Output.EmailSubject%>">
	  <input type="hidden" id="txtEmailAttachAs" name="txtEmailAttachAs" value="<%=objReport.Output.EmailAttachmentName%>">
	  <input type="hidden" id="txtEmailGroupAddr" name="txtEmailGroupAddr" value="">
	  <input type="hidden" id="txtFileName" name="txtFileName" value="<%=objReport.Output.Filename%>">
	  <input type="hidden" id="txtEmailGroupID" name="txtEmailGroupID" value="<%=objReport.Output.EmailGroupID%>">
	  <input type="hidden" id="txtUtilType" name="txtUtilType" value="<%=session("utilType")%>">
	  <input type="hidden" id="txtUtilID" name="txtUtilID" value="<%=Session("utilID")%>">
	  <input type="hidden" id="download_token_value_id" name="download_token_value_id"/>
	  <%=Html.AntiForgeryToken()%>
  </form>

</div>

  <div id="outputoptions" data-framesource="util_run_outputoptions" style="display: none;">
		  <% Html.RenderPartial("~/Views/Home/util_run_outputoptions.ascx")%>
  </div>

  <input type='hidden' id="txtDefn_Name" name="txtDefn_Name" value="<%:objReport.Name.ToString()%>">

<script>
  
  function getChartLine(competency, minScore, prefScore, actualScore) {

    var returnDiv;

    var chartTitleText = "Competency: " + competency +
      "\nMinimum Score: " + minScore +
      "\nPreferred Score: " + prefScore +
      "\nActual Score: " + actualScore;

    returnDiv = '<div title="' + chartTitleText + '" class="form-group"> \
    <div title="' + chartTitleText + '" class="col-sm-10"> \
      <div title="' + chartTitleText + '" class="bullet-graph blue"> \
        <div title="' + chartTitleText + '" class="graph"> \
          <div class="ui-accordion-header ui-state-default ui-accordion-header-collapsed region-1" style="width: ' + minScore + '%;"></div> \
          <div class="ui-accordion-header ui-state-default ui-accordion-header-active ui-state-active region-2" style="width: ' + (prefScore - minScore) + '%;"></div> \
          <div class="region-3" style="width: ' + (100 - Math.max(prefScore, minScore)) + '%;"></div>';

    if (actualScore > 0) {
      returnDiv = returnDiv + '<div class="ui-widget-header measure" style="width: ' + actualScore + '%;"></div>';
    }
          
    returnDiv = returnDiv + '<div class="ui-state-error target-1" style="width: ' + minScore + '%;"></div> \
                <div title="' + chartTitleText + '" class="ui-state-error target-1" style="width: ' + prefScore + '%;"></div> \
        </div> \
      </div> \
    </div> \
  </div>';

    return returnDiv;

  }

  function SnapColumnsToGrid() {

    var colModel = $("#gridReportData").jqGrid('getGridParam', 'colModel');
    var colCount = colModel.length;
    var totalWidth = $("#reportworkframe").width();
    var colWidth = 0;
    var i, nCol;

    colWidth = ((totalWidth) * 0.35);
    var matchWidth = colWidth * 0.2;
    var talentWidth = Math.max(colWidth * 0.70, 420);


    colWidth = (totalWidth - matchWidth - talentWidth) / (colCount - 5);
    for (i = 2, nCol = colCount - 3; i < nCol; i++) {
      $("#gridReportData").jqGrid('setColWidth', i, colWidth - 5, true);
    }

    $("#gridReportData").jqGrid('setColWidth', 'Match Score', matchWidth, false);
    $("#gridReportData").jqGrid('setColWidth', 'Talent Chart', talentWidth - 40, true);

    var gridWidth = $('#reportworkframe').width();
    $("#gridReportData").jqGrid('setGridWidth', gridWidth);

    if (menu_isSSIMode()) {
    	gridHeight = $('#reportworkframe').height() - 100;
    	$("#gridReportData").jqGrid('setGridHeight', gridHeight);
    } else {
    	gridHeight = $('#reportworkframe').height();
    	$("#gridReportData").jqGrid('setGridHeight', gridHeight);
    }
  }


  var grid = $("#gridReportData"),
    getColumnIndexByName = function (columnName) {
      var cm = grid.jqGrid('getGridParam', 'colModel');
      for (var i = 0, l = cm.length; i < l; i++) {
        if (cm[i].name === columnName) {
          return i; // return the index
        }
      }
      return -1;
    },
    gridLoaded = function () {
      setTimeout(function () {

        var index = getColumnIndexByName('Talent Chart');
        $('#gridReportData').find('tr.jqgrow td:nth-child(' + (index + 1) + ')').each(function () {
          var ar;
          try {
            ar = $.parseJSON($(this).text());
            if (ar && ar.length > 0) {
              var talentchartCellObject = this;
              $(talentchartCellObject).html("<table width='100%'></table>");

              $.each(ar, function (index, obj) {

                var cell1Css = "'width:160px;white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-size: small;'";
                var chartLine = getChartLine(obj.Competency, obj.MinScore, obj.PrefScore, obj.ActualScore);

                $(talentchartCellObject).find("table").append("<tr><td title='' style='width:80px;border: 0;'><div style=" + cell1Css + ">" + obj.Competency + "</div></td>" +
                  "<td style='width:150px;border:0;'>" + chartLine + "</td></tr>");
              });

            }
          } catch (e) { }
        });
      }, 100);

      if (menu_isSSIMode()) {
        $(".ui-dialog-buttonpane #cmdClose").show();
      } else {
      	$('.popup').dialog('option', 'buttons', window.newButtons);
      	$(".popup").dialog("open");
      	$("#divReportButtons #cmdClose").hide();
      	$(".ui-dialog-buttonpane #cmdOK").hide();
      	$(".ui-dialog-buttonpane #cmdCancel").removeClass('ui-state-focus');
      	$(".ui-dialog-buttonpane #cmdCancel").button({ disabled: true });
      }

      SnapColumnsToGrid();
    };

  var gridHeight;
  if (menu_isSSIMode()) {
    gridHeight = $('#reportworkframe').height() - 100;
  } else {
    gridHeight = $('#reportworkframe').height();
  }

  gridWidth = $('#reportworkframe').width();

  $.ajax({
    cache: false,
    url: '<%:Url.Action("getTalentReportData", "Home")%>',
    dataType: "json",
    error: function (xhr, st, err) {
      OpenHR.modalPrompt(xhr.responseJSON, 2, "", "");
      closeclick();
    },
    success: function (jsonData) {
	    if (typeof jsonData.colModel == "undefined") {
	    	gotNoData(jsonData);
	    } else {
		    gotData(jsonData);
	    }
    }
  });

	function gotData(jdata) {
  	$("#gridReportData").jqGrid({
  		datatype: "local",
  		data: $.parseJSON(JSON.stringify(jdata)).rows,
  		mtype: 'GET',
  		jsonReader: {
  			root: "rows", //array containing actual data
  			page: "page", //current page
  			total: "total", //total pages for the query
  			records: "records", //total number of records
  			repeatitems: false,
  			id: "ID_1"
  		},
  		colModel: jdata.colModel,
  		rowNum: 100,
  		viewrecords: true,
  		loadComplete: gridLoaded,
  		height: gridHeight,
  		width: gridWidth,
  		shrinkToFit: false,
  		autoWidth: false
  	});
	};

  function gotNoData(data) {
  	OpenHR.modalPrompt(data, 2, "", "");
  	closeclick();
  };

</script>
