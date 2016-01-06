<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="DMI.NET.Models" %>
<%@ Import Namespace="DMI.NET.Repository" %>

<script src="<%:Url.LatestContent("~/Scripts/jquery/jquery.sparkline.min.js") %>"></script>

<%
  dim repository As New ReportRepository

  Dim objReport = repository.LoadTalentReport(1, UtilityActionType.Edit)
  
   %>

<div id="reportworkframe" data-framesource="util_run_talentreport" style="display: inline-block; width:100%; height: 100%;">
	<table id="gridReportData"></table>
</div>

<div id="outputoptions" data-framesource="util_run_outputoptions" style="display: none;">
		<% Html.RenderPartial("~/Views/Home/util_run_outputoptions.ascx")%>
</div>

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



<script>
  
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
                var target = obj.PrefScore,
                  performance = obj.ActualScore,
                  range1 = 100,
                  range2 = target,
                  range3 = obj.MinScore;
                var graphData = [target, performance, range1, range2, range3];

                var cell1Css = "'width:80px;white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-size: small;'";
                var chartTitleText = "Minimum Score: " + obj.MinScore +
                  "\nPreferred Score: " + obj.PrefScore +
                  "\nActual Score: " + obj.ActualScore;

                $(talentchartCellObject).find("table").append("<tr ><td style='width:80px;padding:2px;border: 0;'><div style=" + cell1Css + ">" + obj.Competency + "</div></td>" +
                  "<td style='width:150px;border:0;' title='" + chartTitleText + "' class='graph_" + index + "'></td></tr>");

                //Create the graph and add it to the 2nd cell.
                $(talentchartCellObject).find(".graph_" + index).sparkline(graphData, { type: 'bullet', targetColor: 'red', width: '150px' });
              });

            }
          } catch (e) { }
        });
      }, 100);

      if (menu_isSSIMode()) {
        $(".ui-dialog-buttonpane #cmdClose").show();
      } else {
        $("#divReportButtons #cmdClose").hide();
        setTimeout(resizeGrid, 100);
      }

    };

  var gridHeight;
  if (menu_isSSIMode()) {
    gridHeight = $('#reportworkframe').height() - 100;
  } else {
    gridHeight = 'auto';
  }


  $.ajax({
    cache: false,
    url: '<%:Url.Action("getTalentReportData", "Home")%>',
    dataType: "json",
    success: function(jsonData) {
      $("#gridReportData").jqGrid({
        datatype: "local",
        data: $.parseJSON(JSON.stringify(jsonData)).rows,
        mtype: 'GET',
        jsonReader: {
          root: "rows", //array containing actual data
          page: "page", //current page
          total: "total", //total pages for the query
          records: "records", //total number of records
          repeatitems: false,
          id: "ID_1"
        },
        colModel: jsonData.colModel,
        rowNum: 100,
        sortname: 'Match Score',
        viewrecords: true,
        sortorder: "desc",
        loadComplete: gridLoaded,
        autowidth: true,
        height: gridHeight,
        loadError: function(xhr, st, err) {
          OpenHR.modalPrompt(xhr.responseJSON, 2, "", "");
          closeclick();
        }
      });
    }
  });

</script>
