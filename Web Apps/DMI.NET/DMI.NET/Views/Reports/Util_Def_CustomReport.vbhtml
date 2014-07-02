@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.CustomReportModel)


<style>
  .wrapper {
    width: 100%;
    overflow-x: auto;
    overflow-y: hidden;
  }

  .inner {
    width: 100%;
  }

  .left {
    width: 50%;
    float: left;
  }

  .right {
    width: 50%;
    float: left;
  }
</style>


<div>

@Using (Html.BeginForm("util_def_customreports_submit", "Reports", FormMethod.Get, New With {.id = "frmReportDefintion", .name = "frmReportDefintion",
																																								.defaultbutton = "submitCustomReport"}))

  @Html.HiddenFor(Function(m) m.ID)
                     
  @<div id="tabs">

        <ul>
            <li><a href="#tabs-1">Definition</a></li>
            <li><a href="#tabs-2">Tables</a></li>
            <li><a href="#report_definition_tab_columns">Columns</a></li>
            <li><a href="#report_definition_tab_order">Order</a></li>
            <li><a href="#report_definition_tab_output">Output</a></li>
        </ul>

        <div id="tabs-1">
          @Code                        
						Html.RenderPartial("Definition", Model)
          End Code


					<div>
						Report Options: (MOVED FROM Output tab (discuss))
						<br/>
						@Html.CheckBox("IsSummary", Model.IsSummary)Summary Report
						<br/>
						@Html.CheckBox("IgnoreZerosForAggregates", Model.IgnoreZerosForAggregates)Ignore zeros when calculating aggregates

					</div>

        </div>

        <div id="tabs-2">
          @Code
						Html.RenderPartial("RelatedTables", Model)
          End Code
        </div>

        <div id="report_definition_tab_columns">
          @Code
						Html.RenderPartial("ColumnSelection", Model.Columns)
          End Code
        </div>

     <div id="report_definition_tab_order">
        @Code
					Html.RenderPartial("SortOrder", Model)
        End Code
     </div>

     <div id="report_definition_tab_output">
       @Code
			 	 Html.RenderPartial("Output", Model.Output)
       End Code
     </div>

  <div id="divEmailGroupSelection">




  </div>



    </div>


  


  End Using

  <form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
    @Code
      Html.RenderPartial("~/Views/Shared/gotoWork.ascx")
    End Code
  </form>

    
</div>

<script type="text/javascript">

  $(function () {
    $("#tabs").tabs();
  });


  function submitForm() {

  	var frmSubmit = $("#frmReportDefintion");
    OpenHR.submitForm(frmSubmit);

  }

    $("#workframe").attr("data-framesource", "UTIL_DEF_CUSTOMREPORTS");


    $('#tabs').bind('tabsshow', function (event, ui) {

      var tabPage;

      if (ui.index == "0") {
        tabPage = $("#frmCustomReportsTab1");
      }
      
      if (ui.index == "1") {

        tabPage = $("#frmCustomReportsTab2");

        //  $.getJSON("Util_Def_CustomReports_getColumnInfo", function (data) {

        //    var items = [];
        //    $.each(data, function (key, val) {
        // //     debugger;
        //      items.push("<li id='" + val.TableName + "'>" + val + "</li>");
        //    });

        //    $("<ul/>", {
        //      "class": "my-new-list",
        //      html: items.join("")
        //    }).appendTo("#showColumns");
        //  });
      }

      if (ui.index == "2") {
        tabPage = $("#frmCustomReportsTab3");
      }


      //})

     //   OpenHR.submitForm(tabPage, "tabContent2");

    })


</script>
