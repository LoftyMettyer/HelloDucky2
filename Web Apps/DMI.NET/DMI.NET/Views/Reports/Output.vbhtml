@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports HR.Intranet.Server.Enums
@Inherits System.Web.Mvc.WebViewPage(Of Models.ReportOutputModel)

@Code
  Layout = Nothing
End Code

<br/>
<div>
  <div style="float:left">
    Output Format:
		<br/>
		@Html.RadioButton("Output.Format", 0, Model.Format = OutputFormats.fmtDataOnly)
		Data Only
		<br/>

		@Html.RadioButton("Output.Format", 1, Model.Format = OutputFormats.fmtCSV)
		CSV File
		<br />

		@Html.RadioButton("Output.Format", 2, Model.Format = OutputFormats.fmtHTML)
		HTML Document
		<br />

		@Html.RadioButton("Output.Format", 3, Model.Format = OutputFormats.fmtWordDoc)
		Word Document
		<br />

		@Html.RadioButton("Output.Format", 4, Model.Format = OutputFormats.fmtExcelWorksheet)
		Excel Worksheet
		<br />		

		@Html.RadioButton("Output.Format", 5, Model.Format = OutputFormats.fmtExcelGraph)
		Excel Chart
		<br />

		@Html.RadioButton("Output.Format", 6, Model.Format = OutputFormats.fmtExcelPivotTable)
		Excel Pivot Table

  </div>

  <div style="float:right">
    Output Destinations :
		<br/>
    @Html.CheckBox("Output.IsPreview", Model.IsPreview) Preview on screen
		<br />
		@Html.CheckBox("Output.ToScreen", Model.ToScreen) Display output on screen
		<br />
		@Html.CheckBox("Output.ToPrinter", Model.ToPrinter) Send to printer
		@Html.TextBox("Output.PrinterName", Model.PrinterName)
		<br />

		@Html.CheckBox("Output.SaveToFile", Model.SaveToFile)
		File Name: @Html.TextBox("Output.Filename", Model.Filename)
		If existing file : @Html.EnumDropDownListFor(Function(m) m.SaveExisting)
		<br/>

    @Html.CheckBox("Output.SendAsEmail", Model.SendAsEmail) Send As email
		<input type="text" id="txtEmailGroup" disabled />
		@Html.HiddenFor(Function(m) m.EmailGroupID, New With {.id = "txtEmailGroupID", .name = "Output.EmailGroupID"}))

		<input type="button" class="ui-state-disabled" id="cmdEmailGroup" name="cmdEmailGroup" value="..." style="padding-top: 0;" onclick="selectEmailGroup()" />
		<br />
		Email Subject: @Html.TextBox("Output.EmailSubject", Model.EmailSubject)
		<br/>
    Attach As: @Html.TextBox("Output.EmailAttachAs", Model.EmailAttachAs)

  </div>






</div>




<script type="text/javascript">

  function selectEmailGroup() {
    var sUrl;
    //  var frmOutputDef = OpenHR.getForm("outputoptions", "frmOutputDef");
    //   var frmEmailSelection = $("#frmEmailSelection")[0];
    //frmEmailSelection.EmailSelCurrentID.value = frmOutputDef.txtEmailGroupID.value;

    sUrl = "util_emailSelection" +
        "?EmailSelCurrentID=" + "13";
    openDialog(sUrl, (screen.width) / 3 + 40, (screen.height) / 2 - 50, "no", "no");
  }


</script>