@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports HR.Intranet.Server.Enums

@code
	ViewBag.CustomPrefix = "Output."
End Code

@Inherits System.Web.Mvc.WebViewPage(Of Models.ReportOutputModel)

<br/>
<div>
  <div style="float:left">
		Output Formats :
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

    @Html.CheckBox("Output.IsPreview", Model.IsPreview)
		@Html.LabelFor(Function(m) m.IsPreview)
		<br />

		@Html.CheckBox("Output.ToScreen", Model.ToScreen)
		@Html.LabelFor(Function(m) m.ToScreen)
		<br />

		@Html.CheckBox("Output.ToPrinter", Model.ToPrinter)
		@Html.LabelFor(Function(m) m.ToPrinter)
		@Html.TextBox("Output.PrinterName", Model.PrinterName)
		<br />

		@Html.CheckBox("Output.SaveToFile", Model.SaveToFile)
		@Html.LabelFor(Function(m) m.SaveToFile)

		@Html.LabelFor(Function(m) m.Filename)
		@Html.TextBox("Output.Filename", Model.Filename)

		@Html.LabelFor(Function(m) m.SaveExisting)
		@Html.EnumDropDownListFor(Function(m) m.SaveExisting)
		<br/>

		@Html.CheckBoxFor(Function(m) m.SendToEmail, New With {Key .Name = "Output.SendToEmail"})
		@Html.LabelFor(Function(m) m.SendToEmail)

		<input type="text" id="txtEmailGroup" name="Output.EmailGroupName" disabled value="@Model.EmailGroupName" />

		@Html.HiddenFor(Function(m) m.EmailGroupID, New With {.id = "txtEmailGroupID", Key .Name = "Output.EmailGroupID"})
		<input type="button" class="ui-state-disabled" id="cmdEmailGroup" name="cmdEmailGroup" value="..." style="padding-top: 0;" onclick="selectEmailGroup()" />
		<br />

		@Html.LabelFor(Function(m) m.EmailSubject)
		@Html.TextBoxFor(Function(m) m.EmailSubject, New With {Key .Name = "Output.EmailSubject"})
		<br/>

		@Html.LabelFor(Function(m) m.EmailAttachmentName)
    @Html.TextBoxFor(Function(m) m.EmailAttachmentName, New With {Key .Name = "Output.EmailAttachmentName"})
		<br/>

		@Html.ValidationMessage("Output.EmailGroupID")
		<br/>
		@Html.ValidationMessage("Output.EmailSubject")
		<br />
		@Html.ValidationMessage("Output.EmailAttachmentName")
		<br />
		@Html.ValidationMessage("Output.FileName")
		<br/>

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