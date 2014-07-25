﻿@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports HR.Intranet.Server.Enums
@code
	ViewBag.CustomPrefix = "Output."
End Code

@Inherits System.Web.Mvc.WebViewPage(Of Models.ReportOutputModel)

<br />
<div>
	<fieldset class="border0 width25 floatleft">
		<legend class="fontsmalltitle">Output Formats</legend>
		@Html.RadioButton("Output.Format", 0, Model.Format = OutputFormats.DataOnly, New With {.onclick = "changeOutputType('DataOnly')"})
		Data Only
		<br />
		@Html.RadioButton("Output.Format", 1, Model.Format = OutputFormats.CSV, New With {.onclick = "changeOutputType('CSV')"})
		CSV File
		<br />
		@Html.RadioButton("Output.Format", 2, Model.Format = OutputFormats.HTML, New With {.onclick = "changeOutputType('HTML')"})
		HTML Document		
		<br />
		@Html.RadioButton("Output.Format", 3, Model.Format = OutputFormats.WordDoc, New With {.onclick = "changeOutputType('WordDoc')"})
		Word Document		
		<br />
		@Html.RadioButton("Output.Format", 4, Model.Format = OutputFormats.ExcelWorksheet, New With {.onclick = "changeOutputType('ExcelWorksheet')"})
		Excel Worksheet		
		<br />
		@Html.RadioButton("Output.Format", 5, Model.Format = OutputFormats.ExcelGraph, New With {.onclick = "changeOutputType('ExcelGraph')"})
		Excel Chart		
		<br />
		@Html.RadioButton("Output.Format", 6, Model.Format = OutputFormats.ExcelPivotTable, New With {.onclick = "changeOutputType('ExcelPivotTable')"})
		Excel Pivot Table
	</fieldset>

	<fieldset id="outputdestinatonfieldset" class="border0 floatleft">
		<legend class="fontsmalltitle">Output Destinations</legend>
		@* Preview on Screen Section *@
		<fieldset class="border0">
			<legend>
				@Html.CheckBox("Output.IsPreview", Model.IsPreview)
				@Html.LabelFor(Function(m) m.IsPreview)
			</legend>
		</fieldset>

		@* Display Output On Screen Section *@
		<fieldset class="border0">
			<legend>
				@Html.CheckBox("Output.ToScreen", Model.ToScreen)
				@Html.LabelFor(Function(m) m.ToScreen)
			</legend>
		</fieldset>

		@* Send To Print Section *@
		<fieldset class="border0">
			<legend>
				@Html.CheckBox("Output.ToPrinter", Model.ToPrinter)
				@Html.LabelFor(Function(m) m.ToPrinter)
			</legend>
			@Html.TextBox("Output.PrinterName", Model.PrinterName, New With {.placeholder = "Default Printer", .class = "readonly width100"})
		</fieldset>

		@* Save To File Section *@
		<fieldset class="border0">
			<legend>
				@Html.CheckBox("Output.SaveToFile", Model.SaveToFile)
				@Html.LabelFor(Function(m) m.SaveToFile)
			</legend>
			@Html.TextBox("Output.Filename", Model.Filename, New With {.placeholder = "File Name", .class = "readonly"})
			@Html.LabelFor(Function(m) m.SaveExisting)
			@Html.EnumDropDownListFor(Function(m) m.SaveExisting)
		</fieldset>

		@* Send To Email Section *@
		<fieldset class="border0">
			<legend>
				@Html.CheckBoxFor(Function(m) m.SendToEmail, New With {Key .Name = "Output.SendToEmail"})
				@Html.LabelFor(Function(m) m.SendToEmail)
			</legend>
			<input type="button" class="ui-state-disabled width10" id="cmdEmailGroup" name="cmdEmailGroup" value="..." onclick="selectEmailGroup()" />
			<input type="text" id="txtEmailGroup" name="Output.EmailGroupName" class="width80 floatright" disabled value="@Model.EmailGroupName" />
			<br />
			@Html.LabelFor(Function(m) m.EmailSubject, New With {.class = "display-label_emails"})
			@Html.TextBoxFor(Function(m) m.EmailSubject, New With {Key .Name = "Output.EmailSubject", .class = "display-textbox-emails"})
			<br />
			@Html.LabelFor(Function(m) m.EmailAttachmentName, New With {.class = "display-label_emails"})
			@Html.TextBoxFor(Function(m) m.EmailAttachmentName, New With {Key .Name = "Output.EmailAttachmentName", .class = "display-textbox-emails"})
			@Html.ValidationMessage("Output.EmailGroupID")		<br />
			@Html.ValidationMessage("Output.EmailSubject")		<br />
			@Html.ValidationMessage("Output.EmailAttachmentName")		<br />
			@Html.ValidationMessage("Output.FileName")		<br />
			@Html.HiddenFor(Function(m) m.EmailGroupID, New With {.id = "txtEmailGroupID", Key .Name = "Output.EmailGroupID"})
		</fieldset>
	</fieldset>
</div>

<script type="text/javascript">

	function selectEmailGroup() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#txtEmailGroupID").val();

		OpenHR.modalExpressionSelect("EMAIL", tableID, currentID, function (id, name) {
			$("#txtEmailGroupID").val(id);
			$("#txtEmailGroup").val(name);
		});

	}

	function changeOutputType(type) {

		switch (type) {

			case "DataOnly":
				break;

			case "CSV":
				break;

			case "HTML":
				break;

			case "WordDoc":
				break;

			case "ExcelWorksheet":
				break;

			case "ExcelGraph":
				break;

			case "ExcelPivotTable":
				break;

		}

	}


</script>