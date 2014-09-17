﻿@Imports DMI.NET
@Imports DMI.NET.Helpers

@code
	Html.EnableClientValidation()
End Code

@Inherits System.Web.Mvc.WebViewPage(Of Models.ReportOutputModel)

	<fieldset class="border0 width20 floatleft">
		<legend class="fontsmalltitle">Output Formats</legend>
		<fieldset id="outputformats">
			@Html.RadioButton("Output.Format", 0, Model.Format = OutputFormats.DataOnly, New With {.onchange = "changeOutputType('DataOnly')"})
			Data Only
			<br />

			<div class="hideforcalendarreport">
				@Html.RadioButton("Output.Format", 1, Model.Format = OutputFormats.CSV, New With {.onchange = "changeOutputType('CSV')"})
				<span class="DataManagerOnly">CSV File</span>
				<br />
			</div>

			@Html.RadioButton("Output.Format", 2, Model.Format = OutputFormats.HTML, New With {.onchange = "changeOutputType('HTML')"})
			<span class="DataManagerOnly">HTML Document</span>
			<br />

			@Html.RadioButton("Output.Format", 3, Model.Format = OutputFormats.WordDoc, New With {.onchange = "changeOutputType('WordDoc')"})
			<span class="DataManagerOnly">Word Document</span>
			<br />

			@Html.RadioButton("Output.Format", 4, Model.Format = OutputFormats.ExcelWorksheet, New With {.onchange = "changeOutputType('ExcelWorksheet')"})
			Excel Worksheet
			<br />

			<div class="hideforcalendarreport">
				@Html.RadioButton("Output.Format", 5, Model.Format = OutputFormats.ExcelGraph, New With {.onchange = "changeOutputType('ExcelGraph')"})
				Excel Chart
				<br />
				@Html.RadioButton("Output.Format", 6, Model.Format = OutputFormats.ExcelPivotTable, New With {.onchange = "changeOutputType('ExcelPivotTable')"})
				Excel Pivot Table
			</div>
			<br />
		</fieldset>
	</fieldset>

	<fieldset id="outputdestinatonfieldset" class="border0 floatleft width70">
		<legend class="fontsmalltitle">Output Destinations</legend>

		<fieldset class="border0 reportdefpreview">
			<div>
				@Html.CheckBoxFor(Function(m) m.IsPreview, New With {Key .Name = "Output.IsPreview"})
				@Html.LabelFor(Function(m) m.IsPreview)
			</div>
		</fieldset>

		<fieldset class="border0 reportdefscreen">
			<div>
				@Html.CheckBoxFor(Function(m) m.ToScreen, New With {Key .Name = "Output.ToScreen"})
				@Html.LabelFor(Function(m) m.ToScreen)
			</div>
		</fieldset>

		<fieldset class="border0 reportdefprinter DataManagerOnly">
			<div class="width30 floatleft">
				@Html.CheckBoxFor(Function(m) m.ToPrinter, New With {Key .Name = "Output.ToPrinter"})
				@Html.LabelFor(Function(m) m.ToPrinter, New With {.class = "DataManagerOnly"})
			</div>
			<div class="width70 floatleft">
				@Html.TextBoxFor(Function(m) m.PrinterName, New With {.Name = "Output.PrinterName", .class = "DataManagerOnly width100", .readonly = "true"})
			</div>
		</fieldset>

		<fieldset class="border0">
			<div class="reportdeffile">
				<div class="width30 floatleft">
					@Html.CheckBoxFor(Function(m) m.SaveToFile, New With {Key .Name = "Output.SaveToFile", .onclick = "setOutputToFile();", .class = "DataManagerOnly"})
					@Html.LabelFor(Function(m) m.SaveToFile, New With {.class = "DataManagerOnly"})
				</div>
				<div class="width70 floatleft">
					@Html.TextBoxFor(Function(m) m.Filename, New With {.Name = "Output.Filename", .class = "width100"})
				</div>
			</div>
		</fieldset>

		<fieldset>
			<div id="outputtabfilenametextbox" class="reportdeffile">
				@Html.LabelFor(Function(m) m.SaveExisting, New With {.class = "display-label_file"})
				@Html.CustomEnumDropDownListFor(Function(m) m.SaveExisting, New With {.Name = "Output.SaveExisting", .class = "DataManagerOnly", .readonly = "true"})
			</div>
		</fieldset>

		<div class="reportdefemail">
			<fieldset id="fieldsetsendemail" class="reportdefemail">
				<div class="">
					<div class="width30 floatleft">
						@Html.CheckBoxFor(Function(m) m.SendToEmail, New With {Key .Name = "Output.SendToEmail", .onclick = "setOutputToEmail();"})
						@Html.LabelFor(Function(m) m.SendToEmail)
					</div>
					<div class="width70 floatleft">
						@Html.HiddenFor(Function(m) m.EmailGroupID, New With {.Name = "Output.EmailGroupID", .id = "txtEmailGroupID"})
						@Html.TextBoxFor(Function(m) m.EmailGroupName, New With {.Name = "Output.EmailGroupName", .id = "txtEmailGroup", .readonly = "readonly", .class = "display-textbox-emails", .style = ""})
						<input type="button" class="reportdefemail" id="cmdEmailGroup" name="cmdEmailGroup" value="..." onclick="selectEmailGroup()" />
					</div>
				</div>
			</fieldset>

			<fieldset id="fieldsetsubjectemail">
				@Html.LabelFor(Function(m) m.EmailSubject, New With {.class = "display-label_emails"})
				@Html.TextBox("Output.EmailSubject", Model.EmailSubject, New With {.class = "display-textbox-emails"})
			</fieldset>
			<fieldset id="fieldseattachas">
				@Html.LabelFor(Function(m) m.EmailAttachmentName, New With {.class = "display-label_emails"})
				@Html.TextBoxFor(Function(m) m.EmailAttachmentName, New With {.Name = "Output.EmailAttachmentName", .class = "display-textbox-emails"})
			</fieldset>
		</div>

		<br />
		@Html.ValidationMessage("Output.EmailGroupID")		<br />
		@Html.ValidationMessage("Output.EmailSubject")		<br />
		@Html.ValidationMessage("Output.EmailAttachmentName")		<br />
		@Html.ValidationMessage("Output.FileName")		<br />
	</fieldset>

	<fieldset class="DataManagerOnly width100">
		Note: Options marked in red are unavailable in OpenHR Web.
	</fieldset>


<script type="text/javascript">

	function setOutputToFile() {

		var bSelected = $("#SaveToFile").prop('checked');

		$(".reportdeffile").children().attr("readonly", !bSelected);

		if (!bSelected) {
			$("#Filename").val("");
		}

		saveToFileChecked();

	}

	function setOutputToEmail() {
		var bSelected = $("#SendToEmail").prop('checked');

		$(".reportdefemail").children().attr("readonly", !bSelected);
		button_disable($("#cmdEmailGroup")[0], !bSelected);

		if (!bSelected) {
			$(".reportdefemail").children().val("");
			$("#txtEmailGroupID").val(0);
		}

		sendAsEmailChecked();
	}

	function selectEmailGroup() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#txtEmailGroupID").val();

		OpenHR.modalExpressionSelect("EMAIL", tableID, currentID, function (id, name) {
			$("#txtEmailGroupID").val(id);
			$("#txtEmailGroup").val(name);
		},400,400);

	}

	function selectOutputType(type) {

		$(".reportdefpreview").children().removeAttr("readonly");
		$(".reportdefscreen").children().removeAttr("readonly");
		$(".reportdeffile").children().attr("readonly", "readonly");

		switch (type) {

			case "DataOnly":
				$(".reportdefpreview").children().attr("readonly", "readonly");
				break;

			case "CSV":
				$(".reportdefscreen").children().attr("readonly", "readonly");
				$(".reportdeffile").children().removeAttr("readonly");
				break;

			case "HTML": case "WordDoc":
				$(".reportdeffile").children().removeAttr("readonly");
				$(".reportdefemail").children().removeAttr("readonly");
				break;

			case "ExcelWorksheet": case "ExcelGraph": case "ExcelPivotTable":
				$(".reportdeffile").children().removeAttr("readonly");
				$(".reportdefemail").children().removeAttr("readonly");
				break;

		}

	}

	function refreshOutputOptions() {

		var type = $('#outputformats :checked').val();

		$(".reportdefpreview").attr("disabled", (type == "0"));
		$(".reportdefemail").children().attr("disabled", (type == "0"));
		$(".reportdeffile").children().attr("disabled", (type == "0"));

		$(".reportdefscreen").attr("disabled", (type == "1"));
		$(".reportdefprinter").attr("disabled", (type == "1" || type == "2"));

		if (type == "0") {
			$(".reportdefpreview").css("color", "#A59393");
			$(".reportdefemail").css("color", "#A59393");
			$(".reportdeffile").css("color", "#A59393");		
		} else {
			$(".reportdefpreview").css("color", "#000000");
			$(".reportdefemail").css("color", "#000000");
			$(".reportdeffile").css("color", "#000000");
		}

		if (type == "1") {
			$(".reportdefscreen").css("color", "#A59393");
			$(".reportdefprinter").css("color", "#A59393");
		} else {
			$(".reportdefscreen").css("color", "#000000");
			$(".reportdefprinter").css("color", "#000000");
		}
	}

	function changeOutputType(type) {

		selectOutputType(type);

		$("#IsPreview").prop('checked', true);
		$("#ToScreen").prop('checked', true);
		$("#ToPrinter").prop('checked', false);
		$("#SaveToFile").prop('checked', false);
		$("#Filename").val("");
		$("#SendToEmail").prop('checked', false);
		
		switch (type) {

			case "DataOnly":
				$("#IsPreview").prop('checked', false);
				break;

			case "CSV":
				$("#ToScreen").prop('checked', false);
				$("#SaveToFile").prop('checked', true);
				break;

			default:
				break;
		}

		refreshOutputOptions();
		setOutputToEmail();
		setOutputToFile();

	}

	function saveToFileChecked() {

		var isChecked = $("#SaveToFile").prop('checked');
		$("#SaveExisting").attr('disabled', !isChecked);
		$("#Filename").removeAttr('readonly');

		if (isChecked) {
			$(".display-label_file").css("color", "#000000");
		} else {
			$("#Filename").attr('readonly', 'readonly');
			$(".display-label_file").css("color", "#A59393");
		}


	}

	function sendAsEmailChecked() {

		var isReadonly = $("#SendToEmail").prop('checked') == false ? 'readonly' : '';
		$("#Output_EmailSubject").removeAttr('readonly');
		$("#EmailAttachmentName").removeAttr('readonly');

		if (isReadonly == "readonly") {
			$("#Output_EmailSubject").attr('readonly', isReadonly);
			$("#EmailAttachmentName").attr('readonly', isReadonly);
			$(".display-label_emails").css("color", "#A59393");
			$("#Output_EmailSubject").val('');
			$("#EmailAttachmentName").val('');
			$("#txtEmailGroupID").val(0);
			$("#txtEmailGroup").val('');
		} else {
			$(".display-label_emails").css("color", "#000000");
		}

	}

	$(function () {

		if ('@Model.ReportType' == '@UtilityType.utlCalendarReport') {
			$(".hideforcalendarreport").hide();
		}

		selectOutputType('@Model.Format');
		refreshOutputOptions();
		saveToFileChecked();
		sendAsEmailChecked();

	});
</script>