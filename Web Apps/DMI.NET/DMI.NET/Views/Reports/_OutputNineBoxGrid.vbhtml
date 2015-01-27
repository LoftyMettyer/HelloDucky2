@Imports DMI.NET
@Imports DMI.NET.Helpers

@Code
	Html.EnableClientValidation()
End Code

@Inherits System.Web.Mvc.WebViewPage(Of Models.ReportOutputModel)

<fieldset class="border0 width20 floatleft">
	<legend class="fontsmalltitle">Output Formats</legend>
	<fieldset id="outputformats">
		@Html.RadioButton("Output.Format", 0, Model.Format = OutputFormats.DataOnly, New With {.onchange = "changeOutputType('DataOnly')"})
		Data Only
		<br />

		@Html.RadioButton("Output.Format", 4, Model.Format = OutputFormats.ExcelWorksheet, New With {.onchange = "changeOutputType('ExcelWorksheet')"})
		Excel Worksheet
		<br />

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

<script type="text/javascript">

	function setOutputToFile() {

		var bSelected = $("#SaveToFile").prop('checked');

		$(".reportdeffile").children().prop("readonly", !bSelected);

		if (!bSelected) {
			$("#Filename").val("");
		}

		saveToFileChecked();

	}

	function setOutputToEmail() {		
		var bSelected = $("#SendToEmail").prop('checked');
		var bReadOnly = isDefinitionReadOnly();

		$(".reportdefemail").children().prop("readonly", !bSelected || bReadOnly);
		$('#cmdEmailGroup').prop('disabled', !bSelected || bReadOnly);
		button_disable($("#cmdEmailGroup")[0], !bSelected || bReadOnly);

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
			enableSaveButton();
		},400,400);

	}

	function selectOutputType(type) {		

		$(".reportdefpreview").children().prop("readonly", false);
		$(".reportdefscreen").children().prop("readonly", false);
		$(".reportdeffile").children().prop("readonly", true);

		switch (type) {

			case "DataOnly":
				$(".reportdefpreview").children().prop("readonly", true);
				break;

			case "CSV":
				$(".reportdefscreen").children().prop("readonly", true);
				$(".reportdeffile").children().prop("readonly", false);
				break;

			case "HTML": case "WordDoc":
				$(".reportdeffile").children().prop("readonly", false);
				$(".reportdefemail").children().prop("readonly", false);
				break;

			case "ExcelWorksheet": case "ExcelGraph": case "ExcelPivotTable":
				$(".reportdeffile").children().prop("readonly", false);
				$(".reportdefemail").children().prop("readonly", false);
				break;

		}

	}

	function refreshOutputOptions() {

		var bReadOnly = isDefinitionReadOnly();
		var type = $('#outputformats :checked').val();
		var bSendToEmail = $("#SendToEmail").prop('checked');

		$(".reportdefpreview").children().prop("readonly", false);
		$(".reportdefscreen").children().prop("readonly", false);
		$(".reportdeffile").children().prop("readonly", false);
		$(".reportdefemail").children().prop("readonly", false);

		$(".reportdefpreview :checkbox").prop("disabled", (type == "0") || bReadOnly);
		$(".reportdefemail :checkbox").prop("disabled", (type == "0") || bReadOnly);
		$('#cmdEmailGroup').prop('disabled', (type == "0") || bReadOnly || !bSendToEmail);
		$(".reportdeffile :checkbox").prop("disabled", (type == "0") || bReadOnly);

		$(".reportdefscreen :checkbox").prop("disabled", (type == "1") || bReadOnly);
		$(".reportdefprinter :checkbox").prop("disabled", (type == "1" || type == "2" || bReadOnly));

		if (type == "0") {
			$(".reportdefpreview").children().prop("readonly", true);
			$(".reportdeffile").children().prop("readonly", true);
			$(".reportdefemail").children().prop("readonly", true);

			$(".reportdefpreview").css("color", "#A59393");
			$(".reportdefemail").css("color", "#A59393");
			$(".reportdeffile").css("color", "#A59393");		
		} else {
			$(".reportdefpreview").css("color", "#000000");
			$(".reportdefemail").css("color", "#000000");
			$(".reportdeffile").css("color", "#000000");
		}

		if (type == "1") {
			$(".reportdefpreview").children().prop("readonly", true);
			$(".reportdefscreen").children().prop("readonly", true);
			$(".reportdefscreen").css("color", "#A59393");
			$(".reportdefprinter").css("color", "#A59393");
		} else {
			$(".reportdefscreen").css("color", "#000000");
			$(".reportdefprinter").css("color", "#000000");
		}

		if (type == "2") {
			$(".reportdefprinter").children().prop("readonly", true);
			$(".reportdefprinter").css("color", "#A59393");
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
		$("#SaveExisting").prop("disabled", !isChecked);
		$("#Filename").prop("readonly", false);
		if (isChecked) {
			$(".display-label_file").css("color", "#ff0000");
		} else {
			$("#Filename").prop("readonly", true);
			$(".display-label_file").css("color", "#A59393");
		}
	}

	function sendAsEmailChecked() {

		var isReadonly = $("#SendToEmail").prop('checked') == false ? 'readonly' : '';
		$("#Output_EmailSubject").prop("readonly", false);
		$("#EmailAttachmentName").prop("readonly", false);

		if (isReadonly == "readonly") {
			$("#Output_EmailSubject").prop("readonly", true);
			$("#EmailAttachmentName").prop("readonly", true);
			$(".display-label_emails").css("color", "#A59393");
			$(".display-textbox-emails").css("background", "#EEEEEE");
			$("#Output_EmailSubject").val('');
			$("#EmailAttachmentName").val('');
			$("#txtEmailGroupID").val(0);
			$("#txtEmailGroup").val('');
		} else {
			$(".display-label_emails").css("color", "#000000");
			$(".display-textbox-emails").css("background", "#ffffff");
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