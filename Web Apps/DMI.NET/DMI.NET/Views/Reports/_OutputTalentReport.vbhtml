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
		Preview
		<br />

		@Html.RadioButton("Output.Format", 4, Model.Format = OutputFormats.ExcelWorksheet, New With {.onchange = "changeOutputType('ExcelWorksheet')"})
		Excel Worksheet
		<br />

		<br />
	</fieldset>
</fieldset>

<fieldset id="outputdestinatonfieldset" class="border0 floatleft width70">
	<legend class="fontsmalltitle">Output Destinations</legend>

	<fieldset class="border0 reportdefpreview" style="display:none">
		<div>
			@Html.CheckBoxFor(Function(m) m.IsPreview, New With {Key .Name = "Output.IsPreview"})
			@Html.LabelFor(Function(m) m.IsPreview)
		</div>
	</fieldset>

	<fieldset class="border0 reportdefscreen" style="display:none">
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
</fieldset>

<script type="text/javascript">

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
		}, 400, 400);

	}

	function selectOutputType(type) {

		switch (type) {

			case "DataOnly":
				$(".reportdefemail").children().prop("readonly", true);
				break;
			case "ExcelWorksheet": 
				$(".reportdefemail").children().prop("readonly", false);
				break;
		}
	}

	function refreshOutputOptions() {

		var bReadOnly = isDefinitionReadOnly();
		var type = $('#outputformats :checked').val();
		var bSendToEmail = $("#SendToEmail").prop('checked');

		$(".reportdefemail").children().prop("readonly", false);

		$(".reportdefemail :checkbox").prop("disabled", (type == "0") || bReadOnly);
		$('#cmdEmailGroup').prop('disabled', (type == "0") || bReadOnly || !bSendToEmail);

		if (type == "0") {
			$(".reportdefemail").children().prop("readonly", true);
  		$(".reportdefemail").css("color", "#A59393");

		} else {
			$(".reportdefemail").css("color", "#000000");
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
			default:
				break;
		}

		refreshOutputOptions();
		setOutputToEmail();

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
			//$("#txtEmailGroup").val('None');
		}

	}

	$(function () {

		selectOutputType('@Model.Format');
		refreshOutputOptions();
		sendAsEmailChecked()

	});
</script>