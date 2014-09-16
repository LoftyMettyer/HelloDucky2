﻿@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.MailMergeModel)

<fieldset class="width100">
	<legend class="fontsmalltitle">Options:</legend>
	<fieldset class="">
		<div class="formField">
			<input type="hidden" id="txtEventFilterID" name="FilterID" value="@Model.FilterID" />
			<div class="floatleft"> 
				@Html.LabelFor(Function(m) m.TemplateFileName)
			</div>

			@Html.TextBoxFor(Function(m) m.TemplateFileName, New With {.id = "TemplateFileName", .class = "floatleft", .readonly = "true"})
			
			<input type="button" class="ui-state-disabled floatleft" style="margin-left:2px" id="cmdTemplateSelect" value="..." onclick="requestTemplateSelect()"  />
			<input type="button" class="ui-state-disabled floatleft" id="cmdTemplateClear" value="Clear" onclick="TemplateClear()"  />
		</div>
	</fieldset>
	
	<fieldset class="clearboth" style="padding-left:126px">
		@Html.CheckBoxFor(Function(m) m.PauseBeforeMerge)
		@Html.LabelFor(Function(m) m.PauseBeforeMerge)
		<br />
		@Html.CheckBoxFor(Function(m) m.SuppressBlankLines)
		@Html.LabelFor(Function(m) m.SuppressBlankLines)
	</fieldset>
</fieldset>

<fieldset class="width25 floatleft" style="">
	<legend class="fontsmalltitle">Output Format:</legend>
	<fieldset class="">
		@Html.RadioButton("OutputFormat", 0, Model.OutputFormat = MailMergeOutputTypes.WordDocument, New With {.onclick = "selectMergeOutput('WordDocument')"})
		Word Document
		<br />
		@Html.RadioButton("OutputFormat", 1, Model.OutputFormat = MailMergeOutputTypes.IndividualEmail, New With {.onclick = "selectMergeOutput('IndividualEmail')"})
		Individual Emails
		<br />
		@Html.RadioButton("OutputFormat", 2, Model.OutputFormat = MailMergeOutputTypes.DocumentManagement, New With {.onclick = "selectMergeOutput('DocumentManagement')"})
		<span class="DataManagerOnly">Document Management</span>
	</fieldset>
</fieldset>

<fieldset class="outputmerge_WordDocument width60 floatleft" style="">
	<legend class="fontsmalltitle">Word Document:</legend>
	<fieldset>
		<div class="padbot10">
			@Html.CheckBoxFor(Function(m) m.DisplayOutputOnScreen)
			@Html.LabelFor(Function(m) m.DisplayOutputOnScreen)
			<br />
		</div>

		<div class="reportdefprinter DataManagerOnly">
			<div class="width30 floatleft">
				@Html.CheckBoxFor(Function(m) m.SendToPrinter)
				@Html.LabelFor(Function(m) m.SendToPrinter)
			</div>
			<div class="width70 floatleft padbot5">
				@Html.TextBox("PrinterName", Model.PrinterName, New With {.placeholder = "Default printer", .class = "DataManagerOnly readonly width100"})
			</div>
		</div>

		<div class="padbot5">
			<div class="width30 floatleft">
				@Html.CheckBoxFor(Function(m) m.SaveToFile, New With {.id = "SaveToFile", .onclick = "setOutputToFile();"})
				@Html.LabelFor(Function(m) m.SaveToFile)
			</div>
			<div class="width70 floatleft">
				@Html.TextBoxFor(Function(m) m.Filename, New With {.placeholder = "File Name", .class = "outputfile width100"})
				@Html.ValidationMessageFor(Function(m) m.Filename)
			</div>
		</div>
	</fieldset>
</fieldset>

<fieldset class="outputmerge_IndividualEmail width60 floatleft" style="">
	<legend class="fontsmalltitle">Individual Emails:</legend>

	<fieldset id="fieldsetsubjectemail">
		@Html.LabelFor(Function(m) m.EmailGroupID, New With {.class = "display-label_emails"})
		@Html.EmailGroupDropdown("EmailGroupID", Model.EmailGroupID, Model.AvailableEmails)
		<br />
		@Html.LabelFor(Function(m) m.EmailSubject, New With {.class = "display-label_emails"})
		@Html.TextBox("EmailSubject", Model.EmailSubject, New With {.class = "display-textbox-emails"})
		<br />
		<br />
		@Html.CheckBoxFor(Function(m) m.EmailAsAttachment, New With {.id = "EmailAsAttachment", .onclick = "setOutputSendAsAttachment();"})
		@Html.LabelFor(Function(m) m.EmailAsAttachment)
		<br />
		@Html.LabelFor(Function(m) m.EmailAttachmentName, New With {.class = "display-label_emails"})
		@Html.TextBoxFor(Function(m) m.EmailAttachmentName, New With {.id = "EmailAttachmentName", .Name = "EmailAttachmentName", .class = "display-textbox-emails"})
		<br />
		@Html.ValidationMessageFor(Function(m) m.EmailAttachmentName)
		<br />
		@Html.ValidationMessageFor(Function(m) m.EmailGroupID)
	</fieldset>

</fieldset>

<fieldset class="outputmerge_DocumentManagement width60 floatleft" style="">
	<legend class="fontsmalltitle">	Document Management :</legend>
	<fieldset>
		<div class="padbot5">
			<div class="width30 floatleft">
				@Html.CheckBoxFor(Function(m) m.DisplayOutputOnScreen)
				@Html.LabelFor(Function(m) m.DisplayOutputOnScreen)
			</div>
			<div class="width70 floatleft">
				@Html.TextBox("PrinterName", Model.PrinterName, New With {.placeholder = "Engine", .class = "width100"})
			</div>
		</div>
	</fieldset>
</fieldset>

<fieldset class="DataManagerOnly width100">
	Note: Options marked in red are unavailable in OpenHR Web.
</fieldset>

<div style='height: 0;width:0; overflow:hidden;'>
	<input type="file" name="cmdGetFilename" id="cmdGetFilename" onchange="templateSelect()" />
</div>

<script type="text/javascript">

	function TemplateClear() {
		$("#TemplateFileName").val("");
		button_disable($('#cmdTemplateClear'), true);
	}

	function requestTemplateSelect() {

		var dialog = document.getElementById("cmdGetFilename");
		dialog.accept = "application/msword, application/vnd.openxmlformats-officedocument.wordprocessingml.document";
		dialog.click();

	}

	function templateSelect() {

		var dialog = document.getElementById("cmdGetFilename");
		var sFileName = /([^\\]+)$/.exec(dialog.value)[1];

		if (sFileName.length > 256) {
			OpenHR.messageBox("Path and file name must not exceed 256 characters in length");
			return;
		}

		if (sFileName != "") {
			$("#TemplateFileName").val(sFileName);
			button_disable($('#cmdTemplateClear'), false);
		}

	}

	function setOutputToFile() {
		var bSelected = $("#SaveToFile").prop("checked");
		$(".outputfile").children().attr("readonly", !bSelected);
		if (!bSelected) {
			$(".outputfile").children().val("");
		}
	}

	function setOutputSendAsAttachment() {
		var bSelected = $("#EmailAsAttachment").prop("checked");
		$("#EmailAttachmentName").attr("readonly", !bSelected);

		if (!bSelected) {
			$("#EmailAttachmentName").val("");
		}
	}

	function selectMergeOutput(outputType) {
		$("[class^=outputmerge_]").hide();
		$(".outputmerge_" + outputType).show(500);
	}

	$(function () {
		selectMergeOutput('@Model.OutputFormat');
		setOutputSendAsAttachment();
		//styling for email address under Individual Emails section
		$('#fieldsetsubjectemail select').css({
			width: '70%',
			marginBottom: '4px',
			float: 'left'
		});

		$('fieldset').css("border", "1");

		if ($("#TemplateFileName").val().length == 0) {
			button_disable($('#cmdTemplateClear'), true);
		}

	});

</script>