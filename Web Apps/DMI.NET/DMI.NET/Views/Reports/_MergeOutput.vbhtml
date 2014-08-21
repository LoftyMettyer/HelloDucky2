@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.MailMergeModel)

<fieldset class="width100">
	<legend class="fontsmalltitle">Options:</legend>
	<fieldset class="formField width60 floatleft">
		<input class="floatleft" type="hidden" id="txtEventFilterID" name="FilterID" value="@Model.FilterID" />
		<label class="floatleft">@Html.LabelFor(Function(m) m.TemplateFileName)</label>
		@Html.TextBox("TemplateFileName", Model.TemplateFileName, New With {.placeholder = "Template", .class = "floatleft"})
		<input type="button" class="ui-state-disabled floatleft" id="cmdEmailGroup" name="cmdTemplate" value="..." style="padding-top: 0;" />
	</fieldset>
	<br />
	<fieldset class="clearboth" style="padding-left:125px">
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
				@Html.CheckBoxFor(Function(m) m.SaveToFile, New With {.onclick = "setOutputToFile();"})
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
		@Html.HiddenFor(Function(m) m.EmailGroupID, New With {.Name = "Output.EmailGroupID", .id = "txtEmailGroupID"})
		@Html.EmailGroupDropdown("EmailGroupID", Model.EmailGroupID, Model.AvailableEmails)
		<br />
		@Html.LabelFor(Function(m) m.EmailSubject, New With {.class = "display-label_emails"})
		@Html.TextBox("Output.EmailSubject", Model.EmailSubject, New With {.class = "display-textbox-emails"})
		<br />
		@Html.LabelFor(Function(m) m.EmailAttachmentName, New With {.class = "display-label_emails"})
		@Html.TextBoxFor(Function(m) m.EmailAttachmentName, New With {.Name = "Output.EmailAttachmentName", .class = "display-textbox-emails"})
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

<script type="text/javascript">

	function setOutputToFile() {
		var bSelected = $("#").val();
		$(".outputfile").children().attr("readonly", !bSelected);
		if (!bSelected) {
			$(".outputfile").children().val("");
		}
	}

	function setOutputSendAsAttachment() {
		var bSelected = $("#").val();
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
		//styling for email address under Individual Emails section
		$('#fieldsetsubjectemail select').css({
			width: '70%',
			marginBottom: '4px',
			float: 'left'
			});
	});

</script>