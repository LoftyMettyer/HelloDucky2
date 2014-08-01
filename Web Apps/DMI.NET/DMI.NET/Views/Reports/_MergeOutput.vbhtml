@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.MailMergeModel)

<fieldset class="width100">
	<legend class="fontsmalltitle">Options:</legend>
	<fieldset class="width60 floatleft">
		@Html.LabelFor(Function(m) m.TemplateFileName)
		@Html.TextBox("TemplateFileName", Model.TemplateFileName, New With {.placeholder = "Template", .class = "width70"})
		<input type="button" class="ui-state-disabled" id="cmdEmailGroup" name="cmdTemplate" value="..." style="padding-top: 0;" />
	</fieldset>
	<fieldset class="width30 floatleft">
		@Html.CheckBoxFor(Function(m) m.PauseBeforeMerge)
		@Html.LabelFor(Function(m) m.PauseBeforeMerge)
		<br />
		@Html.CheckBoxFor(Function(m) m.SuppressBlankLines)
		@Html.LabelFor(Function(m) m.SuppressBlankLines)
	</fieldset>
</fieldset>

<fieldset class="width100">
	<legend class="fontsmalltitle">Output Format:</legend>
	<fieldset class="width30 floatleft">
		@Html.RadioButton("OutputFormat", 0, Model.OutputFormat = MailMergeOutputTypes.WordDocument, New With {.onclick = "selectMergeOutput('WordDocument')"})
		Word Document
		<br />
		@Html.RadioButton("OutputFormat", 1, Model.OutputFormat = MailMergeOutputTypes.IndividualEmail, New With {.onclick = "selectMergeOutput('IndividualEmail')"})
		Individual Emails
		<br />
		@Html.RadioButton("OutputFormat", 2, Model.OutputFormat = MailMergeOutputTypes.DocumentManagement, New With {.onclick = "selectMergeOutput('DocumentManagement')"})
		<span class="DataManagerOnly">Document Management</span>
		<br />
	</fieldset>

	<fieldset class="outputmerge_WordDocument width60 floatleft">
		<fieldset>
			<legend class="fontsmalltitle">Word Document:</legend>

			<fieldset class="border0">
				<legend>
					@Html.CheckBoxFor(Function(m) m.DisplayOutputOnScreen)
					@Html.LabelFor(Function(m) m.DisplayOutputOnScreen)
				</legend>
			</fieldset>

			<fieldset class="border0 DataManagerOnly" style="margin-bottom:10px">
				<legend>
					@Html.CheckBoxFor(Function(m) m.SendToPrinter)
					@Html.LabelFor(Function(m) m.SendToPrinter)
				</legend>
				@Html.TextBox("PrinterName", Model.PrinterName, New With {.placeholder = "Default printer", .class = "DataManagerOnly readonly width100"})
			</fieldset>

			<fieldset class="border0" style="margin-bottom:10px">
				<legend>
					@Html.CheckBoxFor(Function(m) m.SaveToFile, New With {.onclick = "setOutputToFile();"})
					@Html.LabelFor(Function(m) m.SaveToFile)
				</legend>
				@Html.TextBoxFor(Function(m) m.Filename, New With {.placeholder = "File Name", .class = "outputfile width100"})
				@Html.ValidationMessageFor(Function(m) m.Filename)
			</fieldset>

		</fieldset>
	</fieldset>

	<fieldset class="outputmerge_IndividualEmail width60 floatleft">
		<legend class="fontsmalltitle">Individual Emails:</legend>
		<fieldset>
			<fieldset>
				<div class="display-label_emails">
					@Html.LabelFor(Function(m) m.EmailGroupID)
				</div>
				@Html.EmailGroupDropdown("EmailGroupID", Model.EmailGroupID, Model.AvailableEmails)

			</fieldset>

			<fieldset>
				<div class="display-label_emails">
					@Html.LabelFor(Function(m) m.EmailSubject)
				</div>
				@Html.TextBox("EmailSubject", Model.EmailSubject, New With {.class = "display-textbox-emails"})
			</fieldset>
		</fieldset>

		<fieldset class="border0">
			<legend>
				@Html.CheckBoxFor(Function(m) m.EmailAsAttachment, New With {.onclick = "setOutputSendAsAttachment();"})
				@Html.LabelFor(Function(m) m.EmailAsAttachment)
			</legend>
			@Html.EditorFor(Function(m) m.EmailAttachmentName)
			<br />
			@Html.ValidationMessageFor(Function(m) m.EmailAttachmentName)
			<br/>
			@Html.ValidationMessageFor(Function(m) m.EmailGroupID)

		</fieldset>
	</fieldset>

	<fieldset class="outputmerge_DocumentManagement width60 floatleft">
		<legend class="fontsmalltitle">	Document Management :</legend>
		<fieldset class="border0 DataManagerOnly" style="margin-bottom:10px">
			<legend>
				@Html.CheckBoxFor(Function(m) m.DisplayOutputOnScreen)
				@Html.LabelFor(Function(m) m.DisplayOutputOnScreen)
			</legend>
			@Html.TextBox("PrinterName", Model.PrinterName, New With {.placeholder = "Engine", .class = "width100"})
		</fieldset>
	</fieldset>
</fieldset>

<fieldset class="DataManagerOnly">
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
	});

</script>