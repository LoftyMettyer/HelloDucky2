@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports HR.Intranet.Server.Enums
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
		@Html.RadioButton("OutputFormat", 0, Model.OutputFormat = MailMergeOutputTypes.WordDocument, New With {.onclick = "selectMergeOutput(0)"})
		Word Document
		<br />
		@Html.RadioButton("OutputFormat", 1, Model.OutputFormat = MailMergeOutputTypes.IndividualEmail, New With {.onclick = "selectMergeOutput(1)"})
		Individual Emails
		<br />
		@Html.RadioButton("OutputFormat", 2, Model.OutputFormat = MailMergeOutputTypes.DocumentManagement, New With {.onclick = "selectMergeOutput(2)"})
		Document Management
		<br />
		<span style="color:red">Some code here to hide and show the relevant sections depending on Ouput format selected</span>
	</fieldset>

	<fieldset class="width60 floatleft">
		<fieldset>
			<legend class="fontsmalltitle">Word Document:</legend>
			@* Display Output on Screen *@
			<fieldset class="border0">
				<legend>
					@Html.CheckBoxFor(Function(m) m.DisplayOutputOnScreen)
					@Html.LabelFor(Function(m) m.DisplayOutputOnScreen)
				</legend>
			</fieldset>

			@* Send To Print Section *@
			<fieldset class="border0" style="margin-bottom:10px">
				<legend>
					@Html.CheckBoxFor(Function(m) m.SendToPrinter)
					@Html.LabelFor(Function(m) m.SendToPrinter)
				</legend>
				@Html.TextBox("PrinterName", Model.PrinterName, New With {.placeholder = "Default printer", .class = "readonly width100"})
			</fieldset>

			@* Save To File Section *@
			<fieldset class="border0" style="margin-bottom:10px">
				<legend>
					@Html.CheckBoxFor(Function(m) m.SaveTofile)
					@Html.LabelFor(Function(m) m.SaveTofile)
				</legend>
				@Html.TextBox("Filename", Model.Filename, New With {.placeholder = "File Name", .class = "readonly width100"})
			</fieldset>
			@*</div>*@
		</fieldset>

		<fieldset>
			<legend class="fontsmalltitle">Individual Emails:</legend>
			<fieldset>
				<fieldset>
					<div class="display-label_emails">
						@Html.LabelFor(Function(m) m.EmailGroupID)
					</div>
					@Html.TextBox("EmailGroupID", Model.EmailGroupID, New With {.class = "display-textbox-emails"})
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
					@Html.CheckBoxFor(Function(m) m.EmailAsAttachment)
					@Html.LabelFor(Function(m) m.EmailAsAttachment)
				</legend>
				@Html.EditorFor(Function(m) m.EmailAttachmentName)
				<br />
				@Html.ValidationMessageFor(Function(m) m.EmailAttachmentName)
			</fieldset>
		</fieldset>

		<fieldset>
			<legend class="fontsmalltitle">	Document Management :</legend>
			<fieldset class="border0" style="margin-bottom:10px">
				<legend>
					@Html.CheckBoxFor(Function(m) m.DisplayOutputOnScreen)
					@Html.LabelFor(Function(m) m.DisplayOutputOnScreen)
				</legend>
				@Html.TextBox("PrinterName", Model.PrinterName, New With {.placeholder = "Engine", .class = "width100"})
			</fieldset>
		</fieldset>
	</fieldset>
</fieldset>

<script type="text/javascript">
	function selectMergeOutput(outputType) {
		$("#divMergeOutput_").hide();
		$("#divMergeOutput_" + outputType).show(500);
	}
</script>